Option Strict Off
Imports NXOpen
Imports NXOpen.BlockStyler
Imports NXOpen.UF
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Math

' Window class
Public Class OutputParameters
    ' Class members
    Private Shared theSession As Session
    Private Shared theUI As UI
    Private theDlxFileName As String
    Private theDialog As NXOpen.BlockStyler.BlockDialog
    Private OpTreeList As NXOpen.BlockStyler.Tree' Block type: Tree Control
    Private PostProcessorBox As NXOpen.BlockStyler.ListBox' Block type: List Box
    Private group1 As NXOpen.BlockStyler.Group' Block type: Group
    Private OutputFolderBox As NXOpen.BlockStyler.FolderSelection' Block type: NativeFolderBrowser

    Public userInput As New Dictionary(Of String, Object) ()
    Public utils As Utils
    Public programRoot As NXOpen.CAM.NCGroup
    Private isFirstRun As Boolean = True

    ' Constructor
    Public Sub New()
        Try
            theSession = Session.GetSession()
            theUI = UI.GetUI()

            '================================================================================
            '================= .dlx FILE RELATIVE LOCATION TO JOURNAL FILE ==================
            '================================================================================
            Dim JournalFolder As String = Path.GetDirectoryName(theSession.ExecutingJournal)
            theDlxFileName = Path.Combine(journalFolder, "DLX\AutoPostprocess.dlx")  ' <------ relative location
            '================================================================================
            '================================================================================
            '================================================================================

            theDialog = theUI.CreateDialog(theDlxFileName)
            theDialog.AddApplyHandler(AddressOf apply_cb)
            theDialog.AddOkHandler(AddressOf ok_cb)
            theDialog.AddUpdateHandler(AddressOf update_cb)
            theDialog.AddInitializeHandler(AddressOf initialize_cb)
            theDialog.AddDialogShownHandler(AddressOf dialogShown_cb)

            ' set default user input values
            userInput("outputDir") = "\\MAZAK-PC\mazak_1_2"
            userInput("machines") = {"Mazak"}

            ' get Program View root object
            utils = New Utils(theUI, theSession)
            programRoot = Me.utils.setup.GetRoot(NXOpen.CAM.CAMSetup.View.ProgramOrder)

            ' gather the information of CAM objects in Program View
            utils.SetInfo(programRoot)
        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
            Throw ex
        End Try
    End Sub

    ' Main: attempt to show the dialog
    Public Shared Sub Main()
        Dim theOutputParameters As OutputParameters = Nothing
        Try
            theOutputParameters = New OutputParameters()
            ' The following method shows the dialog immediately
            theOutputParameters.Show()
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        Finally
            If theOutputParameters IsNot Nothing Then
                theOutputParameters.Dispose()
                theOutputParameters = Nothing
            End If
        End Try
    End Sub

    ' Unload option for the shared library .dll (not the journal) 
    Public Shared Function GetUnloadOption(ByVal arg As String) As Integer
        '------------------------------------------------------------------------------
        ' Immediately : unload the library as soon as the automation program has completed
        ' Explicitly : unload the library from the "Unload Shared Image" dialog
        ' AtTermination : unload the library when the NX session terminates
        '
        ' NOTE:  A program which associates NX Open applications with the menubar
        ' MUST NOT use this option since it will UNLOAD your NX Open application image
        ' from the menubar.
        '------------------------------------------------------------------------------

        'Return CType(Session.LibraryUnloadOption.Explicitly, Integer)
        Return CType(Session.LibraryUnloadOption.Immediately, Integer)
        'Return CType(Session.LibraryUnloadOption.AtTermination, Integer)
    End Function

    ' Clean-up. This method is automatically called by NX.
    Public Shared Sub UnloadLibrary(ByVal arg As String)
        Try
            '---- Enter your clean-up code here -----
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' This method shows the dialog on the screen
    Public Sub Show()
        Try
            theDialog.Show
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' Method to dispose dialogs
    Public Sub Dispose()
        If theDialog IsNot Nothing Then 
            theDialog.Dispose()
            theDialog = Nothing
        End If
    End Sub

    ' Callback Name: initialize_cb
    Public Sub initialize_cb()
        Try
            OpTreeList = CType(
                theDialog.TopBlock.FindBlock("OpTreeList"),
                NXOpen.BlockStyler.Tree
            )
            PostProcessorBox = CType(
                theDialog.TopBlock.FindBlock("PostProcessorBox"),
                NXOpen.BlockStyler.ListBox
            )
            group1 = CType(
                theDialog.TopBlock.FindBlock("group1"),
                NXOpen.BlockStyler.Group
            )
            OutputFolderBox = CType(
                theDialog.TopBlock.FindBlock("OutputFolderBox"),
                NXOpen.BlockStyler.FolderSelection
            )

            ' registration of treeList callbacks
            OpTreeList.SetOnInsertNodeHandler( AddressOf OnInsertNodeCallback )
            OpTreeList.SetOnSelectHandler( AddressOf OnSelectcallback )
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' Callback Name: dialogShown_cb (executed just before the dialog launch)
    Public Sub dialogShown_cb()
        Try
            ' show a Name column in tree
            Dim nameColumnId As Integer = 0
            OpTreeList.InsertColumn(nameColumnId, "Name", -1)
            OpTreeList.SetColumnResizePolicy(nameColumnId, Tree.ColumnResizePolicy.ResizeWithTree)

            ' build the operations tree
            utils.BuildOpTree(OpTreeList, programRoot, Nothing)
            'OpTreeList.SetColumnSortable(nameColumnId, False)  ' disable sort option

            ' preselect postprocessor
            PostProcessorBox.SetSelectedItemStrings(userInput("machines"))

            ' preselect first operation
            If Me.isFirstRun Then
                OpTreeList.SelectNode(OpTreeList.RootNode, True, True)

                Me.isFirstRun = False
            End If

            ' scroll to first selected operation (or else to top)
            If OpTreeList.FirstSelectedNode IsNot Nothing Then
                OpTreeList.FirstSelectedNode.ScrollTo(
                    nameColumnId, BlockStyler.Node.Scroll.Center
                )
            Else
                OpTreeList.RootNode.ScrollTo(
                    nameColumnId, BlockStyler.Node.Scroll.Center
                )
            End If
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' Callback Name: update_cb (executed on user input/action)
    Public Function update_cb(ByVal block As NXOpen.BlockStyler.UIBlock) As Integer
        Try
            If block Is PostProcessorBox Then
                Dim properties As PropertyList = GetBlockProperties("PostProcessorBox")

                ' if options are selected update current values
                if properties.GetStrings("SelectedItemStrings").Length <> 0 Then
                    userInput("machines") = properties.GetStrings("SelectedItemStrings")
                End If
            ElseIf block Is OutputFolderBox Then
                Dim userInputPath As String =
                    GetBlockProperties("OutputFolderBox").GetString("Path")

                ' append "\" at the end if needed
                If Len(userInputPath) <> 0 then
                    If right$(userInputPath, 1) <> "\" then
                        userInputPath = userInputPath & "\"
                    End If
                End If

                userInput("outputDir") = userInputPath
            End If

            ' ' print properties and their types' numerals
            ' For Each propertyName As String In properties.GetPropertyNames()
            '    Guide.InfoWriteLine(propertyName & ": " &
            '        properties.GetPropertyType(propertyName))
            ' Next
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try

        update_cb = 0
    End Function

    ' Callback Name: apply_cb
    Public Function apply_cb() As Integer
        apply_cb = 1
        Try
            ' ' DEBUG info
            ' Guide.InfoWriteLine(" === DEBUG === ")
            ' For Each op As Dictionary(Of String, Object) In utils.info.Values
            '     utils.PrintInfo(op("object"))
            ' Next

            update_cb(PostProcessorBox)
            update_cb(OutputFolderBox)

            ' check first if existing files will be overwritten and ask user
            Me.utils.userAnsweredForOverwrite = False
            For Each op As Dictionary(Of String, Object) In utils.info.Values
                utils.postProcess(op("object"), Me.userInput, False)
            Next

            ' stop and don't overwrite existing files if user declined the overwrite
            If Me.utils.userAnsweredForOverwrite And
                Not Me.utils.userWantsOverwrite Then
                Return apply_cb
            End If

            ' proceed to post process
            Guide.InfoWriteLine("----- AutoPostprocess info -----")

            For Each op As Dictionary(Of String, Object) In utils.info.Values
                utils.postProcess(op("object"), Me.userInput, True)
            Next

            Guide.InfoWriteLine("--------------------------------")

        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    ' Callback Name: ok_cb
    Public Function ok_cb() As Integer
        Dim errorCode as Integer = 0
        Try
            apply_cb()

            ' don't close window if user declined the overwrite
            If Me.utils.userAnsweredForOverwrite And
                Not Me.utils.userWantsOverwrite Then
                errorCode = 1
            End If
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            errorCode = 1
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try

        ok_cb = errorCode
    End Function

    ' Treelist specific callback:
    ' expand new nodes on creation
    Public Sub OnInsertNodeCallback(
        ByVal tree As NXOpen.BlockStyler.Tree,
        ByVal node As NXOpen.BlockStyler.Node
    )
        Try
            node.Expand(Node.ExpandOption.Expand)
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' Treelist specific callback:
    ' select/deselect all child nodes
    Public Sub OnSelectcallback(
        ByVal tree as NXOpen.BlockStyler.Tree,
        ByVal node as NXOpen.BlockStyler.Node,
        ByVal columnID As Integer,
        ByVal selected As Boolean
    )
        Try
            Dim theChildNode As NXOpen.BlockStyler.Node = node.FirstChildNode
            While theChildNode IsNot Nothing
                tree.SelectNode(theChildNode, selected, False)
                theChildNode = theChildNode.NextSiblingNode
            End While

            ' also set the operation info in utils
            Me.utils.info(node.DisplayText)("isSelected") = selected
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    ' Returns the propertylist of the specified BlockID
    Public Function GetBlockProperties(ByVal blockID As String) As PropertyList
        GetBlockProperties = Nothing
        Try
            GetBlockProperties = theDialog.GetBlockProperties(blockID)
        Catch ex As Exception
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function
End Class

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

Public Class Utils
    Public theUI As UI
    Public theSession  As Session
    Public workPart As NXOpen.Part
    Public setup As NXOpen.CAM.CAMSetup
    Public theUfSession As UFSession = UFSession.GetUFSession()
    Public info As New Dictionary(Of String, Dictionary(Of String, Object)) ()
    Public userAnsweredForOverwrite As Boolean = False
    Public userWantsOverwrite As Boolean = False
    Dim machineInfo As New Dictionary(Of String, Dictionary(Of String, String)) ()
    Dim takenScreenshots As New Collection()

    ' constructor
    Public Sub New(ByRef newUI As UI, ByRef newSession As Session)
        theUI = newUI
        theSession = newSession
        workPart = theSession.Parts.Work
        Try
            setup = workpart.CAMSetup
        Catch ex As Exception
            theUI.NXMessageBox.Show(
                "Block Styler",
                NXMessageBox.DialogType.Error,
                "Cannot be executed outside of Manufacturing Application."
            )
        End Try

        machineInfo.Add("Mazak", New Dictionary(Of String, String) ())
        machineInfo("Mazak")("extension") = ".eia"
        machineInfo("Mazak")("PostProcessor") = "mazak_nexus510"

        machineInfo.Add("Haas", New Dictionary(Of String, String) ())
        machineInfo("Haas")("extension") = ".tap"
        machineInfo("Haas")("PostProcessor") = "HAAS_3Axis"

        machineInfo.Add("Doosan", New Dictionary(Of String, String) ())
        machineInfo("Doosan")("extension") = ".mpf"
        machineInfo("Doosan")("PostProcessor") = "Sinumerik_828D"

        '================================================================================
        '============================== NEW MACHINE INFO ================================
        '================================================================================
        'machineInfo.Add("NEW_MACHINE", New Dictionary(Of String, String) ())  ' <------ new machine entry
        'machineInfo("NEW_MACHINE")("extension") = "EXTENSION"  ' <------ new machine extension
        'machineInfo("NEW_MACHINE")("PostProcessor") = "POSTPROCESSOR"  ' <------ new machine postprocessor
        '================================================================================
        '================================================================================
        '================================================================================
    End Sub

    ' check if object is an OpFolder
    Function IsOpFolder(ByVal member As NXOpen.CAM.CAMObject) As Boolean
        If member.Name = "NC_PROGRAM" Then
            Return False
        End If

        ' true if it is a folder containing an operation, inside another folder with operations
        If info(member.Name)("containsOperations") And
            info(info(member.Name)("parent"))("containsOperations") Then
            Return True
        Else
            Return False
        End If
    End Function

    ' check if object is inside an OpFolder
    Function IsInsideOpFolder(ByVal member As NXOpen.CAM.CAMObject) As Boolean
        If member.Name = "NC_PROGRAM" Then
            Return False
        End If

        ' true if is inside some OpFolder
        If IsOpFolder(info(info(member.Name)("parent"))("object")) Or
            IsInsideOpFolder(info(info(member.Name)("parent"))("object"))  Then
            Return True
        Else
            Return False
        End If
    End Function

    ' check if object is an operation or OpFolder (OpItem)
    Function IsOpItem(ByVal member As NXOpen.CAM.CAMObject) As Boolean
        If member.Name = "NC_PROGRAM" Then
            Return False
        End If

        If TypeOf member Is NXOpen.CAM.Operation Or IsOpFolder(member) Then
            Return True
        Else
            Return False
        End If
    End Function

    ' check if object can be postprocessed
    Function CanBePostProcessed(ByVal member As NXOpen.CAM.CAMObject) As Boolean
        '' skip top level root folder
        'If member.Name = "NC_PROGRAM" Then
        '    Return False
        'End If

        ' only postprocess OpItems
        If Not IsOpItem(member) Then
            Return False
        End If

        ' only postprocess outside OpFolders
        If IsInsideOpFolder(member) Then
            Return False
        End If

        ' on any other case it can be postprocessed
        Return True
    End Function

    ' ceck if object is in Unused Items folder of Geometry View
    Function IsInUnusedItems(ByVal item) As Boolean
        ' root item is never in Unused Items
        If item.Name = "GEOMETRY" Or
            item.Name = "NC_PROGRAM" Then
            Return False
        End If

        ' return True for Unused Items folder itself
        If item.Name = "NONE"
            Return True
        End If

        ' get parent of operation of folder accordingly
        Dim parent As NXOpen.CAM.CAMObject
        If TypeOf item Is NXOpen.CAM.Operation Then
            parent = item.GetParent(NXOpen.CAM.CAMSetup.View.Geometry)
        Else
            parent = item.GetParent()
        End If

        ' if the parent is a folder continue recursively
        Return IsInUnusedItems(parent)
    End Function

    ' check if object has (cutting) toolpath
    Function HasToolpath(ByVal member) As Boolean
        ' only check Operations and OpFolders
        If Not IsOpItem(member) Then
            Return False
        End If

        If TypeOf member Is NXOpen.CAM.Operation Then
            If member.GetToolpathCuttingLength() = 0 Then
                Return False
            Else
                Return True
            End If
        Else
            ' check contents of OpFolders
            HasToolpath = True
            For Each submember As NXOpen.CAM.CAMObject In member.GetMembers
                HasToolpath = HasToolpath And HasToolpath(submember)
            Next
        End If
    End Function

    ' get object's phase reference frame as NXOpen.CAM.OrientGeometry
    Function GetPhaseFrame(ByVal item) As NXOpen.CAM.OrientGeometry
        ' skip Unused Items and top level root folder
        If item.Name = "NONE" Or item.Name = "NC_PROGRAM" Then
            Return Nothing
        End If

        ' get parent of operation or folder accordingly
        Dim parent As NXOpen.CAM.CAMObject
        If TypeOf item Is NXOpen.CAM.Operation Then
            parent = item.GetParent(NXOpen.CAM.CAMSetup.View.Geometry)
        Else
            parent = item.GetParent()
        End If

        ' if the parent is a folder continue recursively
        If TypeOf parent Is NXOpen.CAM.OrientGeometry Then
            Return parent
        Else
            Return GetPhaseFrame(parent)
        End If
    End Function

    ' set camera view to given absolute rotation
    Sub SetView(
        ByVal phaseMcs As NXOpen.CartesianCoordinateSystem,
        ByVal rotX As Double,
        ByVal rotY As Double,
        ByVal rotZ As Double
    )
        ' get phase MCS matrix elements
        Dim phaseCsys As Double() = {
            phaseMcs.Orientation.Element().Xx,
            phaseMcs.Orientation.Element().Yx,
            phaseMcs.Orientation.Element().Zx,
            phaseMcs.Orientation.Element().Xy,
            phaseMcs.Orientation.Element().Yy,
            phaseMcs.Orientation.Element().Zy,
            phaseMcs.Orientation.Element().Xz,
            phaseMcs.Orientation.Element().Yz,
            phaseMcs.Orientation.Element().Zz
        }

        ' calculate transpose (inverse) of phase MCS matrix
        Dim phaseCsysT As Double() = {0, 0, 0, 0, 0, 0, 0, 0, 0}
        theUfSession.Mtx3.Transpose(phaseCsys, phaseCsysT)

        ' calculate rotation angles' sines and cosines
        Dim Ca As Double = Cos(rotZ)
        Dim Sa As Double = Sin(rotZ)
        Dim Cb As Double = Cos(rotY)
        Dim Sb As Double = Sin(rotY)
        Dim Cc As Double = Cos(rotX)
        Dim Sc As Double = Sin(rotX)

        ' build Euler rotation matrix
        Dim eulerRot As Double() = {
            Ca*Cb, Ca*Sb*Sc-Sa*Cc, Ca*Sb*Cc+Sa*Sc,
            Sa*Cb, Sa*Sb*Sc+Ca*Cc, Sa*Sb*Sc-Ca*Sc,
              -Sb,          Cb*Sc,          Cb*Cc
        }

        ' calculate resulting absolute rotation matrix
        Dim resultMatrix As Double() = {0, 0, 0, 0, 0, 0, 0, 0, 0}
        theUfSession.Mtx3.Multiply(phaseCsys, phaseCsysT, resultMatrix)
        theUfSession.Mtx3.Multiply(resultMatrix, eulerRot, resultMatrix)
        theUfSession.Mtx3.Multiply(resultMatrix, phaseCsys, resultMatrix)

        ' Guide.InfoWriteLine(phaseCsys(0) & " " & phaseCsys(3) & " " & phaseCsys(6))
        ' Guide.InfoWriteLine(phaseCsys(1) & " " & phaseCsys(4) & " " & phaseCsys(7))
        ' Guide.InfoWriteLine(phaseCsys(2) & " " & phaseCsys(5) & " " & phaseCsys(8))
        ' Guide.InfoWriteLine("-----------------")
        ' Guide.InfoWriteLine(resultMatrix(0) & " " & resultMatrix(3) & " " & resultMatrix(6))
        ' Guide.InfoWriteLine(resultMatrix(1) & " " & resultMatrix(4) & " " & resultMatrix(7))
        ' Guide.InfoWriteLine(resultMatrix(2) & " " & resultMatrix(5) & " " & resultMatrix(8))

        Dim rotMatrix As NXOpen.Matrix3x3 = Nothing
        rotMatrix.Xx = resultMatrix(0)
        rotMatrix.Yx = resultMatrix(1)
        rotMatrix.Zx = resultMatrix(2)
        rotMatrix.Xy = resultMatrix(3)
        rotMatrix.Yy = resultMatrix(4)
        rotMatrix.Zy = resultMatrix(5)
        rotMatrix.Xz = resultMatrix(6)
        rotMatrix.Yz = resultMatrix(7)
        rotMatrix.Zz = resultMatrix(8)

        ' set the view
        Dim translation As NXOpen.Point3d = New NXOpen.Point3d(0, 0, 0)
        Dim zoom As Double = 0.8
        workPart.ModelingViews.WorkView.SetRotationTranslationScale(
            rotMatrix, translation, 1
        )  ' apply rotation
        workPart.ModelingViews.WorkView.Fit()  ' fit view
        workPart.ModelingViews.WorkView.ZoomAboutPoint(
            zoom, translation, translation
        )  ' zoom out a bit
    End Sub

    ' take screenshot of current view
    Sub TakeScreenshot(
        ByVal outputFile As String,
        ByVal phaseFrame As NXOpen.CAM.OrientGeometry,
        ByVal rotX As Double,
        ByVal rotY As Double,
        ByVal rotZ As Double
    )
        ' store current MCS display status
        Dim previsousMcsVisibility As Boolean = theSession.CAMSession.GetMcsDisplay()

        ' store current WCS
        Dim previousWcs As NXOpen.CartesianCoordinateSystem = workPart.Wcs.CoordinateSystem
        Dim previousWcsVisibility As Boolean = workPart.Wcs.Visibility()

        ' store current phase MCS object
        Dim phaseMcs As NXOpen.CartesianCoordinateSystem =
            setup.CAMGroupCollection.CreateMillOrientGeomBuilder(phaseFrame).Mcs

        ' store current view
        Dim previousViewRotation As NXOpen.Matrix3x3 =
            workPart.ModelingViews.WorkView.Matrix
        Dim previousViewOrigin As NXOpen.Point3d =
            workPart.ModelingViews.WorkView.Origin
        Dim previousViewScale As Double =
            workPart.ModelingViews.WorkView.Scale

        ' hide MCS
        theSession.CAMSession.SetMcsDisplay(False)

        ' set WCS to MCS
        workPart.Wcs.SetCoordinateSystemCartesianAtCsys(phaseMcs)

        ' show WCS
        workPart.Wcs.Visibility = True

        ' rotate the view accordingly
        SetView(phaseMcs, rotX, rotY, rotZ)

        ' create builder objects for screenshot
        Dim imageExportBuilder1 As NXOpen.Gateway.ImageExportBuilder =
            theUI.CreateImageExportBuilder()
        Dim imageCaptureBuilder1 As NXOpen.Gateway.ImageCaptureBuilder =
            workPart.ImageCaptureManager.CreateImageCaptureBuilder()

        ' set export options
        imageExportBuilder1.BackgroundOption =
            NXOpen.Gateway.ImageExportBuilder.BackgroundOptions.Original
        imageExportBuilder1.EnhanceEdges = False
        imageExportBuilder1.FileName = outputFile & ".png"
        imageCaptureBuilder1.Size = NXOpen.Gateway.ImageCaptureBuilder.ImageSize.Pixels128
        imageCaptureBuilder1.Format = NXOpen.Gateway.ImageCaptureBuilder.ImageFormat.Png

        ' print screenshot info
        Guide.InfoWriteLine("= Screenshot: " & imageExportBuilder1.FileName)

        ' export image
        Dim exportItem As NXOpen.NXObject = imageExportBuilder1.Commit()

        ' reset WCS
        workPart.Wcs.Visibility = previousWcsVisibility
        workPart.Wcs.SetCoordinateSystemCartesianAtCsys(previousWcs)

        ' reset MCS
        theSession.CAMSession.SetMcsDisplay(previsousMcsVisibility)

        ' reset view
        workPart.ModelingViews.WorkView.SetRotationTranslationScale(
            previousViewRotation, previousViewOrigin, previousViewScale
        )

        ' destroy builder instances
        imageCaptureBuilder1.Destroy()
        imageExportBuilder1.Destroy()
    End Sub

    ' make the screenshots for a single phase
    Sub MakeScreenshots(ByVal operation, ByVal outputDir)
        Dim phaseFrame As NXOpen.CAM.OrientGeometry

        ' get the csys frame object accordingly for operations and folders
        If IsOpFolder(operation) Then
            MakeScreenshots(operation.GetMembers()(0), outputDir)
            Return
        Else
            phaseFrame = GetPhaseFrame(operation)
        End If

        ' if could not get phase frame (Unused Items), then exit
        If phaseFrame Is Nothing
            Return
        End If

        ' if screenshots are already taken for this phase, then skip
        If takenScreenshots.Contains(phaseFrame.Name) Then
            Return
        End If

        Dim outputFileBaseName As String =
            outputDir & split(workPart.Name, "_MFG")(0)

        ' take the screenshots
        TakeScreenshot(
            outputFileBaseName & "_ISO",
            phaseFrame,
            pi/4, 0, pi/4
        )  ' ISO
        TakeScreenshot(
            outputFileBaseName & "_XZ",
            phaseFrame,
            pi/2, 0, 0
        )  ' XZ
        TakeScreenshot(
            outputFileBaseName & "_XY",
            phaseFrame,
            0, 0, 0
        )  ' XY

        '================================================================================
        '================================ NEW SCREENSHOT ================================
        '================================================================================
        'TakeScreenshot(
        '    outputFileBaseName & "_YZ",  ' <----- file name: part_name + _YZ
        '    phaseFrame,
        '    pi/2, 0, pi/2  ' <----- rotation angles: rotX, rotY, rotZ
        ')  ' YZ

        '===== the Euler rotations are applied in the following order:
        '===== 1. rotZ
        '===== 2. rotY
        '===== 3. rotX
        '================================================================================
        '================================================================================
        '================================================================================

        ' mark this phase's screenshots as taken
        takenScreenshots.add(phaseFrame, phaseFrame.Name)
    End Sub

    ' post process (generate) operations
    Sub postProcess(
        ByVal operation As NXOpen.CAM.CAMObject,
        ByVal userInput As Dictionary(Of String, Object),
        ByVal JustPostprocess As Boolean
    )
        ' skip non-postprocessable or unselected objects
        If Not CanBePostProcessed(operation) Or
            Not info(operation.Name)("isSelected") Then
            Return
        End If

        ' only postprocess items not in Unused items
        If JustPostprocess And
            info(operation.Name)("isUnusedItem") Then
            Guide.InfoWriteLine(
                "! Warning: Operation " & operation.Name &
                " was found in 'Unused Items' of Geometry View and was not postprocessed."
            )
            Return
        End If

        ' only postprocess generated cutting toolpaths
        If JustPostprocess And
            Not HasToolpath(operation) Then
            Guide.InfoWriteLine(
                "! Warning: Operation " & operation.Name &
                " had an empty cutting tool path and was not postprocessed."
            )
            Return
        End If

        ' set the output folder path
        Dim outputDir As String =
            userInput("outputDir") & info(operation.Name)("phaseDir")

        ' make screenshots
        If JustPostprocess Then
            MakeScreenshots(operation, outputDir)
        End If

        ' repeat for every machine selected
        For Each machine As String In userInput("machines")
            ' set the output file path for this machine
            Dim outputFile As String =
                outputDir & machine.ToUpper() & "\" &
                operation.Name & machineInfo(machine)("extension")

            ' warn user that existing files will be overwritten and ask for confirmation
            If Not JustPostprocess Then
                If Not userAnsweredForOverwrite And
                    File.Exists(outputFile) Then
                    Dim question  = theUI.NXMessageBox.Show(
                        "Output File exists",
                        NXMessageBox.DialogType.Question,
                        "Files with the same name as the Output File were found." &
                        Environment.NewLine &
                        "Do you want to overwrite them?"
                    )

                    ' store user answer
                    If question = 1 Then
                        userWantsOverwrite = True
                    Else
                        userWantsOverwrite = False
                    End If

                    ' remember that user has answered
                    userAnsweredForOverwrite  = True
                End If

                ' don't proceed with postprocess
                Return
            End If

            ' print file info
            Guide.InfoWriteLine("- Output: " & outputFile)

            ' create the output directory
            Directory.CreateDirectory(outputDir & machine.ToUpper())

            ' post process
            Dim operations(0) As NXOpen.CAM.CAMObject
            operations(0) = operation
            setup.PostprocessWithSetting(
                operations,
                machineInfo(machine)("PostProcessor"),
                outputFile,
                setup.OutputUnits.PostDefined,
                setup.PostprocessSettingsOutputWarning.PostDefined,
                setup.PostprocessSettingsReviewTool.PostDefined
            )
        Next
    End Sub

    ' store object info in a 2d dictionary
    Sub SetInfo(ByVal root As NXOpen.CAM.NCGroup)
        If root.Name = "NC_PROGRAM" Then
            ' set root info
            info.Add(root.Name, New Dictionary(Of String, Object) ())
            info(root.Name)("parent") = "NULL"
            info(root.Name)("phaseDir") = ""
            info(root.Name)("containsOperations") = False
            info(root.Name)("object") = root
            info(root.Name)("isSelected") = False
            info(root.Name)("isUnusedItem") = False
        End If

        For Each member As NXOpen.CAM.CAMObject In root.GetMembers
            ' skip Unused Items folder in Program View
            If member.Name = "NONE" Then
                Continue For
            End If

            ' set parent dir name to part name
            If member.Name = "1234" Then
                member.SetName(split(workPart.Name, "_MFG")(0))
            End If

            ' set member info
            info.Add(member.Name, New Dictionary(Of String, Object) ())  ' add new key
            info(member.Name)("parent") = root.Name  ' parent name
            If root.Name = "NC_PROGRAM" Then  ' phaseDir
                info(member.Name)("phaseDir") = ""
            Else
                info(member.Name)("phaseDir") =
                    info(root.Name)("phaseDir") & root.Name & "\"
            End If
            info(member.Name)("containsOperations") = False  ' true if it contains operations
            info(member.Name)("object") = member  ' store the CAMObject
            info(member.Name)("isSelected") = False  ' true when is selected (set by OpTreeList)
            info(member.Name)("isUnusedItem") =
                IsInUnusedItems(member)  ' true if is in Unused Items folder of Geometry View

            If TypeOf member Is NXOpen.CAM.Operation Then
                ' mark operation parent folder as containsOperations
                info(root.Name)("containsOperations") = True
            Else
                ' continue recursively (depth first) inside folder
                SetInfo(member)
            End If
        Next
    End Sub

    ' print stored object info
    Sub PrintInfo(ByVal operation As NXOpen.CAM.CAMObject)
        Guide.InfoWriteLine("Name: " & operation.Name)
        Guide.InfoWriteLine("parent: " & info(operation.Name)("parent"))
        Guide.InfoWriteLine("phaseDir: " & info(operation.Name)("phaseDir"))
        Guide.InfoWriteLine("containsOperations: " & info(operation.Name)("containsOperations"))
        Guide.InfoWriteLine("IsOpFolder: " & IsOpFolder(operation))
        Guide.InfoWriteLine("IsOpItem: " & IsOpItem(operation))
        Guide.InfoWriteLine("HasToolpath: " & HasToolpath(operation))
        If Not GetPhaseFrame(operation) Is Nothing Then
            Guide.InfoWriteLine("PhaseFrameName: " & GetPhaseFrame(operation).Name)
        End If
        Guide.InfoWriteLine("isSelected: " & info(operation.Name)("isSelected"))
        Guide.InfoWriteLine(" ----------------------- ")
    End Sub

    ' build operations tree
    Sub BuildOpTree(
        ByVal OpTreeList As NXOpen.BlockStyler.Tree,
        ByVal root As NXOpen.CAM.NCGroup,
        ByRef rootNode As NXOpen.BlockStyler.Node
    )
        For Each member As NXOpen.CAM.CAMObject In root.GetMembers
            ' skip Unused Items folder
            If member.Name = "NONE" Then
                Continue For
            End If

            ' insert node in tree
            Dim memberNode As NXOpen.BlockStyler.Node = OpTreeList.CreateNode(member.Name)
            OpTreeList.InsertNode(memberNode, rootNode, Nothing, Tree.NodeInsertOption.Sort)
            OpTreeList.SelectNode(memberNode, info(member.Name)("isSelected"), False)

            ' if this node was a folder continure recursively
            If Not CanBePostProcessed(member) Then
                memberNode.DisplayIcon = "group"
                memberNode.SelectedIcon = "group"
                BuildOpTree(OpTreeList, member, memberNode)
            End If
        Next
    End Sub
End Class

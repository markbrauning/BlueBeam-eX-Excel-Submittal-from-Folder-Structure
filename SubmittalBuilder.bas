Attribute VB_Name = "SubmittalBuilder"
Option Explicit
Public WordApp As Word.Application
Public xFilesList() As String
Public xFoldersList() As String

Sub BuildSubmittalPDF()
    Application.ScreenUpdating = False
    'On Error GoTo Error_Handler
    If Not VerifyContinue("You are about to merge all of the PDFs in the 'Source Folder'." & vbNewLine & _
                        "Note that this function requires the eXtreme version of Bluebeam.") Then Exit Sub
    
    'Building Folder Structure list
    Application.StatusBar = "Building Folder structure in 'Source Folder'..."
    Dim SourceDocsPath As String
    SourceDocsPath = Sheets("Options").Range("SourceDocFolder").Value
    GetSubFolderList SourceDocsPath, 0, 0
    'Sort Folders by depth (deepest first)
    
    Dim xFoldersList_temp As Variant
    
    xFoldersList_temp = TransposeArray(xFoldersList)
    xFoldersList_temp = Two_Dim_Array_Sort(xFoldersList_temp, 2, True, False) '2 is depth column.
    xFoldersList_temp = TransposeArray(xFoldersList_temp)
    
    'Go through each folder
    Dim i As Long, k As Long
    Dim scriptRevu As String
    Dim SectionPDFpath As String
    Dim PDFFilePath As String
    Dim SectionFolder As String
    Dim QuoteCheck As Boolean
    For i = 0 To UBound(xFoldersList_temp, 2)
        SectionFolder = xFoldersList_temp(0, i)
        '---SectionPDFpath = ParentFolderPath + SubFolderName + .pdf
        SectionPDFpath = xFoldersList_temp(3, i) & "\" & xFoldersList_temp(1, i) & ".pdf"
        SectionPDFpath = Replace(SectionPDFpath, "'", "\'")
        scriptRevu = "Open('" & SectionPDFpath & "', '')"
        QuoteCheck = Not GetFilesList(SectionFolder, False, "pdf", 0)
        If QuoteCheck Then GoTo Sub_Exit
        For k = 0 To UBound(xFilesList, 1)
            PDFFilePath = xFilesList(k)
            scriptRevu = scriptRevu & " " & "InsertPages(9999, '" & PDFFilePath & "', true, false, false, false, false)"
        Next k
        scriptRevu = scriptRevu & " Save('" & SectionPDFpath & "', 1)"
        scriptRevu = scriptRevu & " Close(true, 1)"
        'Debug.Print "Send scipt to: " & SectionPDFpath & vbNewLine & _
        "--------------------------" & vbNewLine & _
        scriptRevu & vbNewLine & _
        "--------------------------" & vbNewLine
        RunRevuScript scriptRevu
        
        'if next script is for the next folder depth up, then wait for a few seconds
        If Not (i = UBound(xFoldersList_temp, 2)) Then
            If xFoldersList_temp(2, i) <> xFoldersList_temp(2, i + 1) Then
                waitdepth Int(xFoldersList_temp(2, i + 1)), 7
            End If
        End If
        
    Next i
    
Sub_Exit:
    Exit Sub

Error_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Source: File_Set_DateModified" & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Sub_Exit
    
End Sub
Sub waitdepth(depth As Integer, waitSec As Integer)
    Dim waitTime As String
    waitTime = Replace("0:00:0" & Str(waitSec), " ", "")
    Debug.Print "Waiting for this folder depth's scripts to finish before moving to the next." & _
    "Next Depth: " & Str(depth)
    Application.Wait (Now + TimeValue(waitTime))
End Sub


Sub RunRevuScript(RevuScript As String)
    Dim RevuPath As String
    Dim command As String
    
    RevuPath = "C:\Program Files\Bluebeam Software\Bluebeam Revu\20\Revu\ScriptEngine.exe"
    command = RevuPath & " " & RevuScript
    Shell command, vbNormalFocus
    
End Sub

Sub DocFromTemplate()
    'Future Dev

End Sub

Sub PrintWord2PFD()
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    If Not VerifyContinue("You are about to print all Word docs in the 'Source Folder' to PDF") Then Exit Sub
    
    'Set up connectiont to Word App
    Application.StatusBar = "Setting up Connection to Word App..."
    Set WordApp = New Word.Application
    WordApp.Visible = False
    If WordApp Is Nothing Then
        MsgBox "There was an issue while establishing connection to Word App."
        Exit Sub
    End If
    
    'Get list of Word Files. Results in public var: xFilesList()
    Application.StatusBar = "Getting List of Word Files in 'Source Folder'..."
    Dim SourceDocsPath As String
    SourceDocsPath = Sheets("Options").Range("SourceDocFolder").Value
    Dim QuoteContinue As Boolean
    QuoteContinue = GetFilesList(SourceDocsPath, True, "docx", 0)
    If Not QuoteContinue Then GoTo Sub_Exit
    

    'Define Word varriables
    Dim wordDoc As Word.Document
    Dim FilePath As String
    Dim NewPDFPath As String
    
    Dim i As Long
    For i = 0 To UBound(xFilesList)
        Application.StatusBar = "Converting Word File " & Str(i) & " of " & Str(UBound(xFilesList))
        FilePath = xFilesList(i)
        NewPDFPath = Replace(FilePath, ".docx", ".pdf")
        If FilePath = "" Then GoTo next_i
        Set wordDoc = WordApp.Documents.Open(FilePath)
        WordApp.Visible = False
        wordDoc.ExportAsFixedFormat OutputFileName:=NewPDFPath, ExportFormat:=17 '17 represents PDF format
        wordDoc.Close
        Set wordDoc = Nothing
next_i:
    Next i

    MsgBox "All Done! " & Str(i) & " total PDFs created"
Sub_Exit:
    On Error Resume Next
    WordApp.Quit
    WordApp.DDETerminateAll
    Set WordApp = Nothing
    If Not WordApp Is Nothing Then MsgBox "There was an issue while closing the Word App."
    ReDim xFilesList(0 To -1)
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

Error_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Source: File_Set_DateModified" & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Sub_Exit

End Sub

Sub SectionGenerator()
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    If Not VerifyContinue("You are about to create Word Doc Section pages for each " & _
                          "sub-folder in the 'Source Folder'") Then Exit Sub
    
    'Set up connectiont to Word App
    Application.StatusBar = "Setting up Connection to Word App..."
    Set WordApp = New Word.Application
    WordApp.Visible = False
    If WordApp Is Nothing Then
        MsgBox "There was an issue while establishing connection to Word App."
        Exit Sub
    End If
    
    'Get list of Sub Folders. Results in public var: xFoldersList()
    Application.StatusBar = "Getting List of Sub-Folders in 'Source Folder'..."
    Dim SourceDocsPath As String
    SourceDocsPath = Sheets("Options").Range("SourceDocFolder").Value
    GetSubFolderList SourceDocsPath, 0, 0
    
    'Define Word varriables
    Dim wordTemplateFilePath As String
    Dim wordTemplateSection As String
    Dim wordTemplateTitle As String
    wordTemplateSection = Sheets("Options").Range("WordTemplate_Section").Value
    wordTemplateTitle = Sheets("Options").Range("WordTemplate_Title").Value
    Dim wordDoc As Word.Document
    Dim FilePath As String
    Dim NewDocPath As String
    
    Dim i As Long
    For i = 0 To UBound(xFoldersList, 2)
        If xFoldersList(2, i) = 0 Then 'if folder is Source folder then use Title Template instead
            wordTemplateFilePath = wordTemplateTitle
        Else
            wordTemplateFilePath = wordTemplateSection
        End If
        Application.StatusBar = "Creating Section Page Word File " & Str(i) & " of " & Str(UBound(xFoldersList, 2))
        Set wordDoc = WordApp.Documents.Add(Template:=wordTemplateFilePath, Visible:=False)
        WordApp.Visible = False
        WordFindReplaceAll wordDoc
        WordFindReplace wordDoc, "[Section Title]", xFoldersList(1, i)
        FilePath = xFoldersList(3, i) & "\" & xFoldersList(1, i) & ".docx"
        wordDoc.SaveAs FilePath, wdFormatDocumentDefault
        wordDoc.Close
        Set wordDoc = Nothing
next_i:
    Next i

    MsgBox "All Done! " & Str(i) & " total Word Files created."
Sub_Exit:
    On Error Resume Next
    WordApp.Quit
    WordApp.DDETerminateAll
    Set WordApp = Nothing
    If Not WordApp Is Nothing Then MsgBox "There was an issue while closing the Word App."
    ReDim xFoldersList(0 To -1)
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

Error_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Source: File_Set_DateModified" & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Sub_Exit

End Sub

Sub GetSubFolderList(MainFolder As String, xFolderCount As Long, xDepth As Integer)
    On Error GoTo Sub_Exit
    
    Dim xFSO As FileSystemObject
    Dim xFolder As folder
    Dim xSubFolder As folder
    
    Set xFSO = New FileSystemObject
    Set xFolder = xFSO.GetFolder(MainFolder)
    
    For Each xSubFolder In xFolder.SubFolders
        ReDim Preserve xFoldersList(0 To 3, 0 To xFolderCount)
        xFoldersList(0, xFolderCount) = xSubFolder.Path
        xFoldersList(1, xFolderCount) = xSubFolder.Name
        xFoldersList(2, xFolderCount) = xDepth
        xFoldersList(3, xFolderCount) = xSubFolder.ParentFolder
        xFolderCount = xFolderCount + 1
        GetSubFolderList xSubFolder.Path, xFolderCount, (xDepth + 1)
    Next xSubFolder
    
Sub_Exit:
    Set xFolder = Nothing
    Set xSubFolder = Nothing
    Set xFSO = Nothing
    
End Sub

Function GetFilesList(MainFolder As String, ListSubFolders As Boolean, SearchEXT As String, xFileCount As Long) As Boolean
    GetFilesList = True
    On Error GoTo Sub_Exit
    
    Dim xFSO As FileSystemObject
    Dim xFolder As folder
    Dim xSubFolder As folder
    Dim xFile As File
       
    Set xFSO = New FileSystemObject
    Set xFolder = xFSO.GetFolder(MainFolder)
    
    For Each xFile In xFolder.Files
        If UCase(xFSO.GetExtensionName(xFile)) = UCase(SearchEXT) Then
            ReDim Preserve xFilesList(0 To xFileCount)
            'check for single quotes in file names
            If InStr(1, xFile.Path, "'", vbTextCompare) > 0 Then
                If Not VerifyContinue("There is a single quote in one of the file names." & vbNewLine & _
                "This will stop the Submittal Builder function to not work." & vbNewLine & xFile.Path) Then
                    GetFilesList = False
                    GoTo Sub_Exit
                End If
            End If
            xFilesList(xFileCount) = xFile.Path
            xFileCount = xFileCount + 1
        End If
    Next xFile
    
    If Not ListSubFolders Then GoTo Sub_Exit
    For Each xSubFolder In xFolder.SubFolders
        If Not GetFilesList(xSubFolder.Path, True, SearchEXT, xFileCount) Then GoTo Sub_Exit
    Next xSubFolder
    
GetFilesList = True
Sub_Exit:
    Set xFile = Nothing
    Set xFolder = Nothing
    Set xSubFolder = Nothing
    Set xFSO = Nothing
    
End Function

Function VerifyContinue(WarningMessage As String) As Boolean
    VerifyContinue = True
    Dim xAns As VbMsgBoxResult
    xAns = MsgBox(WarningMessage & vbNewLine & _
                  "Would you like to continue?", vbYesNo)
    If xAns = vbNo Then VerifyContinue = False
End Function

Function WordFindReplaceAll(wdoc As Word.Document)
    Dim txtFind  As String
    Dim txtReplace As String
    Dim FRTable As ListObject
    Dim xRow As ListRow
    Set FRTable = ThisWorkbook.Sheets("Options").ListObjects("FindReplaceTable")
    For Each xRow In FRTable.ListRows
        txtFind = xRow.Range.Cells(1).Value
        txtReplace = xRow.Range.Cells(2).Value
        If txtFind = "" Then GoTo nextRow
        WordFindReplace wdoc, txtFind, txtReplace
nextRow:
    Next xRow
End Function

Function WordFindReplace(wdoc As Word.Document, txtFind As String, txtReplace As String)
    Dim rangeToSearch As Word.Range
    Dim rangeArr(0 To 2) As Object
    
    Set rangeArr(0) = wdoc.Content
    Set rangeArr(1) = wdoc.Sections(1).Headers(wdHeaderFooterPrimary).Range
    Set rangeArr(2) = wdoc.Sections(1).Footers(wdHeaderFooterPrimary).Range
    
    ' Perform find and replace
    Dim i As Integer
    For i = 0 To 2
        rangeArr(i).Find.ClearFormatting
        rangeArr(i).Find.Text = txtFind
        rangeArr(i).Find.Replacement.Text = txtReplace
        ' Loop through each instance of the search text and replace it
        Do While rangeArr(i).Find.Execute
            rangeArr(i).Find.Execute Replace:=wdReplaceAll
        Loop
    Next i
    
End Function

Public Function Two_Dim_Array_Sort(sortArray As Variant, searchCol As Integer, _
                                   numericSort As Boolean, Optional ascendingOrder As Boolean = True) As Variant
    Dim temp As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    firstRow = LBound(sortArray, 1)
    lastRow = UBound(sortArray, 1)
    firstCol = LBound(sortArray, 2)
    lastCol = UBound(sortArray, 2)
    For i = firstRow To lastRow - 1
        For j = i + 1 To lastRow
            If (numericSort And ascendingOrder And sortArray(i, searchCol) > sortArray(j, searchCol)) _
            Or (Not (numericSort) And ascendingOrder And StrComp(sortArray(i, searchCol), sortArray(j, searchCol)) = 1) _
            Or (numericSort And Not (ascendingOrder) And sortArray(i, searchCol) < sortArray(j, searchCol)) _
            Or (Not (numericSort) And Not (ascendingOrder) And StrComp(sortArray(i, searchCol), sortArray(j, searchCol)) = -1) Then
                For k = firstCol To lastCol
                    temp = sortArray(j, k)
                    sortArray(j, k) = sortArray(i, k)
                    sortArray(i, k) = temp
                Next k
            End If
        Next j
    Next i
    
    Two_Dim_Array_Sort = sortArray

End Function

Public Function TransposeArray(myarray As Variant) As Variant
    Dim X As Long
    Dim Y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim tempArray As Variant
        Xupper = UBound(myarray, 2)
        Yupper = UBound(myarray, 1)
        ReDim tempArray(Xupper, Yupper)
        For X = 0 To Xupper
            For Y = 0 To Yupper
                tempArray(X, Y) = myarray(Y, X)
            Next Y
        Next X
        TransposeArray = tempArray
End Function


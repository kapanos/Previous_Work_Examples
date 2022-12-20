Attribute VB_Name = "App_Functions"
Option Explicit

Sub SortDataByColumn(strSheetName, strFirstHeader, strSortHeader) 'sorts sheet data using sheet filters
'// -- 08/07/2022 - KP -- //                                     'pass in sheet name, first header to seach, header to sort
    
    Dim blnSheetProtected As Boolean
    Dim lngHeaderRow, lngSortCol, lngFirstHeaderCol, lngLastHeaderCol As Long
    Dim rngHeaders, rngSortHeader As Range
    Dim ws As Worksheet
    Dim wb As Workbook

    Set ws = ThisWorkbook.Worksheets(strSheetName)
    
    On Error GoTo 0 ' errCatch
    
    blnSheetProtected = ws.ProtectContents
    
    Application.ScreenUpdating = False
    
    lngHeaderRow = 0
    
    'Get the first data row on Report sheet
    lngHeaderRow = ws.UsedRange.Find(strFirstHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
    
    If lngHeaderRow < 1 Then
        MsgBox "Specified sort header row not found!", 65536, "Sort Error"
        GoTo finishUp
    End If
    
    'find first header column
    lngFirstHeaderCol = ws.Rows(lngHeaderRow).Find(strFirstHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, SearchOrder:=xlByRows).Column
    
    'find last header column
    lngLastHeaderCol = ws.Rows(lngHeaderRow).Find("*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Column
       
    'Find strSortHeader on the header row
    lngSortCol = ws.Rows(lngHeaderRow).Find(strSortHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    
    'set header range for sorting
    Set rngHeaders = ws.Range(ws.Cells(lngHeaderRow, lngFirstHeaderCol), ws.Cells(lngHeaderRow, lngLastHeaderCol).Address)
    
    'set target sort header cell as range
    Set rngSortHeader = ws.Range(ws.Cells(lngHeaderRow, lngSortCol).Address)
'
'    ws.Sort.SortFields.Add2 Key:=rngSortHeader, SortOn:=xlSortOnValues, _
'        Order:=xlAscending, DataOption:=xlSortNormal
    

    ws.Activate
    ws.Unprotect
    
    If ws.AutoFilterMode = False Then
        rngHeaders.AutoFilter
    End If
        
    With ActiveSheet
        .AutoFilter.Sort.SortFields.Clear
        .AutoFilter.Sort.SortFields.Add2 Key:= _
            rngSortHeader, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
    End With
    
    With ThisWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

finishUp:
    
    'if the sheet was protected before this function, then protect it again
    If blnSheetProtected Then ws.Protect
    
    Application.ScreenUpdating = True
    Exit Sub

errCatch:
    
    MsgBox "Sort header: '" & strSortHeader & "', was not found. " & chr(10) & chr(10) & _
            "Check for unknown spaces, spelling, or try using wildcards with the header name.", 65536, "Sort Header Not Found"
    blnProcessError = True
    Resume finishUp

End Sub

Sub DeleteExternalLinks(strWbName As String) 'Delete's links to external sources - The workbook does not depend on other sources to open
'// -- 06/22/2022 - KP -- //                 'However, SharePoint keeps creating links in the exported workbook which user must be manually delete.
                                             'Microsoft has no solution that can be built in without changes to Excel application settings.
    Dim i As Long
    Dim aLink, ExtLinks As Variant
    Dim wb As Workbook
        
    Select Case True 'workbook name cannot include it's path, because in this application it is already open
        Case InStr(strWbName, "\")
            strWbName = Mid(strWbName, InStrRev(strWbName, "\") + 1) 'UNC path
        Case InStr(strWbName, "/")
            strWbName = Mid(strWbName, InStrRev(strWbName, "/") + 1) 'https SharePoint path
    End Select
    
    Set wb = Workbooks(strWbName)
    wb.Activate
    
    ExtLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    
        On Error Resume Next 'in case there are no links
    If UBound(ActiveWorkbook.LinkSources(xlExcelLinks)) > 0 Then
        For i = 1 To UBound(ExtLinks)
            wb.BreakLink Name:=ExtLinks(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
    
End Sub

Function VerifyNetResource(strResourceName As String) As Boolean 'Make sure needed resource is available
'// -- 07/28/2021 - KP -- //
        
    On Error GoTo errCatch
    
    Application.StatusBar = "Checking network connection . . ."
    
    If Dir(strResourceName, vbDirectory) <> "" Then
        VerifyNetResource = True
    Else
        GoTo errCatch
    End If
    
    DoEvents
    
    Exit Function

errCatch:
    VerifyNetResource = False 'if there is any kind of error, then the resource is not avaialable.
    MsgBox "Required resource '" & strResourceName & "' is not available." & chr(10) & chr(10) & _
        "Please check your network connection or select a file from a local folder.", 65536, "Network Resource Unavailable"
        
End Function

Function SelectImportFile(strBrowsePath, Optional strExplorerDialogTitle As String, Optional strInitialFile As String) As String 'Browse to specific file and select, Explorer dialog title may be specified if needed.
'// -- 07/03/2022 -- KP //

    Dim strFileSelect As String
    Dim lngItemCount As Long
    Dim objFileDialog As FileDialog 'pop-up dialog object
    
    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)
    
    If Right(strBrowsePath, 1) <> "\" Then strBrowsePath = strBrowsePath & "\"

    If strInitialFile <> "" Then strBrowsePath = strBrowsePath & strInitialFile & "*.xls*"
    'If a file browse dialog window title wasn't specified, then use the one below
    If strExplorerDialogTitle = "" Then strExplorerDialogTitle = "Select File to Import"

    If strBrowsePath = "" Then 'If the strBrowsePath came in empty, the start browsing from Local User profile documents
        If Dir(Environ("OneDrive")) & "\" <> "" Then 'if OneDrive is connected, then selection that user "documents" folder
            strBrowsePath = Environ("OneDrive") & "\Documents\"
        Else 'select the users local drive folder
            strBrowsePath = Environ("USERPROFILE") & "\Documents\"
        End If
    End If
    
    'if path passed in is missing the backslash and proper MS-DOS paths, then add it
    If (Right(strBrowsePath, 1) <> "\" And InStr(strBrowsePath, "\") = 0) Then strBrowsePath = strBrowsePath & "\"
    
    With Application.FileDialog(msoFileDialogOpen) 'open the file browser dialog box
        If Len(strBrowsePath) > 10 Then
            .Filters.Add "Excel Files", "*.xls; *.xlsx", 1
            .InitialFileName = strBrowsePath 'start browsing from the path passed in via strBrowsePath
        End If
        
        .Title = strExplorerDialogTitle
        .AllowMultiSelect = False 'prevent user from selecting multiple files
        .Show
        lngItemCount = .SelectedItems.Count
        
        If lngItemCount = 1 Then
        
            If lngItemCount = 1 Then
                strFileSelect = .SelectedItems(1)
                SelectImportFile = .SelectedItems(1)
            
            Else
                If Dir(strFileSelect) = "" Then 'if the filename or path was changed while browsing, return an error.
                    MsgBox "File not found!", , "Error"
                    blnProcessError = True
                    Exit Function
                
                End If
                
            End If
            
        Else
            blnUserAbort = True
            blnProcessError = True 'user hit cancel when browsing for file
        
        End If
    
    End With
    
    Application.StatusBar = ""
    
End Function

Function CheckFileImported(strFnamePartialCheck, Optional strFnameFull) As Boolean 'Determines partial file match was imported before and recorded on Controls sheet and eliminates duplicates
'// -- 08/06/2022 -- KP //          'Full filename with path must be passed in for strFnameFull
                                    'Set the boundaries of the file import log below(strFileCol, strDateCol, lngFirstRow, lngLastRow)
    Dim i, lngEmptyCount As Long
    Dim strFoundFileName, strFileDate, strFileCol, strDateCol As String
    Dim lngFilesCol, lngDatesCol, lngFirstRow, lngLastRow 'Define import log data range
    Dim lngTargetRow, lngLastFileRow As Long
    Dim c As Range 'cells to search for data source string
    Dim rngFiles, rngFileDates As Range
    Dim cws As Worksheet 'Controls/Status sheet in this workbook
                    
    If (blnProcessError Or blnUserAbort) Then Exit Function
    
    On Error Resume Next
    
    'Make sure the full filename includes a path only if passed in to this function
    If Not Dir(strFnameFull) Like "*.xlsx" Then
        If InStr(strFnameFull, "\") = 0 Then
            MsgBox "CheckFileImported: Full filename does not include a path"
            CheckFileImported = False
            Exit Function
        End If
        If Not strFnameFull Like "*.xls*" Then
            MsgBox "Import file must be '.xlsx' file type.", 65536, "File Select Error:"
            CheckFileImported = False
            blnProcessError = True
            Exit Function
        End If
    End If
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    ' -- Define the boundaries where filenames and import dates are located (for quick changes and cross application compatibility) --
    strFileCol = "H"    'Fully-qualified Filename Column
    strDateCol = "O"    'Date Imported Column
    lngFirstRow = 4     'First data row of file import log range
    lngLastRow = 5     'Last data row of file import log range
    
    lngFilesCol = cws.Range(strFileCol & "1").Column    'Convert column letters to integer
    lngDatesCol = cws.Range(strDateCol & "1").Column    'Convert column letters to integer
    '-- End Import Log boundaries --------------
        
    Set rngFiles = Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastRow, lngFilesCol))
    Set rngFileDates = Range(cws.Cells(lngFirstRow, lngDatesCol), cws.Cells(lngLastRow, lngDatesCol))
    
    lngTargetRow = 0
        
'    'set the filename from the search (partial can be full).  If full filename is empty, then use partial for message display
'    If strFnameFull = "" Then strFnameFull = strFnamePartialCheck
    
    'Get the last row of external workbook references listed on "Controls" sheet
    lngLastFileRow = rngFiles.Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

    'set the last filerow to the first if the log is empty
    If lngLastFileRow < lngFirstRow Then lngLastFileRow = lngFirstRow

    'Look for a partial filename match.  If a partial match is found, set then target row
    lngTargetRow = Range(cws.Cells(4, lngFilesCol), cws.Cells(lngLastRow, lngFilesCol)).Find(strFnamePartialCheck, _
                    LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
             
    On Error GoTo errCatch
    
    'if the partial filename was not found then use the first row
    If lngTargetRow = 0 Then
'        lngTargetRow = lngLastFileRow + 1
        CheckFileImported = False
    End If

    'See if the current file has already been imported.  If not, and a full filename was passed in, then log the file
    If lngTargetRow >= lngFirstRow Then
        
        CheckFileImported = True
        
        cws.Unprotect
        
        For Each c In Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastFileRow, lngFilesCol))
                            
            'get rid of duplicates of any files imported before caused by past process errors
            If Application.WorksheetFunction.CountIf(rngFiles, "*" & strFnamePartialCheck & "*") > 1 Then
            'If Application.WorksheetFunction.CountIf(range(cws.cells(lngfirstrow,lngfilescol),cws.cells(lngLastFileRow,lngfilescol)), "*" & Mid(c, InStrRev(c, "\") + 1, 20) & "*") > 1 Then
                lngTargetRow = rngFiles.Find("*" & strFnamePartialCheck & "*", LookIn:=xlValues, LookAt:=xlWhole, _
                        MatchCase:=False, SearchDirection:=xlNext).Row
                cws.Cells(lngTargetRow, lngFilesCol) = "" 'clear duplicate filename
                cws.Cells(lngTargetRow, lngDatesCol) = "" 'clear duplicate filename date
            Else
                        
                'Find a partial matching filename
                If c Like "*" & strFnamePartialCheck & "*" Then
                    If Application.WorksheetFunction.CountIf(Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastFileRow, lngFilesCol)), "*" & strFnamePartialCheck & "*") = 1 Then
                        strFoundFileName = c 'set filename found by partial match
                        lngTargetRow = rngFiles.Find("*" & strFnamePartialCheck & "*", LookIn:=xlValues, LookAt:=xlWhole, _
                                MatchCase:=False, SearchDirection:=xlPrevious).Row
                    Else
                        lngLastFileRow = rngFiles.Find("*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                        lngTargetRow = lngLastFileRow + 1
                    End If
                End If
            End If
        Next
        
        'Make sure the minim target row is the first data row and not header row
        If lngTargetRow < lngFirstRow Then lngTargetRow = lngLastRow + 1
        
    Else 'set new row to crate a new file entry if it was never imported
        lngTargetRow = lngLastFileRow + 1
    End If
    
    On Error Resume Next
    
    If Not Dir(strFnameFull) Like "*.xlsx" Then 'log file information if a valid full path was passed in
        cws.Cells(lngTargetRow, lngFilesCol) = strFnameFull 'log filename
        strFileDate = Format(FileDateTime(strFnameFull), "m-d-yyyy") 'log file date
        
        If strFileDate = "" Then strFileDate = Format(FileDateTime(strFnameFull), "m-d-yyyy") 'get filedate from file
            
        cws.Cells(lngTargetRow, lngDatesCol) = strFileDate 'write file date to file import log
    
    End If
    
    On Error Resume Next
    
    'Clean out empty rows
    lngLastFileRow = rngFiles.Find("*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    lngEmptyCount = Application.WorksheetFunction.CountIf(Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastFileRow, lngFilesCol)), "")
       
    On Error GoTo errCatch
    
    i = lngFirstRow

    Do While (lngEmptyCount > 0)
        
        If cws.Cells(i, lngFilesCol) = "" Then
            cws.Cells(i, lngFilesCol) = cws.Cells(i + 1, lngFilesCol) 'copy the filename from the cell below
            cws.Cells(i, lngDatesCol) = cws.Cells(i + 1, lngDatesCol) 'copy the date from the cell below
            cws.Cells(i + 1, lngFilesCol) = "" 'clear the cell just copied from
            cws.Cells(i + 1, lngDatesCol) = "" 'clear the cell just copied from
        End If
        
        lngLastFileRow = rngFiles.Find("*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
        If lngLastFileRow < lngFirstRow Then lngLastFileRow = lngFirstRow
        lngEmptyCount = Application.WorksheetFunction.CountIf(Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastFileRow, lngFilesCol)), "")
        
        i = i + 1
        
        If i > (lngLastRow - lngFirstRow) Then Exit Do 'prevent infinite loop
        
    Loop
    
finishUp:

    cws.Protect
    
    Exit Function
    
errCatch:
'    MsgBox "There was an error verifying the filename to import.", 6lngfirstrowlngfirstrow36, "File Verification Error"
'    blnProcessError = True
'    ResetApplication
    Resume finishUp

End Function

Function GetFileLastImported(strSearchString) As String 'Returns the file and location last recorded in the import log by strSearchString
'// -- 07/19/2022 -- KP //                              'Define the boundaries of the file import log below(strFileCol, strDateCol, lngFirstRow, lngLastRow)
                                    
    Dim strFoundFileName, strFileCol, strDateCol As String
    Dim lngFilesCol, lngFirstRow, lngLastRow 'Define import log data range
    Dim lngFileFoundRow, lngLastFileRow As Long
    Dim rngFiles As Range
    Dim cws As Worksheet 'Controls/Status sheet in this workbook
                    
    If (blnProcessError Or blnUserAbort) Then Exit Function

    Set cws = ThisWorkbook.Worksheets("Controls")
    
    ' -- Define the boundaries where filenames and import dates are located (for quick changes and cross application compatibility) --
    strFileCol = "H"    'Fully-qualified Filename Column
    lngFirstRow = 4     'First data row of file import log range
    lngLastRow = 5     'Last data row of file import log range
    
    lngFilesCol = cws.Range(strFileCol & "1").Column    'Convert column letters to integer
    '-- End Import Log boundaries --------------
        
    Set rngFiles = Range(cws.Cells(lngFirstRow, lngFilesCol), cws.Cells(lngLastRow, lngFilesCol))
    
    strFoundFileName = ""
    lngFileFoundRow = 0
    
    On Error Resume Next
    
    'Look for a partial filename match.  If a partial match is found, set then target row
    lngFileFoundRow = rngFiles.Find(strSearchString, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    On Error GoTo errCatch
    
    'if the partial filename was not found then return a blank strfoundfilename
    If lngFileFoundRow >= lngFirstRow Then
        strFoundFileName = cws.Cells(lngFileFoundRow, lngFilesCol)
        If Not Dir(strFoundFileName) Like "*.xlsx" Then 'verify the last path was valid
            strFoundFileName = "" 'erase the filefound if file was not a valid previous import
        End If
    End If
    
    'pass filename and date string back to function
    GetFileLastImported = strFoundFileName
    
finishUp:
    
    On Error Resume Next

    Exit Function
    
errCatch:

    Resume finishUp

End Function

Sub ClearTargetSheetData(strSheetName As String, strFirstHeader As String, strLastHeader As String, Optional lngRowOffset As Long) 'clears all date and dates from specified sheet
'// -- 08/07/2021 - KP --//         'Pass in first and last column headers find within scope of sheet
                                    'Row offset allows for double-row headers w/sub-headers to start clearing below
    Dim blnDataFound As Boolean
    Dim lngDataFirstRow As Long 'First data row to clear
    Dim lngDataLastRow As Long, lngFirstHeaderCol, lngLastHeaderCol As Long
    Dim lngHeaderRow As Long 'Row which contains data headers
    Dim strHeaderFirst, strHeaderLast, strHeaderName  As String
    Dim c, rngData As Range
    Dim ws As Worksheet, wsActive As Worksheet
    
    blnProcessError = False

    'copy passed-in vars to this procedure's scope only
    strHeaderFirst = strFirstHeader
    strHeaderLast = strLastHeader
        
    If strSheetName = "" Then 'If the sheet name was not passed in from another procedure
        MsgBox "Sheet name not specified in procedure call.", 65536, "Incomplete Procedure Call"
        blnProcessError = True
        Exit Sub
    End If
    
    Set wsActive = ThisWorkbook.ActiveSheet 'Remember which sheet was active when button was pressed
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(strSheetName) 'Sheet name passed in with procedure call
    
    On Error Resume Next
    
    'Find the first columm header on the target sheet
    lngHeaderRow = ws.UsedRange.Find(strHeaderFirst, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
    
    'check to see if the sheet is empty when a header isn't found by looking in the first 12 rows and 6 columns
    If lngHeaderRow = 0 Then
        For Each c In ws.Range("A1:F12")
            If Len(c) > 0 Then
                blnDataFound = True
                MsgBox "Data was found, but the data first header: " & strHeaderFirst & " was found not on sheet: " & strSheetName & "'." & _
                        chr(10) & chr(10) & "Check the calling function. ", 65536, "Clear sheet: '" & strSheetName & "' Failed"
                GoTo finishUp
            End If
        Next
    Else
        blnDataFound = True
    End If
    
    If blnDataFound = False Then GoTo finishUp 'sheet was already blank, so don't annoy the user with a useless message
    
    'if a row offset was passed in then shift the starting row by the number passed in - allows for double-row headers w/sub-headers
    lngDataFirstRow = 1 + lngHeaderRow + lngRowOffset
    
    'Find the first columm header on the target sheet
    lngFirstHeaderCol = ws.Rows(lngHeaderRow).Find(strHeaderFirst, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlNext, SearchOrder:=xlByRows).Column
    
    'Find the Last columm header on the target sheet
    lngLastHeaderCol = ws.Rows(lngHeaderRow).Find(strHeaderLast, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Column
    
    If (lngFirstHeaderCol = 0 Or lngLastHeaderCol = 0) Then
        MsgBox "First column header '" & strHeaderFirst & " not found to clear data on target sheet!", 65536, "Headers Not Found"
        GoTo finishUp
    End If

    'find the last row by looking for data in all columns betwee the first column and the last
    lngDataLastRow = ws.Columns(lngLastHeaderCol).Find("*", _
                LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'Protect sheet headers
    If lngDataLastRow <= lngHeaderRow Then lngDataLastRow = lngHeaderRow + 1
    
    'sheet has already been cleared with headers found.  Exit sub
    If lngDataLastRow = lngDataFirstRow Then GoTo finishUp
    
    Application.StatusBar = "Clearing data on " & strSheetName & " sheet . . ."
          
    'Clear information from the last import
    ws.Unprotect
    
    On Error GoTo errCatch
        
    Set rngData = Range(ws.Cells(lngDataFirstRow, lngFirstHeaderCol), ws.Cells(lngDataLastRow, lngLastHeaderCol))
    'clear formats and data, and prevents formats from being cleared on 'Controls' sheet
    With rngData
        If ws.Name = "Controls" Then 'Do not clear formatting on 'Controls' sheet
            .ClearContents
        Else 'clears formats and content on all other sheets within range
            .Clear
        End If
    End With
    
    ws.Activate
    
    'call function to resize scrollbars.  Pass in column letter to use to find last column
    ResizeScrollbars strSheetName, Left(Replace(ws.Cells(lngHeaderRow, lngFirstHeaderCol).Address, "$", ""), 1)
    
    ws.Range("A1").Select
    
    Application.ScreenUpdating = True

finishUp:

    Exit Sub

errCatch:
    
    If blnDataFound = True Then
        Select Case True
            Case strHeaderFirst = ""
                strHeaderName = strHeaderFirst
            Case strHeaderLast = ""
                strHeaderName = strHeaderLast
            Case Else
        End Select
        
        MsgBox "The header: '" & strHeaderName & "' was not found on sheet: '" & strSheetName & "'." _
                    , 65536, "Clear Sheet Error on Sheet: " & strSheetName
    End If
    
    Resume finishUp
    
End Sub

Sub ResizeScrollbars(strSheetName, strColumnLetter) 'Shrinks the worksheet side vertical scroll bars to not include empty rows
'// -- 08/07/2022 - KP -- //                        'Pass in column as column letter, A,B,C,D, etc.

    Dim lngLastRow, lngTargetCol As Long
    Dim strColLetter As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(strSheetName)
    
    strColLetter = UCase(strColumnLetter)
    lngTargetCol = Asc(strColLetter) - 64
    
    'Get the new last row of all rows containing data or formulas
    lngLastRow = ws.Columns(lngTargetCol).Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

    If lngLastRow < 11 Then lngLastRow = 11
    
    'Find the end of all cells below the data that have been touched, causing excessive blank rows to be displayed
'    ws.Range(strColLetter & lngLastRow + 1 & ":" & Replace(Mid(ws.UsedRange.Address, InStrRev(ws.UsedRange.Address, ":") + 1), "$", "")).EntireRow.Delete
    
    ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell) 'Reset the scrollbars to the proper size to end where the cell data ends
    
End Sub

Function CheckLocalFile() As Boolean 'Make sure the file is run from a local PC rather than SharePoint
'// -- 03/08/2022 - KP -- //

    Dim strWbPath As String
    Dim strIntersectPath As String
    Dim strUserPath As String
    Dim lngNamePos As Long 'local user name position
        
    strWbPath = ConvertPathSPtoUNC(ThisWorkbook.Path)
    
    strWbPath = Mid(strWbPath, InStrRev(strWbPath, "\Documents"))
    
    strUserPath = Environ("OneDrive")
    
    lngNamePos = InStr(LCase(ThisWorkbook.Path), strUserPath)
    
    'if the name is found, then check to see the workbook is accessible via local UNC path
    strWbPath = strUserPath & strWbPath & "\"
    
    If Dir(strWbPath) <> "" Then
        CheckLocalFile = True
    Else
        CheckLocalFile = False
    End If

End Function

Function ConvertPathSPtoUNC(strSharePointPath) As String 'Converts shareport path 'https://' path to conventional UNC type format
'// -- 02/16/2022 -- KP //

    Dim strNewPath As String
    
    If InStr(strSharePointPath, "/") Then
        strNewPath = Replace(strSharePointPath, "/", "\")
        strNewPath = Replace(strNewPath, "https:", "")
        strNewPath = Replace(strNewPath, "~%20", " ")
        ConvertPathSPtoUNC = strNewPath
    Else
        ConvertPathSPtoUNC = strSharePointPath 'just return the input path if it is already in proper format since SharePoint is unpredictable for hijacking pathnames
    End If
    
End Function

Sub SetBorderAndFillFormat(strSheetName As String, strFirstHeader As String, strLastHeader As String) 'clears all date and dates from specified sheet
'// -- 08/06/2021 - KP --//           'sets alternate row shade conditional formatting and cell borders for
                                      'Pass in first and last column headers find within scope of sheet

    Dim lngLastRow As Long, lngFirstCol, lngLastCol As Long
    Dim lngDataFirstRow As Long 'First data row to clear
    Dim lngHeaderRow As Long 'Row which contains data headers
    Dim strHeaderFirst, strHeaderLast As String
    Dim ws As Worksheet, wsActive As Worksheet
    
    blnProcessError = False
    
    'copy passed-in vars to this procedure's scope only
    strHeaderFirst = strFirstHeader
    strHeaderLast = strLastHeader
        
    If strSheetName = "" Then 'If the sheet name was not passed in from another procedure
        MsgBox "Sheet name not specified in procedure call.", 65536, "Incomplete Procedure Call"
        blnProcessError = True
        Exit Sub
    End If
    
    Set wsActive = ThisWorkbook.ActiveSheet 'Remember which sheet was active when button was pressed
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(strSheetName) 'Sheet name passed in with procedure call

    'Find the first columm header on the target sheet
    lngHeaderRow = ws.UsedRange.Find(strHeaderFirst, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    lngDataFirstRow = lngHeaderRow + 1
    
    'Find the first columm header on the target sheet
    lngFirstCol = ws.Rows(lngHeaderRow).Find(strHeaderFirst, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Column
    
    'Find the Last columm header on the target sheet
    lngLastCol = ws.Rows(lngHeaderRow).Find(strHeaderLast, LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Column
    
    If lngFirstCol = 0 Then
        MsgBox "First column header not found on target sheet!", 65536, "Headers Not Found"
    End If
    
    If lngLastCol = 0 Then
        MsgBox "Last column header not found on target sheet!", 65536, "Headers Not Found"
    End If
        
    'get the last data row
    lngLastRow = ws.Columns(lngFirstCol).Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

    If lngLastRow < lngDataFirstRow Then lngLastRow = lngDataFirstRow 'Protect sheet headers
    
    Application.StatusBar = "Formatting data . . ."

    Set ws = ThisWorkbook.Worksheets(strSheetName) 'Sheet name to set conditional formatting for alternate shaded rows and vertical borders
    
    On Error Resume Next 'in case column not found or name changed

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    With Range(ws.Cells(lngDataFirstRow, lngFirstCol), ws.Cells(lngLastRow, lngLastCol))
        '.ClearFormats
        .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"
        .FormatConditions(1).Interior.Color = RGB(228, 228, 228)

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            '.ThemeColor = 1
            .Color = RGB(160, 160, 160)
            .Weight = xlThin
        End With

        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            '.ThemeColor = 1
            .Color = RGB(160, 160, 160)
            .Weight = xlThin
        End With

        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            '.ThemeColor = 1
            .Color = RGB(160, 160, 160)
            .Weight = xlThin
        End With

        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            '.ThemeColor = 1
            .Color = RGB(160, 160, 160)
            .Weight = xlThin
        End With

    End With

    'set autofilters
    Range(ws.Cells(lngHeaderRow, lngFirstCol), ws.Cells(lngHeaderRow, lngLastCol)).Select
    With Selection
        .AutoFilter 'clear auto-filters
        .AutoFilter 'set autofilters to include new data range
    End With

    'Center justify all columns but the first
    With Range(ws.Cells(lngDataFirstRow, lngFirstCol + 1), ws.Cells(lngLastRow, lngLastCol))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

finishUp:

    Application.ScreenUpdating = True
    Application.GoTo Reference:=Worksheets(strSheetName).Range("A1"), Scroll:=True 'scroll full left and park on an innocuous cell
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    
    Exit Sub

errCatch:
    MsgBox "There was an error while formatting " & strSheetName & " row shading and borders." & chr(10) & chr(10) & _
        "Please make sure columns headers have not changed.", 65536, "Alternate Shading & Border Formats"
    blnProcessError = True
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Resume finishUp

End Sub

Function GetLocalPathFromSP(strWbPath) As String 'convert personal sharepoint path to local path by workbook name
'// -- 12/22/2021 - KP -- //

    Dim strLocalPath As String
    Dim strUserDocs As String, strUserDocsShPt As String

    On Error Resume Next
    
    'make sure the shaprepoint path is a personal one, and doesn't contain '/temp' caused by running from email or Sharepoint
    If (strWbPath Like "*/personal*" And Not strWbPath Like "*/temp*") Then
        
        'get the user's sharepoint documents path by finding the last occurence of "Documents" in the Sharepoint path
        strWbPath = Replace(Mid(strWbPath, InStrRev(strWbPath, "Documents")), "/", "\")
        
        'get the user's local corporate folder
        strUserDocs = Environ("OneDrive")
        
        'set the local workbook path by combining the user's OneDrive local path with documents path on Sharepoint
        GetLocalPathFromSP = strUserDocs & "\" & strWbPath

        If Dir(strUserDocs & "\" & strWbPath, vbDirectory) = "" Then
            MsgBox "This file must be run on the local PC" & chr(10) & chr(10) & _
                "If you haven't yet, please download or detach from email.", 65536, "Notification"
            GetLocalPathFromSP = ""
        End If
    
    Else 'If it's not a sharepoint path, just return the normal drive/path or UNC path
        If Not strWbPath Like "*/*" Then GetLocalPathFromSP = strWbPath
    
    End If
    
    On Error GoTo 0

End Function

Function ChkProcess(process As String) 'Used to see if external Windows process is running (such as the user's Macro Help Notepad window)
'// -- 01/06/2022 - KP -- //
    
    Dim objList As Object

    Set objList = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        ChkProcess = True
    Else
        ChkProcess = False
    End If

End Function

Function TransposeArray(arrTemp As Variant) As Variant 'Transpose Array
'// -- 02/24/2022 - KP -- //

    Dim r, c As Long
    Dim arr As Variant
    
    ReDim arr(1 To UBound(arrTemp, 2) + 1, 1 To UBound(arrTemp) + 1)
    
    For r = 1 To UBound(arrTemp, 2) + 1
        For c = 1 To UBound(arrTemp) + 1
            arr(r, c) = arrTemp(c - 1, r - 1)
        Next c
    Next r
    
    TransposeArray = arr

End Function






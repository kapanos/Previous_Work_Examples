Attribute VB_Name = "AppFunctions"
Option Explicit

Sub FileScavenger() 'Find all known and required files for this report in the users 'Download' folder, and move them to the last used folders for each respective file
'// -- 07/29/2022 - KP -- //   'this includes reading the filenames stored in the Value Classes(Combined).xlsx file and moves those as well.
    
    Dim blnFilesMoved As Boolean
    Dim uChoice As Integer
    Dim i, ar As Long
    Dim lngFirstFileRow, lngLastFileRow, lngFileCol As Long
    Dim strFilename, strTargetFolder, strTargetFile, strDownloads, strDownloadedFile, strFilesMoved As String
    Dim arrVCFiles As Variant
    Dim rngFiles, c As Range
    Dim cws As Worksheet
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    'Get the user's local 'Downloads' file path
    strDownloads = Environ("USERPROFILE") & "\Downloads"
    
    'find the header row in the file import log, plus 1 for the first data row
    lngFirstFileRow = 1 + cws.Columns("G:G").Find("File Import Log", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
    
    'find last file row in the import log
    lngLastFileRow = cws.Columns("G:G").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'protect headers
    If lngFirstFileRow = 0 Then lngFirstFileRow = 5
    If lngLastFileRow = 0 Then
        lngLastFileRow = 5
        MsgBox "No files are recorded in the file import log.  The Rolling 12 Month Forecast Combined must be run at least once to obtain target file paths", 65536, "Note:"
    End If
    
    'find the first column where files are located
    lngFileCol = cws.Rows(lngFirstFileRow - 1).Find("File Import Log", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, SearchOrder:=xlByRows).Column

    'set range of files to search through
    Set rngFiles = Range(cws.Cells(lngFirstFileRow, lngFileCol), cws.Cells(lngLastFileRow, lngFileCol))
 
    For Each c In rngFiles
        
        strFilename = c
        
        If strFilename Like "*.xls*" Then
            
            strTargetFolder = Left(c, InStrRev(c, "\") - 1) 'get last path of file
            strFilename = Mid(c, InStrRev(c, "\") + 1)

            'strip date from any file with a filename date except value classes
            If Not strFilename Like "*Value Classes*" Then
                
                strFilename = StripDateFromFilename(strFilename)
            
                If Dir(strDownloads & "\" & strFilename & "*.xls*") Like "*.xls*" Then
                    
                    'get the filename available in the downloads folder
                    strTargetFile = Dir(strDownloads & "\" & strFilename & "*.xls*")
                    
                    'see if the same file has already been downloaded since all files to find in this application are named by date
                    If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then
                        uChoice = MsgBox("The file, '" & strTargetFile & "' already exists in your respository folder." & chr(10) & chr(10) & _
                                "Do you want to replace the file in '" & strTargetFolder & "'?", vbYesNo, "Replace Target File?")
                        
                        If uChoice = vbYes Then 'delete the file first, before trying to move it
                            Kill strTargetFolder & "\" & strTargetFile 'Delete the current file already in the working folder
                            FileCopy strDownloads & "\" & strTargetFile, strTargetFolder & "\" & strTargetFile
                            DoEvents
                            If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then _
                                Kill strDownloads & "\" & strTargetFile 'delete the file from downloads after verifying it was copied
                        End If
                    Else
                        FileCopy strDownloads & "\" & strTargetFile, strTargetFolder & "\" & strTargetFile
                        DoEvents
                        If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then _
                            Kill strDownloads & "\" & strTargetFile 'delete the file from downloads after verifying it was copied
                    End If

                    'append filename to array for user message
                    strFilesMoved = strFilesMoved & "  *  " & strTargetFile & chr(10)
                End If
                
                strFilename = ""
            
            Else
'                If strFilename Like "*Value Classes*" Then
                
                    arrVCFiles = GetVCFilePathsFromQuery(strTargetFolder & "\" & strFilename)
                    
                    'make sure the Value classes array is not empty before trying to get the file paths
                    If Not IsEmpty(arrVCFiles) Then
                    
                        'loop through value classes array with fully-qualified filenames
                        For ar = 1 To UBound(arrVCFiles)
                            strTargetFolder = Left(arrVCFiles(ar, 2), InStrRev(arrVCFiles(ar, 2), "\") - 1) 'get last path of file from power query in Value Classes(Combined) file
                            strFilename = Mid(arrVCFiles(ar, 2), InStrRev(arrVCFiles(ar, 2), "\") + 1) 'get the filename only from the query
                        
                            If Dir(strDownloads & "\" & strFilename) Like "*.xls*" Then
                                
                                'get the filename available in the downloads folder
                                strTargetFile = Dir(strDownloads & "\" & strFilename)
                                
                                'see if the same file has already been downloaded since all files to find in this application are named by date
                                If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then
                                    uChoice = MsgBox("The file, '" & strTargetFile & "' already exists in your respository folder." & chr(10) & chr(10) & _
                                            "Do you want to replace the file in '" & strTargetFolder & "'?", vbYesNo, "Replace Target File?")
                                    
                                    If uChoice = vbYes Then 'delete the file first, before trying to move it
                                        Kill strTargetFolder & "\" & strTargetFile 'Delete the current file already in the working folder
                                        FileCopy strDownloads & "\" & strTargetFile, strTargetFolder & "\" & strTargetFile
                                        DoEvents
                                        If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then _
                                            Kill strDownloads & "\" & strTargetFile 'delete the file from downloads after verifying it was copied
                                    End If
                                Else
                                    FileCopy strDownloads & "\" & strTargetFile, strTargetFolder & "\" & strTargetFile
                                    DoEvents
                                    If Dir(strTargetFolder & "\" & strTargetFile) Like "*.xls*" Then _
                                        Kill strDownloads & "\" & strTargetFile 'delete the file from downloads after verifying it was copied
                                End If
                                
                                'append filename to array for user message
                                strFilesMoved = strFilesMoved & "  *  " & strTargetFile & chr(10)
                                strFilename = ""
                            End If
                        
                        Next ar
                        
'                    End If
                End If
    
            End If
            
        End If
    Next
    
    If strFilesMoved <> "" Then
        MsgBox "The following files were found in downloads and moved to their respective folders:" & chr(10) & chr(10) & _
                strFilesMoved, 65536, "Files Found and Moved"
    Else
        MsgBox "No files were found to be moved.", 65536, "No Downloads Found"
    End If

End Sub

Function GetIntFromTextMonth(strMonth) As String 'find a text month in a string and return the 2-digit integer for the month
'// -- 07/22/2022 - KP -- //

    Dim i As Long
    Dim arrMonth As Variant

    If strMonth = "" Then
        strMonth = "00"
        Exit Function
    End If
    
    'create numerical array of month name abbreviations
    ReDim arrMonth(1 To 12)
    arrMonth(1) = "Jan": arrMonth(2) = "Feb": arrMonth(3) = "Mar": arrMonth(4) = "Apr": arrMonth(5) = "May": arrMonth(6) = "Jun"
    arrMonth(7) = "Jul": arrMonth(8) = "Aug": arrMonth(9) = "Sep": arrMonth(10) = "Oct": arrMonth(11) = "Nov": arrMonth(12) = "Dec"
    
    'Retrieve a numerical date for the month stated in the filename
    For i = 1 To 12
        If InStr(strMonth, CStr(arrMonth(i))) Then
            strMonth = Format(i, "00") 'set and format month in 2-digit month
            GetIntFromTextMonth = strMonth
            Exit For
        End If
    Next i

End Function

Function GetNumberFromString(strInput) As String 'Return only a contiguous string of integers
'// -- KP - 07/20/2022 -- //

    On Error Resume Next
    
    Dim i As Long
    Dim chr, strNumber As String
    
    If Len(strInput) = 0 Then 'treat empty strInput as a number so nothing is changed
        Exit Function
    End If
    
    For i = 1 To Len(strInput) 'loop through each character in string

        chr = CStr(Mid(strInput, i, 1))
        If (Asc(chr) > 47 And Asc(chr) < 58) Then 'if character is an ascii value 48-57 (0-9)
            strNumber = strNumber & chr
        End If
        
    Next i
    
    GetNumberFromString = strNumber
    
    Exit Function
    
End Function

Sub testdateformat()

    Debug.Print GetDateFromFileName("Rolling 12 Mo Forecast All Customer 07-02-22 (June Test).xlsx", "mm-dd-yy", "yyyy_mm_dd")

End Sub

Function GetDateFromFileName(strInput, strDateFormat As String) As String 'Find a date string in multiple formats in a filename
'// -- KP - 07/21/2022 -- //                  'pass date format in as string.  If date is not in file string, the return file's last-modified date

    On Error Resume Next
    
    Dim i As Long
    Dim ch, strDate, strDelimiter As String
    
    'treat empty intput string as a number so nothing is changed
    If Len(strInput) = 0 Then
        GetDateFromFileName = ""
        Exit Function
    End If

    'strip directory path name if passed in so only the filename itself is used.  (eliminates problems with folders named with numbers)
    If InStr(strInput, "\") Then
        strInput = Mid(strInput, InStrRev(strInput, "\") + 1)
    End If
    
    For i = 1 To Len(strInput) 'loop through each character in string
        
        ch = CStr(Mid(strInput, i, 1))
        
        If ((Asc(ch) > 47 And Asc(ch) < 58) Or (Asc(ch) = 45) And Left(strDate, 1) <> "-") Then 'if character is an ascii value 48-57 (0-9) or a "-" hyphen
            strDate = strDate & ch
        End If
        
        'blank the current string if underscores or integers are present in the file name, but not part of the date
        If (Len(strDate) <> Len(strDateFormat) And (InStr(ch, " ") Or InStr(ch, chr(95)))) Then
            strDate = ""
        End If
        
        'exit loop if a hyphen is present in date string, and 'ch' is not numerical
        If (InStr(strDate, "-") And ((Asc(ch) > 47 And Asc(ch) < 58) Or Asc(ch) = 45) = False) Then
            Exit For
        End If
    Next i
                
    If strDate <> "" Then strDate = Format(strDate, strDateFormat)
    
    'Check for proper hyphen formatting - no numerical date will be longer than 10 digits, and must contain 2 hyphens or 2 slashes
    If (Len(strDate) > 10 Or Len(strDate) - Len(Replace(Replace(strDate, "-", ""), "/", "")) <> _
            Len(strDateFormat) - Len(Replace(Replace(strDateFormat, "-", ""), "/", ""))) Then
        strDate = ""
        'Exit Function '-- 'not used when returning actual file last modified date
    End If

'    Don't return anything if date was not in string
'    If strDate = "" Then Exit Function
    
    'Option - Use file last-modified date if date not in filename string not found - Added 03/30/2022 - KP
    If strDate = "" Then strDate = Format(FileDateTime(strInput), strDateFormat)
                
'    strDate = Format(strDate, strDateFormat)

    GetDateFromFileName = strDate

End Function

Function StripAlphaChars(strInput) As String 'Strip apha-characters from a mixed type string to look for carts and displays (i.e. 800000-series #'s)
'// -- KP - 07/15/2022 -- //                  'pass in the string to find only numbers

    On Error Resume Next
    
    Dim i As Long
    Dim ch, strNumbers As String
    
    'treat empty intput string as a number so nothing is changed
    If Len(strInput) = 0 Then
        StripAlphaChars = ""
        Exit Function
    End If
    
    strInput = Replace(Replace(Replace(Replace(strInput, "-", ""), " ", ""), ".", ""), ",", "")
    
    For i = 1 To Len(strInput) 'loop through each character in string
        
        ch = CStr(Mid(strInput, i, 1))
        
        If ((Asc(ch) > 47 And Asc(ch) < 58)) Then 'if character is an ascii value 48-57 (0-9)
            strNumbers = strNumbers & ch
        End If

    Next i

    StripAlphaChars = strNumbers

End Function

Function StripDateFromFilename(strInput) 'Strip numbers and hyphens from filename for bare filename without date string
'// -- KP - 07/29/2022 -- //                  'pass in the string to find only numbers

    On Error Resume Next
    
    Dim i As Long
    Dim ch, strChars, strInputString As String
    
    'if a full filepath was passed in, then strip the path and leave just the filename
    If InStr(strInput, "\") Then strInputString = Mid(strInput, InStrRev(strInput, "\") + 1)
    
    strInputString = strInput
    
    'treat empty intput string as a number so nothing is changed
    If Len(strInputString) = 0 Then
        StripDateFromFilename = ""
        Exit Function
    End If
    
    If InStr(strInputString, "_") Then
        strInputString = Left(strInputString, InStrRev(strInputString, "_") - 1)
    ElseIf InStr(strInputString, " ") Then
        strInputString = Left(strInputString, InStrRev(strInputString, " ") - 1)
    End If
    
    StripDateFromFilename = strInputString
    
End Function

Function GetMapicsSCRFilename() As String 'Get the SCR Filename for MAPICS form types
'// -- 03/07/2022 - KP -- //

    Dim lngSCRFileRow As Long
    Dim strSCRFileName As String
    
    Dim cws As Worksheet
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    On Error GoTo errCatch

    lngSCRFileRow = cws.Columns("G:G").Find("SCR_Standard_", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    If lngSCRFileRow > 0 Then
        strSCRFileName = cws.Cells(lngSCRFileRow, "G") 'get the last filename of the current month selected in the 12 Mo Rolling All Customer
    Else
        MsgBox "SCR File matching MAPICS reporting month not found." & chr(10) & chr(10) & _
        "Please download the full Oracle Analytics SCR report to " & chr(10) & _
        "  populate 'Form Type' for exploded Vita-Paks.", 65536, "SCR File Not Found"
        
        strSCRFileName = SelectImportFile(ThisWorkbook.Path, "Select the current Supply Chain Report", "SCR_Standard")
        'strSCRFilename = ""
    End If
    
    GetMapicsSCRFilename = strSCRFileName
    
    Exit Function
    
errCatch:
    blnProcessError = True
    
End Function

Sub CheckNewPC() 'Check if file is being run on a new PC. If true, then clear the file log download paths to prevent looking for non-existent files
'// -- 03/07/2022 - KP -- //

    Dim c As Range
    Dim lngLastRow, lngUserProfileRow As Long
    Dim strUserProfile As String
    Dim strSCRFileName As String
    
    Dim cws As Worksheet
    
    Set cws = ThisWorkbook.Worksheets("Controls")

    On Error Resume Next
    
    strUserProfile = Environ("USERPROFILE") 'Get the user profile from the Windows(DOS) environment
        
    'Get the last row in the file import log
    lngLastRow = cws.Columns("G:G").Find("*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'Find user profile in the file import log and get the last row
    lngUserProfileRow = cws.Columns("G:G").Find(strUserProfile, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    If lngLastRow < 5 Then Exit Sub      'no files found in the file import log
    
    If lngUserProfileRow = 0 Then
        
        For Each c In cws.Range(cws.Cells(5, "G"), cws.Cells(lngLastRow, "G"))
        
            c = "" ''Find a partial matching filename of the current month selected in the 12 Mo Rolling All Customer
            
        Next
        
    End If
    
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

Sub testGetVCFilePathsFromQuery()
    
    Dim i As Long
    Dim arrTest As Variant
    
    arrTest = GetVCFilePathsFromQuery("C:\Users\kimberly.panos\OneDrive - ivcinc.com\Documents\Automated Reports\Automated Rolling 12 Mo Forecast Combined\Value Classes(Combined).xlsx")
    
    If IsEmpty(arrTest) = False Then
        For i = 1 To UBound(arrTest)
            Debug.Print "Query: " & arrTest(i, 1), "File: " & arrTest(i, 2)
        Next i
    Else
        MsgBox "An error occurred while getting Value Class file paths.", 65536, "Get files from Queries Error"
    
    End If
    
End Sub

Function GetVCFilePathsFromQuery(strFile As String) As Variant 'find a query in a workbook to return the file it is pointed at
'// -- 07/29/2021 - KP -- //
    
    Dim blnFileValidated As Boolean
    Dim i, ac, ar As Long
    Dim lngLastSourceRow, arrFileCol As Long
    Dim strWbName, strFilePath As String
    Dim strQueryName As String
    Dim strFileNamePartial, strCurrentSourceFile As String
    Dim strPQueryName, strSheetName, strNewSourceFile As String
    Dim strFormula As String
    Dim arrSources As Variant
    Dim pQuery 'do not specify type
    Dim rngFiles, c As Range
    Dim wb As Workbook
    Dim cws, ws As Worksheet
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    Application.ScreenUpdating = False
    
    
    strFilePath = Left(strFile, InStrRev(strFile, "\") - 1)
    strWbName = Mid(strFile, InStrRev(strFile, "\") + 1)
    
    'Open the target workbook as read-only and hidden
    Set wb = Workbooks.Open(strFilePath & "\" & strWbName, False, True, , , , True)
    wb.Windows(1).Visible = False
    
    'Set value class worksheet
    Set ws = wb.Worksheets("Value Class")
    
    On Error Resume Next
    
    lngLastSourceRow = cws.Range("E3:E7").Find("*e*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'no data was found, so quit trying to get paths
    If lngLastSourceRow = 0 Then
        blnProcessError = True
        GoTo finishUp
    End If
    
    On Error GoTo 0 ' errCatch
    
    Set rngFiles = ws.Range("E3:F" & lngLastSourceRow) 'get the last row of listed files in the Value Class sheet of Value Classes(Combined) workbook
    
    arrSources = rngFiles
    
    'clear the url's from the array
    For i = 1 To UBound(arrSources)
        arrSources(i, 2) = ""
    Next i

   'Find the target query related to worksheet
    For Each pQuery In wb.Queries


        With pQuery
                                    
            strPQueryName = pQuery.Name 'actual name of the Power Query in 'Data Sources'
                
            'Loop through all to set sheet name and file source
            For ar = 1 To UBound(arrSources)
                
                'Set the filename & path if the sheet name is like the query name
                If Left(strPQueryName, InStr(strPQueryName, " ") - 1) Like Left(arrSources(ar, 1), InStr(arrSources(ar, 1), " ") - 1) & "*" Then
                    strFormula = pQuery.Formula
                    strCurrentSourceFile = Left(Mid(strFormula, InStr(strFormula, "File.Contents(") + 15), InStr(Mid(strFormula, InStr(strFormula, "File.Contents(") + 15), ".xlsx") + 4) 'get path from formula
                    arrSources(ar, 2) = strCurrentSourceFile
                    Exit For
                End If
                
            Next ar

            strCurrentSourceFile = ""
            strFormula = ""

        End With
    
    Next

    
finishUp:

    GetVCFilePathsFromQuery = arrSources
    
    On Error Resume Next
    wb.Saved = True
    wb.Close
    
    Set wb = Nothing
    Application.ScreenUpdating = True
    Exit Function
    
errCatch:
    
    MsgBox "Unable to retrieve file path from query in target workbook", 65536, "Filename From Query Error"
    Resume finishUp
    wb.Windows(1).Visible = True
    
End Function

Function GetTableRange(ByVal strTablePartName As String, strWbName As String, Optional strColumnHeaderTitle As String) As String
'// -- 12/23/2021 - KP -- //  'find a named table in a workbook by partial or whole table name in in a workbook
                              'provide table partial name, target workbook name, header row to start on if needed
    Dim oListObject As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim strTableRange As String 'string representation of range found
    Dim lngHeaderRow As Long
    Dim lngLastRow As Long, lngLastCol As Long
    Dim lngQtyFirstCol As Long 'first column where qty1 is to locate date above
    If blnProcessError Then Exit Function
    
    On Error Resume Next
    
    Set wb = Workbooks(Mid(strWbName, InStrRev(strWbName, "\") + 1))

    DoEvents

    Set wb = Workbooks.Open(strWbName, False, True, , , , True)
    
    For Each ws In wb.Sheets
        For Each oListObject In ws.ListObjects
            If oListObject.Name Like "*" & strTablePartName & "*" Then
            
                strTableRange = "[" & ws.Name & "$" & Replace(oListObject.Range.Address, "$", "") & "]"
                'If a header string for column "A" was passed in, then get the row it's on and set the public variable range for those headers
                
                If strColumnHeaderTitle <> "" Then
                    'find the header row based on the column header name passed into this Function
                    lngHeaderRow = ws.Columns("A:A").Find(strColumnHeaderTitle, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                Else 'Find header row by wildcard - Look in column "B" due to potential notes above column "A" headers
                    lngHeaderRow = ws.Columns("B:B").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                End If
                
                'find the last data row based on the string passed into this Function
                lngLastRow = ws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                'find the last used column
                lngLastCol = ws.UsedRange.Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                'put all header names into an array so the import workbook may be closed
                arrTempData = Range(ws.Cells(lngHeaderRow, 1), ws.Cells(lngLastRow, lngLastCol)).Value
                
                If LCase(wb.Name) Like "rolling*" Then
                    'Find first qty column to get date (Rolling 12 Mo Forecast has it as "M1")
                    lngQtyFirstCol = ws.UsedRange.Find("M1", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                    'Set public var for pulling matching MAPICS data when running all reports at once
                    strOracleDate = Format(ws.Cells(lngHeaderRow - 1, lngQtyFirstCol), "yyyy-mm") 'get datea above header "M1" in imported Oracle Rolling Forecast
                End If
                
            Else
                strTableRange = ""
            End If
        Next oListObject
    Next ws
    
    GetTableRange = strTableRange 'write table range found back to Function as a string
    
    wb.Saved = True
    wb.Close
    Set oListObject = Nothing
    
    Set wb = Nothing

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
    If IsMissing(strFnameFull) = False Then
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
    strFileCol = "G"    'Fully-qualified Filename Column
    strDateCol = "M"    'Date Imported Column
    lngFirstRow = 5     'First data row of file import log range
    lngLastRow = 15     'Last data row of file import log range
    
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
    
    If Not Dir(strFnameFull) Like "*.xlsx" Then  'log file information if a valid full path was passed in
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
    strFileCol = "G"    'Fully-qualified Filename Column
    lngFirstRow = 5     'First data row of file import log range
    lngLastRow = 15     'Last data row of file import log range
    
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
    If lngFileFoundRow >= 5 Then
        strFoundFileName = cws.Cells(lngFileFoundRow, lngFilesCol)
        If Not Dir(strFoundFileName) Like "*.xlsx" Then 'verify the last path was valid
            strFoundFileName = "" 'erase the filefound if file was not a valid previous import
        End If
    End If
    
    'pass found filename string back to function
    GetFileLastImported = strFoundFileName
    
finishUp:
    
    On Error Resume Next

    Exit Function
    
errCatch:

    Resume finishUp

End Function

Function GetDirFileMatch(strSourcePath, strFileNamePartial) As String
'Returns newest filename series based on filename.  Files named with dates NOT as yyyy-mm-dd format will not be in order due to alphabetical sorting
'// -- 06/06/2022 - KP -- //
    
    Dim strFileSearch, strFileFound As String
    Dim sExtension, oShellApp, oFolder, oFolderItems, oFolderItem
    
    On Error GoTo errCatch
    
    sExtension = "*.xlsx"

    If Right(strSourcePath, 1) = "\" Then strSourcePath = Left(strSourcePath, Len(strSourcePath) - 1)
    Set oShellApp = CreateObject("Shell.Application")
    Set oFolder = oShellApp.Namespace(CStr(strSourcePath))
    Set oFolderItems = oFolder.items()
                 
    oFolderItems.Filter 64 + 128, sExtension ' 32 - folders, 64 - not folders, 128 - hidden
    For Each oFolderItem In oFolderItems
        strFileSearch = strSourcePath & "\" & strFileNamePartial & "*"
       ' If InStr(oFolderItem.Path, strFileNamePartial) Then
        
        If oFolderItem.Path Like strFileSearch Then
            strFileFound = oFolderItem.Path
        End If
   
    Next
    
    strFileFound = Mid(strFileFound, InStrRev(strFileFound, "\") + 1)
    GetDirFileMatch = strFileFound
    
    Application.StatusBar = ""
    
    Exit Function
errCatch:
    MsgBox "Criteria: " & strFileNamePartial, 65536, "Error getting date from file"
    GetDirFileMatch = ""

End Function

Function GetDirLastFileByDate(strSourcePath, strFileNamePartial) As String
'Returns newest filename series based on filename with date formatted as yyyy-mm-dd due to alphabetical sorting
'// -- 07/19/2022 - KP -- //
    
    Dim strFileDate As String
    
    Dim strFileFound As String
    Dim sExtension, oShellApp, oFolder, oFolderItems, oFolderItem

    strFileDate = "1/1/2010" 'initialize date string

    sExtension = "*.xlsx"

    Set oShellApp = CreateObject("Shell.Application")
    Set oFolder = oShellApp.Namespace(CStr(strSourcePath))
    Set oFolderItems = oFolder.items()
    
    oFolderItems.Filter 64 + 128, sExtension ' 32 - folders, 64 - not folders, 128 - hidden
    
    For Each oFolderItem In oFolderItems
        If InStr(oFolderItem.Path, strFileNamePartial) Then
            
            'filename
            strFileFound = oFolderItem.Path
              
            'Compare one file date with the last, and get the newest file
            If DateDiff("n", strFileDate, FileDateTime(strFileFound)) > 0 Then
              
              strFileDate = FileDateTime(strFileFound) 'record the last file found date for the next loop
              strFileFound = oFolderItem.Path 'filename
            
            Else
              strFileFound = oFolderItem.Path
            
            End If
            
        End If
    Next
    
    strFileFound = Mid(strFileFound, InStrRev(strFileFound, "\") + 1)
    GetDirLastFileByDate = strFileFound

End Function


Function ClearSheetData(strSheetName As String) 'clears all date and dates from specified sheet
'// -- 12/20/2021 - KP --//

    Dim i As Long
    Dim uChoice
    Dim lngLastRow As Long, lngLastCol As Long
    Dim lngDataFirstRow As Long 'First data row to clear
    Dim lngHeaderRow As Long 'Row which contains data headers
    Dim ws As Worksheet, wsTemp As Worksheet
    
    blnProcessError = False
    
    If strSheetName = "" Then 'If the sheet name was not passed in from another procedure
        MsgBox "Sheet name not specified in procedure call.", 65536, "Incomplete Procedure Call"
        blnProcessError = True
        Exit Function
    End If
    
    Set wsTemp = ThisWorkbook.ActiveSheet 'Remember which sheet was active when button was pressed
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(strSheetName) 'Sheet name passed in with procedure call

    'IVC data first row below header row for the first source sheet
    lngHeaderRow = ws.UsedRange.Find("qty1", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    lngDataFirstRow = lngHeaderRow + 1
       
    lngLastRow = ws.UsedRange.Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
        
    lngLastCol = ws.UsedRange.Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
  
    If lngLastRow < 5 Then lngLastRow = 5 'Protect sheet headers
    
    Application.StatusBar = "Clearing " & strSheetName & " data . . ."
    DoEvents
       
    'Clear information from the last import
    ws.Range("A1:A3") = ""
    
    For i = 10 To lngLastCol
    
        If ((LCase(ws.Cells(lngHeaderRow, i)) Like "*qty*" Or LCase(ws.Cells(lngHeaderRow, i)) Like "val*") And _
                (LCase(ws.Cells(lngHeaderRow - 1, i)) <> "backlog" And Not LCase(ws.Cells(lngHeaderRow, i)) Like "*total*")) Then

            ws.Cells(lngHeaderRow - 1, i).ClearContents
        End If
    Next i
    
    ws.Activate
    ws.Range("A1").Select
    
    DoEvents
       
    'clear all IVC Total contents and formatting from used range, including SUM totals below data range
    Range(ws.Cells(lngDataFirstRow, 1), ws.Cells(lngLastRow, lngLastCol)).Clear
    
    ResizeScrollbars (strSheetName)
    
    wsTemp.Activate
    
    Application.ScreenUpdating = True
        
    DoEvents

End Function

Function ResizeScrollbars(strSheetName) 'Deletes empty rows below last used row and resets the worksheet side vertical scroll bars to not include empty rows
'// -- 11/24/2021 - KP -- //

    Dim lngLastRow As Long
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(strSheetName)
    
    'Get the new last row of all rows containing data or formulas
    lngLastRow = ws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

    If lngLastRow < 5 Then lngLastRow = 5
    
    'Find the end of all cells below the data that have been touched, causing excessive blank rows to be displayed
    ws.Range("A" & lngLastRow + 1 & ":" & Replace(Mid(ws.UsedRange.Address, InStrRev(ws.UsedRange.Address, ":") + 1), "$", "")).EntireRow.Delete
    
    ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell) 'Reset the scrollbars to the proper size to end where the cell data ends
    
End Function

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

Function IsNum(strInput) As Boolean 'Determine if input string contains numbers
'// -- KP - 07/26/2021 -- //

    On Error Resume Next
    
    Dim i As Long
    Dim strVal As String
    
    If Len(strInput) = 0 Then 'treat empty strInput as a number so nothing is changed
        IsNum = True
        Exit Function
    End If
    
    For i = 1 To Len(strInput) 'loop through each character in string

        strVal = CStr(Mid(strInput, i, 1))
        If (Asc(strVal) > 47 And Asc(strVal) < 58) Then 'if character is an ascii value 48-57 (0-9)
            IsNum = True
        Else
            IsNum = False
            Exit For
        End If
        
    Next i
        
    Exit Function
    
End Function

Sub testPillSums()
    StopWatch ("Start")
    
    SetPillsSumFormulas ("IVC Total")

    Application.StatusBar = "Pill Sum Formula Test Took: " & StopWatch("Stop")
End Sub

Sub SetPillsSumFormulas(strSheetName As String) 'Create totals formulas down the rows for each Pills qtyX and Valx columns on specified sheet
'// -- 07/29/2022 - KP -- //

    Dim qtyXcol, r, c As Long
    Dim lngCTcol As Long
    Dim lngQtyFirstCol As Long, lngQty12col As Long
    Dim lngPillQtyFirstCol As Long, lngPillQty12Col As Long
    Dim lngValFirstCol, lngVal12Col As Long
    Dim lngLastRow As Long, lngLastCol As Long, lngPillTotCol As Long 'last header column
    Dim ws, cws As Worksheet

    If blnProcessError Then Exit Sub 'exit if there was an error upstream of this process
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    Set ws = ThisWorkbook.Worksheets(strSheetName)

    On Error Resume Next
    
    Application.ScreenUpdating = False 'turn off screen updating for speed
    
    Application.Calculation = xlCalculationManual 'turn off auto calculation for speed

    'Get the last data row
    lngLastRow = ws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'Get the last data column
    lngLastCol = ws.Range("A4").CurrentRegion.Columns.Count

    'Protect header titles and formats
    If lngLastRow < 5 Then lngLastRow = 5

    'Find the first qty column
    lngQtyFirstCol = ws.Rows("4:4").Find("qty", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlNext).Column
    
    'get the last qty column before total
    lngQty12col = ws.Rows("4:4").Find("qty12", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlToLeft).Column
    
    'get the first column with $ values
    lngValFirstCol = ws.Rows("4:4").Find("val", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlNext).Column
    
    'get the last column with $ values
    lngVal12Col = ws.Rows("4:4").Find("val12", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlToLeft).Column
    
    'get the first Pill Qty1 column
    lngPillQtyFirstCol = ws.Rows("4:4").Find("Pills Qty*", LookIn:=xlValues, MatchCase:=False, LookAt:=xlWhole, SearchDirection:=xlNext).Column ' First pills qty 0 column
    
    'get the last Pill Qty column
    lngPillQty12Col = ws.Rows("4:4").Find("lls Qty12", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlToLeft).Column 'last pills qty 12 column

    'get the CT count column
    lngCTcol = ws.Rows("4:4").Find("CT", LookIn:=xlValues, MatchCase:=False, LookAt:=xlPart, SearchDirection:=xlToLeft).Column

    'Set pill qty total column
    lngPillTotCol = lngPillQty12Col + 1

    If lngQtyFirstCol = 0 Then GoTo errCatch 'if the pill column can't be found, then exit sub

    Application.StatusBar = "Inserting Pill Qty sub-total formulas . . . "
    
    DoEvents

    On Error GoTo errCatch 'in case column not found or name changed
    
    'Copy all pill qty and total formulas except IVC Totals sheet
    If Not ws.Name Like "*IVC?Total" Then
    
        'only if explode VP's/MULTIPLE packets is selected
        If cws.Range("M1") = True Then
        
            'insert the pill quantity0 forumla and number format to copy across and down the page
            ws.Cells(5, lngPillQtyFirstCol).Formula = "=$" & Replace(ws.Cells(5, lngCTcol).Address, "$", "") & "*" & Replace(ws.Cells(5, lngQtyFirstCol).Address, "$", "")
        
            'copy Pill Qty1 formula and format across to Pill Qty12
            ws.Cells(5, lngPillQtyFirstCol).Copy
                 Range(ws.Cells(5, lngPillQtyFirstCol + 1), ws.Cells(5, lngPillQty12Col)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone
    
            ws.Cells(5, lngPillTotCol).Formula = "=SUM(" & Replace(ws.Cells(5, lngPillQtyFirstCol).Address & ":" & ws.Cells(5, lngPillQty12Col).Address, "$", "") & ")"
            Range(ws.Cells(5, lngPillQtyFirstCol), ws.Cells(5, lngPillTotCol)).Copy
            Range(ws.Cells(6, lngPillQtyFirstCol), ws.Cells(lngLastRow + 2, lngPillTotCol)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone
            Application.CutCopyMode = False
            
            'insert missing row subtotals for qty and $ values for new Oracle Backlog items
            If ws.Name = "Oracle" Then
                ws.Range("E1").Activate 'de-select pasted cells
                For r = 5 To lngLastRow
                    If ws.Cells(r, lngQty12col + 1) = "" Then
                        ws.Cells(r, lngQty12col + 1) = Application.WorksheetFunction.SUM(Range(ws.Cells(r, lngQtyFirstCol), ws.Cells(r, lngQty12col)))
                    End If
                    
                    If ws.Cells(r, lngVal12Col + 1) = "" Then
                       ws.Cells(r, lngVal12Col + 1) = Application.WorksheetFunction.SUM(Range(ws.Cells(r, lngValFirstCol), ws.Cells(r, lngVal12Col)))
                    End If
                Next r
            End If
                    
        Else
            Range(ws.Cells(5, lngPillQtyFirstCol), ws.Cells(5, lngPillTotCol)).ClearContents
        End If
                
    End If

    'put SUM formula below entire dataset qty and val columns and autofit columns if they are too narrow for sums to be read
    For c = lngQtyFirstCol To lngLastCol
            ws.Cells(lngLastRow + 2, c).Formula = "=SUM(" & ws.Cells(5, c).Address & ":" & ws.Cells(lngLastRow, c).Address & ")"
            If c < lngLastCol Then ws.Columns(c).EntireColumn.AutoFit 'don't auto-fit the last column
    Next c

    Range(ws.Cells(lngLastRow + 1, 1), ws.Cells(lngLastRow + 1, lngLastCol)).Clear 'remove any formulas or formatting between the last data row and column totals

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Exit Sub

errCatch:
    MsgBox "There was an error in " & strSheetName & " Pill Qty & Totals formula fill process." & chr(10) & chr(10) & _
        "Please make sure data columns or names have not changed.", 65536, "Pill Calc Formula Error" & Err.Description
    blnProcessError = True
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub SetColumnFormat(strSheetName) 'formats all quantity and dollar value columns
'// -- 05/17/2021 - KP -- //

    Dim c As Range
    Dim lngLastRow As Long, lngLastCol As Long
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(strSheetName)

    On Error GoTo errCatch 'in case column not found or name changed

    Application.StatusBar = "Setting number formats . . ."
    DoEvents
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    'Get the last data row
    lngLastRow = ws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

    'Protect header titles and formats
    If lngLastRow < 5 Then lngLastRow = 5

    'Find the last column on the sheet
    lngLastCol = ws.Range("A4").CurrentRegion.Columns.Count

    'Set number formats for each column

    For Each c In Range(ws.Cells(4, 2), ws.Cells(4, lngLastCol)) 'start in column to to skip the "V" in IVC as a formatting case
        
        c.Value = Replace(c, "$", "")

        Select Case True
        
            Case (LCase(c) Like "*value class*") 'format cost and price columns as dollars with comma
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).HorizontalAlignment = xlCenter
                c.ColumnWidth = "6.0"
                
            Case (LCase(c) Like "*cost*" Or LCase(c) Like "*price*") 'format cost and price columns as dollars with comma
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).NumberFormat = "$ #,##0.00_);[Red]($ #,##0.00); - "
            
            Case LCase(c) Like "child*" 'set the "Children" format with "-" hash for zero
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).NumberFormat = " #,##_);; - "
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).HorizontalAlignment = xlCenter
                c.ColumnWidth = "9.1"
            
            Case LCase(c) Like "abc code" 'set ABC Code column with and align = center
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).HorizontalAlignment = xlCenter
                c.ColumnWidth = "6.7"
                
            Case LCase(c) Like "ct" 'set the "CT" format with "-" hash for zero
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).NumberFormat = " #,##_);; - "
                c.ColumnWidth = "9.1"
            
                'formats every Qty column to number with commas, including totals with "-" hash for zeros
            Case (LCase(c) Like "*q*" Or (LCase(c) Like "*total*" And LCase(Cells(c.Row, c.Column - 1) Like "*q*")))
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).NumberFormat = "#,##0_);[Red]-#,##0; - "
                If Not LCase(c) Like "*total pills*" Then c.EntireColumn.AutoFit 'Don't fit "12 month pill total", because it the header name is so long
               
                'formats every "Val#" and value Value total column as currency with "-" hash for zero dollars
            Case (LCase(c) Like "val*" Or (LCase(c) Like "*vtotal*"))
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).NumberFormat = "$ #,##0.00_);[Red]($ #,##0.00); - "
                c.EntireColumn.AutoFit
            
                'Aligns all matching colums to center
            Case (LCase(c) Like "* loc" Or LCase(c) Like "*class" Or LCase(c) Like "*type" Or _
                        LCase(c) Like "*code" Or LCase(c) = "uom" Or LCase(c) = "uom" Or LCase(c) Like "mto")
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).HorizontalAlignment = xlCenter
                
                'Indents matching columns 1
            Case (LCase(c) Like "cust*" Or LCase(c) = "type" Or LCase(c) Like "part*" Or _
                        LCase(c) Like "*bulk*" Or LCase(c) = "uom" Or LCase(c) = "uom")
                Range(ws.Cells(5, c.Column), ws.Cells(lngLastRow + 2, c.Column)).IndentLevel = 1
            
            Case Else

        End Select
        
        DoEvents
        
    Next

    Range(ws.Cells(lngLastRow + 1, 1), ws.Cells(lngLastRow + 1, lngLastCol)).Clear 'remove any formulas or formatting between the last data row and column totals

    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ws.Activate
    ws.Range("A1").Activate 'scroll all the way to the left again
    ws.Range("F4").Activate 'Then move off import stats

    Exit Sub

errCatch:
    MsgBox "Error while formatting " & strSheetName & "." & chr(10) & chr(10) & _
        "Error code: " & Err.Description, 65536, "Number Format Error"
    blnProcessError = True
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub SetBorderAndFillFormat(strSheetName As String) 'sets alternate row shade conditional formatting and cell borders for
'// -- 04/08/2022 - KP -- //                       'any sheet with headers starting on row 4 - Pass in sheet name as string

    Dim lngLastRow As Long, lngLastCol As Long
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(strSheetName) 'Sheet name to set conditional formatting for alternate shaded rows and vertical borders
    
    On Error Resume Next 'in case column not found or name changed

    'Get the last data row
    lngLastRow = ws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    'Protect header titles and formats
    If lngLastRow < 5 Then lngLastRow = 5

    'get the first Pill qty0 column
    lngLastCol = ws.Range("A4").CurrentRegion.Columns.Count

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    With Range(ws.Cells(5, 1), ws.Cells(lngLastRow, lngLastCol))
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

    Application.ScreenUpdating = True
    Application.Goto Reference:=Worksheets(strSheetName).Range("E3"), Scroll:=True 'scroll full left and park on an innocuous cell
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

End Sub

Sub UpdateWorksheetList() 'Update list of worksheets in this workbook on 'Controls' worksheet
'// -- 07/07/2022 -- KP //

    Dim ws As Worksheet 'worksheets used in loop to identify and rename sheet
    Dim cws As Worksheet 'Controls worksheet
    Dim hr As Long 'Row to drop each worksheet name
    Dim c As Range 'cells used to loop worksheet
    Dim strHeader As String 'cell where sheet list header is found
    Dim lngHcol As Long
    Dim i As Long
    
    Set cws = ThisWorkbook.Worksheets("Controls")
    
    On Error GoTo errCatch 'Jump to error message to notify user
    
    ThisWorkbook.Activate 'Make sure focus is on this workbook
            
    DoEvents 'Wait for sheets to be created before renaming them

    cws.Unprotect
    
    For Each c In cws.UsedRange 'Find the column where the sheets list is
    
        If LCase(c) Like "*sheets*" Then
            strHeader = c.Address
            hr = Mid(strHeader, InStrRev(strHeader, "$") + 1) + 1
            lngHcol = c.Column
        End If
    
    Next

    'clean up empty rows in list
    Range(cws.Cells(hr, lngHcol), cws.Cells(30, lngHcol)).ClearContents

    For Each ws In ThisWorkbook.Worksheets
        'write the worksheet name on the sheet list except the name of the "controls" sheet
        If Not LCase(ws.Name) Like "control*" Then
            cws.Cells(hr, Replace(Left(strHeader, InStrRev(strHeader, "$")), "$", "")) = ws.Name 'write the worksheet name on the sheet list
            hr = hr + 1
        End If
    Next
    
    cws.Protect
    
    Exit Sub
    
errCatch:
    MsgBox "There was an error creating the workbook worksheets layout." & chr(10), 65536, "Worksheet Creation Error"

End Sub

Function ButtonsShow(blnShow As Boolean, Optional strWb As String) 'Toggle all buttons in workbook visible or not - for exporting to non VBA model
'// -- 07/26/2022 - KP -- //

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    If strWb = "" Then strWb = ThisWorkbook.Name
    
    Set wb = Workbooks(strWb)
        
    For Each ws In wb.Worksheets
        If Not (LCase(ws.Name) Like "*controls*" Or _
            LCase(ws.Name) Like "sop*" Or _
            LCase(ws.Name) Like "instr*" Or _
            LCase(ws.Name) Like "test*") Then
            For Each shp In ws.Shapes 'only toggle buttons, as charts are also detected as shapes
                If (LCase(shp.Name) Like "shp*" Or _
                    LCase(shp.Name) Like "rect*" Or _
                    LCase(shp.Name) Like "group*" Or _
                    LCase(shp.Name) Like "btn*" Or _
                    LCase(shp.Name) Like "chk*") Then
                    
                    If blnShow = True Then shp.Visible = msoTrue
                    If blnShow = False Then shp.Visible = msoFalse
                End If
            Next
        End If
    Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Function

Function StopWatch(strStartStop As String) As String 'pass in "start" or "stop".  "Stop" will return the number of minutes & seconds
'// -- 12/22/2021 - KP -- //
        
    On Error Resume Next
    
    Dim lngTime
    Dim strTime
    
    Select Case LCase(strStartStop)
        
        Case "start" 'get the start time and divide by number of seconds in the day
            dtTimeStart = Date + CDate(Timer / 86400)
        
        Case "stop" 'subtract the start time from the end time, and multipy by seconds in the day, round to 1 decimal place
            dtTimeStop = Date + CDate(Timer / 86400)
            lngTime = Round((dtTimeStop - dtTimeStart) * 86400, 1) 'return seconds with tenths of second
            If lngTime > 59 Then
                strTime = Int(lngTime / 60) & " Min " & lngTime Mod 60 & " Sec"
            Else
                strTime = CStr(lngTime) & " Seconds"
            End If
            
            StopWatch = strTime
            
    End Select
    
End Function

Function GetLocalPathFromSP(strWbPath As String) As String 'convert personal sharepoint path to local path by workbook name
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

Sub TestUpdatePivotTables()

    blnAutoProcess = False
    
    UpdatePivotTables (ThisWorkbook.Name)
    
End Sub

Function UpdatePivotTables(Optional strWbName As String) 'updates all pivot tables pointed at IVC Total sheet in the workbook specified
'// -- 07/08/2022 - KP -- //
    
    Dim sCount As Long
    Dim lngIVCTotLastRow  As Long
    Dim lngWsLastRow, lngWsLastCol As Long 'worksheets in loop
    Dim strPvtSource As String 'Get worksheet name from pivot table
    Dim ivc As Worksheet
    Dim strPvtRange As String 'store new concatenated pivot table range
    Dim strDataDate As String
    
    Dim pvt As PivotTable 'for pivot table looping
    Dim wb As Workbook
    Dim ws As Worksheet 'for worksheet loop
    Dim pws As Worksheet 'Sheet where pivot table is pointed
    Dim iws As Worksheet 'IVC Total sheet
       
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Updating all pivot tables . . ."
    Application.ScreenUpdating = False
    
    'if workbook name not specified, then update pivots in this workbook
    If strWbName = "" Then strWbName = ThisWorkbook.Name

    Set wb = Workbooks(strWbName)
    Set iws = wb.Worksheets("IVC Total")
    
    'allow manual proccessing of only MAPICS or Oracle
    If blnAutoProcess = True Then
        On Error GoTo errCatch
    Else
        On Error Resume Next
    End If
    
    'get last data row on IVC sheet to make sure there is data before updating tables
    lngIVCTotLastRow = iws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
'    If lngIVCTotLastRow < 400 Then 'don't do anything with the pivot tables if data is missing from the IVC Total sheet
'        blnProcessError = True
'        Exit Function
'    End If
    
    strDataDate = Format(iws.Cells(3, iws.Rows("4:4").Find("qty1", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, _
            SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column), "mmm-yyyy")
      
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For sCount = 1 To wb.Sheets.Count 'Loop through all sheets in the workbook (for some reason 'For Each ws' isn't working here')
    
        'Set worksheet for loop
        Set ws = wb.Sheets(sCount)



        If ws.PivotTables.Count > 0 Then 'When a pivot table is found

            For Each pvt In ws.PivotTables 'loop through each pivot table on the current sheet
                
'                If LCase(pvt.Name) Like "pvt_*" Then 'If the pivot table matches "pvt_ivct_" to only update named pivot tables pointed at IVC Total sheet

                    'Get and current data source sheet name from the individual pivot table
                    If InStr(pvt.PivotCache.SourceData, "]") > 0 Then 'Get sheet name without path regardless of SharePoint or not
                        strPvtSource = Mid(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "]") + 1)
                        strPvtSource = Replace(Left(strPvtSource, InStrRev(strPvtSource, "!") - 1), "'", "")
                    Else
                        strPvtSource = Replace(Left(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "!") - 1), "'", "")
                    End If
                    
                    Set pws = wb.Worksheets(strPvtSource)
                    
                    'get last data column and row on sheet where data for the pivot table was found
                    lngWsLastRow = pws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                    lngWsLastCol = pws.UsedRange.Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                    
                    'get new source data range for pivot tables pointed at its source sheet
                    strPvtRange = Left(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "R")) & lngWsLastRow & "C" & lngWsLastCol
                    
                    'Strip external workbook path and workbook name from source reference
                    If InStrRev(strPvtRange, "]") > 0 Then strPvtRange = "'" & Mid(strPvtRange, InStrRev(strPvtRange, "]") + 1)
                    
                    'update pivot table range to include any changes to data range IVC Total sheet
                    pvt.ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=strPvtRange)
                    
                    'refresh the current pivot table data source values
                    pvt.PivotCache.Refresh
                    
                    'Filter blanks from "Bulk_IVCN" since the blanks represent non-IVCN items
                    If pvt.Name Like "*Bulk*" And pvt.Name Like "*IVCN*" Then 'Use dual wildcards in case someone removes the underscore from the pivot name
                        pvt.PivotFields("IVCN Bulk").PivotItems("(blank)").Visible = False
                    End If
                    
                    DoEvents
                    
'                End If
                Set pws = Nothing
            
            Next
            'write date of data source and timestamp for last run
            ws.Range("D1") = "Source Data For: " & strDataDate & "  ||  Last run: " & Trim(Mid(iws.Range("A3"), InStr(iws.Range("A3"), ":") + 1))
            
        End If
    
    Next sCount
 
finishUp:

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Function
    
errCatch:
    blnProcessError = True
    MsgBox "There was an error while updating pivot table data ranges.", 65536, "Error: " & Err.Description
    ResetApplication
    Resume finishUp
    
End Function

Sub UpdatePivotTablesManually()

    blnAutoProcess = False
    
    StopWatch "Start"
    
    UpdatePivotsClearExtReference (ThisWorkbook.Name)
    
    Application.StatusBar = "Pivot Table Refresh completed: " & Format(Now(), "h:mm AM/PM") & "  ||  Process Time: " & StopWatch("Stop")
    
    
End Sub

Function UpdatePivotsClearExtReference(strWbName As String) 'updates all pivot tables by workbook name
'// -- 02/16/2022 - KP -- //
    
    Dim sCount As Long
    Dim lngIVCTotLastRow  As Long
    Dim lngWsLastRow, lngWsLastCol As Long 'worksheets in loop
    Dim strPvtSource As String 'Get worksheet name from pivot table
    Dim ivc As Worksheet
    Dim strPvtRange As String 'store new concatenated pivot table range
    Dim strDataDate As String
    
    Dim pvt As PivotTable 'for pivot table looping
    Dim wb As Workbook
    Dim ws As Worksheet 'for worksheet loop
    Dim pws As Worksheet 'Sheet where pivot table is pointed
        Dim iws As Worksheet 'IVC Total sheet
       
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'if workbook name not specified, then update pivots in this workbook
    If strWbName = "" Then strWbName = ThisWorkbook.Name

    Set wb = Workbooks(Replace(strWbName, Application.PathSeparator, ""))
    Set iws = wb.Worksheets("IVC Total")
    
    On Error GoTo errCatch
    
    'get last data row on IVC sheet to make sure there is data before updating tables
    lngIVCTotLastRow = iws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    
    If lngIVCTotLastRow < 400 Then 'don't do anything with the pivot tables if data is missing from the IVC Total sheet
        blnProcessError = True
        MsgBox "Data must be combined before pivot tables can be updated!", 65536, "Combine Data First"
        Exit Function
    End If
    
    strDataDate = Format(iws.Cells(3, iws.Rows("4:4").Find("qty1", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, _
            SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column), "mmm-yyyy")
      
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For sCount = 1 To wb.Sheets.Count 'Loop through all sheets in the workbook (for some reason 'For Each ws' isn't working here')
    
        'Set worksheet for loop
        Set ws = wb.Sheets(sCount)
        
        If ws.PivotTables.Count > 0 Then 'When a pivot table is found

            For Each pvt In ws.PivotTables 'loop through each pivot table on the current sheet
                
                'Get pivot table data source name from pivot table
                'strpvtsource =

                If LCase(pvt.Name) Like "pvt_*" Then 'If the pivot table matches "pvt_ivct_" to only update named pivot tables pointed at IVC Total sheet

                    'Get and current data source sheet name from the individual pivot table
                    If InStr(pvt.PivotCache.SourceData, "]") > 0 Then 'Get sheet name without path regardless of SharePoint or not
                        strPvtSource = Mid(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "]") + 1)
                        strPvtSource = Replace(Left(strPvtSource, InStrRev(strPvtSource, "!") - 1), "'", "")
                    Else
                        strPvtSource = Replace(Left(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "!") - 1), "'", "")
                    End If
                    
                    strPvtSource = Replace(strPvtSource, ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name, "")
                    
                    Set pws = wb.Worksheets(strPvtSource)
                    
                    'get last data column and row on sheet where pivot table was found
                    lngWsLastRow = pws.Columns("A:A").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
                    lngWsLastCol = pws.UsedRange.Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                    
                    'get new source data range for pivot tables pointed at its source sheet
                    strPvtRange = Left(pvt.PivotCache.SourceData, InStrRev(pvt.PivotCache.SourceData, "R")) & lngWsLastRow & "C" & lngWsLastCol
                    
                    'Strip external workbook path and workbook name from source reference
                    If InStrRev(strPvtRange, "]") > 0 Then strPvtRange = "'" & Mid(strPvtRange, InStrRev(strPvtRange, "]") + 1)
                    
                    'update pivot table range to include any changes to data range IVC Total sheet
                    pvt.ChangePivotCache wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=strPvtRange)
                    
                    'refresh the current pivot table data source values
                    pvt.PivotCache.Refresh
                    
                    DoEvents
                    
                End If
            
            Next
            'write date of data source and timestamp for last run
            ws.Range("D1") = "Source Data For: " & strDataDate & "  ||  Last run: " & Trim(Mid(iws.Range("A3"), InStr(iws.Range("A3"), ":") + 1))
            
        End If
    
    Next sCount
 
finishUp:

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Function
    
errCatch:
    blnProcessError = True
    MsgBox "There was an error while updating pivot table data ranges.", 65536, "Error: " & Err.Description
    ResetApplication
    Resume finishUp
    
End Function

Function ChkProcess(process As String) 'Used to see if external process is running (such as the user's Macro Help Notepad window)
'// -- 01/06/2022 - KP -- //
    
    Dim objList As Object

    Set objList = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        ChkProcess = True
    Else
        ChkProcess = False
    End If

End Function

Sub ProcessAll_EnableDisable() 'Hides or shows buttons on each sheet so users don't get confused with separate buttons on sheets
'// -- 07/26/2022 - KP -- //

    Dim strTextDialog As String
    Dim strSheetName As String
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim shp As Shape
    Dim sws As Worksheet
        
    Set wb = Workbooks(ThisWorkbook.Name)
    
    Set sws = wb.Worksheets("SOP")
    
    sws.Unprotect
    
    'Toggle SOP messages to auto or manual indication
    If (sws.Shapes("btn_ProcessAll").TextFrame.Characters.Text Like "*Manual*" Or _
            sws.Shapes("btn_ProcessAll").TextFrame.Characters.Text Like "You have selected*") Then  'Toggle manual Oraclem, MAPICS, and Combined Buttons off
        blnManualProcessing = False
        'text message under Run All Process button
        sws.Shapes("btn_ProcessAll").TextFrame.Characters.Text = "Current Mode: Automatic"
        strTextDialog = "Automatic Mode: Buttons for individual processing of Oracle, MAPICS, and IVC Total are hidden." & chr(10) & chr(10) & _
                "Manual mode is typically used to provide separate MAPICS or Oracle reports. or identify potential data integrity issues.  " & _
                "This allows Oracle or MAPICS data to manipulated manually before combining the modified data with the " & _
                "'Combine Data' button on the IVC Total sheet, which still updates all pivot tables when combined."
                
        'Set standard color for Import and Combine All button (Orange) and unlock the button so it can be pressed
        With sws.Shapes("btn_AllReports")
            .Locked = False
            .Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Fill.Transparency = 0
        End With

    Else
        blnManualProcessing = True
        blnAutoProcess = False
        ButtonsShow (True) 'restore the manual processing buttons on the Oracle, Mapics, and IVC Totals sheets
        
        'text message under Run All Process button
        sws.Shapes("btn_ProcessAll").TextFrame.Characters.Text = "Current Mode: Manual"
        strTextDialog = "You have selected to process Oracle, MAPICS and IVC Total combined sheets separately.  Individual 'Import' buttons are now visible " & _
                    "for Oracle and MAPICS sheets, and 'Combine Data' on the IVC Total sheet.  This will allow you to import MAPICS and Oracle data separately to make manual , " & _
                    "changes before combining on the 'IVC Total' sheet." & _
                    chr(10) & chr(10) & "Note: If you import Oracle or MAPICS data separately, it will clear the IVC Total sheet to " & _
                    "maintain data consistency between the IVC Total and the Oracle and MAPICS sheets. " & _
                    "  You may then click the 'Combine Data' button on the 'IVC Total' sheet, which also updates all pivot tables."

        
        'Gray out the 'Import and Combine All' button and lock it so it can't be pressed while in manual mode
        With sws.Shapes("btn_AllReports")
            .Locked = False
            .Fill.ForeColor.RGB = RGB(180, 180, 180)
            .Fill.Transparency = 0.8
        End With
        
    End If

    sws.Shapes("TxBx_RunAllPrompt").TextFrame.Characters.Text = strTextDialog

    For Each ws In wb.Worksheets
        
        If (LCase(ws.Name) Like "*mapics*" Or LCase(ws.Name) Like "*oracle*" Or LCase(ws.Name) Like "*ivc total*") Then
                
            If blnManualProcessing = True Then

                For Each shp In ws.Shapes
                    If (LCase(shp.Name) Like "btn*" Or LCase(shp.Name) Like "chk*") Then
                        shp.Visible = msoTrue
                    End If
                Next
                
'                For Each ctl In ws.Controls
'                    If LCase(ctl.Name) Like "chk*" Then
'                        shp.Visible = msoTrue
'                    End If
'                Next
                
            End If
            
            If blnManualProcessing = False Then
                        
                For Each shp In ws.Shapes
                    If (LCase(shp.Name) Like "btn*" Or LCase(shp.Name) Like "chk*") Then
                        shp.Visible = msoFalse
                    End If
                Next
                
            End If
                            
        End If
    Next
        
    sws.Protect
        
End Sub

Sub TestLOCsPresent()

    StopWatch ("Start")
    
    Debug.Print CheckRollingForLOC("C:\Users\kimberly.panos\OneDrive - ivcinc.com\Documents\_Imported Data\Rolling 12 " & _
                "Mo Forecast & SCR\Rolling 12 Mo Forecast All Customer 04-06-22.xlsx", "[FCST$A8:AM9]")
    

    Application.StatusBar = "Time to Test: " & StopWatch("Stop")
    
    
End Sub

Function CheckRollingForLOC(strSourceFile, strTableRange) As Boolean 'Determine if '12 Mo Rolling All Customer' contains MFG and PKG locations
'// -- 04/08/2022 - KP -- //                          'If so, then use those instead of SCR locations

    Dim i As Long
    Dim lngLocCount As Long
    Dim ArrTemp As Variant
    Dim strSQL As String

    strSQL = "SELECT Top 1 * FROM " & strTableRange
                                                
    dbOpen (strSourceFile) 'initialize connection string for database(workbook) and open connection
    
    If blnProcessError = False Then
        Application.StatusBar = "Checking for MFG & PKG LOC in source rolling forecast:  " & Mid(strSourceFile, InStrRev(strSourceFile, "\") + 1)
        DoEvents
    End If

    'run the SQL query with the database connection and return results to ADODB recordset
    rs.Open strSQL, cnn, adOpenStatic, adLockOptimistic

    DoEvents
    
    ' make sure the recordset isn't empty before continuing
    On Error Resume Next
    
    If rs.RecordCount < 1 Then 'make sure the query returned results
        Set rs = Nothing
        CheckRollingForLOC = False
        Exit Function
    End If

    'Loop through headers verify header names
    For i = 0 To rs.Fields.Count - 1
    
        'Verify resordset contains location information
        If UCase(rs.Fields(i).Name) Like "*LOC" Then 'make sure to not wipe out the header names
                lngLocCount = lngLocCount + 1
        End If
        
'        'Check for the presence of backlog in Oracle data  - (Concept for adding Oracle Backlog - Using a different strategy for now - KP 07/18/2022)
'        If UCase(rs.Fields(i).Name) = "M0" Then
'            blnOracleBackLog = True '(Public Var)
'        End If

    Next i
        
    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    
    If lngLocCount = 2 Then
        CheckRollingForLOC = True
    Else
        CheckRollingForLOC = False
    End If

End Function

Function TransposeArray(ArrTemp As Variant) As Variant 'Transpose Array
'// -- 02/24/2022 - KP -- //

    Dim r, c As Long
    Dim arr As Variant
    
    ReDim arr(1 To UBound(ArrTemp, 2) + 1, 1 To UBound(ArrTemp) + 1)
    
    For r = 1 To UBound(ArrTemp, 2) + 1
        For c = 1 To UBound(ArrTemp) + 1
            arr(r, c) = ArrTemp(c - 1, r - 1)
        Next c
    Next r
    
    TransposeArray = arr

End Function


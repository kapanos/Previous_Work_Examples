Sub FileScavenger() 'Find all known and required files for this report in the users 'Download' folder, and move them to the last used folders for each respective file
'// -- 07/29/2022 - KP -- //   'this includes reading the filenames stored in the Value Classes(Combined).xlsx file and moves those as well.
    
    Dim blnFilesMoved As Boolean
    Dim uChoice As Integer
    Dim i, ar As Long
    Dim lngFirstFileRow, lngLastFileRow, lngFileCol As Long
    Dim strFileName, strTargetFolder, strTargetFile, strDownloads, strDownloadedFile, strFilesMoved As String
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
        
        strFileName = c
        
        If strFileName Like "*.xls*" Then
            
            strTargetFolder = Left(c, InStrRev(c, "\") - 1) 'get last path of file
            strFileName = Mid(c, InStrRev(c, "\") + 1)

            'strip date from any file with a filename date except value classes
            If Not strFileName Like "*Value Classes*" Then
                
                strFileName = StripDateFromFilename(strFileName)
            
                If Dir(strDownloads & "\" & strFileName & "*.xls*") Like "*.xls*" Then
                    
                    'get the filename available in the downloads folder
                    strTargetFile = Dir(strDownloads & "\" & strFileName & "*.xls*")
                    
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
                
                strFileName = ""
            
            Else
'                If strFilename Like "*Value Classes*" Then
                
                    arrVCFiles = GetVCFilePathsFromQuery(strTargetFolder & "\" & strFileName)
                    
                    'make sure the Value classes array is not empty before trying to get the file paths
                    If Not IsEmpty(arrVCFiles) Then
                    
                        'loop through value classes array with fully-qualified filenames
                        For ar = 1 To UBound(arrVCFiles)
                            strTargetFolder = Left(arrVCFiles(ar, 2), InStrRev(arrVCFiles(ar, 2), "\") - 1) 'get last path of file from power query in Value Classes(Combined) file
                            strFileName = Mid(arrVCFiles(ar, 2), InStrRev(arrVCFiles(ar, 2), "\") + 1) 'get the filename only from the query
                        
                            If Dir(strDownloads & "\" & strFileName) Like "*.xls*" Then
                                
                                'get the filename available in the downloads folder
                                strTargetFile = Dir(strDownloads & "\" & strFileName)
                                
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
                                strFileName = ""
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
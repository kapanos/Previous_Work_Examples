Option Explicit

Sub Manual_Report_Calculation()

    Call CreateOutputReport
    Call ExportToPDF
    
End Sub

Sub PopulateMissingValues() 'Calculate missing values for Coal Consumption report only
'// -- 09/03/2020 KP -- //

    Dim c As Long, r As Long
    Dim dblLastValue As Double, dblNextValue As Double, dblDivisor As Double
    Dim dblPadValue As Double 'differential value for incremental padded adjustments
    Dim lngFirstEmptyRow As Long, lngLastEmptyRow As Long
    Dim lngDivisorAdd As Long
    Dim lngLastRow As Long
    Dim lngLastCol As Long
    Dim cpws As Worksheet
    Dim arrColArray As Variant
        
    Set cpws = Worksheets("Control_Panel")

    lngLastRow = cpws.Columns("C:C").Find("*", LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row
    
    lngLastCol = cpws.Cells(1, cpws.Range("1:1").Columns.Count).End(xlToLeft).Column
            
    If lngLastRow < 4 Then Exit Sub 'exit if no rows were below tag headers
    
    For c = 4 To lngLastCol 'Start on first tag after timestamps
    
        If Application.WorksheetFunction.CountIf(Range(cpws.Cells(3, c), cpws.Cells(lngLastRow, c)), "") Then 'only loop if blanks are found
            
            ReDim arrColArray(1 To lngLastRow - 2) As Variant 'Dim array - rowcount needs to be offset for starting below headers
            
            arrColArray = Application.WorksheetFunction.Transpose(Range(cpws.Cells(3, c), cpws.Cells(lngLastRow, c)).value)

            r = 1
    
            Do While r <= UBound(arrColArray) 'Loop through array to find missing values
    
                If Len(arrColArray(r)) = 0 Then 'If empty value is found
    
                    dblLastValue = arrColArray(r - 1) 'record last non-empty value
                    
                    lngFirstEmptyRow = r 'row where first empty value was found
                                  
                    lngDivisorAdd = 1 'initialize counter for averaging
                    
                    Do While Len(arrColArray(r)) = 0 'Process blank values - continue to loop to find last empty row in current range
                                           
                        lngLastEmptyRow = r 'keep incrementing row number as long as values are empty
                   
                        r = r + 1 'increment empty row indicator
                    
                    Loop
                    
                    dblNextValue = arrColArray(r) 'first value found after empty values
                    
                    dblDivisor = 2 + lngLastEmptyRow - lngFirstEmptyRow 'keep track of number of blank counts
                    
                    dblPadValue = (dblNextValue - dblLastValue) / dblDivisor
                    
                    For r = lngFirstEmptyRow To lngLastEmptyRow 'Ramp empty values from last detected value to first value found after empties
                        
                        arrColArray(r) = arrColArray(r - 1) + dblPadValue 'add the calculated value to the previous value
    
                    Next r
    
                End If
    
                r = r + 1 'increment empty row indicator
                
            Loop 'resume looping after last padded empty values
    
            Range(cpws.Cells(3, c), cpws.Cells(lngLastRow, c)) = Application.WorksheetFunction.Transpose(arrColArray)
            
        End If
        
        Application.StatusBar = "Find and populate empty values: " & Round(100 * ((c - 3) / (lngLastCol - 3))) & "%"
    
    Next c
    
    Set arrColArray = Nothing

End Sub

Public Sub ExportToPDF() 'Copies report output and charts to new workbook
'// -- KP - 09/14/2020 KP // -- //
    
    Dim i As Integer
    Dim strSaveFilePath As String 'filepath to export file
    Dim strSaveFileName As String 'filename of export file
    Dim strFileDate As String 'last date in report data
    Dim sws As Worksheet, cpws As Worksheet
    Dim xws As Worksheet
    Dim boolSingleSelect As Boolean
    Dim intChartCount As Integer
    
    Set sws = ThisWorkbook.Worksheets("Settings")
    Set cpws = ThisWorkbook.Worksheets("Control_Panel")

    Application.StatusBar = "Saving export file . . ."
    Application.DisplayAlerts = False
    
    strFileDate = Format(DateAdd("n", -2, sws.Range("E9")), "yyyy-mm-dd") 'get end date from Settings sheet

    sws.Calculate 'Make sure the filename with date is refreshed before saving
    
    strSaveFilePath = sws.Range("E3") 'Server directory where to save the file

    'Make sure a properly formatted target path is defined
    If Right(strSaveFilePath, 1) <> "\" Then
        strSaveFilePath = strSaveFilePath & "\"
        sws.Range("E3") = strSaveFilePath
    End If
    
    strSaveFileName = sws.Range("E3") & sws.Range("F3") & "_" & strFileDate & ".pdf" 'Root file name with formatted date appended
    
    sws.Range("G3") = strSaveFileName
    
    If Dir(strSaveFilePath, vbDirectory) = "" Then
        MsgBox "Export path doesn't exist!" & Chr(10) & Chr(10) & _
            "Please check your remote use directory.", , "File Write Failure!"
        Exit Sub
    End If
        
    'Use for Charts - Creates array of chart sheets and save to single PDF file
'
'    boolSingleSelect = True
'
'    For Each objChart In Charts
'
'        objChart.Select boolSingleSelect
'        boolSingleSelect = False
'
'    Next
    
    ThisWorkbook.Sheets(Array("Report")).Select

    If strSaveFilePath <> "False" Then
        ActiveSheet.ExportAsFixedFormat xlTypePDF, strSaveFileName, xlQualityStandard
        ActiveSheet.DisplayPageBreaks = False
    End If
     
    'Set objChart = Nothing
    
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Ready"
    
End Sub


Sub CreateOutputReport() 'Build Coal Consumption output report using dynamic arrays
'// -- 9/11/2020 KP -- //

    Dim i As Long, u As Long 'Unit number
    Dim d As Long 'Day
    Dim intInterval As Integer
    Dim ar As Long, ac As Long 'array rows and columns
    Dim lngDayStartRow As Long, lngDayEndRow As Long 'roving range for daily totals per Unit
    Dim lngAverageCount As Long
    Dim dblUnitDailyTotal As Double, dblIntervalSum As Double
    Dim dblUnitMonthlyTotal As Double
    Dim lngDataLastCol As Long
    Dim lngDivisor As Long
    Dim strReportLastDate As String
    Dim strReportRowDate As String
    Dim lngReportLastRow As Long, lngReportLastCol As Long
    Dim lngDayCount As Long, lngTimestampCount As Long
    Dim lngDataStartRow As Long, lngDataLastRow As Long
    Dim lngUnitFirstCol As Long, lngUnitLastCol As Long
    Dim intFirstUnit As Integer, intLastUnit As Integer, intUnitCount As Integer
    Dim strUnitName As String
    Dim arrData As Variant, arrReport As Variant
    Dim sws As Worksheet, cpws As Worksheet, rws As Worksheet
    
    Set sws = ThisWorkbook.Worksheets("Settings")
    Set cpws = ThisWorkbook.Worksheets("Control_Panel")
    Set rws = ThisWorkbook.Worksheets("Report")
    
    If ProcessError = True Then Exit Sub
    
    Application.StatusBar = "Building Report . . ."
    
    Application.ScreenUpdating = False 'set TRUE for testing
    
    rws.Unprotect
    
    'clear previous report print range
    rws.PageSetup.PrintArea = ""
    
    'Data time interval
    intInterval = sws.Range("F6")
    
    lngDivisor = Round(60 / intInterval, 0)
    
    'Find the last row of data
    lngDataLastRow = cpws.Columns("C:D").Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row

    'Find the last report row
    lngReportLastRow = rws.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
    'Find the last report column
    lngReportLastCol = rws.Cells(4, Columns.Count).End(xlToLeft).Column
        
    'protect report base formatting
    If lngReportLastRow < 6 Then lngReportLastRow = 6
    If lngReportLastCol < 6 Then lngReportLastCol = 6
    
    'Clear report of previous data
    Range(rws.Cells(6, 3), rws.Cells(lngReportLastRow, lngReportLastCol)).ClearContents 'clear data
    Range(rws.Cells(4, 5), rws.Cells(5, lngReportLastCol)).ClearContents 'clear totals and unit headers
    Range(rws.Cells(8, 3), rws.Cells(lngReportLastRow + 2, lngReportLastCol + 2)).ClearFormats
    
    'Number of days of data
    If InStr(cpws.Range("C3"), "TimeStamp") = 0 Then
        lngDayCount = DateDiff("d", cpws.Range("C3"), Application.WorksheetFunction.Max(cpws.Range("C3:C" & lngDataLastRow)))
    Else
        MsgBox "Data not present.", 65536, "Query Error"
        Exit Sub
    End If
    
    'Last column of data
    lngDataLastCol = cpws.Cells(1, cpws.Columns.Count).End(xlToLeft).Column
    
    'find first and last unit numbers based on selected tags
    intFirstUnit = Mid(cpws.Range("E1"), InStr(cpws.Range("E1"), ".") - 1, 1)
    intLastUnit = Mid(cpws.Cells(1, lngDataLastCol), InStr(cpws.Cells(1, lngDataLastCol), ".") - 1, 1)
    
    'write first and last unit numbers found in tags to Settings to ensure no discrepancy between tags and units displayed in Report Mananger
    sws.Range("E12") = CStr(intFirstUnit)
    sws.Range("F12") = CStr(intLastUnit)
    
    'derive number of units from first and last unit numbers regardless of the order they were found.
    If Len(sws.Range("F12")) > 0 Then
        intUnitCount = 1 + Application.WorksheetFunction.Max(CInt(sws.Range("E12")), CInt(sws.Range("F12"))) - Application.WorksheetFunction.Min(CInt(sws.Range("E12")), CInt(sws.Range("F12")))
    Else
        intUnitCount = 1
    End If
    
    'Define dimensions of arrays - using "ReDim" to allow use of varibles for dimensions
    ReDim arrData(3 To lngDataLastRow, 1 To lngDataLastCol) 'column offset for 2 columns before timestamp column
    ReDim arrReport(1 To lngDayCount + 1, 1 To intUnitCount + 2) 'dim array for report dimensions plus date columns
    
    'copy data from Canary query to array.  keep range definition the same as worksheet
    arrData = Range(cpws.Cells(1, 3), cpws.Cells(lngDataLastRow, lngDataLastCol)).value
    
    'Populate report array dates in column 1 of report array to find data and then copy results report
    For i = 0 To lngDayCount - 1 'start at zero since days are added
        strReportLastDate = Format(DateAdd("d", i, sws.Range("D9")), "m/d/yyyy")
        arrReport(i + 1, 2) = strReportLastDate 'Date column to show DoW and day
        If Format(arrReport(i + 1, 2), "d") = 1 Then
            arrReport(i + 1, 1) = strReportLastDate 'duplicate date column for month indicator
        End If
        
    Next i

    u = intFirstUnit 'set first unit number
    
    lngUnitFirstCol = 3 'Set Fuel_Flow start column for first unit
    
    For i = 2 To (intUnitCount + 1) 'used as the report target column enumerator - Date column is first column

        strUnitName = Left(arrData(1, lngUnitFirstCol), InStr(arrData(1, lngUnitFirstCol), ".") - 1) 'Set unit name from for column to find matching data
        
        'find first data column of current unit adjusted for array column position - 3
        For ac = lngUnitFirstCol To lngDataLastCol - 3
        
            If (Left(arrData(1, ac), InStr(arrData(1, ac), ".") - 1) = strUnitName And InStr(arrData(1, ac), "Mill_A") > 0) Then
                
                lngUnitFirstCol = ac 'set the unit's first feeder column
                                
            End If
        
        Next ac
        
        lngUnitLastCol = lngUnitFirstCol + 1 'set starting point to look for last fuel flow tag for current unit
        
        'find the last column of data for the current Unit
        For ac = lngUnitLastCol + 1 To lngDataLastCol - 2
        
            If (Left(arrData(1, ac), InStr(arrData(1, ac), ".") - 1) = strUnitName And InStr(arrData(1, ac), "Mill_F") > 0) Then
                
                lngUnitLastCol = ac 'set the unit's last feeder column
                
                Exit For 'exit loop when last feeder is found
            
            End If
        
        Next ac
                        
        lngDayStartRow = 3 'Set the unit data start row
        
        ar = lngDayStartRow
        
        'Write to Report array
        For d = 1 To (lngDayCount) 'Used as the aggregate report date target row enumerator
            
            If d > lngDayCount Then Exit For
                            
            'Get date on the current row
            strReportRowDate = Format(arrReport(d, 2), "m/d/yyyy") 'Get the report target row by date

            'skip shifting start row count at first since row above is tag name
            If lngDayStartRow > 3 Then lngDayStartRow = lngDayEndRow + 1
            
            lngDayEndRow = lngDayStartRow
                                                    
            'Find last row of current day for unit feeders daily aggregates - Limit timestamp search range to 1450 rows each for speed since there should be only 1440 timestamps each day
            While (Format(arrData(lngDayEndRow, 1), "m/d/yyyy") = strReportRowDate And lngDayEndRow <= UBound(arrData, 1))
                
                lngDayEndRow = lngDayEndRow + 1
            
            Wend

            'Count Timestamp rows in data set for the day for averaging by number of minutes in day
            lngTimestampCount = (lngDayEndRow - lngDayStartRow) + 1
            
            lngAverageCount = 0
            
            'Loop through unit array columns and rows for current day
            For ar = lngDayStartRow To lngDayEndRow 'Loop through rows within current day
                
                For ac = lngUnitFirstCol To lngUnitLastCol 'loop through fuel array columns for current unit

                    If arrData(ar, ac) > 0 Then 'only sum and set average count for values > 0
                        
                        dblUnitDailyTotal = dblUnitDailyTotal + arrData(ar, ac) 'sum all array cells in daily aggregate
                        
                        lngAverageCount = lngAverageCount + 1 'increment to average values > 0

                    End If
        
                Next ac

            Next ar
            
            lngDayStartRow = ar 'set row to start row in 'For' loop for next day to speed up report
  
            ' >>>> Average adjusted totals to compensate for any missing timestamps
            dblUnitDailyTotal = dblUnitDailyTotal / lngTimestampCount * 1440 / 60

            'Set the first row for the next day's aggregate loop
            lngDayStartRow = lngDayEndRow + 1
            
            'write the daily aggregates to the array column
            If arrReport(d, 2) > 0 Then arrReport(d, i + 1) = dblUnitDailyTotal 'Copy aggregate value to report array on the current date row - offset 1 for 2 date columns
     
            Application.StatusBar = "Calculating " & Left(strUnitName, Len(strUnitName) - 1) & " " & u & " daily totals:  " & Round(d / (lngDayCount) * 100, 0) & "%"
            
            dblUnitMonthlyTotal = dblUnitMonthlyTotal + dblUnitDailyTotal 'add daily totals to monthly each day's loop
            
            dblUnitDailyTotal = 0
            
            lngAverageCount = 0

        Next d
        
        rws.Cells(4, i + 3) = dblUnitMonthlyTotal 'monthly total top of unit column
        
        rws.Cells(5, i + 3) = Left(strUnitName, Len(strUnitName) - 1) & " #" & u 'unit name in report colum
        
        dblUnitDailyTotal = 0 'clear the sum for the next date or Unit
        
        dblUnitMonthlyTotal = 0
        
        lngUnitFirstCol = lngUnitLastCol + 2
        
        u = u + 1

    Next i
    
    'copy report array data to report worksheet
    Range(rws.Cells(6, "C"), rws.Cells(d + 4, i + 2)) = arrReport
    
    'Find the updated report last row
    lngReportLastRow = rws.Columns("C:D").Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row

    'Find the updated report last column
    lngReportLastCol = rws.Cells(4, Columns.Count).End(xlToLeft).Column
    
    'protect report base formatting
    If lngReportLastRow < 7 Then lngReportLastRow = 7
    If lngReportLastCol < 6 Then lngReportLastCol = 6
     
    rws.Activate
     
    'Copy formats down the report page
    Range(rws.Cells(7, "C"), rws.Cells(7, lngReportLastCol)).Copy
        Range(rws.Cells(7, "C"), rws.Cells(lngReportLastRow, lngReportLastCol)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        rws.Range("A1").Activate 'de-select pasted format range
   
    On Error GoTo 0

    Set arrData = Nothing
    Set arrReport = Nothing
      
    Application.ScreenUpdating = True

    rws.PageSetup.PrintArea = Range(rws.Cells(3, 3), rws.Cells(lngReportLastRow, lngReportLastCol)).Address
    
    rws.DisplayPageBreaks = False

    rws.Protect
    
    cpws.Activate
    
    Application.StatusBar = "Ready"

End Sub

Sub CreateOutputReport_Old() 'Build Coal Consumption output report - reads and writes directly to and from report worksheet
'// -- 9/11/2020 KP -- //

    Dim f As Long 'Facility number
    Dim d As Long 'Day
    Dim intInterval As Integer
    Dim i As Long, r As Long, s As Long, x As Long, ar As Long
    Dim lngAggStartRow As Long, lngAggEndRow As Long 'roving range for daily totals per facility
    Dim rngAggRange As Range
    Dim dblTotal As Double, dblIntervalSum As Double
    Dim lngDivisor As Long
    Dim strReportRowDate As String, strDataRowDate As String
    Dim lngDayCount As Long, lngAggRowCount As Long
    Dim lngDataStartRow As Long, lngDataLastRow As Long
    Dim lngDataFirstCol As Long, lngDataLastCol As Long
    Dim lngReportLastRow As Long
    Dim intFacilityCount As Integer
    Dim intFirstFacility As Integer, intLastFacility As Integer
    Dim strFacilityName As String
    Dim rngFacilityFlow As Range
    Dim rngTargetCol As Range, rngTargetCell As Range
    Dim strAggFormula As String
    Dim sws As Worksheet, cpws As Worksheet, rws As Worksheet
    
    Set sws = ThisWorkbook.Worksheets("Settings")
    Set cpws = ThisWorkbook.Worksheets("Control_Panel")
    Set rws = ThisWorkbook.Worksheets("Report")
    
    If ProcessError = True Then Exit Sub
    
    Application.StatusBar = "Building Report . . ."
    
    Application.Calculation = xlCalculationManual
    
    Application.ScreenUpdating = False ' TRUE for testing
    
    rws.Unprotect
    
    'Clear report of previous data
    rws.Range("E6:M500").ClearContents
    
    'Data time interval
    intInterval = sws.Range("F6")
    
    lngDivisor = 60 / intInterval
    
    'Find the last row of data
    lngDataLastRow = cpws.Columns("C:D").Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row
    
    'Number of days of data
    If InStr(cpws.Range("C3"), "TimeStamp") = 0 Then
        lngDayCount = DateDiff("d", cpws.Range("C3"), Application.WorksheetFunction.Max(cpws.Range("C3:C" & lngDataLastRow)))
    Else
        MsgBox "Data not present.", 65536, "Query Error"
        Exit Sub
    End If
    
    'Last column of data
    lngDataLastCol = cpws.Cells(1, cpws.Columns.Count).End(xlToLeft).Column
    
    'number of facilities
    intFacilityCount = Application.WorksheetFunction.CountIf(rws.Range("E5:M5"), "??*")
    
    'Loop to find first and last facilities with data in this report
    intFirstFacility = Mid(cpws.Range("E1"), InStr(cpws.Range("E1"), ".") - 1, 1)
    intLastFacility = Mid(cpws.Cells(1, lngDataLastCol), InStr(cpws.Cells(1, lngDataLastCol), ".") - 1, 1)
    
    On Error Resume Next
        
    For i = 5 To (intFacilityCount + 4) 'used as the target column enumerator starting column 'E'
    
        For f = intFirstFacility To intLastFacility 'accomodate facility skipped numbers
            
            If Right(rws.Cells(5, i), 1) = f Then 'If the facility enumerator matches the target column

                strFacilityName = "Cross" & f 'Set facility name for column to find matching data
                
                'get the first column of data for the current facility
                lngDataFirstCol = 1 + cpws.Cells(1, 1).EntireRow.Find(What:=strFacilityName, LookIn:=xlValues, LookAt:=xlPart, _
                    searchorder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
                
                'Get the last column of date for the current facility
                lngDataLastCol = cpws.Cells(1, 1).EntireRow.Find(What:=strFacilityName, LookIn:=xlValues, LookAt:=xlPart, _
                    searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
                
                r = 3 'Set the data start row - lngAggStartRow adds 1, so this needs to start one row above data
                lngAggStartRow = 3
                
                'Write to Report
                For d = 6 To (lngDayCount + 5) 'Used as the aggregate date target row enumerator on report starting row 6
                                    
                    'Get date on the current row
                    strReportRowDate = Format(rws.Cells(d, 4), "m/d/yyyy") 'Get the date from the current report row

                    'Check if query data date matches current report date row -1 second to capture the next 12:00 AM timestamp
                    strDataRowDate = Format(DateAdd("s", -1, cpws.Cells(r, 3)), "m/d/yyyy")

                    'Begin row of unit feeders daily aggregate
                    If lngAggStartRow > 3 Then lngAggStartRow = lngAggEndRow + 1
                                                            
                    'Find last row of current day for unit feeders daily aggregates - Limit search range to 1450 rows each for speed since there should be only 1440 timestamps each day
                    lngAggEndRow = cpws.Range("C" & r & ":C" & r + 1450).Find(What:=DateValue(Format(DateAdd("n", intInterval, rws.Cells(d, 4)) + 1, "m/d/yyyy h:mm AM/PM")), _
                        LookIn:=xlFormulas, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Row
                    
                    'Report runs 40% faster by not starting at the top of data each time to find the current aggregate end row
                    r = lngAggEndRow
                    
                    'Count Timestamp rows in current day's aggregate
                    lngAggRowCount = (lngAggEndRow - lngAggStartRow) + 1

                    'Set range for daily aggregate
                    dblTotal = Application.WorksheetFunction.Sum(Range(cpws.Cells(lngAggStartRow, lngDataFirstCol), cpws.Cells(lngAggEndRow, lngDataLastCol)))
                    
                    ' >>>> Average adjusted totals to compensate for missing data rows
 '                   dblTotal = dblTotal / lngAggRowCount * 1440 '>>> Comment out this line to prevent averaging of missing data
                    
                    dblTotal = Application.WorksheetFunction.AverageIf(Range(cpws.Cells(lngAggStartRow, lngDataFirstCol), cpws.Cells(lngAggEndRow, lngDataLastCol)), "<>" & "") * 1440 * (lngDataLastCol - lngDataFirstCol + 1)
                    
                    'Set the first row for the next day's aggregate loop
                    lngAggStartRow = lngAggEndRow + 1
                    
                    'write the daily aggregates to the report columns
                    If rws.Range("D" & d) <> "" Then rws.Cells(d, i) = dblTotal / lngDivisor 'Set the sum in the current facilty column for the current date row

                    Application.StatusBar = "Calculating " & Left(strFacilityName, Len(strFacilityName) - 1) & " " & Right(strFacilityName, 1) & " daily totals:  " & Round(d / (lngDayCount + 5) * 100, 0) & "%"
                    
                    dblTotal = 0 'reset daily totals for next day

                Next d

                dblTotal = 0 'clear the sum for the next date or facility
 
            End If
            
            r = 3

        Next f
        
    Next i
    
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
      
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Ready"
    
    rws.PageSetup.PrintArea = Range(rws.Cells(3, 3), rws.Cells(d - 1, i - 1)).Address
    
    rws.DisplayPageBreaks = False
    
End Sub

Option Explicit

Sub ProcessCharts() 'Set chart data ranges and formats
'// -- KP 10/12/2020 -- //
   
    Dim strChartName As String
    Dim i As Long
    Dim lng1WeekLastRow, lng4WeekLastRow, lngWeeksLastCol As Long 'last rows and column on weeks data sheets
    Dim lngMaxMwOutput As Long 'max generating unit MW output range
    Dim lngMajorUnit As Long, lngMinorUnit As Long
    Dim dblStdIntercept As Double, dblStdAval As Double, dblStdBval As Double
    Dim lngStandardsRow As Long 'standards trendline equation
    Dim str4WeekChartLegend As String, str1WeekChartLegend As String
    Dim intLegendCount As Integer, strLegendText As String, lgdLegendEntry As LegendEntry 'chart legend entries
    Dim strStandardsSeriesName As String 'standards series name
    Dim lngStandardsCol 'standards for each chart on Standards sheet
    Dim lngStandardsLastRow As Long 'last row of data for standards on Standards sheet
    Dim strTagName As String, lngTagCol As Long  'dynamic tag column for looping through charts
    Dim lngMwGrossCol As Long, strSCR_Int_A As String, strSCR_Int_B As String, strAH_DP_A As String, strAH_DP_B As String
    Dim strSH_Abs As String, strRH_Abs As String, strWW_Econ_Abs As String
    Dim chtChart As Chart, intChartIndex As Integer 'Chart and chart index
    Dim cpws As Worksheet, stws As Worksheet, ws As Worksheet
    Dim wks1 As Worksheet, wks4 As Worksheet

    Set cpws = ThisWorkbook.Worksheets("Control_Panel")
    Set stws = ThisWorkbook.Worksheets("Standards")
    
    Application.ScreenUpdating = False
    
    Application.Calculation = xlCalculationAutomatic 'make sure all sheets have been calculated before generating charts
    Application.Calculation = xlCalculationManual
    
    'get last row with standards values (all charts use same number of rows)
    lngStandardsLastRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(stws.Range("B1:B100")), stws.Range("B1:B100"), 0)

    For Each ws In Worksheets 'set One and Four week worksheets
        If ws.Name Like "*" & "One" & "*" Then _
            Set wks1 = Worksheets(ws.Name) 'dynamically set 1-week worksheet
        If ws.Name Like "*" & "Four" & "*" Then _
            Set wks4 = Worksheets(ws.Name) 'dynamically set 4-week worksheet
    Next
    
    'Set Megawatts output column on weeks sheets since both sheets are the same
    lngMwGrossCol = wks1.Cells(1, 1).EntireRow.Find(What:="*Gross*", LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column

    'get last rows on weeks data sheets
    lng1WeekLastRow = wks1.Columns("C:C").Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row
    lng4WeekLastRow = wks4.Columns("C:C").Find("*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row
    
    'get standards equation row
    lngStandardsRow = 1 + stws.Columns("A:B").Find("*Month*", LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlPrevious, searchorder:=xlByRows).Row

    'get last column of data used by both 1 and 4 weeks sheets
    lngWeeksLastCol = wks1.Range("A1:ZZ1").Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    
    For Each chtChart In Charts
    
        With chtChart
    
            If .Name <> "" Then 'only process chart if it exists
            
                intChartIndex = intChartIndex + 1 'seriescollection index in looping all charts
                
                intLegendCount = .Legend.LegendEntries.Count 'Count legend entries.  There should only be 3

                lngMaxMwOutput = Application.WorksheetFunction.Max(stws.Range("B5:B40")) 'Unit maximum MW output range

                If lngMaxMwOutput > 400 Then 'dynamically set units for major and minor series lines
                    lngMajorUnit = 50
                    'lngMinorUnit = 5
                Else
                    lngMaxMwOutput = 350 'dynamically set units for major and minor series lines
                    lngMajorUnit = 50
                    'lngMinorUnit = 4
                End If

                'get the standards data column for current chart
                lngStandardsCol = stws.Range("A1:ZZ10").Find(What:="*" & .Name & "*", LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column

                'match tag with chart - adjust tag names to accommodate tag name variations between units or facilities
                Select Case .Name   'Case is Chart name  - use either partial or full unique tag names to associate with chart names
                    Case "SCR_Int_A"                                    'Note: When multiple SH tags are agregated use the tag alias "SH_123_Abs"
                        strTagName = "SCR_A_In_T" 'SCR A Inlet Temp
                    Case "SCR_Int_B"
                        strTagName = "SCR_B_In_T" 'SCR B Inlet Temp
                    Case "AH_DP_A"
                        strTagName = "AH_A_Gas_DP" 'Air Heater Gas A DP
                    Case "AH_DP_B"
                        strTagName = "AH_B_Gas_DP" 'Air Heater Gas B DP
                    Case "SH1_Abs"
                        strTagName = ".SH1\.A" 'SH Absorp Horizontal and Pendants (Before Attemporator)
                    Case "SH2_Abs"
                        strTagName = ".SH2\.A" 'SH Absorp Division Panels and Platens (After Attemporator)
                    Case "RH_Abs"
                        strTagName = ".RHX\.A" 'RH Total Absorption
                    Case "WW_Econ_Abs"
                        strTagName = ".FHX\.A" 'C3 Total economizer - water wall heat absorption
                    Case Else
                        MsgBox "There was an errror processing charts due to tag mismatch." & Chr(10) & Chr(10) & "Verify tag and chart matches in 'ProcessCharts()'", 65536, "Chart/Tag Mismatch"
                        ProcessError = True
                        Exit Sub
                End Select

                'get unit tag name and column common to both 4 and 1 week sheets
                strTagName = wks1.Cells(1, 1).EntireRow.Find(What:="*" & strTagName & "*", LookIn:=xlFormulas, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
                lngTagCol = wks1.Cells(1, 1).EntireRow.Find(What:=strTagName, LookIn:=xlFormulas, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
                str1WeekChartLegend = wks1.Range("A3") & " - " & Replace(wks1.Range("A2"), "-", "to") 'Set 1-week chart legend date range
                str4WeekChartLegend = wks4.Range("A3") & " - " & Replace(wks4.Range("A2"), "-", "to") 'Set 4-week chart legend date range

                DoEvents

                'Get standards first column and values for each standards vertical axis tag for each chart
                lngStandardsCol = stws.Range("A1:ZZ10").Find(What:="*" & .Name & "*", LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
                dblStdIntercept = stws.Cells(lngStandardsRow, lngStandardsCol + 2)
                dblStdAval = stws.Cells(lngStandardsRow, lngStandardsCol)
                dblStdBval = stws.Cells(lngStandardsRow, lngStandardsCol + 1)
                
                DoEvents
                
                'set chart plot area dimensions and position (unable to use 'with' without selecting plot area, which slows down trendline processing)
                .Activate
                .PlotArea.Left = 28
                .PlotArea.Top = 58
                .PlotArea.Height = 390
                .PlotArea.Width = 620

                DoEvents

                'set legend dimensions and font size
                With .Legend
                    .Left = 44.978
                    .Top = 21
                    .Width = 192.75
                    .Height = 38
                    .Font.Size = 8
                End With
                
                Application.StatusBar = "Updating Chart Trendlines :  " & intChartIndex & " of " & Charts.Count & "  ||  " & .Name
                '======= set scatter chart series collection data ranges and formats =============
                '---- Series Collection 1 - Four Weeks ---------------------------
                If .SeriesCollection.Count < 1 Then .SeriesCollection.NewSeries
                With .SeriesCollection(1)
                    chtChart.Legend.LegendEntries(1).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
                    .AxisGroup = xlPrimary
                    .Name = str4WeekChartLegend
                                                             
                    With .Format 'gets rid of plot line for scatter data
                        .Fill.Visible = msoTrue
                        .Fill.ForeColor.RGB = RGB(255, 255, 0)
                        .Line.Visible = msoFalse
                        .Line.Transparency = 1
                    End With
                    
                    .Border.LineStyle = -4142 'no marker border
                    .MarkerStyle = 2
                    .MarkerSize = 3
                    
                    .XValues = "=" & wks4.Name & "!" & Range(wks4.Cells(2, lngMwGrossCol), wks4.Cells(lng1WeekLastRow, lngMwGrossCol)).Address & ""
                    .values = "=" & wks4.Name & "!" & Range(wks4.Cells(2, lngTagCol), wks4.Cells(lng1WeekLastRow, lngTagCol)).Address & ""
                                        
                    If .Trendlines.Count > 1 Then .Trendlines(2).Delete
                    If .Trendlines.Count < 1 Then .Trendlines.Add
                    
                    With .Trendlines(1)
                         .Format.Line.Visible = msoTrue
                         .Type = xlPolynomial
                         .Order = 2
                         .Intercept = stws.Cells(lngStandardsRow, lngStandardsCol + 2)
                         .Format.Line.ForeColor.RGB = RGB(255, 255, 0)
                         .Format.Line.Weight = 0.75
                         .Backward = 0
                         .Forward = 30
                         .Name = ""
                     End With
                    
                End With
                
                DoEvents
                
                '---- Series Collection 2 --- One week ------------------------
                If .SeriesCollection.Count < 2 Then .SeriesCollection.NewSeries
                With .SeriesCollection(2)
                    chtChart.Legend.LegendEntries(2).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(30, 210, 234)
                                                             
                    With .Format 'gets rid of plot line for scatter data
                        .Fill.Visible = msoTrue
                        .Fill.ForeColor.RGB = RGB(255, 255, 0)
                        .Line.Visible = msoFalse
                        .Line.Transparency = 1
                    End With
                    
                    .AxisGroup = xlPrimary
                    .Name = str1WeekChartLegend
                    
                    .Border.LineStyle = -4142 'no marker border
                    .MarkerStyle = 2
                    .MarkerSize = 3
                    
                    .XValues = "=" & wks1.Name & "!" & Range(wks1.Cells(2, lngMwGrossCol), wks1.Cells(lng4WeekLastRow, lngMwGrossCol)).Address & ""
                    .values = "=" & wks1.Name & "!" & Range(wks1.Cells(2, lngTagCol), wks1.Cells(lng4WeekLastRow, lngTagCol)).Address & ""

                    With .Format 'gets rid of plot line on scatter chart
                        .Fill.Visible = msoTrue
                        .Fill.ForeColor.RGB = RGB(30, 210, 234)
                        .Line.Visible = msoFalse
                        .Line.Transparency = 1
                    End With

                    If .Trendlines.Count > 1 Then .Trendlines(2).Delete
                    If .Trendlines.Count < 1 Then .Trendlines.Add
                    
                    With .Trendlines(1)
                         .Format.Line.Visible = msoTrue
                         .Type = xlPolynomial
                         .Order = 2
                         .Intercept = stws.Cells(lngStandardsRow, lngStandardsCol + 2)
                         .Format.Line.ForeColor.RGB = RGB(30, 210, 234)
                         .Format.Line.Weight = 0.75
                         .Backward = 200
                         .Forward = 60
                         .Name = ""
                     End With
                    
                End With
                
                '---- Series Collection 3 ----------------------------
                If .SeriesCollection.Count < 3 Then .SeriesCollection.NewSeries
                With .SeriesCollection(3) 'standard curve values
                    .AxisGroup = xlSecondary 'Set standards curve to secondary axis so it is on top/in front of scatter data
                    chtChart.Legend.LegendEntries(3).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(204, 153, 255) 'legend text color
                    .XValues = "=" & stws.Name & "!" & Range(stws.Cells(4, lngStandardsCol), stws.Cells(lngStandardsLastRow, lngStandardsCol)).Address & ""
                    .values = "=" & stws.Name & "!" & Range(stws.Cells(4, lngStandardsCol + 1), stws.Cells(lngStandardsLastRow, lngStandardsCol + 1)).Address & ""
                    .MarkerStyle = -4142 'hide markers to use line for standards
                    .Format.Fill.Visible = msoFalse 'hide markers to use line for standards
                    .Format.Line.Visible = msoTrue
                    .Format.Line.ForeColor.RGB = RGB(235, 112, 244)
                    .Format.Line.Weight = 1.5
                    .PlotOrder = 1 'bring standandards curve line to front on chart
                    If .Trendlines.Count > 0 Then .Trendlines(1).Delete
                    
                    'set standards series name/date range from Standards worksheet
                    strStandardsSeriesName = "Standards: " & Format(stws.Cells(lngStandardsRow, lngStandardsCol - 1), "m/d/yyyy") & _
                        " - " & Format(stws.Cells(lngStandardsRow + 1, lngStandardsCol - 1), "m/d/yyyy")

                    If .Name = "" Then 'if standards series doesn't have a name, then add it
                        .Name.Add = strStandardsSeriesName
                    Else 'otherwise, just update the standards series name
                        .Name = strStandardsSeriesName
                    End If

                End With
                '====== End series collection processing ===================

                DoEvents
                  
                'delete trendline legend entries if they were created
                If .Legend.LegendEntries.Count > 3 Then
                    For i = .Legend.LegendEntries.Count To 4 Step -1
                            .Legend.LegendEntries(i).Delete
                    Next i
                End If
                
                'set secondary axis for trendline equal to primary axis and hide secondary axis scale since it matches primary
                .Axes(xlValue, xlSecondary).MinimumScale = .Axes(xlValue, xlPrimary).MinimumScale
                .Axes(xlValue, xlSecondary).MaximumScale = .Axes(xlValue, xlPrimary).MaximumScale
                .Axes(xlValue, xlSecondary).MajorUnit = .Axes(xlValue, xlPrimary).MajorUnit
                .Axes(xlValue, xlSecondary).MinorUnit = .Axes(xlValue, xlPrimary).MinorUnit
                .Axes(xlValue, xlSecondary).TickLabelPosition = xlNextToAxis
                .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "General"
                .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 8
                .ChartArea.Select 'select the entire chart so nothing will appear to be selected after processing

            End If
        
        DoEvents

        .Deselect 'de-select all items on chart
        .Refresh 'refresh chart
            
        End With

    Next
    
    Set chtChart = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    DoEvents
    
    Application.StatusBar = "Ready"
            
End Sub

Sub ClearChartsDataRange() 'clears all chart data ranges leaving blank charts
'// -- 10/07/2020 -- KP //

    Dim i As Long, x As Long
    Dim chtChart As Chart

    For Each chtChart In Charts
        With chtChart
            For i = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection.Count > 0 Then
                    .SeriesCollection(i).Delete
                End If
            Next i
        End With
    Next

End Sub



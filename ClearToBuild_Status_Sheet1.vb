Option Explicit
Option Base 1

'-- Used for Mouse Hover
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
'-----------------------
Private lngLastItmIndex As Long
Private blnSortForward As Boolean 'sort forward or backwards

Private Sub LstVw_BOM_Click()

    SetFormPositions

End Sub

Private Sub LstVw_OpenWorkOrdersFG_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
'// -- 05/31/2022 - KP -- // 'pop up the BOM components list when clicking on a FG item

    Dim strItem As String
    Dim itm As MSComctlLib.ListItem
    
    On Error Resume Next
    
    With LstVw_OpenWorkOrdersFG
        .SelectedItem.Selected = False ' unselect a previous selected subitem
        On Error GoTo 0
        ConvertPixelsToTwips x, y  'make the necessary units conversion
        Set itm = .HitTest(x, y) 'set the object using the converted coordinates
    End With
    
    'if the db connection is already open, don't queue a series of queries that won't be seen
    If cnn.State = 1 Then Exit Sub

    DoEvents
    
    'Pop up the BOM list when a valid FG is selected
    If Not itm Is Nothing Then
        strItemNo = itm
        BOM_LvPopUp (strItemNo)
    End If
    
End Sub

Private Sub lstvw_BOM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) 'Click Sort Exploded BOM Column
'// -- 04/21/2022 - KP -- //

    'Sort by column clicked on
    With LstVw_BOM
        SortExplodedBOM (.ColumnHeaders.Item(ColumnHeader.Index).Index - 1)
    End With

End Sub

Sub SortExplodedBOM(col As Long) 'Sorts any column with text or numbers
'// -- 04/27/2022 - KP -- //     'Date column is very specific to its position since
                                 'there is special code to handle dates
    Dim sTemp As String * 15
    Dim nTemp As String
    Dim strColVal As String
    Dim lvCount As Long
    Dim i As Long
    Dim lngASAPdateCol As Long, strASAPfieldDate As String
    
    strASAPfieldDate = "Date" 'Sort any Date column containing this string
    
    With LstVw_BOM
        
        .sortKey = col
        lvCount = .ListItems.Count
        
        'Format date as "yyyymmdd" for proper sorting
        If Trim(.ColumnHeaders(.sortKey + 1)) Like "*" & strASAPfieldDate & "*" Then
            lngASAPdateCol = .sortKey 'set the column to be ignored as a number for number sorting
            'if the column is a date, then convert it to a reverse-format integer first
            For i = 1 To lvCount
                 .ListItems(i).SubItems(.sortKey) = Format(.ListItems(i).SubItems(.sortKey), "yyyymmdd")
            Next i
        End If
                    
        'get current sort then invert sort order
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        
        'if the column was a date, then reverse-parse and convert it back to a formatted date
        If Trim(.ColumnHeaders(.sortKey + 1)) Like "*" & strASAPfieldDate & "*" Then
            For i = 1 To lvCount 'Reverse date text string (yyyymmdd) and format back to date (m/d/yyyy)
                .ListItems(i).SubItems(.sortKey) = Format(Left(.ListItems(i).SubItems(.sortKey), 4) & "/" & _
                    Mid(.ListItems(i).SubItems(.sortKey), 5, 2) & "/" & Right(.ListItems(i).SubItems(.sortKey), 2), "m/d/yyyy")
            Next i
            Exit Sub
        End If
        
        'Set numerical columns to sort.  DO NOT include date columns
        If (col > 3 And col <> lngASAPdateCol) Then 'sort as number anything beyond the description column as they are all numerical values and sort as a number

            lvCount = .ListItems.Count 'get the number of rows in the ListView
            
            For i = 1 To lvCount
            
                sTemp = vbNullString
                
                If .sortKey Then
                    ''RSet - right align a string within a string variable.
                    strColVal = CLng(.ListItems(i).SubItems(.sortKey)) + 10000000000000# 'Add a number impossible to reach with real numbers 1 with 14 zeros
                    RSet sTemp = strColVal   '.ListItems(i).SubItems(.sortKey)           'needed for sorting negative numbers since sort only works with strings
                    .ListItems(i).SubItems(.sortKey) = sTemp
                Else
                
                    RSet sTemp = .ListItems(i)
                    .ListItems(i).Text = sTemp
                    
                End If
                
            Next
            
            .Sorted = True 'sort column as text with the added 100000000000000
            
            For i = 1 To lvCount
                If .sortKey Then 'write values back to sorted column and subract 100000000000000, then format as #,##0
                    strColVal = LTrim$(Format(CLng(.ListItems(i).SubItems(.sortKey) - 10000000000000#), "#,##0")) 'Subract the 100000000000000 to restore original values, including negative
                   .ListItems(i).SubItems(.sortKey) = strColVal
                   ' .ListItems(i).SubItems(.sortKey) = LTrim$(.ListItems(i).SubItems(.sortKey))
                Else
                    .ListItems(i).Text = LTrim$(.ListItems(i))
                End If
                
            Next
            
        Else 'Column is text so sort here
            .Sorted = True
        End If

    End With
    
End Sub
Private Sub LstVw_OpenWorkOrdersFG_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'// -- 04/21/2022 - KP -- //

    'Sort by column clicked on
    With LstVw_OpenWorkOrdersFG
        SortOpenWorkOrdersFG (.ColumnHeaders.Item(ColumnHeader.Index).Index - 1)
    End With

End Sub

Sub SortOpenWorkOrdersFG(col As Long) 'Sorts any column with text or numbers
'// -- 04/21/2022 - KP -- //

    Dim sTemp As String * 15
    Dim nTemp As String
    Dim strColVal As String
    Dim lvCount As Long
    Dim i As Long
    
    With LstVw_OpenWorkOrdersFG
        
        .sortKey = col
        
        'get and invert sort order direction
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If

        If col > 2 Then 'sort as number anything beyond the description column as they are all numerical values and sort as a number

            lvCount = .ListItems.Count 'get the number of rows in the ListView
            
            For i = 1 To lvCount
                sTemp = vbNullString
                
                If .sortKey Then
                    ''RSet - right align a string within a string variable.
                    strColVal = CLng(.ListItems(i).SubItems(.sortKey)) + 10000000000000# 'Add a number impossible to reach with real numbers 1 with 14 zeros
                    RSet sTemp = strColVal   '.ListItems(i).SubItems(.sortKey)           'needed for sorting negative numbers since sort only works with strings
                    .ListItems(i).SubItems(.sortKey) = sTemp
                Else
                
                    RSet sTemp = .ListItems(i)
                    .ListItems(i).Text = sTemp
                    
                End If
                
            Next
            
            .Sorted = True 'sort column as text with the added 100000000000000
            
            For i = 1 To lvCount
                If .sortKey Then 'write values back to sorted column and subract 100000000000000, then format as #,##0
                    strColVal = LTrim$(Format(CLng(.ListItems(i).SubItems(.sortKey) - 10000000000000#), "#,##0")) 'Subract the 100000000000000 to restore original values, including negative
                   .ListItems(i).SubItems(.sortKey) = strColVal
                   ' .ListItems(i).SubItems(.sortKey) = LTrim$(.ListItems(i).SubItems(.sortKey))
                Else
                    .ListItems(i).Text = LTrim$(.ListItems(i))
                End If
            Next
            
        Else 'Column is text so sort here
            .Sorted = True
        End If

    End With
    
End Sub

Public Sub PopulateFGData() 'populates list view on Materials_Status sheet
'// -- 08/01/2022 - KP -- //

    Dim i, r, c As Long
    Dim lngDbRow As Long
    Dim lngViewWidth, lngHeight As Long
    Dim strDbName, strSourcePath As String 'workbook source
    Dim strColHeaders, strFieldName As String
    Dim li As ListItem
    
    Dim cws As Worksheet

    Set cws = ThisWorkbook.Worksheets("Controls")

    blnProcessError = False
    
    'close the components pop-up if visible
    If Me.LstVw_BOM.Visible = True Then Me.LstVw_BOM.Visible = False
    LstVw_OpenWorkOrdersFG.Visible = False
    
    Application.StatusBar = "Importing FG Items data . . "
    
    On Error Resume Next
    
    'get the data filename from the Controls sheet file import log
    lngDbRow = cws.Range("E5:E10").Find("*Data-ClearToBuild*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False).Row
    
    If lngDbRow = 0 Then
        MsgBox "Data must be imported and processed before running this application.", 65536, "Processing Error"
        
        'verify all source files are present and accessible.
        Call InitializeDataArrays
        
        lngDbRow = cws.Range("E5:E10").Find("*Data-ClearToBuild.xlsx", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False).Row
        
        'if the primary still doesn't exist, then exit sub
        If lngDbRow = 0 Then
            MsgBox "Unable to proceed without source data", 65536, "Source Data Needed"
            Call Main.ResetApplication
            Exit Sub
        End If
    End If
    
    On Error GoTo 0 'ErrCatch

    If lngDbRow > 4 Then
        strDbName = GetLocalPathFromSP(cws.Cells(lngDbRow, "E"))
        If (Dir(strDbName) <> "" And InStrRev(strDbName, "\")) Then
            strSourcePath = Left(strDbName, InStrRev(strDbName, "\") - 1)
        Else
            strSourcePath = GetLocalPathFromSP(ThisWorkbook.Path)
            strDbName = strSourcePath & "\" & Mid(strDbName, InStrRev(strDbName, "\") + 1)
        End If
    Else
        strSourcePath = GetLocalPathFromSP(ThisWorkbook.Path)
        strDbName = SelectImportFile(strSourcePath, "Please select the file containing All Open W/O's,Open P/O's, SCR, and Exploded Bom data", "*.xlsx")
        strSourcePath = Left(strDbName, InStrRev(strDbName, "\") - 1)
        strDbName = strSourcePath & "\" & Mid(strDbName, InStrRev(strDbName, "\") + 1)
    End If
    
    If Dir(strDbName) = "" Then
        MsgBox "The database file: " & Chr(10) & Chr(10) & strDbName & Chr(10) & Chr(10) & _
            "could not be found.", 65536, "Database Not Found!"
        blnProcessError = True
        Exit Sub
    End If
    
    DoEvents
    
    Application.StatusBar = "Querying source data workbook . . ."
    
    dbOpen (strDbName)

    strSQL = " SELECT DISTINCT(w.[Item Number]) as [FG Item#], w.[Item Description], " & _
            "   SUM(s.[Global Available On Hand Qty]) as [FG On-Hand], " & _
            "   SUM(w.[Order Qty]) as [W/O Qty], COUNT(w.[Order Number]) as [W/O Ct],  " & _
            "   SUM(s.[Total Forecast Sets 1]) as [Fcst 1 Mo Qty], SUM(s.[Total Forecast Sets 2]) as [Fcst 2 Mo Qty],  " & _
            "   SUM(s.[Total Forecast Sets 3]) as [Fcst 3 Mo Qty],  " & _
            "   (ROUND(SUM(s.[Global Available On Hand Qty]),0) + ROUND(SUM(w.[Order Qty]),0)) -  " & _
            "   (ROUND(SUM(s.[Total Forecast Sets 1]),0) + ROUND(SUM(s.[Total Forecast Sets 2]),0) +  " & _
            "   SUM(s.[Total Forecast Sets 3])) as [Excess Order Qty ] " & _
            " FROM ([OpenWorkOrders$] w " & _
            " INNER JOIN ( " & _
            "       SELECT DISTINCT([Item Number]) as [Item Number], [Global Available On Hand Qty], [Available OH In Quar], " & _
            "            [Total Forecast Sets 1], [Total Forecast Sets 2], [Total Forecast Sets 3], [Available On Hand Qty] " & _
            "       FROM [SupplyChainReport$] " & _
            "       WHERE [Branch Code] = 'IVP'" & _
            "       GROUP BY [Item Number],[Total Forecast Sets 1], [Total Forecast Sets 2], [Total Forecast Sets 3], " & _
            "           [Available On Hand Qty], [Global Available On Hand Qty], [Available OH In Quar] " & _
            "       ) s on s.[Item Number] = w.[Item Number]) " & _
            " WHERE [Order Status] = 'Pending' AND w.[Item Type] LIKE 'FINISHED G%' AND w.[Item Number] <> 'REWORK-PRODUCT' " & _
            " GROUP BY w.[Item Number], w.[Item Description]"

    rs.Open strSQL, cnn, adUseClient 'open connection and execute query

    arrFGItems = Application.WorksheetFunction.Transpose(rs.GetRows) 'load SQL recordset results into array
    
    ReDim arrTemp(0 To UBound(arrFGItems, 2))
'
'    'Position Multi-Page select relative position from tab strip then offset to keep listview window inside tab strip area
'    x = Me.MultipleTabSelect.Left + 10
'    y = Me.MultipleTabSelect.Top + 50
'    lngHeight = Me.MultipleTabSelect.height - 60
'
'    DoEvents
    
'    Me.MultipleTabSelect.Width = Me.LstVw_OpenWorkOrdersFG.Width + (x / 3) 'keep centered within tabs area
    
    'Populate FG Listbox
    With LstVw_OpenWorkOrdersFG
        .Font.Size = 10
        .HideColumnHeaders = False
        
'        .Top = y 'set vertical position on sheet based on Multi-Page Select
'        .Left = x 'set horizontal position on sheet based on Multi-Page Select
'        .height = lngHeight
        .View = lvwReport
        .Gridlines = True
        .ListItems.Clear
        
        'make invisible if in the wrong place on the sheet by default
        .Visible = False
        
        With .ColumnHeaders
        
            .Clear
                
            'add and hide a dummy column to allow right or left justification (Column 1 does not allow justification-bug with Microsoft ListView 6.0)
            .Add , , "", 0
                    
            For i = 0 To rs.Fields.Count - 1
                
                strFieldName = rs.Fields(i).Name
                                
                Select Case True 'set column widths
                
                    Case strFieldName Like "*FG Item*" 'FG Item#W
                        .Add , , strFieldName, 100, 2
                        
                    Case strFieldName Like "*Descrip*" 'Description
                        .Add , , strFieldName, 260, lvwColumnLeft
                    
                    Case LCase(strFieldName) Like "*ct*" 'W/O Count
                        .Add , , strFieldName, 70, 2
                    
                    Case LCase(strFieldName) Like "w/o qty" 'W/O qty
                        .Add , , strFieldName, 70, 1
                    
                    Case LCase(strFieldName) Like "*excess*" 'Excess Order Qty
                        .Add , , strFieldName, 110, 2
    
                    Case Else 'All other columns
                        .Add , , strFieldName, 90, 1
                
                End Select
                
                'set total item list view width by the sum of all column widths together
                If i < rs.Fields.Count - 1 Then
                    lngViewWidth = lngViewWidth + .Item(i + 1).Width
                Else
                    lngViewWidth = Round(lngViewWidth * 0.97, 1) 'view window offset width
'                    lngViewWidth = lngViewWidth - 40 'view window offset width
                End If
            Next i
        
        End With
        
        'list view total width added from field name column widths
        .Width = lngViewWidth
        
        'Populate data fields
        For r = 1 To rs.RecordCount

            c = 1 'must be instantiated for first item
        
            Set li = .ListItems.Add(, , arrFGItems(r, c))

            For c = 1 To rs.Fields.Count

                If c < 3 Then 'set field formats
                    li.ListSubItems.Add , , arrFGItems(r, c) 'text for FG Item# & Description
                Else
                    li.ListSubItems.Add , , Format(arrFGItems(r, c), "#,##0") 'Quantities
                End If
               
                '//-- This is where to set the row or individual cell colors based on certain criteria
                If r Mod 2 = 1 Then 'alternate row color
                    li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                Else
                    li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                End If

            Next c
                
        Next r
        
    End With
    
FinishUp:
    
    On Error Resume Next

    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    
    If blnAutoProcess = False Then
        Call SetFormPositions
        Application.StatusBar = ""
    End If
        
    Call PopulateFlagC2BArray 'always call flag FG items procedure as refresh will clear colors
    
    Exit Sub
    
ErrCatch:
    MsgBox Err.Description
    Resume FinishUp
    Call Main.ResetApplication


End Sub

Private Sub MultipleTabSelect_GotFocus()
'// -- 04/26/2022 - KP -- //
    
    Me.LstVw_BOM.Visible = False
    
End Sub

Private Sub MultipleTabSelect_Click(ByVal Index As Long)
    '// -- 06/01/2022 - KP -- //
    
    'close the pop-up listview component list if open
    If Me.LstVw_BOM_PopUp.Visible Then Me.LstVw_BOM_PopUp.Visible = False
    
    'Global var needed to control other mouseover events
    intTabSelected = Index
    
    Call MultiSelectControl(Index)
       
    'Hide the components pop=up when Exploded BOM components are visible
'    Me.LstVw_BOM.Visible = False

End Sub

Private Sub MultiSelectControl(tabSelected) 'format tab strip and text objects
'// -- 04/26/2022 - KP -- //
    
    Dim i As Long
    
    With Me.MultipleTabSelect
    
        i = .Value
        .SendToBack
                   
        .Pages(0).Caption = "FG Items"
        .Pages(1).Caption = "Components"
        
        If i = 0 Then
            Me.LstVw_OpenWorkOrdersFG.Visible = True
            Me.LstVw_BOM.Visible = False
            Call FlagFGClearToBuild 'flag items clear to build in bold green
            
        Else
            Me.LstVw_BOM.Visible = True
            Me.LstVw_OpenWorkOrdersFG.Visible = False
            Me.LstVw_BOM_PopUp.Visible = False
            
        End If
        
    End With
    
    'align ListView and multi-Page forms
    Call SetFormPositions
    
End Sub
Public Sub PopulateFlagC2BArray() 'Populates Exploded BOM ListView Data
'// -- 06/01/2022 - KP -- //

    Dim i, r, c, L, x, y As Long
    Dim strPad As String
    Dim lngViewWidth As Long
    Dim strColHeaders, strFieldName As String
    Dim strWbName As String
    Dim li As ListItem
    
    ThisWorkbook.Activate
    
    blnProcessError = False

    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    
    Application.StatusBar = "Loading Exploded BOM data . . ."
        
    strWbName = GetLocalPathFromSP(ThisWorkbook.Path) & "\" & ThisWorkbook.Name
    dbOpen (strWbName)

    strSQL = " SELECT DISTINCT(p.[Component]) AS [Component], p.[Item Type], p.[Component Desc] as [Description], p.[WO Req Qty], " & _
             "    p.[Demand Inside Lead Time Qty], p.[Net Available] as [Net Avail], p.[OH in Quar], p.[ASAP_Qty], p.[ASAP_Date], o.[Comp Ct], p.[Parent]  " & _
             " FROM ([OpenPO_Components$] p  " & _
             " LEFT JOIN (  " & _
             "    SELECT DISTINCT([Parent]) AS [Parent], COUNT([Component]) AS [Comp Ct]  " & _
             "    FROM [OpenPO_Components$] WHERE [Component] <> ''    " & _
             "    GROUP BY [Parent]  " & _
             "    ) o on o.[Parent] = p.[Parent]  " & _
             " )  " & _
             " WHERE p.[Component] <> '' " & _
             " GROUP BY p.[Component], p.[Item Type], p.[Component Desc], p.[Parent], p.[Demand Inside Lead Time Qty], " & _
             "    p.[Net Available], p.[OH in Quar], p.[ASAP_Qty], p.[ASAP_Date], o.[Comp Ct], p.[WO Req Qty] " & _
             "    ORDER BY p.[ASAP_Date] ASC "

    DoEvents
    
    rs.Open strSQL, cnn, adUseClient 'open connection and execute query

    DoEvents
    'Load recordset into array
    ArrFlagC2B = Application.WorksheetFunction.Transpose(rs.GetRows) 'load SQL recordset results into array
    
FinishUp:
    
    On Error Resume Next
    
    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
        
    Exit Sub
    
ErrCatch:

    Resume FinishUp

End Sub

Private Sub FlagFGClearToBuild() 'Flag FG Items ListView as clear to build by text color
'// -- 06/01/2022 - KP -- //

    Dim i, ar, c As Long 'components build array rows/cols
    Dim strWoFgItem, strC2BItem As String
    Dim lngSCR_DemandInsideLTQty, lngAvailOHQty, lngSCR_AvailOHQuarr As Long
    Dim lngPO_OpenQty, lngWO_OrderQty As Long
    Dim lngCompC2B, lngCompsNeeded As Long
    Dim blnClearToBuild As Boolean

    Me.Unprotect

    blnClearToBuild = False 'assume nothing is clear to build until component count is checked

    On Error Resume Next 'no consistent way to handle an empty array
    
    'if FG pop-up array is empty then populate it from local data since this is also the array used here
    If UBound(ArrFlagC2B) = 0 Then
        Call PopulateFlagC2BArray
    End If
    
    On Error GoTo 0

    With Me.LstVw_OpenWorkOrdersFG
        
        For i = 1 To .ListItems.Count 'Loop through all Item #'s in FG Items List View
        
            strWoFgItem = CStr(.ListItems(i)) 'FG Item Number

            'Loop through ArrFlagC2B rows
            For ar = 1 To UBound(ArrFlagC2B)

                'get the BOM (Parent# always last column in array is why Ubound is used)
                strC2BItem = CStr(ArrFlagC2B(ar, UBound(ArrFlagC2B, 2)))
            
                'Match Parent# to FG Item# from FG Item ListView array
                If CStr(strC2BItem) = CStr(strWoFgItem) Then

                    lngWO_OrderQty = ArrFlagC2B(ar, 4)
                    lngSCR_DemandInsideLTQty = ArrFlagC2B(ar, 5)
                    lngAvailOHQty = ArrFlagC2B(ar, 6) 'Net Available
                    lngSCR_AvailOHQuarr = ArrFlagC2B(ar, 7)
                    lngCompsNeeded = ArrFlagC2B(ar, 10) 'number of components needed

                    'Check if W/O qty exceeds QOH and available components available are greater than the W/O qty
                    If lngSCR_DemandInsideLTQty <= (lngAvailOHQty - lngSCR_AvailOHQuarr) Then
                        
                        'aggregate components with enough on hand
                        lngCompC2B = lngCompC2B + 1
                        
                    End If
                    
                    If (lngCompC2B = lngCompsNeeded And lngCompsNeeded > 0) Then
                        blnClearToBuild = True
                    End If

                End If
            
            Next ar
       
        'Set Green Color for Clear-20Build
        If blnClearToBuild = True Then
            For c = 1 To .ColumnHeaders.Count - 1 'don't count hidden column
              .ListItems.Item(i).ListSubItems(c).ForeColor = RGB(0, 150, 40)
              .ListItems.Item(i).ListSubItems(c).Bold = True
            Next c
            
        'Set others to standard color(black)
        Else
            For c = 1 To .ColumnHeaders.Count - 1 'don't count hidden column
              .ListItems.Item(i).ListSubItems(c).ForeColor = RGB(0, 0, 0)
              .ListItems.Item(i).ListSubItems(c).Bold = False
            Next c
            
        End If
                    
        blnClearToBuild = False
        lngCompsNeeded = 0
        lngCompC2B = 0
        
        Next i
        
FinishUp:
        
        .Refresh
    
    End With

    Exit Sub
    
ErrCatch:


    Resume FinishUp:

End Sub

Private Sub MultipleTabSelect_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    lngQueryQueue = 0
    
 '   If lngLastX < 3 And lngLastY < 3 Then Exit Sub

    On Error Resume Next 'in case of first time run
    lngLastX = Abs(x - lngLastX)
    lngLastY = Abs(y - lngLastY)
    
    On Error GoTo 0
    
'    If (Me.LstVw_BOM.Visible = True And x > 10 And x < 366 And y < 27) Then
    
    If (Me.LstVw_BOM_PopUp.Visible = True And x > 10 And x < 366 And y < 25) Then
        Me.LstVw_BOM_PopUp.Visible = False
'        DoEvents
    End If
    
End Sub

Public Sub PopulateExplodedBOM_LV() 'Populates Exploded BOM ListView Data
'// -- 04/28/2022 - KP -- //
    Dim i, r, c, L, x, y As Long
    Dim strPad As String
    Dim lngViewWidth As Long
    Dim strColHeaders, strFieldName As String
    Dim strWbName As String
    Dim li As ListItem
    
    ThisWorkbook.Activate
    
    blnProcessError = False

    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    
    Application.StatusBar = "Loading Exploded BOM data . . ."
        
    strWbName = GetLocalPathFromSP(ThisWorkbook.Path) & "\" & ThisWorkbook.Name
    dbOpen (strWbName)

    strSQL = " SELECT Distinct([Component]) as [Component#], [Item Type], [Component Desc] as [Description], COUNT([Parent]) as [Common FG's], " & _
             "     SUM([WO Req Qty]) as [W/O Qty], SUM([Demand Inside Lead Time Qty]) as [Dmd ILT Qty], " & _
             "     SUM([Total Forecast Sets 1]) as [Forecast 1], SUM([Total Forecast Sets 2]) as [Forecast 2], " & _
             "     SUM([Total Forecast Sets 3]) as [Forecast 3], SUM([Net Available]) as [Net Avail], SUM([ASAP_Qty]) as [ASAP Qty], " & _
             "     MIN([ASAP_Date]) as [ASAP Date], SUM([Total_Qty]) as [Tot Qty], SUM([OH in Quar]) as [Av OH Quar]  " & _
             " FROM [OpenPO_Components$] " & _
             " WHERE [Parent] Is Not Null " & _
             " GROUP BY [Component], [Item Type], [Component Desc] " & _
             " ORDER BY [Component] ASC"

    DoEvents
    
    rs.Open strSQL, cnn, adUseClient 'open connection and execute query

    DoEvents
    'Load recordset into array
    ArrExpBOM = Application.WorksheetFunction.Transpose(rs.GetRows) 'load SQL recordset results into array

    On Error GoTo ErrCatch
        
    'Skip populating the Exp BOM Components tab if only needing to populate the pulic array, 'ArrExpBOM'
    If blnDisableEvents = True Then Exit Sub
    
    'Populate Exploded BOM ListView
    With Me.LstVw_BOM

        .Font.Size = 10
        
        .View = lvwReport
        .Gridlines = True
        .ListItems.Clear
    
        With .ColumnHeaders
        
            .Clear
                
            'add and hide a dummy column to allow right or left justification
            .Add , , "", 0
                    
            For i = 0 To rs.Fields.Count - 2 'Do not include Avail OH in Quarr (count is normally -1)
                
                strFieldName = rs.Fields(i).Name

                Select Case True 'set column widths

                    Case (strFieldName Like "Component*" Or strFieldName Like "*Item Type*")
                        .Add , , strFieldName, 100, lvwColumnLeft '2 centers field

                    Case strFieldName Like "*Type*" 'Description
                        .Add , , strFieldName, 140, lvwColumnLeft
                    
                    Case strFieldName Like "*Desc*" 'Description
                        .Add , , strFieldName, 250, lvwColumnLeft
                    
                    Case strFieldName Like "*FG*" 'Description
                        .Add , , strFieldName, 92, lvwColumnRight

                    Case LCase(strFieldName) Like "*need by*" 'W/O Count
                        .Add , , strFieldName, 85, 2 '2 is for center alignment

                    Case ((LCase(strFieldName) Like "*qty*" And Not LCase(strFieldName) Like "*lt*") Or LCase(strFieldName) Like "forecast*")
                        .Add , , strFieldName, 85, lvwColumnRight
    
                    Case Else 'All other columns
                        .Add , , strFieldName, 90, 1
                
                End Select
                
                'Align headers regardless of column alignment by padding with spaces
                If i > 0 Then
                    If .Item(i).Alignment = 1 Then
                        strPad = ""
                        
                        For L = 1 To Int((17 - Len(strFieldName)) / 2) - 1
                            strPad = strPad & " "
                        Next L
                        .Item(i).Text = .Item(i).Text & strPad
                    End If
                End If
                
                'set total item list view width by the sum of all column widths together
                If i < rs.Fields.Count - 1 Then
                    lngViewWidth = lngViewWidth + .Item(i + 1).Width
                Else
                    lngViewWidth = Round(lngViewWidth * 0.88, 1)  'view window offset width
'                    lngViewWidth = lngViewWidth - 40 'view window offset width
                End If
            Next i
            
            'Avail OH in Quarr invisible, and only used to flag clear to build
            .Add , , strFieldName, 0, 1
        
        End With
        
        'list view total width added from field name column widths
        .Width = lngViewWidth
        
        'Populate data fields
        For r = 1 To UBound(ArrExpBOM) 'row loop

            c = 1 'must be instantiated for first item
        
            Set li = .ListItems.Add(, , ArrExpBOM(r, c))

            For c = 1 To UBound(ArrExpBOM, 2) 'column loop
                
                Select Case True
                    Case c < 4 'set field formats
                        li.ListSubItems.Add , , ArrExpBOM(r, c) 'text for FG Item# & Description
                    Case c = 12
                        li.ListSubItems.Add , , Format(ArrExpBOM(r, c), "m/d/yyyy")
                    Case Else
                        li.ListSubItems.Add , , Format(ArrExpBOM(r, c), "#,##0") 'Quantities
                End Select
                
                
                '//-- This is where to set the row or individual cell colors based on certain criteria
                If r Mod 2 = 1 Then 'alternate row color
                    li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                Else
                    li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                End If

            Next c

        Next r
        
    End With
    
    SetFormPositions
    
    blnProcessError = False

FinishUp:

    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    If blnAutoProcess = False Then Application.StatusBar = ""
    
    Exit Sub
    
ErrCatch:
 
    MsgBox Err.Description, 65536, "Populate Exploded BOM List View"
    Resume FinishUp

End Sub

Private Sub PopulateFG_PopUpArray() 'Populates FG Items pop-up component list view
'// -- 04/27/2022 - KP -- //
    
    Dim i As Long
    Dim strDbName, strSourcePath As String 'workbook source

    Dim cws As Worksheet

    Set cws = ThisWorkbook.Worksheets("Controls")
    
    blnProcessError = False
    
    If cnn.State = 1 Then cnn.Close
    If rs.State = 1 Then rs.Close
    
    strDbName = GetLocalPathFromSP(ThisWorkbook.Path) & "\" & ThisWorkbook.Name
    
    DoEvents
    
    On Error GoTo 0
    
    dbOpen (strDbName)

    strSQL = " SELECT DISTINCT(p.[Component]) AS [Component], p.[Item Type], p.[Component Desc] as [Description],  " & _
             "   p.[Net Available] as [Net Avail], p.[ASAP_Qty], p.[ASAP_Date], IIF(o.[Shared FG's]>1,o.[Shared FG's],0) as [Shared FG's], p.[Parent]  " & _
             " FROM ([OpenPO_Components$] p " & _
             " LEFT JOIN ( " & _
             "      SELECT DISTINCT([Component]) AS [Component], COUNT([Parent]) AS [Shared FG's] " & _
             "      FROM [OpenPO_Components$] WHERE [Component] <> ''     " & _
             "      GROUP BY [Component] " & _
             "      ) o on o.[Component] = p.[Component] " & _
             "   ) " & _
             " WHERE p.[Component] <> '' " & _
             " GROUP BY p.[Component], p.[Item Type], p.[Component Desc], o.[Shared FG's], " & _
             "   p.[Parent], p.[Net Available], p.[ASAP_Qty], p.[ASAP_Date] " & _
             " ORDER BY p.[ASAP_Date] ASC "

    rs.Open strSQL, cnn, adUseClient 'open connection and execute query
        
    DoEvents

    On Error GoTo ErrCatch
    
    'Set full recordset into ArrBOM array
    ArrBOM = Application.WorksheetFunction.Transpose(rs.GetRows) 'load SQL recordset results into array

    'Dim the fields array
    ReDim arrHeaders(1 To rs.Fields.Count)

    For i = 1 To UBound(arrHeaders)

        arrHeaders(i) = rs.Fields(i - 1).Name

    Next i

FinishUp:
    
    On Error Resume Next
    If cnn.State = 1 Then cnn.Close
    If rs.State = 1 Then rs.Close

    If blnAutoProcess = False Then Application.StatusBar = ""
    
    Exit Sub

ErrCatch:
    Resume FinishUp

End Sub

Sub BOM_LvPopUp(strItemNo) 'Populates the pop-up List View components from the FG Items ListView
'// -- 04/27/2022 - KP -- //  'Triggered by procedure: LstVw_OpenWorkOrdersFG_MouseMove

    Dim i, r, c, p As Long
    Dim lngViewWidth, lngViewHeight, lngHeightOffset As Long
    Dim lngColWidth, lngCompCount As Long
    Dim strColHeaders, strFieldName, strPadding As String
    Dim li As ListItem
       
    On Error Resume Next
    
    blnDisableEvents = False
           
    'if array is empty then populate it from local data
    If UBound(ArrBOM) < 1 Then Call PopulateFG_PopUpArray
    
    DoEvents

    On Error GoTo 0
    
    'Populate FG Listbox
    With LstVw_BOM_PopUp
        .BringToFront
        .Visible = False
        .View = lvwReport
        .Gridlines = True
        .ListItems.Clear
        .Refresh
        
        With .ColumnHeaders
        
            .Clear
                
            'add and hide a dummy column to allow right or left justification
            .Add , , "", 0
                    
            For i = 1 To UBound(arrHeaders) - 1 'don't include parent #.  Only used for matching components
                
                strFieldName = Replace(arrHeaders(i), "_", " ")
                                
                Select Case True 'set column widths
                
                    Case strFieldName Like "Component" 'FG Item#W
                        strFieldName = "FG#: " & strItemNo
                        .Add , , strFieldName, 120, 2 '2 aligns centered
                        
                    Case strFieldName Like "Item Type*" 'Item Type
                        .Add , , strFieldName, 130, lvwColumnLeft
                        
                    Case strFieldName Like "Desc*" 'Description
                        lngColWidth = 340
                        .Add , , strFieldName, lngColWidth, lvwColumnLeft
                        
                    Case LCase(strFieldName) Like "*shared*" 'Shared FG's
                        .Add , , strFieldName, 80, 2 '2 aligns centered
                    
                    Case LCase(strFieldName) Like "*parent*" 'FG Dependants
                       'Don't add a column since Parent is only used to count dependancies

                    Case Else 'All other columns
                        .Add , , strFieldName, 80, 2
                
                End Select
                
                'Calculate item list view width by the sum of all column widths together
                If i < UBound(arrHeaders) + 1 Then
                    lngViewWidth = lngViewWidth + .Item(i).Width
                End If
            Next i
        
        End With
        
        'width offset due to variation in pixel vs. inches
        lngViewWidth = Round(lngViewWidth * 0.832, 1)  'view window offset width
'        lngViewWidth = 570
        
        'Loop to populate data fields 'Note: List view column uses Base 0.  Skip using 0 column since it can't be formatted (bug in VBA ListView 6.0 objects)
        For r = 1 To UBound(ArrBOM)

            'Match Item selected with Parent number to get component list
            If CStr(ArrBOM(r, UBound(ArrBOM, 2))) = strItemNo Then 'Use Ubound column count since Parent will always be the last field
                                                            
                i = i + 1
                c = 1 'must be instantiated for first item
                
                lngCompCount = lngCompCount + 1

                Set li = .ListItems.Add(, , ArrBOM(r, c))
                
                'Populate data fields
                For c = 1 To UBound(ArrBOM, 2) - 1 'do not include the parent #.  It is only used for matching components
                    
                    'Format component#, description, and
                    If c < 4 Then
                        li.ListSubItems.Add , , ArrBOM(r, c) 'text for FG Item#, Type & Description
                        
                    Else
                        'Format dates and quantities - Column count is offset +1 due to starting LstVw col 1 instead of col 0
                        If LCase(.ColumnHeaders.Item(c + 1).Text) Like "*date*" Then
                            li.ListSubItems.Add , , Format(ArrBOM(r, c), "m/d/yyyy") 'Format Dates
                        Else
                            li.ListSubItems.Add , , Format(ArrBOM(r, c), "#,##0;;-") 'Format Quantities (hyphen zeros)
                        End If
                        
                    End If
                   
                    '//-- This is where to set the row or individual cell colors based on certain criteria
                    If r Mod 2 = 1 Then 'alternate row color
                        li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                    Else
                        li.ListSubItems.Item(c).ForeColor = RGB(0, 0, 0)
                    End If
    
                Next c
                
                .ColumnHeaders(4).Text = "Description"
                
'                'pad spacing between 'Description' and Component Count in Description header
'                For p = 1 To Int(lngColWidth / 7)  'pad spacing between 'Description' and Shared FG count in Description header(might use later)
'
'                    strPadding = strPadding & " "
'
'                Next p
'
'                p = p - (Len(CStr(lngCompCount)) * 10)
'
'                .ColumnHeaders(4).Text = "Description" & strPadding & "Total Components: " & lngCompCount
'
'                strPadding = ""

            End If
                
        Next r
        
        'Adjust for linecount to pixel conversion since it is not proportional
        Select Case True
            Case .ListItems.Count < 11
                lngHeightOffset = 18
            Case Else
                lngHeightOffset = 12
        End Select
        
        lngViewHeight = lngHeightOffset + Round(.ListItems.Count * 12.5, 0)

        If lngCompCount > 0 Then
            .Visible = True
            
            'HORIZONTAL position on sheet
            .Left = 360
            
            'VERTICAL position on sheet - shift up if list is too long to see in current position.
            
            If lngViewHeight < 450 Then
                .Top = 105 'set absolute vertical position on sheet
            Else
                .Top = 5 'shift top position and total height if list is too long for current 'top' position
                lngViewHeight = 450
            End If
            
            'list view height & widths dynamic values
            .Width = lngViewWidth
            .Height = lngViewHeight
                    
            .sortKey = 2
            .SortOrder = lvwAscending
            .Sorted = True
            .Refresh
            .Activate
            '.Select 'keeps cursor from landing in FG Item# column in edit mode
            .Font.Size = 10
        Else
            .Visible = False
        End If
        
    End With

    Exit Sub

FinishUp:

    If rs.State = 1 Then rs.Close
    If cnn.State = 1 Then cnn.Close
    
    On Error Resume Next
    Exit Sub
    
ErrCatch:
    MsgBox "Query or array error: " & Err.Description, 65536, "ListView Components Pop-Up"
    blnDisableEvents = True
    Sleep 1000 '1000 = 1 second - Allows enough time to move cursor off listview area
    Resume FinishUp

End Sub

Private Sub DevFindObjects() 'Dev-testing only to find hidden objects on sheet
'// -- 04/25/2022 - KP -- //   (Author of base code unknown -
    Dim obj As Object
    
    Debug.Print Me.OLEObjects.Count
    Stop
    For Each obj In Me.OLEObjects
        
        Debug.Print obj.Name
        obj.Visible = True
        Stop
        obj.Visible = False
        obj.Visible = True
        
    Next
End Sub

Private Sub LstVw_BOM_LostFocus()
'// -- 04/26/2022 - KP -- //

   ' Me.LstVw_BOM.SendToBack
    Me.LstVw_BOM.Visible = False

End Sub

Private Sub TxBx_ScreenBackgroundTitle_click()

    Me.LstVw_BOM.Visible = False
    
End Sub


Private Sub Worksheet_Activate() 'Populate all FG items List View window
 '// -- 05/31/2022 - KP -- //
    
    'prevent inadvertent event-triggered loop
    If blnDisableEvents = True Then Exit Sub
    
    'default to FG items list
    Me.MultipleTabSelect.Value = 0
    Me.LstVw_OpenWorkOrdersFG.Visible = True
    Me.LstVw_BOM.Visible = False
    Call PopulateFGData
    Call PopulateExplodedBOM_LV
    Call FlagFGClearToBuild 'Change FG Items text color to green if clear to build
    
    'bring FG Items to foreground if that was the last tab selected
    If Me.MultipleTabSelect.Value = 0 Then
        With Me.LstVw_OpenWorkOrdersFG
            .Visible = True
            .BringToFront
        End With
    End If
    
    ThisWorkbook.Saved = True
    blnDisableEvents = False
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If blnDisableEvents Then Exit Sub
'   ' Me.LstVw_BOM.SendToBack
'    Me.LstVw_BOM.Visible = False
'    Me.LstVw_OpenWorkOrdersFG.BringToFront
    
    Call SetFormPositions

End Sub

Public Sub SetFormPositions() 'keep form positions where they belong on the sheet
'// -- 06/01/2022 - KP -- //
    
    Dim i, lngFormWidthSum As Long
    Dim ws As Worksheet

    Me.Activate 'set focus to Materials_Status sheet to prevent event-triggered errors in forms
    
    Dim xLeft, yTop, lngHeight, lngWidth, lngTabSelected As Long
    Dim dblFontSize As Double
    
    lngFormWidthSum = 0
    
    'make sure no sheets are locked that might cause errors
    For Each ws In ThisWorkbook.Worksheets
        With ws
            If ws.Name = "Materials_Status" Then
                .EnableSelection = xlNoRestrictions
                .Protect contents:=False, UserInterfaceOnly:=False
            End If
        End With
    Next

    On Error GoTo ErrCatch
    
    xLeft = 49
    yTop = 137
    lngHeight = 280
    lngWidth = 1030 'Total Multi-Page selector
    dblFontSize = 10

    Me.Unprotect
    
    'Multi-Page/Page Selector
    With MultipleTabSelect
        If Round(.Left, 0) <> xLeft Then .Left = xLeft
        If Round(.Top, 0) <> yTop Then .Top = yTop 'set vertical position on sheet
        If Round(.Height, 0) <> 340 Then .Height = 340
        If Round(.Width, 0) <> lngWidth Then .Width = lngWidth
        
        'get coordinates for other objects
        xLeft = .Left + 10 'set vertical position on sheet
        yTop = .Top + 50 'set horizontal position on sheet
        lngHeight = .Height - 60
        lngTabSelected = .Value
'        .SendToBack
    End With
    
    If MultipleTabSelect.Value = 0 Then
        'Finished Goods Page
        With LstVw_OpenWorkOrdersFG
            .Font.Size = dblFontSize
            If Round(.Top, 0) <> yTop Then .Top = yTop  'set vertical position on sheet based on Multi-Page Select
            If Round(.Left, 0) <> xLeft Then .Left = xLeft 'set horizontal position on sheet based on Multi-Page Select
            If Round(.Height, 0) <> lngHeight Then .Height = lngHeight
            
            'If Round(.Width, 0) <> 766 Then .Width = 766
            
            For i = 2 To .ColumnHeaders.Count
                lngFormWidthSum = lngFormWidthSum + (.ColumnHeaders(i).Width - 25)
            Next i
            
            .Width = lngFormWidthSum

    '        If lngTabSelected = 0 Then
    '            .Visible = True
    '
    ''            'Make checkbox visible only on FG Items tab (Won't stay visible at the moment)
    ''            With Me.Shapes("ChkBx_HighlightC2B")
    ''                .Visible = msoCTrue
    ''                .ZOrder msoBringToFront
    ''            End With
    ''
    '        Else
    '            .Visible = False
    '        End If
        End With
    Else
        'Exploded BOM Page
        With Me.LstVw_BOM
            .Font.Size = dblFontSize
            .Top = Me.LstVw_OpenWorkOrdersFG.Top
            .Left = Me.LstVw_OpenWorkOrdersFG.Left
            .Height = lngHeight
            .Width = Me.MultipleTabSelect.Width - Int(Me.MultipleTabSelect.Left / 2) + 2
    '        If lngTabSelected = 1 Then
    '            .Visible = True
    '        Else
    '            .Visible = False
    '        End If
            For i = 2 To .ColumnHeaders.Count
                lngFormWidthSum = lngFormWidthSum + (.ColumnHeaders(i).Width - 22)
            Next i
            
            .Width = lngFormWidthSum
            
        End With
    
    End If
    
    'this is usually the last procedure whether manual or automated, so a good place to set it
    blnProcessError = False
    
'    'Lock sheets where appropriate
'    For Each ws In ThisWorkbook.Worksheets
'        With ws
'            If ws.Name = "Materials_Status" Then
'                .EnableSelection = xlNoRestrictions
'                .Protect Contents:=False, UserInterfaceOnly:=False
'            End If
'        End With
'    Next
'

    Exit Sub
    
ErrCatch:
    MsgBox Err.Description, 65536, "Error Setting Form Positions"
    
End Sub

'Private Sub LstVw_OpenWorkOrdersFG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
'                    ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)  'Used for Mouse hover ListView
''// -- 04/27/2022 - KP -- //   'Unknown author of base code: https://stackoverflow.com/questions/66686689/mousehover-in-listview-vba/72001159#72001159
'
'    Dim i As Long
'    Dim itm As MSComctlLib.ListItem
'
'    'exit sub if FG Items tab is not selected (tab 0)
'    If intTabSelected > 0 Then Exit Sub
'
'    'limit excessive queries and screen flickering by creating a null hover area
'    If (y Mod 10 > 2 And y Mod 10 < 8) Then
'        If (x Mod 5 > 3 And x Mod 5 < 1) Then Exit Sub
'    End If
'
'    'decrement query stack until it equals 1 to prevent overrun of queries outside of hover area
'    If lngQueryQueue > 1 Then
'        lngQueryQueue = lngQueryQueue - 1
'        Exit Sub
'    End If
'
'    If lngQueryQueue > 2 Then Exit Sub 'Don't perform multiple queries as the mouse cursor sweeps across items
'
'    If LstVw_OpenWorkOrdersFG.ListItems.Count < 1 Then Exit Sub ' Don't show anything if component count is < 1
'
'    DoEvents
'
'    'protect from crashes when mouse moves outside range of ListView window
'    On Error Resume Next
'
'    LstVw_OpenWorkOrdersFG.SelectedItem.Selected = False ' unselect a previous selected subitem
'
'    ConvertPixelsToTwips x, y  'make the necessary units conversion
'
'    Set itm = LstVw_OpenWorkOrdersFG.HitTest(x, y) 'set the object using the converted coordinates
'
'    If Not itm Is Nothing Then
'        itm.Selected = True
'    End If
'
'    'Only perform action when the item index changes so multiple events aren't triggered while hovering over the same item
'    'If (itm.Index = lngLastItmIndex Or (lngLastItmIndex - itm.Index) > 10) Then Exit Sub
'
'    If itm.Index = lngLastItmIndex Then Exit Sub
'
'    On Error GoTo errCatch
'
'    'if the db connection is already open, don't queue a series of queries that won't be seen
''    If cnn.State = 1 Then Exit Sub
'
'    DoEvents
'
'    lngLastItmIndex = itm.Index
'
'    strItemNo = LstVw_OpenWorkOrdersFG.ListItems(lngLastItmIndex)
'
'    If strLastItemNo = strItemNo Then Exit Sub
'
'    On Error GoTo errCatch
'
'    'Make visible the component listbox over the Open Workorders FG Item list
'    If x < 7550 Then 'limit to work only over the left few columns
'        If Me.LstVw_BOM.Visible = False Then
'            Me.LstVw_BOM.Visible = True
'            Me.LstVw_BOM.BringToFront
'        End If
'
'        'prevent multiple and duplicate queries stacking-up
'        If strLastItemNo <> strItemNo Then
'
'    'POP-UP ComponentsView
'            BOM_LvPopUp (strItemNo)
'
'            If lngQueryQueue >= 0 Then
'                lngQueryQueue = lngQueryQueue + 1 'keep query queue from hitting zero
'                Sleep 50
'            End If
'            strLastItemNo = strItemNo 'track last item number to prevent duplicate queries
'
'        End If
'
'        'prevent error if stopped mid-query with nothing selected
'        'If Not LstVw_OpenWorkOrdersFG.Selected Is Nothing Then LstVw_OpenWorkOrdersFG.Selected = False
'
'    Else
'        If Me.LstVw_BOM.Visible = True Then
'            Me.LstVw_BOM.Visible = False
'           ' Me.LstVw_BOM.SendToBack
'        End If
'
'    End If
'
'    Exit Sub
'
'FinishUp:
'
'    DoEvents
'
'    If cnn.State = 1 Then cnn.Close
'
'errCatch:
'    If (Err.Description = "" Or Err.Description Like "*without*") Then Exit Sub
'    MsgBox Err.Description
'    Resume FinishUp
'
'End Sub

Private Sub ConvertPixelsToTwips(ByRef x As stdole.OLE_XPOS_PIXELS, ByRef y As stdole.OLE_YPOS_PIXELS) 'used for mouse hover
'// -- 04/25/2022 - KP - //    'Unknown author of base code: https://stackoverflow.com/questions/66686689/mousehover-in-listview-vba/72001159#72001159

    Dim hDC As Long, RetVal As Long, TwipsPerPixelX As Long, TwipsPerPixelY As Long
    Const LOGPIXELSX = 88
    Const LOGPIXELSY = 90
    Const TWIPSPERINCH = 1440
 
    hDC = GetDC(0)
    TwipsPerPixelX = TWIPSPERINCH / GetDeviceCaps(hDC, LOGPIXELSX)
    TwipsPerPixelY = TWIPSPERINCH / GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    x = x * TwipsPerPixelX:  y = y * TwipsPerPixelY
End Sub


Sub UpdatePQDataPath(strQueryName, strSourcePath) 'Clear and replace previous data set with updated data
'// -- 03/28/2022 - KP -- //          'Specify fully-qualified filename to update power query where query is the same name as target sheet
        
    Dim strFormula As String
    Dim strFileNamePartial, strPrevPath As String
    Dim pQuery
    Dim strPQueryName As String
    Dim strTargetSheet As String
    Dim ws As Worksheet
    
    'extrapolate report name from file path to address proper PQ formula
    strFileNamePartial = Left(Mid(strSourcePath, InStrRev(strSourcePath, "\") + 1), Len(Mid(strSourcePath, InStrRev(strSourcePath, "\") + 1)))
'
'    If strFileNamePartial Like "SCR*" Then
'        strFileNamePartial = "SupplyChainReport"
'    End If
'
    strPQueryName = Replace(Replace(strFileNamePartial, "_", ""), ".xlsx", "")

    'Find the target query by name passed in
    For Each pQuery In thisworkbook.Queries
    
        If pQuery.Name Like "*" & strPQueryName & "*" Then
             
            With pQuery
            
                strFormula = .Formula 'get the power query formula
                
                strPrevPath = Left(Mid(strFormula, InStr(strFormula, "C:")), InStr(Mid(strFormula, InStr(strFormula, "C:")), ".xlsx") + 4) 'get path from formula
                
                strFormula = Replace(strFormula, strPrevPath, strSourcePath) 'replace old path with new path
                
                .Formula = strFormula 'set the path name in the saved query
                
                .Refresh 'get the most recent data

            End With
        
        End If
        
    Next
    
End Sub

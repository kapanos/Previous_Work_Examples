Option Explicit

Sub testEmail()

    Email_Reports (Worksheets("Settings").Range("G3"))
    
End Sub

Sub Email_Reports(strSaveFile As String)
'// -- 10/02/2020 KP -- // 'Generates email and sends to any number of recipients contained 
						   'in a single line string delimited by semi-colon
    Dim strSenderName As String
    Dim strSenderEmail As String
    Dim strReplyAddress As String
    Dim strRecipients As String
    Dim strSubject As String
    Dim strMessageBody As String
    Dim strDaysWeeksBack As String
    Dim strUnitDescription As String
    Dim strReportName As String, intUnitNumber As Integer
    Dim chtChart As ChartObject
    Dim strCalendarInterval As String, strCalendarIntervalType As String, intDateSpan As Integer
    Dim Mail As New Message
    Dim Cfg As Configuration
    Dim cpws As Worksheet, sws As Worksheet
    
    Set cpws = ThisWorkbook.Worksheets("Control_Panel")
    Set sws = ThisWorkbook.Worksheets("Settings")
    
    DoEvents
    
    strDaysWeeksBack = sws.Range("D6") 'number of days or weeks of historical data back
    strUnitDescription = sws.Range("E6")  'Verbose location of facility for email body
    strReportName = cpws.Shapes.Range(Array("SC_ReportTitle")).TextFrame.Characters.Text 'Report name from yellow report button
    
    cpws.Activate
    
    strCalendarInterval = sws.Range("E15") 'user-selected calendar interval(i.e. Days, Weeks, Months)
    intDateSpan = sws.Range("D6") 'quantity of days, weeks, months
    
    Select Case strCalendarInterval
        Case "d"
            strCalendarIntervalType = "days"
        Case "ww"
            strCalendarIntervalType = "Weeks"
        Case "m"
            strCalendarIntervalType = "months"
        Case Else
    End Select
    
    strSenderName = ""
    strSenderEmail = sws.Range("D18")
    strReplyAddress = ""
    strRecipients = GetEmailList
    
    If ProcessError = False Then
        strSubject = strReportName & " Report"
        strMessageBody = "Attached is the " & strReportName & " report for " & sws.Range("D12") & " " & sws.Range("E12") & " for the past " & _
            intDateSpan & " " & strCalendarIntervalType & " " & ". <br><br>" & _
            "Report Name: " & strReportName
    Else
        strSubject = strReportName & " - Report error!"
        strMessageBody = "There was an error creating the " & strReportName & " report.  Please check the report target folder in 'Report Manager'."
    End If
    
    If Len(strRecipients) > 1 Then
    
        Application.StatusBar = "Generating and sending email with attachments . . ."
    
        Set Cfg = Mail.Configuration 'instantiate Outlook mail configuration

        Cfg(cdoSendUsingMethod) = cdoSendUsingPort 'send using specified SMTP port
        Cfg(cdoSMTPServer) = "smtp.SanteeCooper.Com"
        Cfg(cdoSMTPServerPort) = 25
        Cfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
       'Cfg(cdoSMTPAuthenticate) = cdoBasic
        'Cfg(cdoSMTPUseSSL) = True  'Use SSL
        'Cfg(cdoSendUserName) = "" 'User login name when required
        'Cfg(cdoSendPassword) = "" 'Email password when required
        Cfg.Fields.Update 'define and configure email fields to use to construct message defined by MS schemas

        With Mail 'Prepare Email Message
            .From = strSenderEmail
            .ReplyTo = strReplyAddress
            .To = strRecipients
            .Subject = strSubject
            .HTMLBody = strMessageBody
            If ProcessError = False Then .AddAttachment strSaveFile
        
        End With
                    
        Mail.Send

        strRecipients = ""
        
        'MsgBox "Reports submitted", , "Notification"
    
    Else
    
        MsgBox "No email recipients specified." & Chr(10) & Chr(10) & _
            "Please add recipients and re-submit.", , "Email Failed!"
    End If
    
    Call SetVersion
    
    Set Mail = Nothing
    Set Cfg = Nothing

End Sub

Sub HeightWidth()
    
    With Worksheets("Sheet1")
    
        'Adjust Row hieght
        .Rows("2:3").RowHeight = 15
        .Rows("4:51").RowHeight = 12.75
          
        'Adjust Column Width
        .Columns("B:B").ColumnWidth = 16
        .Columns("C:T").ColumnWidth = 10
        .Columns("U:AB").ColumnWidth = 11
        
    End With
    
    Range("C1").Select
    
End Sub

Sub Delete_Files()

    Dim Dir_Path As String
    Dim oFSO As Object
    Dim oFile As Object
    Dim sws As Worksheet
    
    Set sws = ThisWorkbook.Worksheets("Settings")
    
    Dir_Path = sws.Range("E3")
    
    'Make sure a properly formatted path is stated ' // --  04/15/2020 KP -- //
    If Right(Dir_Path, 1) <> "\" Then
        Dir_Path = Dir_Path & "\"
        sws.Range("E3") = Dir_Path
    End If
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    If oFSO.FolderExists(Dir_Path) Then 'Check that the folder exists
        For Each oFile In oFSO.GetFolder(Dir_Path).Files
            If Now() - FileDateTime(oFile) > 60 Then  ' // --  04/15/2020 KP -- //
                oFile.Delete
            End If
        Next
    End If

End Sub

Sub navReportReturn() 'Jump back to report page
' // -- KP 04/06/2020 -- //
    
    Dim cpws As Worksheet, dsws As Worksheet
    Set cpws = ThisWorkbook.Worksheets("Control_Panel")
    Set dcws = ThisWorkbook.Worksheets("Settings")
    
    cpws.Activate
    cpws.Range("C1").Activate
    
    Call GetEmailList
    
    dsws.Visible = xlSheetHidden
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

End Sub

Sub QuitExcel()
    
    'copy one cell to clear the clipboard
    Sheets("Control_Panel").Select
    Range("C1").Select
    Selection.Copy
    
    DoEvents
    
    'mark the workbook as saved so that Excel does not prompt
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Saved = True
    
    'close Excel
    Application.Quit
    ThisWorkbook.Close

    
End Sub



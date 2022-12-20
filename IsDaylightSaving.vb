Option Explicit

Sub testDLS()

    Debug.Print ""
    Debug.Print ""
    
    Debug.Print dstHourOffset("10/09/2019", "11/20/2019")
    
End Sub

Public Function dstHourOffset(dtStart As Date, dtEnd As Date) As Integer
    'Returns +1/-1 hour offset if start/end dates span start or end of daylight saving
    'Returns 0 if both dates fall within or outside DST
    
    Application.ScreenUpdating = False
    '// -- KP 05/08/2020 -- //
    Dim boolDateStart As Boolean
    Dim boolDateEnd As Boolean
    
    boolDateStart = isDaylightSaving(dtStart)
    boolDateEnd = isDaylightSaving(dtEnd)
    
    'If start date outside DST and end date inside DST
    If (boolDateStart = False And boolDateEnd = True) Then dstHourOffset = -1
    
    'If date span doesn't transition through dst change
    If ((boolDateStart = False And boolDateEnd = False) _
        Or (boolDateStart = True And boolDateEnd = True)) Then dstHourOffset = 0
    
    'If start date inside DST and end date outside DST
    If (boolDateStart = True And boolDateEnd = False) Then dstHourOffset = 1

End Function

 Public Function isDaylightSaving(chkDate As Date) As Boolean 'Determines if date is within daylight saving time
    
    '// -- KP 05/08/2020 -- //
    Dim dtDaylightStart As Date
    Dim dtDaylightEnd As Date
    
    dtDaylightStart = DateSerial(Year(chkDate), 3, 1) + 14 - Weekday(DateSerial(Year(chkDate), 3, 1) - 1)
    dtDaylightEnd = DateSerial(Year(chkDate), 11, 1) + 7 - Weekday(DateSerial(Year(chkDate), 11, 1) - 1)

    If (chkDate > dtDaylightStart And chkDate < dtDaylightEnd) Then
        isDaylightSaving = True
    Else
        isDaylightSaving = False
    End If

End Function



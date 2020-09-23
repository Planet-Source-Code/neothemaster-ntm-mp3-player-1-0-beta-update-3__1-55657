Attribute VB_Name = "Mod"
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Type POINT_TYPE
  X As Long
  Y As Long
End Type
Declare Function GetTickCount& Lib "kernel32" ()
Public Function GetUptime()
Dim Secs, Mins, Hours, Days
Dim TotalMins, TotalHours, TotalSecs, TempSecs
Dim CaptionText
TotalSecs = Int(GetTickCount / 1000)
Days = Int(((TotalSecs / 60) / 60) / 24)
TempSecs = Int(Days * 86400)
TotalSecs = TotalSecs - TempSecs
TotalHours = Int((TotalSecs / 60) / 60)
TempSecs = Int(TotalHours * 3600)
TotalSecs = TotalSecs - TempSecs
TotalMins = Int(TotalSecs / 60)
TempSecs = Int(TotalMins * 60)
TotalSecs = (TotalSecs - TempSecs)
If TotalHours > 23 Then
Hours = (TotalHours - 23)
Else
Hours = TotalHours
End If
If TotalMins > 59 Then
Mins = (TotalMins - (Hours * 60))
Else
Mins = TotalMins
End If
GetUptime = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " Seconds"
End Function

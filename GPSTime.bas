Attribute VB_Name = "CalGPSTime"
Public Function GPSTIME(Year As Integer, Month As Integer, Day As Integer, Hour As Integer, Min As Integer, Sec As Double) As Double
Dim N As Integer, flag As Integer, d1 As Long, d2 As Long, Z As Long, M As Double
'n is the cycle of the leap year, flag means whether it's leap year, d1 and d2 is the number of days, Z is the number of week, M is the number of seconds

N = (Year - 1980) \ 4
If (Year - 1980 - 4 * N) = 0 Then                                        'whether it's leap year
  flag = 1
  d1 = 1461 * N
Else
  flag = 0
  d1 = 1461 * N + (Year - 1980 - 4 * N) * 365 + 1
End If

Select Case Month
  Case 1: d2 = d1 + Day - 6
  Case 2: d2 = d1 + Day + 25
  Case 3: d2 = d1 + Day + flag + 53
  Case 4: d2 = d1 + Day + flag + 84
  Case 5: d2 = d1 + Day + flag + 114
  Case 6: d2 = d1 + Day + flag + 145
  Case 7: d2 = d1 + Day + flag + 175
  Case 8: d2 = d1 + Day + flag + 206
  Case 9: d2 = d1 + Day + flag + 237
  Case 10: d2 = d1 + Day + flag + 267
  Case 11: d2 = d1 + Day + flag + 298
  Case 12: d2 = d1 + Day + flag + 328
End Select

Z = d2 \ 7
M = ((d2 - 7 * Z) * 24 + Hour) * CDbl(3600) + Min * 60 + Sec
GPSTIME = M

End Function

<div align="center">

## DaysWorks


</div>

### Description

WorkingDays counts days except Weekends and Holidays
 
### More Info
 
WorkingDays()

for call

x=WorkingDays("dd/mm/yy", "dd/mm/yy", xArray())

where xArray contents holidays "dd/mm"

integer WorkingDays


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[eduardo Alvarez Bastida](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eduardo-alvarez-bastida.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eduardo-alvarez-bastida-daysworks__1-9393/archive/master.zip)





### Source Code

```
Put this in a CommandButton
'
Dim aH(8)
aH(1) = "1/1"
aH(2) = "5/2"
aH(3) = "21/3"
aH(4) = "1/5"
aH(5) = "5/5"
aH(6) = "16/9"
aH(7) = "20/10"
aH(8) = "25/12"
debug.print = WorkingDays("01/01/00", "01/01/01", aH())
'
Public Function WorkingDays(dBeginDate As Date, dEndDate As Date, ByRef aHolidays As Variant) As Integer
  Dim intTotalDays As Integer
  Dim intHoliday As Integer
  Dim booWeekend As Boolean
  Dim intSatSun As Integer
  Dim strCDayMonth As String
  Dim strNDayMonth As String
  Dim i As Integer
  Dim dNewDate As Date
  If dBeginDate>=dEndDate then exit Function
  intTotalDays = DateDiff("d", dBeginDate, dEndDate)
  For i = 1 To intTotalDays
    dNewDate = DateAdd("d", i, dBeginDate)
    If isWeekEnd(dNewDate) Then
      booWeekend = True
    Else
      booWeekend = False
    End If
    strNDayMonth = Day(dNewDate) & "/" & Month(dNewDate)
    For n = 1 To UBound(aHolidays)
'      strMonth = Mid(aHolidays(h), istr("/", aHolidays(h)) + 1)
      If (strNDayMonth = aHolidays(n)) And Not booWeekend Then
        intHoliday = intHoliday + 1
        booWeekend = False
        Exit For
      End If
    Next n
    If booWeekend Then
      intSatSun = intSatSun + 1
    End If
  Next i
  WorkingDays = intTotalDays - intSatSun - intHoliday
End Function
Private Function isWeekEnd(ByRef dCheck As Date) As Boolean
  If DatePart("w", dCheck) = 1 Or DatePart("w", dCheck) = 7 Then isWeekEnd = True
End Function
```


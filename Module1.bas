Attribute VB_Name = "Module1"
Option Compare Database

Public Function ISOWeekNum(AnyDate As Date, Optional WhichFormat As Variant) As Integer
' WhichFormat: missing or <> 2 then returns week number,
'                                = 2 then YYWW
'
Dim ThisYear As Integer
Dim PreviousYearStart As Date
Dim ThisYearStart As Date
Dim NextYearStart As Date
Dim YearNum As Integer

ThisYear = Year(AnyDate)
ThisYearStart = YearStart(ThisYear)
PreviousYearStart = YearStart(ThisYear - 1)
NextYearStart = YearStart(ThisYear + 1)
Select Case AnyDate
    Case Is >= NextYearStart
        ISOWeekNum = (AnyDate - NextYearStart) \ 7 + 1
        YearNum = Year(AnyDate) + 1
    Case Is < ThisYearStart
        ISOWeekNum = (AnyDate - PreviousYearStart) \ 7 + 1
        YearNum = Year(AnyDate) - 1
    Case Else
        ISOWeekNum = (AnyDate - ThisYearStart) \ 7 + 1
        YearNum = Year(AnyDate)
End Select

If IsMissing(WhichFormat) Then Exit Function
If WhichFormat = 2 Then
    ISOWeekNum = CInt(Format(Right(YearNum, 2), "00") & _
    Format(ISOWeekNum, "00"))
End If

End Function

Public Function YearStart(WhichYear As Integer) As Date

Dim WeekDay As Integer
Dim NewYear As Date

NewYear = DateSerial(WhichYear, 1, 1)
WeekDay = (NewYear - 2) Mod 7 'Generate weekday index where Monday = 0

If WeekDay < 4 Then
    YearStart = NewYear - WeekDay
Else
    YearStart = NewYear - WeekDay + 7
End If

End Function

Public Function WeekStart(WhichWeek As Integer, WhichYear As _
                    Integer) As Date

WeekStart = YearStart(WhichYear) + ((WhichWeek - 1) * 7)

End Function

Private Sub test()
Dim gaas As String
gaas = """"
MsgBox (gaas)

End Sub

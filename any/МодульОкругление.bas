Attribute VB_Name = "МодульОкругление"
Option Explicit

Function Округлить(ByVal Число As Double, Optional ByVal ЧиселПослеТочки As Integer = 0) As Double
    Округлить = Format(Abs(Число), "0." & Replace(Space(ЧиселПослеТочки), " ", "0")) * Sgn(Число)
End Function

Function ОкруглитьВверх(ByVal Число As Double, Optional ByVal ЧиселПослеТочки As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(ЧиселПослеТочки), " ", "0") & "5")
    If Число = Округлить(Число, ЧиселПослеТочки) Then
        ОкруглитьВверх = Число
    Else
        ОкруглитьВверх = Округлить(Число + s, ЧиселПослеТочки)
    End If
End Function

Function ОкруглитьВниз(ByVal Число As Double, Optional ByVal ЧиселПослеТочки As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(ЧиселПослеТочки), " ", "0") & "5")
    ОкруглитьВниз = Округлить(Abs(Число) - s, ЧиселПослеТочки) * Sgn(Число)
End Function


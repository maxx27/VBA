Attribute VB_Name = "ћодульќкругление"
Option Explicit

Function ќкруглить(ByVal „исло As Double, Optional ByVal „иселѕосле“очки As Integer = 0) As Double
    ќкруглить = Format(Abs(„исло), "0." & Replace(Space(„иселѕосле“очки), " ", "0")) * Sgn(„исло)
End Function

Function ќкруглить¬верх(ByVal „исло As Double, Optional ByVal „иселѕосле“очки As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(„иселѕосле“очки), " ", "0") & "5")
    If „исло = ќкруглить(„исло, „иселѕосле“очки) Then
        ќкруглить¬верх = „исло
    Else
        ќкруглить¬верх = ќкруглить(„исло + s, „иселѕосле“очки)
    End If
End Function

Function ќкруглить¬низ(ByVal „исло As Double, Optional ByVal „иселѕосле“очки As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(„иселѕосле“очки), " ", "0") & "5")
    ќкруглить¬низ = ќкруглить(Abs(„исло) - s, „иселѕосле“очки) * Sgn(„исло)
End Function


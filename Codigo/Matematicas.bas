Attribute VB_Name = "Matematicas"
Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) / 100
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)
End Function

Function Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Double
    Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
End Function

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Long
    Randomize GetTickCount()
    RandomNumber = Int((UpperBound - LowerBound + 1) * Rnd) + LowerBound
    If RandomNumber > UpperBound Then RandomNumber = UpperBound
End Function

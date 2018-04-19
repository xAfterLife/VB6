Attribute VB_Name = "GerRandomNumber"
Option Explicit

Public Function Random(MinValue As Integer, MaxValue As Integer) As Integer
Randomize
Random = Int((MaxValue - MinValue + 1) * Rnd + MinValue)
End Function

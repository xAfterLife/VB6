Attribute VB_Name = "GetRandomNumber"
Option Explicit

Public Function Random(MinValue As Integer, MaxValue As Integer) As Integer
Randomize
Random = Int((MaxValue - MinValue) * Rnd + MinValue)
End Function

Attribute VB_Name = "basGUID"
Option Explicit

Private Type GUID
    PartOne As Long
    PartTwo As Integer
    PartThree As Integer
    PartFour(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
                                      (ptrGuid As GUID) As Long
Attribute CoCreateGuid.VB_UserMemId = 1610809345

Public Function GUID() As String
    Dim lRetVal As Long
    Dim udtGuid As GUID
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String

    sAns = ""

    lRetVal = CoCreateGuid(udtGuid)

    If lRetVal = 0 Then
        sPartOne = Hex$(udtGuid.PartOne)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.PartOne)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") _
                 & Trim$(sPartOne)

        sPartTwo = Hex$(udtGuid.PartTwo)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.PartTwo)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") _
                 & Trim$(sPartTwo)

        sPartThree = Hex$(udtGuid.PartThree)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.PartThree)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") _
                   & Trim$(sPartThree)

        For iCtr = 0 To 7
            sPartFour = sPartFour & _
                        Format$(Hex$(udtGuid.PartFour(iCtr)), "00")
        Next
        sAns = sPartOne & "-" & sPartTwo & "-" & sPartThree _
             & "-" & sPartFour
    End If

    GUID = sAns

    Exit Function
End Function

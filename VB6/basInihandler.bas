Attribute VB_Name = "Modul_Inihandler"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" _
                                                   Alias "WritePrivateProfileStringA" _
                                                   (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long
Attribute WritePrivateProfileString.VB_UserMemId = 1073741825

Private Declare Function GetPrivateProfileString Lib "kernel32" _
                                                 Alias "GetPrivateProfileStringA" _
                                                 (ByVal lpApplicationName As String, _
                                                  ByVal lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long
Attribute GetPrivateProfileString.VB_UserMemId = 1610809344

Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
Attribute INIWrite.VB_UserMemId = 1073938437

10        Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
20        INIWrite = (Err.Number = 0)
End Function

Public Function INIRead(sSection As String, sKeyName As String, sINIFileName As String) As String
          Dim sRet As String

10        sRet = String(255, Chr(0))
20        INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function

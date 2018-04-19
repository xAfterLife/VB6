Attribute VB_Name = "CheckWord"
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lptype As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const WM_QUIT As Long = &H12
Private Const ERROR_SUCCESS = 0

Private Function GetRegString(hKey As Long, strSubKey As String, strValueName As String) As String

    Dim strSetting As String
    Dim lngDataLen As Long
    Dim lngRes As Long

    If RegOpenKey(hKey, strSubKey, lngRes) = ERROR_SUCCESS Then
        strSetting = Space(255)
        lngDataLen = Len(strSetting)
        If RegQueryValueEx(lngRes, strValueName, ByVal 0, REG_EXPAND_SZ, ByVal strSetting, lngDataLen) = ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                GetRegString = Left(strSetting, lngDataLen - 1)
            End If
        End If
        If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
            MsgBox "RegCloseKey Failed: " & strSubKey, vbCritical
        End If
    End If

End Function

Private Function IsAppPresent(strSubKey$, strValueName$) As Boolean
Attribute IsAppPresent.VB_UserMemId = 1677721603
    IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)))
End Function

Public Sub CheckWord()
    If IsAppPresent("Word.Document\CurVer", "") = False Then
        MsgBox "Please Install MS Office", vbCritical, "Exam Finder Error"
    Else
        MsgBox "Word Is Installed!"
    End If
End Sub

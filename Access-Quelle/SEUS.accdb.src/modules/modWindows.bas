Attribute VB_Name = "modWindows"
Option Compare Database
Option Explicit

Dim OldWinUserName As String

Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" _
                                 Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                                       nSize As Long) As Long

Function mWin_UserName() As String
    If OldWinUserName <> "" Then
        mWin_UserName = OldWinUserName
        Exit Function
    End If

    Dim lngLen As Long, lngX As Long
    Dim strUserName As String

    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If lngX <> 0 Then
        mWin_UserName = Left$(strUserName, lngLen - 1)
    Else
        mWin_UserName = ""
    End If

    OldWinUserName = mWin_UserName
End Function

Function User()
    MsgBox "Zur Zeit angemeldet ist: " & mWin_UserName
End Function

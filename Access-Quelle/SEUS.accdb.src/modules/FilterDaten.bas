Attribute VB_Name = "FilterDaten"
Option Compare Database
Option Explicit

Dim FlagID As Long

Sub mFD_SetFlagID(NewID As Long)
    FlagID = NewID
End Sub

Function mFD_GetFlagID() As Long
    mFD_GetFlagID = FlagID
End Function

Sub myString(strInput As String)
    If strInput <> "" Then
        strInput = strInput + "Hier kommen" & _
                   "die weiteren" & _
                   ""

        MsgBox (strInput)
    End If
End Sub

Attribute VB_Name = "BerichtsDruck"
Option Compare Database
Option Explicit

Dim mBR_ID As Long
Dim mBR_ShowLogo As Boolean
Dim mBR_KarteiID As Long, mBR_Datum As Date

Sub mBR_SetKarteiID(ID As Long)
    mBR_KarteiID = ID
End Sub

Function mBR_GetKarteiID() As Long
    mBR_GetKarteiID = mBR_KarteiID
End Function

Sub mBR_SetDatum(Datum As Date)
    mBR_Datum = Datum
End Sub

Function mBR_GetDatum() As Date
    mBR_GetDatum = mBR_Datum
End Function

Sub mBR_SetID(ID As Long)
    mBR_ID = ID
End Sub

Function mBR_GetID() As Long
    mBR_GetID = mBR_ID
End Function

Sub mBR_SetShowLogo(SetBoolean As Boolean)
    mBR_ShowLogo = SetBoolean
End Sub

Function mBR_GetShowLogo() As Boolean
    mBR_GetShowLogo = mBR_ShowLogo
End Function

Function mBR_Mahntext(MahnStufe As Integer) As String
    mBR_Mahntext = mED_GetMemo("MahnText" & MahnStufe)
End Function

Function mBR_MahnFussText(MahnDatum As Date, LetzterBankbelegDatum As String) As String
    Dim txt As String

    txt = mED_GetMemo("MahnFussText")

    txt = Replace(txt, "<Mahndatum+14>", Format(MahnDatum + 7, "DD.MM.YYYY"))
    txt = Replace(txt, "<LetzterBankbelegDatum>", Format(LetzterBankbelegDatum, "DD.MM.YYYY"))

    mBR_MahnFussText = txt
End Function

Function mBR_RechungsFussText() As String
    mBR_RechungsFussText = mED_GetMemo("RechnungsFussText")
End Function

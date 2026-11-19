Attribute VB_Name = "Nummern"
Option Compare Database
Option Explicit

Function mNr_NewNr(Feldname As String, Optional Datum As Date, Optional StellenLaufzahl As Long) As String
    Dim MaxNr As Long, Jahr As Long, NewNr As Long, TabName As String, Multi As Long, Form As String, i As Long


    If StellenLaufzahl = 0 Then StellenLaufzahl = 4
    Multi = 1
    For i = 1 To StellenLaufzahl
        Multi = Multi * 10
        Form = Form & "0"
    Next

    If Datum = 0 Then Datum = Date
    Jahr = Year(Datum)                           ' * 100 + Month(Datum)


    '** Umstellung auf Max-Suche in Tabelle **
    If Feldname = "LieferscheinNr" Then TabName = "T_Lieferscheine"
    If Feldname = "RechnungNr" Then TabName = "T_Rechnungen"

    NewNr = (Nz(DMax("Nr", TabName, "Year(datum)=" & Format(Year(Datum))), 0) + 1) Mod Multi

    mNr_NewNr = Format((Year(Datum) Mod 100), "00") & Format(NewNr, Form)
End Function

Sub testmNr_NewNr()
    MsgBox mNr_NewNr("RechnungNr", Date)
End Sub

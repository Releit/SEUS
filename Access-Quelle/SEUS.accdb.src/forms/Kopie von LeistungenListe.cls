Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim LstFilter As String

Function pv_Filter()
    Dim txt As String

    LstFilter = ""

    If Not IsNull(cmbJahrMonat) Then
        If IsNull(cmbJahrMonatBis) Then
            txt = "year([Datum])*100+month([Datum])=" & Str(cmbJahrMonat)
            LstFilter = "year([Datum])*100+month([Datum])=" & Format(cmbJahrMonat)
        Else
            txt = "year([Datum])*100+month([Datum]) BETWEEN " & Str(cmbJahrMonat) & " AND " & Str(cmbJahrMonatBis)
            LstFilter = "year([Datum])*100+month([Datum]) BETWEEN " & Str(cmbJahrMonat) & " AND " & Str(cmbJahrMonatBis)
        End If
    End If

    If Not IsNull(SuchKunde) Then
        txt = txt & IIf(txt = "", "", " AND ") & "T_Leistungen.KarteiID= " & Str(SuchKunde)
        LstFilter = LstFilter & IIf(LstFilter = "", "", " AND ") & "T_Leistungen.KarteiID= " & Str(SuchKunde)
    End If

    If Not IsNull(SuchProjektID) Then
        txt = txt & IIf(txt = "", "", " AND ") & "T_Leistungen.ProjektID= " & Str(SuchProjektID)
        LstFilter = LstFilter & IIf(LstFilter = "", "", " AND ") & "T_Leistungen.ProjektID= " & Str(SuchProjektID)
    End If

    If Not IsNull(cmbMitarbeiterID) Then
        txt = txt & IIf(txt = "", "", " AND ") & "T_Leistungen.MitarbeiterID= " & Str(cmbMitarbeiterID)
        LstFilter = LstFilter & IIf(LstFilter = "", "", " AND ") & "T_Leistungen.MitarbeiterID= " & Str(cmbMitarbeiterID)
    End If

    If optVerrechnet = 2 Then                    '** nur unverrechnete
        txt = txt & IIf(txt = "", "", " AND ") & "AufRechnungID IS NULL"
        LstFilter = LstFilter & IIf(LstFilter = "", "", " AND ") & "AufRechnungID IS NULL"
    End If

    UF.Form.Filter = txt
    If txt <> "" Then
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If

    'MsgBox LstFilter

    pv_GesamtSumme
End Function

Sub pv_GesamtSumme()
    If UF.Form.FilterOn Then
        GesamtExkl = DSum("Menge*VKExkl", "T_Leistungen", LstFilter)
    Else
        GesamtExkl = Null
    End If
End Sub

Private Sub Befehl41_Click()
    DoCmd.OpenReport "Leistungsverzeichnis", acViewPreview, , UF.Form.Filter
End Sub

Private Sub cmbJahrMonat_AfterUpdate()
    SuchKunde.Requery
    SuchKunde = Null
    SuchProjektID.Requery
    cmbMitarbeiterID.Requery
    pv_Filter
End Sub

Private Sub cmbJahrMonatBis_AfterUpdate()
    SuchKunde.Requery
    SuchKunde = Null
    SuchProjektID.Requery
    cmbMitarbeiterID.Requery
    pv_Filter
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdRequery_Click()
    '** Leistungen auf gelöschten Rechnungen korrigieren **
    CurrentDb.Execute "UPDATE T_Leistungen LEFT JOIN T_Rechnungen ON T_Leistungen.AufRechnungID = T_Rechnungen.ID SET T_Leistungen.AufRechnungID = Null" & _
                      " WHERE AufRechnungID Is Not Null AND T_Rechnungen.ID Is Null"

    UF.Requery
    pv_GesamtSumme
End Sub

Private Sub cmdVerrechnen_Click()
    Dim NewID As Long

    NewID = mVR_FillVerrechnungKunden(UF.Form.Filter)

    If NewID Then DoCmd.OpenForm "Verrechnung", , "ID=" & Str(NewID)
End Sub

Private Sub Form_Load()

    cmbJahrMonat = Year(Date) * 100 + Month(Date)
    pv_Filter
End Sub

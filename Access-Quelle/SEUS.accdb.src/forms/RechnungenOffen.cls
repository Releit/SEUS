Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim InitSummen As Boolean

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdRequery_Click()
    DoCmd.Hourglass True
    If Register1 = 0 Then UF.Requery             '** Liste **
    If Register1 = 1 Then pv_Summen              '** offen Gesamtbeträge **
    DoCmd.Hourglass False
End Sub

Private Sub Form_Load()
    BetragMahnung1.Visible = mED_GetLong("MahnAnzahl") > 0
    BetragMahnung2.Visible = mED_GetLong("MahnAnzahl") > 1
    BetragMahnung3.Visible = mED_GetLong("MahnAnzahl") > 2
End Sub

Private Sub Register1_Change()
    If Register1 = 1 And Not InitSummen Then
        pv_Summen
        InitSummen = True
    End If
End Sub

Sub pv_Summen()
    DoCmd.Hourglass True
    BetragGesamt = DSum("Saldo", "V_RechnungenMitSaldo")
    If Not IsNull(Tage) Then BetragÄlterAlsTage = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "Datum<" & CDate2SQL(Date - Val(Tage))), 0) Else BetragÄlterAlsTage = Null
    If mED_GetLong("MahnAnzahl") > 0 Then BetragMahnung1 = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "Mahnstufe=1"), 0)
    If mED_GetLong("MahnAnzahl") > 1 Then BetragMahnung2 = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "Mahnstufe=2"), 0)
    If mED_GetLong("MahnAnzahl") > 2 Then BetragMahnung3 = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "Mahnstufe=3"), 0)
    BetragRA = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "NOT RechtsanwaltAm IS NULL"), 0)
    DoCmd.Hourglass False
End Sub

Private Sub Tage_AfterUpdate()
    If Not IsNull(Tage) Then
        DoCmd.Hourglass True
        BetragÄlterAlsTage = Nz(DSum("Saldo", "V_RechnungenMitSaldo", "Datum<" & CDate2SQL(Date - Val(Tage))), 0)
        DoCmd.Hourglass False
    Else
        BetragÄlterAlsTage = Null
    End If

End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Betrag_AfterUpdate()
    If IsNull(Netto) Then
        If Not IsNull(Betrag) And Not IsNull(UstBetrag) Then
            Netto = Betrag - UstBetrag
        Else
            If Not IsNull(Betrag) And Not IsNull(Ust10Proz) And Not IsNull(Ust20Proz) Then
                Netto = Betrag - Ust10Proz - Ust20Proz
            End If
        End If
    End If
End Sub

Private Sub cmdDateiAuswählen_Click()
    Dim sPath As String

    sPath = mED_ER_Pfad(Nz(Datum, Date))


    Set DateiAuswahlStartForm = Me

    DoCmd.OpenForm "Dateiauswahl"
    With Forms.Dateiauswahl
        !SearchFolder = sPath
        .pv_ReadFolder sPath
    End With

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If InnerGemErwerb And IsNull(UstProz) Then
        MsgBox "Bei einem Innergemeinschaftlichen Erwerb ist der Ust%-Satz für die USt- Berechnung/-Voranmeldung erforderlich!", , "fehlende Eingabe"
        Cancel = True
    End If
End Sub

Private Sub Form_Current()
    LstZA.Requery
End Sub

Private Sub InnerGemErwerb_AfterUpdate()
    If InnerGemErwerb Then
        UstBetrag = 0
        If IsNull(UstProz) Then MsgBox "Bitte auch den regulär anzuwendenten Ust%-Satz links eingeben!"
    End If
End Sub

Private Sub LiefKarteiID_AfterUpdate()
    LstZA.Requery
End Sub

Private Sub LstZA_DblClick(Cancel As Integer)
    Dim Sql As String, Betrag As Currency

    If Brutto - Nz(Skonto, 0) < Val(LstZA.Column(2)) Then
        Betrag = Brutto - Nz(Skonto, 0)
    Else
        Betrag = LstZA.Column(2)
    End If

    Sql = "INSERT INTO T_EingangsRechnungen_ZA(ZAID, ERID, Betrag) VALUES(" & Str(LstZA) & "," & Str(Me!ID) & "," & Str(Betrag) & ")"
    CurrentDb.Execute Sql

    UF_ZE.Requery

    BetragBezahlt = mDF_ER_BetragBezahlt(Me!ID)
    BezahltAm = mDF_ER_BezahltAm(Me!ID)
End Sub

Private Sub cmdSkonto_Click()
    Dim Sql As String

    If IsNull(BetragBezahlt) Then BetragBezahlt = mDF_ER_BetragBezahlt(Me!ID)
    If Nz(DCount("ID", "T_EingangsRechnungen_ZA", "ERID=" & Str(Me!ID)), 0) = 1 Then
        Sql = "UPDATE T_EingangsRechnungen_ZA SET Betrag=mDF_ZABetrag(ZAID) WHERE ERID=" & Str(Me!ID)
        'Debug.Print Sql
        CurrentDb.Execute Sql
    End If
    If Abs((BetragBezahlt / Betrag - 1)) < 0.05 Then Skonto = Betrag - BetragBezahlt 'maximal 5 % Skonto automatisch

End Sub

Private Sub PfadDatei_DblClick(Cancel As Integer)
    Dim FName As String
    If Not IsNull(PfadDatei) Then
        FName = mED_ER_Pfad(Nz(Datum, Date)) & "\" & PfadDatei
        Debug.Print FName
        mFM_OpenFile FName
    End If
End Sub

Private Sub Ust10Proz_AfterUpdate()
    Betrag_AfterUpdate
End Sub

Private Sub Ust20Proz_AfterUpdate()
    Betrag_AfterUpdate
End Sub

Private Sub UstBetrag_AfterUpdate()

    If IsNull(UstBetrag) Then UstBetrag = 0

    If Not IsNull(Betrag) And Not InnerGemErwerb Then
        UstProz = UstBetrag / (Betrag - UstBetrag)
    End If
End Sub

Private Sub UstBetrag_BeforeUpdate(Cancel As Integer)
    If InnerGemErwerb And Nz(UstBetrag, 0) <> 0 Then
        MsgBox "Bei einem Innergemeinschaftlichen Erwerb muß der Ust-Betrag der Rechnung 0 sein!", , ""
        Cancel = True
    End If
End Sub

Private Sub UStProz_AfterUpdate()
    If Not IsNull(UstProz) And UstProz > 1 Then UstProz = UstProz / 100
    If Not IsNull(UstProz) And Not InnerGemErwerb Then UstBetrag = Round(Betrag - Betrag / (1 + UstProz), 2)

    If UstProz = 0.2 And Not IsNull(Betrag) Then
        Ust10Proz = 0
        Ust20Proz = UstBetrag
        Netto = Betrag - UstBetrag
    End If

    If UstProz = 0.1 And Not IsNull(Betrag) Then
        Ust10Proz = UstBetrag
        Ust20Proz = 0
        Netto = Betrag - UstBetrag
    End If

End Sub

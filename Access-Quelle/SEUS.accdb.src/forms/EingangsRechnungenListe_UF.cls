Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    Dim FormName As String

    FormName = "EingangsRechnungen"

    DoCmd.OpenForm FormName, , , , , acHidden
    With Forms(FormName)
        .RecordsetClone.FindFirst "ID=" & Str(Me!ID)
        If Not .RecordsetClone.EOF Then
            .Recordset.Bookmark = .RecordsetClone.Bookmark
            .Visible = True
        Else
            MsgBox "Datensatz nicht gefunden!", , FormName
            DoCmd.Close acForm, FormName
        End If
    End With
End Function

Private Sub Brutto_AfterUpdate()
    If Not IsNull(Brutto) Then
        If IsNull(Netto) And Not IsNull(UstProz) Then
            Netto = Brutto / (1 + UstProz)
            If IsNull(Ust20Proz) Then Ust20Proz = Brutto - Netto
        End If

        If IsNull(Netto) And (Not IsNull(Ust20Proz) Or Not IsNull(Ust10Proz)) Then
            Netto = Brutto - Nz(Ust20Proz, 0) - Nz(Ust10Proz, 0)
        End If
    End If
End Sub

Private Sub Netto_AfterUpdate()
    If IsNull(Ust20Proz) And Not IsNull(UstProz) Then
        Ust20Proz = Netto * UstProz
        If IsNull(Brutto) Then Brutto = Netto + Ust20Proz
    End If
End Sub

Private Sub PfadDatei_DblClick(Cancel As Integer)
    Dim sPath As String, LfName As String

    If IsNull(PfadDatei) Then

        sPath = mED_ER_Pfad(Nz(BezahltAm, Nz(Datum, Date)))

        Set DateiAuswahlStartForm = Forms.EingangsRechnungenListe!UF.Form

        DoCmd.OpenForm "Dateiauswahl"
        With Forms.Dateiauswahl
            !SearchFolder = sPath
            .pv_ReadFolder sPath
            LfName = cmbLiefKarteiID.Column(1)
            If InStr(LfName, "") > 0 Then
                !FilterTxt = Left(LfName, InStr(LfName, " ") - 1)
            Else
                !FilterTxt = LfName
            End If
            .FilterTxt_AfterUpdate
        End With
    Else
        mFM_OpenFile mED_ER_Pfad(Nz(BezahltAm, Nz(Datum, Date))) & "\" & PfadDatei
    End If
End Sub

Private Sub Ust20Proz_AfterUpdate()
    If Not IsNull(Brutto) Then
        Netto = Brutto - (Nz(Ust20Proz, 0) + Nz(Ust10Proz, 0))
    End If
End Sub

Private Sub UStProz_AfterUpdate()
    If Nz(UstProz, 0) > 1 Then UstProz = UstProz / 100
    Netto = Brutto / (1 + UstProz)
    Ust20Proz = Brutto - Netto

End Sub

Private Sub UstProz_BeforeUpdate(Cancel As Integer)
    If Nz(UstProz, 0) > 20 Then
        MsgBox "max 20%", , "ungültige Eingabe"
        Cancel = True
    End If
End Sub

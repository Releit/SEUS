Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub pv_Summen()
    If Not IsNull(Me!ID) Then
        BetragExkl = Round(DSum("Vkexkl*Menge", "T_RechnungsPos", "RechnungID=" & Str(Me!ID)), 4)
        BetragInkl = Round(DSum("Vkexkl*Menge*(1+UstProz)", "T_RechnungsPos", "RechnungID=" & Str(Me!ID)), 4)
        Saldo = Round(DSum("Betrag", "T_Buchungen", "RechnungID=" & Str(Me!ID)), 4)
    Else
        BetragExkl = 0
        BetragInkl = 0
        Saldo = 0
    End If
End Sub

Private Sub AdressenID_AfterUpdate()
    pv_SetAnschrift
End Sub

Sub pv_SetAnschrift()
    If Not IsNull(AdressenID) Then
        Anschriftname = mDF_AdressAnschriftName(Nz(AdressenID, 0))
        Adresse = DLookup("Adresse", "T_Adressen", "ID=" & Str(Nz(AdressenID, 0)))
        LandCode = DLookup("LandCode", "T_Adressen", "ID=" & Str(Nz(AdressenID, 0)))
        PLZ = DLookup("PLZ", "T_Adressen", "ID=" & Str(Nz(AdressenID, 0)))
        Ort = DLookup("Ort", "T_Adressen", "ID=" & Str(Nz(AdressenID, 0)))
    Else
        Anschriftname = mDF_KarteiAnschriftName(Nz(KarteiID, 0))
        Adresse = DLookup("Adresse", "T_Kartei", "ID=" & Str(Nz(KarteiID, 0)))
        LandCode = DLookup("LandCode", "T_Kartei", "ID=" & Str(Nz(KarteiID, 0)))
        PLZ = DLookup("PLZ", "T_Kartei", "ID=" & Str(Nz(KarteiID, 0)))
        Ort = DLookup("ort", "T_Kartei", "ID=" & Str(Nz(KarteiID, 0)))
    End If
End Sub

Private Sub AdressenID_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", Nz(Me!KarteiID, 0)
End Sub

Private Sub cmdAddAdresse_Click()
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", Nz(Me!KarteiID, 0)
End Sub

Private Sub cmdAddKartei_Click()
    mSU_KarteiAdd = True
    mSU_StartForm = "Rechnungen"
    DoCmd.OpenForm "Kartei", , , "ID=0"
End Sub

Private Sub cmdDel_Click()
    Dim Sql As String, DelID As Long, PrevID As Long

    If Not IsNull(Me!ID) Then
        If Not IsNull(Nr) Then
            MsgBox "Rechnungen mit Nummer können nicht gelöscht werden!", , "Löschen"
            Exit Sub
        End If

        If MsgBox("Datensatz wirklich löschen?", vbYesNo + vbDefaultButton2, "Datensatz löschen") = vbYes Then
            DelID = Me!ID
            PrevID = Nz(DMax("ID", "T_Rechnungen", "ID<" & Str(DelID)), 0)
            If PrevID = 0 Then PrevID = Nz(DMax("ID", "T_Rechnungen", "ID>" & Str(DelID)), 0)

            '** Zuordnung der Leistungen entfernen **"
            Sql = "UPDATE T_Leistungen SET AufRechnungID=NULL WHERE AufRechnungID=" & Str(DelID)
            CurrentDb.Execute Sql

            '** Positionen Löschen **"
            Sql = "DELETE * FROM T_RechnungsPos WHERE RechnungID=" & Str(DelID)
            CurrentDb.Execute Sql

            '** Buchungen Löschen **"
            Sql = "DELETE * FROM T_RechnungsPos WHERE RechnungID=" & Str(DelID)
            CurrentDb.Execute Sql

            If PrevID Then
                Me.Filter = "ID=" & Str(PrevID)
                Me.FilterOn = True
            Else
                DoCmd.Close acForm, Me.Name
            End If

            '** Rechnung Löschen **"
            Sql = "DELETE * FROM T_Rechnungen WHERE ID=" & Str(DelID)
            CurrentDb.Execute Sql
        End If
    End If

End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name

    On Error Resume Next
    Forms.rechnungenliste!UF.Requery
End Sub

Private Sub cmdKarteiSuche_Click()
    mSU_KarteiSuche = True
    mSU_StartForm = "Rechnungen"
    DoCmd.OpenForm "Karteiliste"
End Sub

Private Sub cmdMail_Click()
    Dim eMail As String, BodyText As String, RS As Object

    On Error GoTo cmdMailErr

    mBR_SetID Me!ID
    mBR_SetShowLogo True

    If IsNull(Datum) Then
        If MsgBox("Rechnung buchen?" & vbCrLf & vbCrLf & "Datum und Nr werden eingetragen", vbYesNo, "Belegdruck") = vbYes Then
            Datum = Date
            Datum_AfterUpdate
        Else
            'Exit Sub
        End If
    Else
        If IsNull(Nr) Then
            If MsgBox("Rechnung buchen?" & vbCrLf & vbCrLf & "Nr wird eingetragen", vbYesNo, "Belegdruck") = vbYes Then
                Datum_AfterUpdate
            Else
                'Exit Sub
            End If
        End If
    End If

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    eMail = Nz(DMin("eMail", "T_Kartei", "ID=" & Str(Me!KarteiID)), "")
    If eMail <> "" Then
        Set RS = CurrentDb.OpenRecordset("SELECT * FROM T_Einstellungen", dbOpenSnapshot)
        BodyText = RS!EmailBody1 & vbCrLf & vbCrLf & RS!EmailBody2


        DoCmd.SendObject acSendReport, "Rechnung", acFormatPDF, eMail, , , "Rechnung " & Format(Nr), _
                         BodyText
    Else
        MsgBox "keine eMailadresse beim Kunden eingetragen!", , "eMail senden"
    End If

cmdMailExit:
    Exit Sub

cmdMailErr:
    MsgBox Err.Description
    Resume cmdMailExit

End Sub

Private Sub cmdMakeGutschrift_Click()
    Dim NewID As Long

    If MsgBox("Gutschrift zu dieser Rechnung anlegen?", vbYesNo, "Gutschrift anlegen") = vbYes Then

        If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
        If Not IsNull(Me!ID) Then
            NewID = mUW_RechnungZuGutschrift(Me!ID, False)
            If NewID <> 0 Then
                DoCmd.SearchForRecord , , acFirst, "ID=" & Str(NewID)
            End If
        End If
    End If
End Sub

Private Sub cmdPaketaufkleber_Click()
    mBR_SetID Me!ID

    DoCmd.OpenReport "Paketaufkleber", acViewPreview
End Sub

Private Sub cmdPrint_Click()
    pv_OpenRechnung True
End Sub

Private Sub cmdShow_Click()
    If BetragExkl > 10000 Then
        MsgBox "Bitte die UID des Kunden angeben!", , "UID"
    End If
    pv_OpenRechnung False
End Sub

Sub pv_OpenRechnung(ToPrinter As Boolean)

    mBR_SetID Me!ID
    mBR_SetShowLogo False

    If IsNull(ZahlArtID) Then
        MsgBox "Bitte Zahlart auswählen", , "Fehlende Eingabe"
        Exit Sub
    End If

    If IsNull(Datum) Then
        If MsgBox("Rechnung buchen?" & vbCrLf & vbCrLf & "Datum und Nr werden eingetragen, Betrag wird bebucht", vbYesNo, "Belegdruck") = vbYes Then
            Datum = Date
            Datum_AfterUpdate
        Else
            'Exit Sub
        End If
    Else
        If IsNull(Nr) Then
            If MsgBox("Rechnung buchen?" & vbCrLf & vbCrLf & "Nr wird eingetragen, Betrag wird bebucht", vbYesNo, "Belegdruck") = vbYes Then
                Datum_AfterUpdate
            Else
                'Exit Sub
            End If
        End If
    End If

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    '** gebuchte Rechnungen sperren **
    pv_SetEnabled

    mBR_SetID Me!ID

    If DCount("ID", "T_Leistungen", "AufRechnungID=" & Str(Me!ID)) > 0 Then
        If MsgBox("mit Leistungsverzeichnis?", vbYesNo, "Rechnung öffnen") = vbYes Then
            DoCmd.OpenReport "Rechnung_Mit_Leistungsverz", IIf(ToPrinter, acNormal, acViewPreview)
        Else
            DoCmd.OpenReport "Rechnung", IIf(ToPrinter, acNormal, acViewPreview)
        End If
    Else
        DoCmd.OpenReport "Rechnung", IIf(ToPrinter, acNormal, acViewPreview)
    End If
End Sub

Private Sub Datum_AfterUpdate()
    If Not IsNull(Datum) Then
        If IsNull(Nr) Then
            Nr = mNr_NewNr("RechnungNr", Datum)
        End If
        mUW_RechnungBuchen Me!ID, Datum
        UF_Buchungen.Requery
        pv_Summen
        mBR_SetDatum Datum
    End If
End Sub

Private Sub Datum_DblClick(Cancel As Integer)
    If IsNull(Datum) Then Datum = Date
End Sub

Private Sub Form_Activate()
    AdressenID.Requery
End Sub

Private Sub Form_Current()
    pv_Summen

    '** gebuchte Rechnungen sperren **
    pv_SetEnabled

    ZuHandenKontPersID.Requery

    mBR_SetKarteiID Nz(KarteiID, 0)
    mBR_SetDatum Nz(Datum, 0)
    UF_ZE.Form!ZEID.Requery


End Sub

Sub pv_SetEnabled()
    '** gebuchte Rechnungen sperren **
    KarteiID.Locked = Not IsNull(Nr)
    Datum.Enabled = IsNull(Nr)
    cmdKarteiSuche.Enabled = IsNull(Nr)
    cmdAddKartei.Enabled = IsNull(Nr)
    UF_Pos.Enabled = IsNull(Nr)
    Nr.Locked = Not IsNull(Nr)
End Sub

Private Sub KarteiID_AfterUpdate()
    AdressenID = Null
    AdressenID.Requery
    Me!Kunde = KarteiID.Column(1)
    ZuHandenKontPersID.Requery
    mBR_SetKarteiID Nz(KarteiID, 0)
    pv_SetAnschrift
End Sub

Private Sub KarteiID_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", Nz(Me!KarteiID, 0)
End Sub

Private Sub KarteiID_NotInList(NewData As String, Response As Integer)
    Dim ctl As Control

    Set ctl = Me!KarteiID

    If MsgBox("Soll ihre Eingabe als neuer Kunde angelegt werden?", vbYesNo, "Kunde nicht vorhanden") = vbYes Then
        '** neue anlegen **
        CurrentDb.Execute "INSERT INTO T_Kartei(Name1) VALUES('" & NewData & "')"
        m_GoToID "Kartei", DMax("ID", "T_Kartei")
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        ctl.Undo
    End If
End Sub

Private Sub Nr_DblClick(Cancel As Integer)
    If Not IsNull(Datum) Or Not IsNull(Nr) Then
        If MsgBox("Rechungsnummer und -datum löschen?", vbYesNo + vbDefaultButton2, "Buchung aufheben") = vbYes Then
            Nr = Null
            Datum = Null

            '** eventuelle alte Buchung löschen
            CurrentDb.Execute "DELETE * FROM T_Buchungen WHERE Buchungsart='Rechnung' AND RechnungID=" & Str(Me!ID)
            UF_Buchungen.Requery

            pv_SetEnabled
            pv_Summen
        End If
    End If
End Sub

Private Sub ZuHandenKontPersID_AfterUpdate()
    If Not IsNull(ZuHandenKontPersID) Then
        Me!ZuHanden = ZuHandenKontPersID.Column(1)
    Else
        Me!ZuHanden = Null
    End If
End Sub

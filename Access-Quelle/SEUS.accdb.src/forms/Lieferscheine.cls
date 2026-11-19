Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub pv_Summen()
    If Not IsNull(Me!ID) Then
        BetragExkl = Round(DSum("Vkexkl*Menge", "T_LieferscheinPos", "LieferscheinID=" & Str(Me!ID)), 4)
        BetragInkl = Round(DSum("Vkexkl*Menge*(1+UstProz)", "T_LieferscheinPos", "LieferscheinID=" & Str(Me!ID)), 4)
    Else
        BetragExkl = 0
        BetragInkl = 0
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
    mSU_StartForm = "Lieferscheine"
    DoCmd.OpenForm "Kartei", , , "ID=0"
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name

    On Error Resume Next
    Forms.Lieferscheineliste!UF.Requery
End Sub

Private Sub cmdKarteiSuche_Click()
    mSU_KarteiSuche = True
    mSU_StartForm = "Lieferscheine"
    DoCmd.OpenForm "Karteiliste"
End Sub

Private Sub cmdMail_Click()
    Dim eMail As String, BodyText As String, rs As Object

    On Error GoTo cmdMailErr

    mBR_SetID Me!ID
    mBR_SetShowLogo True

    If IsNull(Datum) Then
        If MsgBox("Datum und Nr eingetragen?", vbYesNo, "Belegdruck") = vbYes Then
            Datum = Date
            Datum_AfterUpdate
        Else
            'Exit Sub
        End If
    Else
        If IsNull(Nr) Then
            If MsgBox("Nr eingetragen?" & vbCrLf & vbCrLf & "Nr wird eingetragen", vbYesNo, "Belegdruck") = vbYes Then
                Datum_AfterUpdate
            Else
                'Exit Sub
            End If
        End If
    End If

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    eMail = Nz(DMin("eMail", "T_Kartei", "ID=" & Str(Me!KarteiID)), "")
    If eMail <> "" Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM T_Einstellungen", dbOpenSnapshot)
        BodyText = rs!EmailBody1 & vbCrLf & vbCrLf & rs!EmailBody2


        DoCmd.SendObject acSendReport, "Lieferschein", acFormatPDF, eMail, , , "Lieferschein " & Format(Nr), _
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

Private Sub cmdMakeRechnung_Click()
    Dim NewID As Long

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    If Not IsNull(Me!ID) And Not IsNull(Me!KarteiID) Then
        '** Prüfen ob noch ungebuchte Rechnungen vorhanden sind **
        If DCount("ID", "T_Rechnungen", "Nr IS NULL AND KarteiID=" & Str(Me!KarteiID)) > 0 Then
            m_GoToID "Rechnungsauswahl", Me.KarteiID
            Exit Sub
        End If

        NewID = mUW_LiefersZuRechnung(Me!ID, Me!KarteiID)
        If NewID <> 0 Then
            UF_Pos.Requery
            m_GoToID "Rechnungen", NewID
            DoCmd.Close acForm, Me.Name
        End If

    Else
        MsgBox "Bitte zuerst einen Kunden auswählen", , "fehlende Eingabe"
    End If

End Sub

Private Sub cmdPrint_Click()
    pv_OpenBeleg True
End Sub

Private Sub cmdShow_Click()
    pv_OpenBeleg False
End Sub

Sub pv_OpenBeleg(ToPrinter As Boolean)

    mBR_SetID Me!ID
    mBR_SetShowLogo False


    If IsNull(Datum) Then
        If MsgBox("Datum und Nr eingetragen?", vbYesNo, "Belegdruck") = vbYes Then
            Datum = Date
            Datum_AfterUpdate
        Else
            'Exit Sub
        End If
    Else
        If IsNull(Nr) Then
            If MsgBox("Nr eingetragen?" & vbCrLf & vbCrLf & "Nr wird eingetragen", vbYesNo, "Belegdruck") = vbYes Then
                Datum_AfterUpdate
            Else
                'Exit Sub
            End If
        End If
    End If

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    DoCmd.OpenReport "Lieferschein", IIf(ToPrinter, acNormal, acViewPreview)
End Sub

Private Sub Datum_AfterUpdate()
    If Not IsNull(Datum) Then
        If IsNull(Nr) Then
            Nr = mNr_NewNr("LieferscheinNr", Datum)
        End If
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
End Sub

Private Sub KarteiID_AfterUpdate()
    Dim txt As String

    If Not IsNull(KarteiID) Then
        txt = Nz(DMin("Anmerkungen", "T_Kartei", "ID=" & Str(KarteiID)), "")
        If txt <> "" Then
            MsgBox txt, , "Anmerkung zum Kunden"
        End If
    End If

    AdressenID.Requery


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

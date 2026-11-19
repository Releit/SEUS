Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbOrtAusPLZ_AfterUpdate()
    If Not IsNull(cmbOrtAusPLZ) Then Ort = cmbOrtAusPLZ.Column(1)
End Sub

Private Sub cmbProjektID_AfterUpdate()
    pv_SetFilter
End Sub

Private Sub cmdCheckUID_Click()
    Dim UIDLZ As String, UID As String, LaenderID As Long

    LaenderID = Nz(DMax("LaenderID", "T_FirmenStandorte", "Hauptsitz=TRUE AND FirmenID=" & Str(Nz(Me!FIRID, 0))), 0)

    If LaenderID <> 0 Then

        UIDLZ = Nz(DLookup("UIDLZ", "Laender", "AID=" & Str(LaenderID)), "")

        If Not IsNull(Me!UID) Then
            UID = Right(UID, Len(Me!UID) - Len(UIDLZ))
            'UIDCheck = True
            'UIDCheckAm = Date
            'UIDCheckVon = mWin_UserName()
            If UIDLZ <> "" Then
                'If mUID_Check(UIDLZ, UID) Then
                    'UIDGueltig = True
                'Else
                    'UIDGueltig = False
                'End If
                'UIDName = Left(mUID_Name(), 100)
                'UIDAddress = Left(mUID_Address(), 100)
            Else
                MsgBox "Nur Länder mit einem UIDLZ-Eintrag in der Ländertabelle können auf eine gültige UID geprüft werden!", , "UID-Check"
                'UIDGueltig = True
                'UIDName = "keine Überprüfung möglich!"
            End If
        Else

        End If
    Else
        MsgBox "Eine UID-Nummern Überprüfung ist nur möglich, wenn dem Hauptsitz ein Land zugeordnet ist!", , "fehlende Eingabe"
    End If
End Sub

Private Sub cmdClose_Click()

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    If mSU_KarteiAdd And Not IsNull(Me!ID) Then
        Forms(mSU_StartForm).KarteiID.Requery
        Forms(mSU_StartForm).KarteiID = Me!ID
        mSU_KarteiSuche = False
        mSU_KarteiAdd = False
    End If

    DoCmd.Close acForm, Me.Name

    On Error Resume Next
    Forms.Karteiliste!UF.Requery
End Sub

Private Sub cmdNewNr_Click()
    If IsNull(Nr) Then
        Nr = Nz(DMax("Nr", "T_Kartei"), 0) + 1
    End If
End Sub

Private Sub cmdPaketaufkleber_Click()
    mBR_SetID Me!ID

    DoCmd.OpenReport "Paketaufkleber_Kartei", acViewPreview
End Sub

Private Sub Form_Current()
    cmbJahrMonat.Requery
End Sub

Private Sub Form_Load()
    Schweißmeister.Visible = mED_GetBoolean("Schweißprüfungen")
    Register1.Pages("Projekte").Visible = mED_GetBoolean("Projekte")
    Register1.Pages("Produktpreise").Visible = mED_GetBoolean("KarteiProduktpreise")
End Sub

Private Sub cmbJahrMonat_AfterUpdate()
    pv_SetFilter
End Sub

Sub pv_SetFilter()
    Dim txt As String

    If Not IsNull(cmbJahrMonat) Then
        txt = "year([Datum])*100+month([Datum])=" & Str(cmbJahrMonat)
    End If

    If Not IsNull(cmbProjektID) Then
        txt = txt & IIf(Len(txt) > 0, " AND ", "") & "ProjektID=" & Str(cmbProjektID)
    End If

    UF_Leistungen.Form.Filter = txt
    UF_Leistungen.Form.FilterOn = True
End Sub

Private Sub PLZ_AfterUpdate()
    cmbOrtAusPLZ.Requery
    Ort = cmbOrtAusPLZ.Column(1, 0)
End Sub

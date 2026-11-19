Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdReCalc_Click()
    Dim NewID As Long

    DoCmd.Hourglass True

    NewID = mVR_FillVerrechnungKunden(Forms.LeistungenListe!UF.Form.Filter, Nz(chkProjektGruppierung, 0))

    Me.Filter = "ID=" & Str(NewID)
    Me.FilterOn = True

    UF_Kunden.Requery
    UF_Positionen.Requery

    DoCmd.Hourglass False
End Sub

Private Sub cmdRequery_Click()
    UF_Kunden.Requery
End Sub

Private Sub cmdVerrechnen_Click()
    Dim VerrID As Long

    VerrID = Nz(Me!ID, 0)

    If VerrID <> 0 Then
        DoCmd.Close acForm, Me.Name

        mUW_VerrechnungZuRechnungen VerrID

        DoCmd.OpenForm "RechnungenListe"
        With Forms.rechnungenliste.UF.Form
            .Filter = "VerrechnungID=" & Str(VerrID)
            .FilterOn = True
        End With
    End If
End Sub

Private Sub Form_Current()
    UF_Kunden.Requery
    UF_Positionen.Requery
End Sub

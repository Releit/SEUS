Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub pv_Filter()
    Dim txt As String

    If Not IsNull(cmbJahr) Then
        txt = "Jahr=" & Str(cmbJahr)
    End If

    If Not IsNull(SuchName) Then
        txt = txt & IIf(txt = "", "", " AND ") & "KundenName LIKE '" & SuchName & "*'"
    End If

    If Not IsNull(SuchNr) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr>=" & Str(SuchNr)
    End If

    If Not IsNull(SuchNrBis) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr<=" & Str(SuchNrBis)
    End If

    UF.Form.Filter = txt

    If txt <> "" Then
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If

    pv_Summe
End Sub

Private Sub cmbJahr_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbMandant_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdNew_Click()
    Dim FormName As String

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , "ID=0"

End Sub

Private Sub cmdPrint_Click()
    DoCmd.OpenReport "Rechnungsliste", acViewPreview, , UF.Form.Filter
End Sub

Private Sub cmdRequery_Click()
    UF.Requery
End Sub

Private Sub Form_Load()

    cmbJahr = Year(Date)
    pv_Filter
End Sub

Private Sub SuchName_AfterUpdate()
    pv_Filter
End Sub

Private Sub SuchNr_AfterUpdate()
    pv_Filter
End Sub

Private Sub SuchNrBis_AfterUpdate()
    pv_Filter
End Sub

Sub pv_Summe()
    Summe = DSum("BetragExkl", "V_Rechnungen", UF.Form.Filter)
    SummeInkl = DSum("BetragInkl", "V_Rechnungen", UF.Form.Filter)

End Sub

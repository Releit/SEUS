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

    If Not IsNull(cmbQuartal) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Quartal=" & Str(cmbQuartal)
    End If

    If Not IsNull(cmbLieferant) Then
        txt = txt & IIf(txt = "", "", " AND ") & "KarteiID=" & Str(cmbLieferant)
    End If

    If Not IsNull(SuchNr) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr>=" & Str(SuchNr)
    End If

    If Not IsNull(SuchNrBis) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr<=" & Str(SuchNrBis)
    End If

    If txt <> "" Then
        UF.Form.Filter = txt
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If
End Sub

Private Sub cmbJahr_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbMandant_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbLieferant_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbQuartal_Click()
    pv_Filter
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdNew_Click()
    Dim FormName As String

    FormName = "EingangsRechnungen"

    DoCmd.OpenForm FormName, , , "ID=0"

End Sub

Private Sub cmdLfdNr_Click()
    If MsgBox("Sollen die fehlenden Laufenden Nummern vergeben werden?", vbYesNo + vbDefaultButton2, "Nummervergabe") = vbYes Then

        Dim RS As Object, Nr As Long

        Set RS = UF.Form.RecordsetClone

        With RS
            .MoveFirst
            'If IsNull(!LfdNr) Then
            Nr = 1

            Do While Not .EOF
                'If IsNull(!LfdNr) Then
                .Edit
                !LfdNr = Nr
                .Update
                Nr = Nr + 1
                'Else
                '    Nr = !LfdNr + 1
                'End If
                .MoveNext
            Loop
        End With
    End If
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

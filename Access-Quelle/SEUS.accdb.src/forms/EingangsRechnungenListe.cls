Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_Filter()
    Dim txt As String

    If Not IsNull(cmbJahr) Then
        txt = "Jahr=" & Str(cmbJahr)
    End If

    If Not IsNull(cmbQt) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Qt=" & Str(cmbQt)
    End If

    If Not IsNull(cmbLieferant) Then
        txt = txt & IIf(txt = "", "", " AND ") & "LiefKarteiID =" & Str(cmbLieferant)
    End If

    If Not IsNull(SuchNr) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr>=" & Str(SuchNr)
    End If

    If Not IsNull(SuchNrBis) Then
        txt = txt & IIf(txt = "", "", " AND ") & "Nr<=" & Str(SuchNrBis)
    End If

    If chkIGE Then
        txt = txt & IIf(txt = "", "", " AND ") & "InnerGemErwerb=True"
    End If

    If chkOhneDatei Then
        txt = txt & IIf(txt = "", "", " AND PfadDatei IS NULL")
    End If

    If txt <> "" Then
        UF.Form.Filter = txt
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If

    SummeNetto = DSum("Netto", "A_EingangsRechnungen", txt)
End Function

Private Sub chkIGE_AfterUpdate()
    pv_Filter
End Sub

Private Sub chkOhneDatei_AfterUpdate()
    pv_Filter
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

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdNew_Click()
    Dim FormName As String

    FormName = "EingangsRechnungen"

    DoCmd.OpenForm FormName, , , "ID=0"

End Sub

Private Sub cmdOhneZuordnungVerschieben_Click()
    Dim Pfad As String
    Dim rs As DAO.Recordset

    If IsNull(cmbJahr) Then
        MsgBox "Jahr auswählen", , "fehlende Eingabe"
    End If

    Pfad = mED_ER_Pfad(DateSerial(cmbJahr, 1, 1))
    MsgBox Pfad
    mFM_FillTempDateien Pfad

    Set rs = CurrentDb.OpenRecordset("TempDateien", dbOpenSnapshot)
    With rs
    Do While Not .EOF
        If DCount("*", "T_Eingangsrechnungen", "PfadDatei='" & !FileName & "'") = 0 Then
            mFM_MoveFile Pfad & "\" & !FileName, Pfad & "\Ohne Zuordnung\" & !FileName
        End If
        .MoveNext
    Loop
    End With

End Sub

Private Sub cmdPrint_Click()
    DoCmd.OpenReport "Rechnungsliste", acViewPreview, , UF.Form.Filter
End Sub

Private Sub cmdRequery_Click()
    pv_Filter
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

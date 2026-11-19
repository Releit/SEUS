Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub pv_Filter()
    Dim txt As String

    Select Case Status
    Case 1                                       '** offen **
        txt = IIf(txt <> "", txt & " AND ", "") & "erledigt=False AND obsolet=False"

    Case 2                                       '** erledigt **
        txt = IIf(txt <> "", txt & " AND ", "") & "erledigt=True"

    Case 3                                       '** obsolet **
        txt = IIf(txt <> "", txt & " AND ", "") & "obsolet=True"
    End Select

    Select Case ZeitFilter
    Case 1                                       '** bereits fällig **
        txt = IIf(txt <> "", txt & " AND ", "") & "(FaelligAm IS NULL OR FaelligAm<=" & CDate2SQL(Date) & ")"
    End Select

    If SuchText <> "" Then
        txt = IIf(txt <> "", txt & " AND ", "") & "(Aufgabe LIKE '*" & SuchText & "*' OR Anmerkungen LIKE '*" & SuchText & "*')"
    End If

    If Nz(SuchKarteiID, 0) <> 0 Then
        txt = IIf(txt <> "", txt & " AND ", "") & " KarteiID =" & Str(SuchKarteiID)
    End If

    If txt <> "" Then
        'MsgBox Txt
        UF.Form.Filter = txt
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If
End Sub

Private Sub cmbAnlass_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbWer_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdMenü_Click()
    DoCmd.OpenForm "Hauptmenü"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdLeistungen_Click()
    DoCmd.Close acForm, Me.Name
    DoCmd.OpenForm "LeistungenListe"
    Forms("LeistungenListe").Visible = True
End Sub

Private Sub cmdPrint_Click()
    DoCmd.OpenReport "ToDoListe", acViewPreview, , UF.Form.Filter
End Sub

Private Sub cmdRequery_Click()
    UF.Requery
End Sub

Private Sub Form_Load()
    'If mPD_PersonenNrAusWindowsName <> 0 Then cmbWer = mPD_PersonenNrAusWindowsName
    pv_Filter
End Sub

Private Sub Status_AfterUpdate()
    pv_Filter
End Sub

Private Sub SuchKarteiID_AfterUpdate()
    pv_Filter
End Sub

Private Sub SuchText_AfterUpdate()
    pv_Filter
End Sub

Private Sub ZeitFilter_AfterUpdate()
    pv_Filter
End Sub

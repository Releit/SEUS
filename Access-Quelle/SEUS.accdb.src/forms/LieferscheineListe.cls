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

    If Not IsNull(SuchName) Then
        txt = txt & IIf(txt = "", "", " AND ") & "KundenName LIKE '" & SuchName & "*'"
    End If

    If optVerrechnet = 2 Then                    '** nur unverrechnete
        txt = txt & IIf(txt = "", "", " AND ") & "Unverrechnet<>0"
    End If

    If txt <> "" Then
        UF.Form.Filter = txt
        UF.Form.FilterOn = True
    Else
        UF.Form.FilterOn = False
    End If
End Function

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdNew_Click()
    Dim FormName As String

    FormName = "Lieferscheine"

    DoCmd.OpenForm FormName, , , "ID=0"

End Sub

Private Sub cmdRequery_Click()
    UF.Requery
End Sub

Private Sub Form_Load()

    cmbJahr = Year(Date)
    pv_Filter
End Sub

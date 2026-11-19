Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_Filter()
    Dim txt As String, Sql As String

    If Not IsNull(SuchName) Then
        txt = "KundenName LIKE '" & SuchName & "*'"
    End If

    If Not IsNull(SuchNr) Then
        txt = txt & IIf(txt <> "", " AND ", "") & "Nr =" & Str(SuchNr)
    End If

    If Not IsNull(SuchPLZ) Then
        txt = txt & IIf(txt <> "", " AND ", "") & "PLZ Like '" & SuchPLZ & "*'"

    End If

    If Not IsNull(SuchPLZ) Then
        txt = txt & IIf(txt <> "", " AND ", "") & "PLZ Like '" & SuchPLZ & "*'"
    End If

    If txt <> "" Then
        UF.Form.Filter = txt
        UF.Form.FilterOn = True
    Else
        UF.Form.Filter = ""
        UF.Form.FilterOn = False
    End If
End Function

Private Sub cmbJahr_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbMandant_AfterUpdate()
    pv_Filter
End Sub

Private Sub cmbFlagID_AfterUpdate()
    mFD_SetFlagID Nz(cmbFlagID, 0)
    UF.Form.Requery
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
    If mSU_KarteiSuche Then
        Forms(mSU_StartForm).KarteiID.Requery
        Forms(mSU_StartForm).KarteiID = mSU_KarteiID
        mSU_KarteiSuche = False
    End If
End Sub

Private Sub cmdExcel_Click()
    Dim Sql As String

    Sql = CurrentDb.QueryDefs("A_Kundenliste").Sql
    Sql = Replace(Sql, ";", "")
    CurrentDb.QueryDefs("A_Kundenliste_Export").Sql = Sql & IIf(UF.Form.Filter <> "", " WHERE " & UF.Form.Filter, "")
    DoCmd.OutputTo acQuery, "A_Kundenliste_Export", acFormatXLS, "D:\Firma\Access\Exporte\Export" & Format(Date, "YYYYMMDD") & ".xls", True
End Sub

Private Sub cmdNew_Click()
    Dim FormName As String

    FormName = "Kartei"

    DoCmd.OpenForm FormName, , , "ID=0"

End Sub

Private Sub cmdOpenReport_Click()
    DoCmd.OpenReport "Kundenliste", acViewPreview, , UF.Form.Filter
End Sub

Private Sub cmdRequery_Click()
    UF.Requery
End Sub

Private Sub Form_Load()
    If mFD_GetFlagID() Then cmbFlagID = mFD_GetFlagID()
    pv_Filter
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim ReportName As String

Private Sub cmbMandant_Click()
    cmdRequery_Click
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdMahnung1_Click()
    Dim Sql As String, sIDList As String

    sIDList = GetIDString(LstZuMahnen1)
    If sIDList = "" Then MsgBox "Bitte mindestens eine Rechnung auswählen !": Exit Sub
    pv_SetMahnStufe LstZuMahnen1, 1

    cmdRequery_Click
End Sub

Public Function GetIDString(ctl As Access.Control) As String
    On Error Resume Next

    Dim item As Variant, sIDList As String, sID

    For Each item In ctl.ItemsSelected
        sID = ctl.ItemData(item)
        sIDList = sIDList & "," & sID
    Next

    sIDList = Mid(sIDList, 2)
    GetIDString = sIDList

End Function

Public Sub pv_SetMahnStufe(ctl As Access.Control, Stufe As Integer)
    'On Error Resume Next

    Dim item As Variant, Sql As String

    For Each item In ctl.ItemsSelected
        pv_MahnungenEintragen ctl.ItemData(item), Stufe
    Next
End Sub

Sub pv_MahnungenEintragen(RechnungID As Long, MahnStufe As Integer)
    Dim Sql As String

    Sql = "UPDATE T_Rechnungen SET MahnStufe=" & Str(MahnStufe) & _
          IIf(MahnStufe < 4, ", Mahnung" & Format(MahnStufe) & "Am=Date()", ", RechtsanwaltAm=Date()") & _
          " WHERE ID =" & Str(RechnungID)
    CurrentDb.Execute Sql
End Sub

Private Sub cmdMahnung2_Click()
    Dim Sql As String, sIDList As String

    sIDList = GetIDString(LstZuMahnen2)
    If sIDList = "" Then MsgBox "Bitte mindestens eine Rechnung auswählen !": Exit Sub
    pv_SetMahnStufe LstZuMahnen2, 2

    cmdRequery_Click
End Sub

Private Sub cmdMahnung3_Click()
    Dim Sql As String, sIDList As String

    sIDList = GetIDString(LstZuMahnen3)
    If sIDList = "" Then MsgBox "Bitte mindestens eine Rechnung auswählen !": Exit Sub
    pv_SetMahnStufe LstZuMahnen3, 3

    cmdRequery_Click
End Sub

Private Sub cmdMahnung4_Click()
    Dim Sql As String, sIDList As String

    sIDList = GetIDString(LstZuMahnen4)
    If sIDList = "" Then MsgBox "Bitte mindestens eine Rechnung auswählen !": Exit Sub
    pv_SetMahnStufe LstZuMahnen4, 4

    cmdRequery_Click
End Sub

Private Sub cmdPrintMahnung1_Click()
    ReportName = "Mahnung1"
    DoCmd.OpenReport ReportName, acViewPreview, , "MahnStufe=1" ' AND Mahnung1Am=#" & cdate2sql(mac_serverdate()) & "#"
End Sub

Private Sub cmdPrintMahnung2_Click()
    ReportName = "Mahnung2"
    DoCmd.OpenReport ReportName, acViewPreview, , "MahnStufe=2" ' AND Mahnung1Am=#" & cdate2sql(mac_serverdate()) & "#"
End Sub

Private Sub cmdPrintMahnung3_Click()
    ReportName = "Mahnung3"
    DoCmd.OpenReport ReportName, acViewPreview, , "MahnStufe=3" ' AND Mahnung1Am=#" & cdate2sql(mac_serverdate()) & "#"
End Sub

Private Sub cmdPrintMahnung4_Click()
    ReportName = "RAListe"
    DoCmd.OpenReport ReportName, acViewPreview, , "MahnStufe=4" ' AND Mahnung1Am=#" & cdate2sql(mac_serverdate()) & "#"
End Sub

Private Sub cmdRequery_Click()
    LstMahnung1.Requery
    LstMahnung2.Requery
    LstMahnung3.Requery
    LstMahnung4.Requery
    LstZuMahnen1.Requery
    LstZuMahnen2.Requery
    LstZuMahnen3.Requery
    LstZuMahnen4.Requery
End Sub

Private Sub Form_Load()

    cmdRequery_Click
    Mahnung1.Visible = mED_GetLong("MahnAnzahl") > 0
    Mahnung2.Visible = mED_GetLong("MahnAnzahl") > 1
    Mahnung3.Visible = mED_GetLong("MahnAnzahl") > 2

End Sub

Private Sub LstMahnung1_DblClick(Cancel As Integer)
    If Not IsNull(LstMahnung1.Column(0)) Then pv_GoToRechnung Val(LstMahnung1.Column(0))
End Sub

Private Sub LstZuMahnen1_DblClick(Cancel As Integer)
    If Not IsNull(LstZuMahnen1.Column(0)) Then pv_GoToRechnung Val(LstZuMahnen1.Column(0))
End Sub

Private Sub LstMahnung2_DblClick(Cancel As Integer)
    If Not IsNull(LstMahnung2.Column(0)) Then pv_GoToRechnung Val(LstMahnung2.Column(0))
End Sub

Private Sub LstZuMahnen2_DblClick(Cancel As Integer)
    If Not IsNull(LstZuMahnen2.Column(0)) Then pv_GoToRechnung Val(LstZuMahnen2.Column(0))
End Sub

Private Sub LstMahnung3_DblClick(Cancel As Integer)
    If Not IsNull(LstMahnung3.Column(0)) Then pv_GoToRechnung Val(LstMahnung3.Column(0))
End Sub

Private Sub LstZuMahnen3_DblClick(Cancel As Integer)
    If Not IsNull(LstZuMahnen3.Column(0)) Then pv_GoToRechnung Val(LstZuMahnen3.Column(0))
End Sub

Private Sub LstMahnung4_DblClick(Cancel As Integer)
    If Not IsNull(LstMahnung4.Column(0)) Then pv_GoToRechnung Val(LstMahnung4.Column(0))
End Sub

Private Sub LstZuMahnen4_DblClick(Cancel As Integer)
    If Not IsNull(LstZuMahnen4.Column(0)) Then pv_GoToRechnung Val(LstZuMahnen4.Column(0))
End Sub

Function pv_GoToRechnung(RechnungID As Long)
    Dim FormName As String

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , , , acHidden
    With Forms(FormName)
        .RecordsetClone.FindFirst "ID=" & Str(RechnungID)
        If Not .RecordsetClone.EOF Then
            .Recordset.Bookmark = .RecordsetClone.Bookmark
            .Visible = True
        Else
            MsgBox "Datensatz nicht gefunden!", , FormName
            DoCmd.Close acForm, FormName
        End If
    End With
End Function

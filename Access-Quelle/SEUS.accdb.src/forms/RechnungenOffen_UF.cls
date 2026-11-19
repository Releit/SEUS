Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbMandant_Click()
    cmdRequery_Click
End Sub

Private Sub cmdBuchen_Click()
    DoCmd.OpenForm "RechnungenVerbuchen", , , "RechnungID=" & Str(Me!RechnungID)
End Sub

Function pv_GoToRechnung()
    Dim FormName As String

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , , , acHidden
    With Forms(FormName)
        .RecordsetClone.FindFirst "ID=" & Str(Me!RechnungID)
        If Not .RecordsetClone.EOF Then
            .Recordset.Bookmark = .RecordsetClone.Bookmark
            .Visible = True
        Else
            MsgBox "Datensatz nicht gefunden!", , FormName
            DoCmd.Close acForm, FormName
        End If
    End With
End Function

Function pv_GoToKunde()
    Dim FormName As String, ToID As Long

    If Not IsNull(Me!FirmenID) Then
        ToID = Me!FirmenID
        FormName = "Rechnungen"
    Else

    End If

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , , , acHidden
    With Forms(FormName)
        .RecordsetClone.FindFirst "ID=" & Str(ToID)
        If Not .RecordsetClone.EOF Then
            .Recordset.Bookmark = .RecordsetClone.Bookmark
            .Visible = True
        Else
            MsgBox "Datensatz nicht gefunden!", , FormName
            DoCmd.Close acForm, FormName
        End If
    End With
End Function

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdRequery_Click()
    Me.Requery
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Function pv_GoToID()
    Dim FormName As String

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , , , acHidden
    With Forms(FormName)
        .RecordsetClone.FindFirst "ID=" & Str(Me!ID)
        If Not .RecordsetClone.EOF Then
            .Recordset.Bookmark = .RecordsetClone.Bookmark
            .Visible = True
        Else
            MsgBox "Datensatz nicht gefunden!", , FormName
            DoCmd.Close acForm, FormName
        End If
    End With
End Function

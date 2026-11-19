Attribute VB_Name = "GoTo"
Option Compare Database
Option Explicit

Function m_GoToID(FormName As String, ToID As Long)
    Dim RS As Object

    DoCmd.OpenForm FormName

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

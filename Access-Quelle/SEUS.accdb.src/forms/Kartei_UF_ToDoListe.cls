Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Function pv_GoToGF()
    Dim FormName As String

    FormName = "GF"

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

Private Sub Bewertung_AfterUpdate()
    If Not IsNull(Bewertung) Then
        Me!BewertungProz = Bewertung.Column(1)
    Else
        Me!BewertungProz = Null
    End If
End Sub

Private Sub Erledigt_AfterUpdate()
    If erledigt And IsNull(ErledigtAm) Then ErledigtAm = Date
    If erledigt And IsNull(Me!ErledigtVon) Then Me!ErledigtVon = mWin_UserName()
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    AngelegtVon = mWin_UserName()
End Sub

Private Sub obsolet_AfterUpdate()
    If obsolet And IsNull(obsoletAm) Then obsoletAm = Date
    If obsolet And IsNull(obsoletVon) Then obsoletVon = mWin_UserName()
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!BearbeitetAm = Date
    Me!BearbeitetVon = mWin_UserName
End Sub

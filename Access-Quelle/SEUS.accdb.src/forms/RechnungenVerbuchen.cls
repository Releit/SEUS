Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Sql As String

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name

    On Error Resume Next
    Forms.rechnungenoffen.Requery
End Sub

Private Sub cmdRequery_Click()
    Me.Requery
End Sub

Private Sub cmdSaldoTotal_Click()
    pv_Buchen -Saldo, "Zahlungseingang"
End Sub

Private Sub cmdSkonto_Click()
    pv_Buchen -Saldo, "Skonto"
End Sub

Sub pv_Buchen(Betrag As Currency, Buchungsart As String)

    If IsNull(Buchungsdatum) Then
        MsgBox "Bitte das Buchungsdatum eingeben!", , "Fehlende Eingabe"
        Exit Sub
    End If

    mUW_RechnungBuchen Me!RechnungID, Buchungsdatum, Betrag, Buchungsart

    Me.Requery
End Sub

Private Sub cmdTeilZahlung_Click()

    If IsNull(Buchungsdatum) Then
        MsgBox "Bitte das Buchungsdatum eingeben!", , "Fehlende Eingabe"
        Exit Sub
    End If

    If Nz(TZBetrag, 0) = 0 Then
        MsgBox "Bitte den Teilzahlungsbetrag eingeben", , "Fehlende Eingabe"
        Exit Sub
    End If

    If Nz(TZBetrag, 0) < 0 Then
        MsgBox "Der Teilzahlungsbetrag muß grösser als 0 sein", , "ungültige Eingabe"
        Exit Sub
    End If

    If Nz(TZBetrag, 0) > -Saldo Then
        MsgBox "Der Teilzahlungsbetrag darf nicht grösser als der offene Saldo sein", , "ungültige Eingabe"
        Exit Sub
    End If

    pv_Buchen TZBetrag, "Zahlungseingang"
    TZBetrag = 0
End Sub

Private Sub Form_Current()
    LstBuchungen.Requery
    Buchungsdatum.SetFocus
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

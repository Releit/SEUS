Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Sql As String

Private Sub Betrag_AfterUpdate()

    Sql = "SELECT RechnungID, KarteiID, V_RechnungenMitSaldo.Datum, BetragInkl, KundenName, Saldo" & _
          " FROM V_RechnungenMitSaldo" & _
          " GROUP BY RechnungID, KarteiID, V_RechnungenMitSaldo.Datum, BetragInkl, KundenName, Saldo"

    If Not IsNull(Betrag) And Not IsNull(Datum) Then
        Sql = Sql & " HAVING V_RechnungenMitSaldo.Datum<=" & CDate2SQL(Datum) & " AND Abs(Saldo+" & Str(Betrag) & ")<0.05"
    Else
        Sql = Sql & " HAVING RechnungID=0"
    End If

    Sql = Sql & " ORDER BY Datum;"
    Debug.Print Sql
    lstBelege.RowSource = Sql
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Current()
    lstBelege.Requery
End Sub

Private Sub lstBelege_DblClick(Cancel As Integer)
    Dim NewID As Long, OldID As Long

    If IsNull(ZahlungsartID) Then
        MsgBox "Bitte auch die Zahlart angeben!", , "fehlede Eingabe"
        Exit Sub
    End If


    If MsgBox("Rechnung dem Zahlungseingang zuordnen?", vbYesNo, "Buchen") = vbYes Then
        OldID = DMax("ID", "T_Rechnungen_ZE")

        Sql = "INSERT INTO T_Rechnungen_ZE(ZEID, RechnungID, Betrag, AngelegtVon)" & _
              " VALUES(" & Format(Me!ID) & _
              "," & Format(lstBelege) & _
              "," & Str(mRD_BetragInkl(lstBelege)) & _
              ",'" & mWin_UserName() & "')"

        CurrentDb.Execute Sql

        NewID = Nz(DMax("ID", "T_Rechnungen_ZE"), 0)

        If NewID > OldID Then
            Sql = "UPDATE T_Rechnungen SET ZahlungseingangID=" & Format(Me!ID) & " WHERE ID=" & Str(lstBelege)
            CurrentDb.Execute Sql
            lstBelege.Requery
            'Betrag = Null
        Else
            MsgBox "Fehler beim Anlegen eines neuen Zahlungseingangs", , "Fehler"
        End If
    End If

    UF_ZE.Requery
End Sub

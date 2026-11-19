Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    LstZA.Requery
    UF_ZA.Requery
End Sub

Private Sub LiefKarteiID_AfterUpdate()
    LstZA.Requery
    UF_ZA.Requery
End Sub

Private Sub LstZA_DblClick(Cancel As Integer)
    Dim NewID As Long, OldID As Long, Sql As String

    If IsNull(ZahlungsartID) Then
        MsgBox "Bitte auch die Zahlart angeben!", , "fehlede Eingabe"
        Exit Sub
    End If

    If Not IsNull(LstZA) Then
        If MsgBox("Rechnung dem Zahlungsausgang zuordnen?", vbYesNo, "Buchen") = vbYes Then
            OldID = DMax("ID", "T_EingangsRechnungen_ZA")

            Sql = "INSERT INTO T_EingangsRechnungen_ZA(ZAID, ERID, Betrag, AngelegtVon)" & _
                  " VALUES(" & Format(Me!ID) & _
                  "," & Format(LstZA) & _
                  "," & Str(LstZA.Column(2)) & _
                  ",'" & mWin_UserName() & "')"

            CurrentDb.Execute Sql

            NewID = Nz(DMax("ID", "T_EingangsRechnungen_ZA"), 0)

            If NewID > OldID Then
                LstZA.Requery
                UF_ZA.Requery
                'Betrag = Null
            Else
                MsgBox "Fehler beim Anlegen der Zuordung in der Tabelle T_EingangsRechnungen_ZA", , "Fehler"
            End If
        End If
    End If
End Sub

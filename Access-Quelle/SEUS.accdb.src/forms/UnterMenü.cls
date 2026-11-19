Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl22_Click()
    DoCmd.OpenForm "SimpleSQL"
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdEinstellungen_Click()
    DoCmd.OpenForm "Einstellungen"
End Sub

Private Sub cmdKartei_Click()
    DoCmd.OpenForm "KarteiListe"
End Sub

Private Sub cmdLieferscheine_Click()
    DoCmd.OpenForm "LeistungenListe"
End Sub

Private Sub cmdNewLFS_Click()
    DoCmd.OpenForm "Lieferscheine", , , "Id=0"
End Sub

Private Sub cmdNewRechnung_Click()
    DoCmd.OpenForm "Rechnungen", , , "Id=0"
End Sub

Private Sub cmdER_Click()
    DoCmd.OpenForm "EingangsrechnungenListe"
End Sub

Private Sub cmdProdukte_Click()
    DoCmd.OpenForm "Produkte"
End Sub

Private Sub cmdRechnungen_Click()
    DoCmd.OpenForm "RechnungenListe"
End Sub

Private Sub cmdRechnungenOffen_Click()
    DoCmd.OpenForm "RechnungenOffen"
End Sub

Private Sub cmdToDoListe_Click()
    DoCmd.OpenForm "ToDoListe"
End Sub

Private Sub cmdZA_Click()
    DoCmd.OpenForm "ZahlungsausgangsListe"
End Sub

Private Sub cmdZahlungseingänge_Click()
    DoCmd.OpenForm "ZahlungseingangsListe"
End Sub

Private Sub mdMahnungen_Click()
    DoCmd.OpenForm "Mahnungen"
End Sub

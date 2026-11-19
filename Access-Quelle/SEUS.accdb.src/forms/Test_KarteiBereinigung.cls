Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Test-Formular: Karteieinträge, die seit mehr als 7 Jahren keine Rechnung erhalten haben,
' und Löschen eines Eintrags samt allen abhängigen Datensätzen.

Private Const JAHRE As Integer = 7

' Liste: letzte Ausgangsrechnung älter als 7 Jahre (ohne Rechnung zählt das Anlagedatum),
' und auch keine Eingangsrechnung in den letzten 7 Jahren.
Private Function ListenSql() As String
    Dim Grenze As String, Sql As String
    Grenze = Format(DateAdd("yyyy", -JAHRE, Date), "\#mm\/dd\/yyyy\#")

    Sql = "SELECT K.ID, K.Nr, Trim(K.Name1 & ' ' & K.Name2) AS Kundenname, K.PLZ, K.Ort, K.AngelegtAm, " & _
          "(SELECT Max(R.Datum) FROM T_Rechnungen AS R WHERE R.KarteiID=K.ID) AS LetzteRechnung, " & _
          "(SELECT Count(*) FROM T_Rechnungen AS R WHERE R.KarteiID=K.ID) AS AnzahlRechnungen, " & _
          "(SELECT Max(E.Datum) FROM T_EingangsRechnungen AS E WHERE E.LiefKarteiID=K.ID) AS LetzteEingangsrechnung " & _
          "FROM T_Kartei AS K " & _
          "WHERE Nz((SELECT Max(R.Datum) FROM T_Rechnungen AS R WHERE R.KarteiID=K.ID), Nz(K.AngelegtAm, #1/1/1900#)) < " & Grenze & _
          " AND Nz((SELECT Max(E.Datum) FROM T_EingangsRechnungen AS E WHERE E.LiefKarteiID=K.ID), #1/1/1900#) < " & Grenze & _
          " ORDER BY K.Name1, K.Name2"
    ListenSql = Sql
End Function

Private Sub Form_Load()
    Me.RecordSource = ListenSql()
End Sub

Private Sub cmdLoeschen_Click()
    Dim KarteiID As Long, Bezeichnung As String, Hindernis As String, Uebersicht As String

    If IsNull(Me!ID) Then Exit Sub
    KarteiID = Me!ID
    Bezeichnung = "Nr " & Nz(Me!Nr, "-") & ": " & Nz(Me!Kundenname, "")

    Hindernis = FremdeVerweise(KarteiID)
    If Hindernis <> "" Then
        MsgBox "Der Karteieintrag " & Bezeichnung & " kann nicht gelöscht werden, " & _
               "weil Datensätze anderer Karteieinträge darauf verweisen:" & vbCrLf & vbCrLf & Hindernis, _
               vbExclamation, "Löschen nicht möglich"
        Exit Sub
    End If

    Uebersicht = Zusammenfassung(KarteiID)
    If MsgBox("Karteieintrag " & Bezeichnung & " endgültig löschen?" & vbCrLf & vbCrLf & _
              "Mitgelöscht werden:" & vbCrLf & Uebersicht & vbCrLf & _
              "Das kann nicht rückgängig gemacht werden.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Karteieintrag löschen") <> vbYes Then Exit Sub

    If LoescheKartei(KarteiID) Then
        Me.Requery
        MsgBox "Karteieintrag " & Bezeichnung & " wurde gelöscht.", vbInformation
    End If
End Sub

' --- Teilmengen eines Karteieintrags als Unterabfragen

Private Function Rech(KarteiID As Long) As String
    Rech = "(SELECT ID FROM T_Rechnungen WHERE KarteiID=" & KarteiID & ")"
End Function

Private Function Lief(KarteiID As Long) As String
    Lief = "(SELECT ID FROM T_Lieferscheine WHERE KarteiID=" & KarteiID & ")"
End Function

Private Function ZE(KarteiID As Long) As String
    ZE = "(SELECT ID FROM [T_Zahlungseingänge] WHERE KarteiID=" & KarteiID & ")"
End Function

Private Function ZA(KarteiID As Long) As String
    ZA = "(SELECT ID FROM [T_Zahlungsausgänge] WHERE KarteiID=" & KarteiID & ")"
End Function

Private Function ER(KarteiID As Long) As String
    ER = "(SELECT ID FROM T_EingangsRechnungen WHERE LiefKarteiID=" & KarteiID & ")"
End Function

Private Function Proj(KarteiID As Long) As String
    Proj = "(SELECT ID FROM T_Projekte WHERE KarteiID=" & KarteiID & ")"
End Function

' Zählt Datensätze einer Tabelle zu einer Bedingung
Private Function Anzahl(Tabelle As String, Bedingung As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & Tabelle & " WHERE " & Bedingung, dbOpenSnapshot)
    Anzahl = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' Hängt eine Zeile an, wenn die Anzahl größer 0 ist
Private Sub Zeile(ByRef Text As String, Anz As Long, Beschreibung As String)
    If Anz > 0 Then Text = Text & "  - " & Anz & " " & Beschreibung & vbCrLf
End Sub

' Prüft, ob Datensätze anderer Karteieinträge auf Daten dieses Eintrags verweisen.
' Dann wird nicht gelöscht, damit keine fremden Daten verloren gehen oder ins Leere zeigen.
Private Function FremdeVerweise(K As Long) As String
    Dim T As String
    Zeile T, Anzahl("T_Rechnungen", "GegenRechnungID IN " & Rech(K) & " AND Nz(KarteiID,0)<>" & K), "Rechnungen anderer Kunden mit Gegenrechnung auf diesen Kunden"
    Zeile T, Anzahl("T_Rechnungen", "ZahlungseingangID IN " & ZE(K) & " AND Nz(KarteiID,0)<>" & K), "Rechnungen anderer Kunden mit Zahlungseingang dieses Kunden"
    Zeile T, Anzahl("T_Buchungen", "(RechnungID IN " & Rech(K) & " OR GegenRechnungID IN " & Rech(K) & " OR ZahlungseingangID IN " & ZE(K) & ") AND Nz(KarteiID,0)<>" & K), "Buchungen anderer Kunden"
    Zeile T, Anzahl("T_Rechnungen_ZE", "(RechnungID IN " & Rech(K) & " AND NOT ZEID IN " & ZE(K) & ") OR (ZEID IN " & ZE(K) & " AND NOT RechnungID IN " & Rech(K) & ")"), "Zahlungszuordnungen zu anderen Kunden"
    Zeile T, Anzahl("T_RechnungsPos", "LieferscheinID IN " & Lief(K) & " AND NOT RechnungID IN " & Rech(K)), "Rechnungspositionen anderer Kunden mit Lieferschein dieses Kunden"
    Zeile T, Anzahl("T_LieferscheinPos", "RechnungID IN " & Rech(K) & " AND NOT LieferscheinID IN " & Lief(K)), "Lieferscheinpositionen anderer Kunden auf Rechnungen dieses Kunden"
    Zeile T, Anzahl("T_EingangsRechnungen_ZA", "(ERID IN " & ER(K) & " AND NOT ZAID IN " & ZA(K) & ") OR (ZAID IN " & ZA(K) & " AND NOT ERID IN " & ER(K) & ")"), "Zahlungsausgänge, die auch Eingangsrechnungen anderer Lieferanten betreffen"
    Zeile T, Anzahl("T_Leistungen", "(AufRechnungID IN " & Rech(K) & " OR AufHonorarRechnungID IN " & Rech(K) & " OR ProjektID IN " & Proj(K) & ") AND Nz(KarteiID,0)<>" & K), "Leistungen anderer Kunden auf Rechnungen oder Projekten dieses Kunden"
    Zeile T, Anzahl("T_Leistungen_alt", "(AufRechnungID IN " & Rech(K) & " OR AufHonorarRechnungID IN " & Rech(K) & ") AND Nz(KarteiID,0)<>" & K), "alte Leistungen anderer Kunden auf Rechnungen dieses Kunden"
    FremdeVerweise = T
End Function

' Listet, was beim Löschen alles mit entfernt wird
Private Function Zusammenfassung(K As Long) As String
    Dim T As String
    Zeile T, Anzahl("T_Rechnungen", "KarteiID=" & K), "Rechnungen"
    Zeile T, Anzahl("T_RechnungsPos", "RechnungID IN " & Rech(K)), "Rechnungspositionen"
    Zeile T, Anzahl("T_Buchungen", "KarteiID=" & K & " OR RechnungID IN " & Rech(K) & " OR ZahlungseingangID IN " & ZE(K)), "Buchungen"
    Zeile T, Anzahl("[T_Zahlungseingänge]", "KarteiID=" & K), "Zahlungseingänge"
    Zeile T, Anzahl("T_Lieferscheine", "KarteiID=" & K), "Lieferscheine"
    Zeile T, Anzahl("T_EingangsRechnungen", "LiefKarteiID=" & K), "Eingangsrechnungen"
    Zeile T, Anzahl("[T_Zahlungsausgänge]", "KarteiID=" & K), "Zahlungsausgänge"
    Zeile T, Anzahl("T_Leistungen", "KarteiID=" & K), "Leistungen (SQL-Server)"
    Zeile T, Anzahl("T_Projekte", "KarteiID=" & K), "Projekte (SQL-Server)"
    Zeile T, Anzahl("T_KundenProduktpreise", "KarteiID=" & K), "Produktpreise (SQL-Server)"
    Zeile T, Anzahl("T_Adressen", "KarteiID=" & K), "Adressen"
    Zeile T, Anzahl("T_Ansprechpartner", "KarteiID=" & K), "Ansprechpartner"
    Zeile T, Anzahl("T_ToDoListe", "KarteiID=" & K), "ToDo-Einträge"
    Zeile T, Anzahl("T_KarteiFlags", "KarteiID=" & K), "Markierungen"
    Zeile T, Anzahl("T_VerrechnungKunden", "KarteiID=" & K), "Verrechnungszuordnungen"
    If T = "" Then T = "  (keine abhängigen Datensätze)" & vbCrLf
    Zusammenfassung = T
End Function

' Löscht den Karteieintrag und alle abhängigen Datensätze in einer Transaktion.
' Reihenfolge: zuerst die abhängigen Tabellen, damit die Beziehungen nicht verletzt werden.
Private Function LoescheKartei(K As Long) As Boolean
    Dim ws As DAO.Workspace, db As DAO.Database, Schritt As String
    Dim Befehle As Variant, i As Integer

    ' Die Unterabfragen müssen ausgewertet werden, solange die Ausgangsdaten noch existieren,
    ' deshalb kommen Rechnungen, Lieferscheine, Zahlungen und Projekte erst am Schluss.
    Befehle = Array( _
        "DELETE * FROM T_Rechnungen_ZE WHERE RechnungID IN " & Rech(K) & " OR ZEID IN " & ZE(K), _
        "DELETE * FROM T_Buchungen WHERE KarteiID=" & K & " OR RechnungID IN " & Rech(K) & " OR GegenRechnungID IN " & Rech(K) & " OR ZahlungseingangID IN " & ZE(K), _
        "DELETE * FROM T_RechnungsPos WHERE RechnungID IN " & Rech(K) & " OR LieferscheinID IN " & Lief(K), _
        "DELETE * FROM T_LieferscheinPos WHERE LieferscheinID IN " & Lief(K) & " OR RechnungID IN " & Rech(K), _
        "DELETE * FROM T_Leistungen WHERE KarteiID=" & K, _
        "DELETE * FROM T_Leistungen_alt WHERE KarteiID=" & K, _
        "DELETE * FROM T_EingangsRechnungen_ZA WHERE ERID IN " & ER(K) & " OR ZAID IN " & ZA(K), _
        "DELETE * FROM T_EingangsRechnungen WHERE LiefKarteiID=" & K, _
        "DELETE * FROM [T_Zahlungsausgänge] WHERE KarteiID=" & K, _
        "DELETE * FROM [T_Zahlungseingänge] WHERE KarteiID=" & K, _
        "UPDATE T_Rechnungen SET GegenRechnungID=Null WHERE KarteiID=" & K, _
        "DELETE * FROM T_Rechnungen WHERE KarteiID=" & K, _
        "DELETE * FROM T_Lieferscheine WHERE KarteiID=" & K, _
        "DELETE * FROM T_Projekte WHERE KarteiID=" & K, _
        "DELETE * FROM T_Projekte_alt WHERE KarteiID=" & K, _
        "DELETE * FROM T_KundenProduktpreise WHERE KarteiID=" & K, _
        "DELETE * FROM T_KundenProduktpreise_alt WHERE KarteiID=" & K, _
        "DELETE * FROM T_Adressen WHERE KarteiID=" & K, _
        "DELETE * FROM T_Ansprechpartner WHERE KarteiID=" & K, _
        "DELETE * FROM T_ToDoListe WHERE KarteiID=" & K, _
        "DELETE * FROM T_KarteiFlags WHERE KarteiID=" & K, _
        "DELETE * FROM T_VerrechnungKunden WHERE KarteiID=" & K, _
        "DELETE * FROM T_Kartei WHERE ID=" & K)

    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    On Error GoTo Fehler
    ws.BeginTrans
    For i = LBound(Befehle) To UBound(Befehle)
        Schritt = Befehle(i)
        db.Execute Schritt, dbFailOnError + dbSeeChanges
    Next i
    ws.CommitTrans
    LoescheKartei = True
    Exit Function

Fehler:
    Dim Meldung As String
    Meldung = Err.Description
    On Error Resume Next
    ws.Rollback
    MsgBox "Löschen abgebrochen, es wurde nichts gelöscht." & vbCrLf & vbCrLf & _
           "Fehler: " & Meldung & vbCrLf & vbCrLf & "Bei: " & Schritt, vbCritical, "Fehler beim Löschen"
    LoescheKartei = False
End Function

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Karteieinträge, die seit mehr als 7 Jahren keine Rechnung erhalten haben,
' und Löschen eines Eintrags samt allen abhängigen Datensätzen.
'
' Geschwindigkeit: Die ID-Listen (Rechnungen, Lieferscheine, Zahlungen ...) werden einmal
' gelesen und als feste Werte "IN (1,2,3)" eingesetzt. Unterabfragen über lokale und
' SQL-Server-Tabellen gemischt würden Access zwingen, ganze Server-Tabellen zu laden.

Private Const JAHRE As Integer = 7
Private Const KEINE As String = "(0)"   ' leere Liste, IDs sind nie 0

' ID-Listen des gerade bearbeiteten Karteieintrags
Private K As Long
Private Rech As String, Lief As String, ZE As String, ZA As String, ER As String, Proj As String

' Liste: letzte Ausgangsrechnung älter als 7 Jahre (ohne Rechnung zählt das Anlagedatum),
' und auch keine Eingangsrechnung in den letzten 7 Jahren. Gruppiert statt je Zeile abgefragt.
Private Function ListenSql() As String
    Dim Grenze As String
    Grenze = Format(DateAdd("yyyy", -JAHRE, Date), "\#mm\/dd\/yyyy\#")

    ListenSql = "SELECT K.ID, K.Nr, Trim(K.Name1 & ' ' & K.Name2) AS Kundenname, K.PLZ, K.Ort, K.AngelegtAm, " & _
          "R.LetzteRechnung, Nz(R.Anzahl, 0) AS AnzahlRechnungen, E.LetzteEingangsrechnung " & _
          "FROM (T_Kartei AS K " & _
          "LEFT JOIN (SELECT KarteiID, Max(Datum) AS LetzteRechnung, Count(*) AS Anzahl FROM T_Rechnungen GROUP BY KarteiID) AS R ON R.KarteiID = K.ID) " & _
          "LEFT JOIN (SELECT LiefKarteiID, Max(Datum) AS LetzteEingangsrechnung FROM T_EingangsRechnungen GROUP BY LiefKarteiID) AS E ON E.LiefKarteiID = K.ID " & _
          "WHERE Nz(R.LetzteRechnung, Nz(K.AngelegtAm, #1/1/1900#)) < " & Grenze & _
          " AND Nz(E.LetzteEingangsrechnung, #1/1/1900#) < " & Grenze & _
          " ORDER BY K.Name1, K.Name2"
End Function

Private Sub Form_Load()
    Me.RecordSource = ListenSql()
End Sub

Private Sub cmdLoeschen_Click()
    Dim Bezeichnung As String, Hindernis As String, Uebersicht As String

    If IsNull(Me!ID) Then Exit Sub
    Bezeichnung = "Nr " & Nz(Me!Nr, "-") & ": " & Nz(Me!Kundenname, "")

    DoCmd.Hourglass True
    LiesIdListen Me!ID
    Hindernis = FremdeVerweise()
    If Hindernis = "" Then Uebersicht = Zusammenfassung()
    DoCmd.Hourglass False

    If Hindernis <> "" Then
        MsgBox "Der Karteieintrag " & Bezeichnung & " kann nicht gelöscht werden, " & _
               "weil Datensätze anderer Karteieinträge darauf verweisen:" & vbCrLf & vbCrLf & Hindernis, _
               vbExclamation, "Löschen nicht möglich"
        Exit Sub
    End If

    If MsgBox("Karteieintrag " & Bezeichnung & " endgültig löschen?" & vbCrLf & vbCrLf & _
              "Mitgelöscht werden:" & vbCrLf & Uebersicht & vbCrLf & _
              "Das kann nicht rückgängig gemacht werden.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Karteieintrag löschen") <> vbYes Then Exit Sub

    DoCmd.Hourglass True
    If LoescheKartei() Then
        Me.Requery
        DoCmd.Hourglass False
        MsgBox "Karteieintrag " & Bezeichnung & " wurde gelöscht.", vbInformation
    End If
    DoCmd.Hourglass False
End Sub

' --- ID-Listen

' Liest die IDs einer Abfrage als "(1,2,3)"; ohne Treffer "(0)"
Private Function IdListe(Sql As String) As String
    Dim rs As DAO.Recordset, T As String
    Set rs = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot, dbSeeChanges)
    Do While Not rs.EOF
        T = T & "," & rs.Fields(0).Value
        rs.MoveNext
    Loop
    rs.Close
    If T = "" Then IdListe = KEINE Else IdListe = "(" & Mid(T, 2) & ")"
End Function

Private Sub LiesIdListen(KarteiID As Long)
    K = KarteiID
    Rech = IdListe("SELECT ID FROM T_Rechnungen WHERE KarteiID=" & K)
    Lief = IdListe("SELECT ID FROM T_Lieferscheine WHERE KarteiID=" & K)
    ZE = IdListe("SELECT ID FROM [T_Zahlungseingänge] WHERE KarteiID=" & K)
    ZA = IdListe("SELECT ID FROM [T_Zahlungsausgänge] WHERE KarteiID=" & K)
    ER = IdListe("SELECT ID FROM T_EingangsRechnungen WHERE LiefKarteiID=" & K)
    Proj = IdListe("SELECT ID FROM T_Projekte WHERE KarteiID=" & K)
End Sub

' Zählt Datensätze einer Tabelle zu einer Bedingung
Private Function Anzahl(Tabelle As String, Bedingung As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM " & Tabelle & " WHERE " & Bedingung, dbOpenSnapshot, dbSeeChanges)
    Anzahl = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' Hängt eine Zeile an, wenn die Anzahl größer 0 ist
Private Sub Zeile(ByRef Text As String, Anz As Long, Beschreibung As String)
    If Anz > 0 Then Text = Text & "  - " & Anz & " " & Beschreibung & vbCrLf
End Sub

' Prüft, ob Datensätze anderer Karteieinträge auf Daten dieses Eintrags verweisen.
' Dann wird nicht gelöscht, damit keine fremden Daten verloren gehen oder ins Leere zeigen.
Private Function FremdeVerweise() As String
    Dim T As String, Fremd As String
    Fremd = " AND Nz(KarteiID,0)<>" & K

    If Rech <> KEINE Then
        Zeile T, Anzahl("T_Rechnungen", "GegenRechnungID IN " & Rech & Fremd), "Rechnungen anderer Kunden mit Gegenrechnung auf diesen Kunden"
        Zeile T, Anzahl("T_Buchungen", "(RechnungID IN " & Rech & " OR GegenRechnungID IN " & Rech & ")" & Fremd), "Buchungen anderer Kunden auf Rechnungen dieses Kunden"
        Zeile T, Anzahl("T_LieferscheinPos", "RechnungID IN " & Rech & " AND NOT LieferscheinID IN " & Lief), "Lieferscheinpositionen anderer Kunden auf Rechnungen dieses Kunden"
        Zeile T, Anzahl("T_Leistungen_alt", "(AufRechnungID IN " & Rech & " OR AufHonorarRechnungID IN " & Rech & ")" & Fremd), "alte Leistungen anderer Kunden auf Rechnungen dieses Kunden"
        Zeile T, Anzahl("T_Rechnungen_ZE", "RechnungID IN " & Rech & " AND NOT ZEID IN " & ZE), "Zahlungszuordnungen zu Zahlungen anderer Kunden"
    End If
    If ZE <> KEINE Then
        Zeile T, Anzahl("T_Rechnungen", "ZahlungseingangID IN " & ZE & Fremd), "Rechnungen anderer Kunden mit Zahlungseingang dieses Kunden"
        Zeile T, Anzahl("T_Buchungen", "ZahlungseingangID IN " & ZE & Fremd), "Buchungen anderer Kunden auf Zahlungen dieses Kunden"
        Zeile T, Anzahl("T_Rechnungen_ZE", "ZEID IN " & ZE & " AND NOT RechnungID IN " & Rech), "Zahlungszuordnungen zu Rechnungen anderer Kunden"
    End If
    If Lief <> KEINE Then
        Zeile T, Anzahl("T_RechnungsPos", "LieferscheinID IN " & Lief & " AND NOT RechnungID IN " & Rech), "Rechnungspositionen anderer Kunden mit Lieferschein dieses Kunden"
    End If
    If ER <> KEINE Or ZA <> KEINE Then
        Zeile T, Anzahl("T_EingangsRechnungen_ZA", "(ERID IN " & ER & " AND NOT ZAID IN " & ZA & ") OR (ZAID IN " & ZA & " AND NOT ERID IN " & ER & ")"), "Zahlungsausgänge, die auch Eingangsrechnungen anderer Lieferanten betreffen"
    End If
    If Rech <> KEINE Or Proj <> KEINE Then
        ' Server-Tabelle: nur feste Werte, damit SQL Server filtert
        Zeile T, Anzahl("T_Leistungen", "(AufRechnungID IN " & Rech & " OR AufHonorarRechnungID IN " & Rech & " OR ProjektID IN " & Proj & ")" & Fremd), "Leistungen anderer Kunden auf Rechnungen oder Projekten dieses Kunden"
    End If
    FremdeVerweise = T
End Function

' Listet, was beim Löschen alles mit entfernt wird
Private Function Zusammenfassung() As String
    Dim T As String
    Zeile T, Anzahl("T_Rechnungen", "KarteiID=" & K), "Rechnungen"
    Zeile T, Anzahl("T_RechnungsPos", "RechnungID IN " & Rech & " OR LieferscheinID IN " & Lief), "Rechnungspositionen"
    Zeile T, Anzahl("T_Buchungen", "KarteiID=" & K & " OR RechnungID IN " & Rech & " OR ZahlungseingangID IN " & ZE), "Buchungen"
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
Private Function LoescheKartei() As Boolean
    Dim ws As DAO.Workspace, db As DAO.Database, Schritt As String
    Dim Befehle As Variant, i As Integer

    Befehle = Array( _
        "DELETE * FROM T_Rechnungen_ZE WHERE RechnungID IN " & Rech & " OR ZEID IN " & ZE, _
        "DELETE * FROM T_Buchungen WHERE KarteiID=" & K & " OR RechnungID IN " & Rech & " OR GegenRechnungID IN " & Rech & " OR ZahlungseingangID IN " & ZE, _
        "DELETE * FROM T_RechnungsPos WHERE RechnungID IN " & Rech & " OR LieferscheinID IN " & Lief, _
        "DELETE * FROM T_LieferscheinPos WHERE LieferscheinID IN " & Lief & " OR RechnungID IN " & Rech, _
        "DELETE * FROM T_Leistungen WHERE KarteiID=" & K, _
        "DELETE * FROM T_Leistungen_alt WHERE KarteiID=" & K, _
        "DELETE * FROM T_EingangsRechnungen_ZA WHERE ERID IN " & ER & " OR ZAID IN " & ZA, _
        "DELETE * FROM T_EingangsRechnungen WHERE LiefKarteiID=" & K, _
        "DELETE * FROM [T_Zahlungsausgänge] WHERE KarteiID=" & K, _
        "DELETE * FROM [T_Zahlungseingänge] WHERE KarteiID=" & K, _
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
    DoCmd.Hourglass False
    MsgBox "Löschen abgebrochen, es wurde nichts gelöscht." & vbCrLf & vbCrLf & _
           "Fehler: " & Meldung & vbCrLf & vbCrLf & "Bei: " & Schritt, vbCritical, "Fehler beim Löschen"
    LoescheKartei = False
End Function

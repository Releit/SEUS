Attribute VB_Name = "BelegUmwandlungen"
Option Compare Database
Option Explicit

Dim NewNr As Long

'** Lieferschein in Rechnung umwandeln **
'** Rückgabe der neuen ID **
Function mUW_LiefersZuRechnung(LieferscheinID As Long, KarteiID As Long, Optional ZuRechnungID As Long, Optional Buchen As Boolean) As Long
    Dim Sql As String, NewID As Long

    On Error GoTo mUW_LiefersZuRechnungErr

    If DCount("ID", "T_LieferscheinPos", "RechnungID IS NULL AND LieferscheinID=" & Str(LieferscheinID)) = 0 Then
        MsgBox "Es sind keine Postionen vorhanden die noch nicht auf einen Rechnung sind!", , "Rechnung anlegen"
        Exit Function
    End If

    If ZuRechnungID = 0 Then
        '** Rechnung erstellen **
        If Buchen Then NewNr = mNr_NewNr("RechnungNr", Date)
        Sql = "INSERT INTO T_Rechnungen (KarteiID, Nr, Datum, ZuHanden, AngelegtVon, Kunde )" & _
              " SELECT KarteiID" & _
              ", " & IIf(Buchen, Str(NewNr), "NULL") & _
              "," & IIf(Buchen, CDate2SQL(Date), "NULL") & _
              ", ZuHanden" & _
              ",'" & mWin_UserName() & "'" & _
              ",'" & mUW_Kunde(KarteiID) & "'" & _
              " FROM T_Lieferscheine WHERE ID=" & Str(LieferscheinID)

        CurrentDb.Execute Sql
        ZuRechnungID = DMax("ID", "T_Rechnungen")

        '** Kunde und Anschrift eintragen **

    End If

    mUW_LiefersZuRechnung = ZuRechnungID

    '** Positionen übernehmen **
    Sql = "INSERT INTO T_RechnungsPos ( RechnungID, LieferscheinID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
          " SELECT " & Str(ZuRechnungID) & _
          "," & Str(LieferscheinID) & _
          ", ProduktID" & _
          ", Beschreibung" & _
          ", Menge" & _
          ", Einheit" & _
          ", VKexkl" & _
          ", UstProz" & _
          ", Anmerkungen" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_LieferscheinPos WHERE RechnungID IS NULL AND LieferscheinID=" & Str(LieferscheinID)

    CurrentDb.Execute Sql

    '** RechnungID in Positionen eintragen **
    Sql = "UPDATE T_LieferscheinPos SET RechnungID=" & Str(ZuRechnungID) & _
          " WHERE RechnungID IS NULL AND LieferscheinID=" & Str(LieferscheinID)
    CurrentDb.Execute Sql

    If Buchen Then
        mUW_RechnungBuchen ZuRechnungID
    End If

mUW_LiefersZuRechnungExit:
    Exit Function

mUW_LiefersZuRechnungErr:
    MsgBox "Fehler in der Prozedur mUW_LiefersZuRechnung: " & Err.Description
    Resume mUW_LiefersZuRechnungExit
End Function

'** Lieferschein in Rechnung umwandeln **
'** Rückgabe der neuen ID **
Function mUW_RechnungZuLiefers(RechnungID As Long) As Long
    Dim Sql As String, NewID As Long

    On Error GoTo mUW_RechnungZuLiefersErr

    If Nz(DCount("ID", "T_RechnungsPos", "LieferscheinID IS NULL AND RechnungID=" & Str(RechnungID)), 0) = 0 Then
        MsgBox "Es sind keine Postionen vorhanden die noch nicht auf einen Lieferschein sind!", , "Lieferschein anlegen"
        Exit Function
    End If

    '** Rechnung erstellen **
    Sql = "INSERT INTO T_Lieferscheine (KarteiID, AdressenID, Nr, Datum, ZuHanden, AngelegtVon )" & _
          " SELECT KarteiID" & _
          ", AdressenID" & _
          ", " & Str(mNr_NewNr("LieferscheinNr", Date)) & _
          "," & CDate2SQL(Date) & _
          ", ZuHanden" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_Rechnungen WHERE ID=" & Str(RechnungID)

    CurrentDb.Execute Sql
    NewID = DMax("ID", "T_Lieferscheine")
    mUW_RechnungZuLiefers = NewID

    '** Positionen übernehmen **
    Sql = "INSERT INTO T_LieferscheinPos ( LieferscheinID, RechnungID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
          " SELECT " & Str(NewID) & _
          "," & Str(RechnungID) & _
          ", ProduktID" & _
          ", Beschreibung" & _
          ", Menge" & _
          ", Einheit" & _
          ", VKexkl" & _
          ", UstProz" & _
          ", Anmerkungen" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_RechnungsPos WHERE LieferscheinID IS NULL AND RechnungID=" & Str(RechnungID)

    CurrentDb.Execute Sql

    Sql = "UPDATE T_RechnungsPos SET LieferscheinID=" & Str(NewID) & " WHERE LieferscheinID IS NULL AND RechnungID=" & Str(RechnungID)
    CurrentDb.Execute Sql

mUW_RechnungZuLiefersExit:
    Exit Function

mUW_RechnungZuLiefersErr:
    MsgBox "Fehler in der Prozedur mUW_RechnungZuLiefers: " & Err.Description
    Resume mUW_RechnungZuLiefersExit
End Function

'** Gutschrift aus Rechnung erstellen **
'** Rückgabe der neuen ID **
Function mUW_RechnungZuGutschrift(RechnungID As Long, Optional Buchen As Boolean) As Long
    Dim Sql As String, NewID As Long

    On Error GoTo mUW_RechnungZuGutschriftErr

    '** Gutschrift erstellen **
    If Buchen Then NewNr = mNr_NewNr("RechnungNr", Date)
    Sql = "INSERT INTO T_Rechnungen (GegenRechnungID, KarteiID, AdressenID, Nr, Datum, ZuHanden, AngelegtVon )" & _
          " SELECT " & Str(RechnungID) & _
          ", KarteiID" & _
          ", AdressenID" & _
          ", " & IIf(Buchen, Str(NewNr), "NULL") & _
          "," & IIf(Buchen, CDate2SQL(Date), "NULL") & _
          ", ZuHanden" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_Rechnungen WHERE ID=" & Str(RechnungID)

    CurrentDb.Execute Sql
    NewID = DMax("ID", "T_Rechnungen")
    mUW_RechnungZuGutschrift = NewID

    '** Positionen übernehmen **
    Sql = "INSERT INTO T_RechnungsPos ( RechnungID, LieferscheinID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
          " SELECT " & Str(NewID) & _
          ",NULL" & _
          ", ProduktID" & _
          ", Beschreibung" & _
          ", -Menge" & _
          ", Einheit" & _
          ", VKexkl" & _
          ", UstProz" & _
          ", Anmerkungen" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_RechnungsPos WHERE RechnungID=" & Str(RechnungID)

    CurrentDb.Execute Sql

    If Buchen Then
        mUW_RechnungBuchen NewID, Date, DMin("BetragInkl", "V_Rechnungen", "RechnungID=" & Str(NewID)), "Gutschrift "
    End If

mUW_RechnungZuGutschriftExit:
    Exit Function

mUW_RechnungZuGutschriftErr:
    MsgBox "Fehler in der Prozedur mUW_RechnungZuGutschrift: " & Err.Description
    Resume mUW_RechnungZuGutschriftExit
End Function

Sub mUW_RechnungBuchen(RechnungID As Long, Optional Buchungsdatum As Date, Optional Betrag As Currency, Optional Buchungsart As String)
    Dim Sql As String

    If Buchungsdatum = 0 Then Buchungsdatum = Nz(DMin("Datum", "T_Rechnungen", "ID=" & Str(RechnungID)), 0)

    If Betrag = 0 Then
        '** eventuelle alte Buchung löschen
        CurrentDb.Execute "DELETE * FROM T_Buchungen WHERE Buchungsart='Rechnung' AND RechnungID=" & Str(RechnungID)
        '** neue Rechnung buchen
        Sql = "INSERT INTO T_Buchungen ( KarteiID, RechnungID, GegenRechnungID,  Buchungsart, Datum, Betrag, AngelegtVon )" & _
              " SELECT V_Rechnungen.KarteiID, V_Rechnungen.RechnungID, V_Rechnungen.GegenRechnungID, 'Rechnung'," & CDate2SQL(Buchungsdatum) & ",-V_Rechnungen.BetragInkl" & _
              ",'" & CurrentUser & "' FROM V_Rechnungen" & _
              " WHERE V_Rechnungen.RechnungID=" & Str(RechnungID)
    Else
        If Buchungsdatum = 0 Then Buchungsdatum = Date

        Sql = "INSERT INTO T_Buchungen ( KarteiID, RechnungID, Buchungsart, Datum, Betrag, AngelegtVon )" & _
              " SELECT V_Rechnungen.KarteiID, V_Rechnungen.RechnungID, '" & Buchungsart & "'," & CDate2SQL(Buchungsdatum) & "," & Str(Betrag) & _
              ",'" & CurrentUser & "' FROM V_Rechnungen" & _
              " WHERE V_Rechnungen.RechnungID=" & Str(RechnungID)
    End If

    CurrentDb.Execute Sql
    'Debug.Print Sql
End Sub

'** Leistungen in Rechnungen umwandeln **
'** Rückgabe der neuen ID **
Function mUW_LeistungenVerrechnen(Von As Date, Bis As Date, Optional KarteiID As Long, Optional ZuRechnungID As Long, Optional Buchen As Boolean) As Long
    Dim Sql As String, NewID As Long

    On Error GoTo mUW_LeistungenVerrechnenErr

    Sql = "AufRechnungID IS NULL AND Datum>=" & CDate2SQL(Von) & " AND Datum<=" & CDate2SQL(Von)
    If Nz(KarteiID, 0) <> 0 Then Sql = Sql & " AND KarteiID=" & Str(KarteiID)

    '** Prüfen ob Leistungen ohne Menge oder ohne Preis vorhanden sind
    If DCount("ID", "T_Leistungen", Sql & " AND (Menge IS NULL OR VKExkl IS NULL)") > 0 Then
        MsgBox "Bitte für Leistungen eine Menge und einen 'VK exkl.' eingeben!" & vbCrLf & vbCrLf & _
               "Für Leistungen ohne Verrechnung kann 'VK exkl.' = 0 eingeben werden.", , "Leistungen verrechnen"
        Exit Function
    End If

    '** Prüfen ob noch unverrechnte Leistungen mit Menge und Preis >0 vorhanden sind
    If DCount("ID", "T_Leistungen", Sql) = 0 Then
        MsgBox "Es sind keine unverrechneten Leistungen vorhanden!", , "Leistungen verrechnen"
        Exit Function
    End If

    If ZuRechnungID = 0 Then
        '** Rechnung erstellen **
        If Buchen Then NewNr = mNr_NewNr("RechnungNr", Date)
        Sql = "INSERT INTO T_Rechnungen (KarteiID, Nr, Datum, AngelegtVon )" & _
              " VALUES(" & Str(KarteiID) & _
              ", " & IIf(Buchen, Str(NewNr), "NULL") & _
              "," & IIf(Buchen, CDate2SQL(Date), "NULL") & _
              ",'" & mWin_UserName() & "')"

        CurrentDb.Execute Sql
        ZuRechnungID = DMax("ID", "T_Rechnungen")
    End If

    If KarteiID <> 0 Then mUW_LeistungenVerrechnen = ZuRechnungID

    '** Positionen übernehmen **
    Sql = "INSERT INTO T_RechnungsPos ( RechnungID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
          " SELECT " & Str(ZuRechnungID) & _
          ", ProduktID" & _
          ", Beschreibung" & _
          ", Menge" & _
          ", Einheit" & _
          ", VKexkl" & _
          ", UstProz" & _
          ", Anmerkungen" & _
          ",'" & mWin_UserName() & "'" & _
          " FROM T_Leistungen WHERE AufRechnungID IS NULL AND KundenID=" & Str(KarteiID)

    CurrentDb.Execute Sql

    '** RechnungID in Positionen eintragen **
    'Sql = "UPDATE T_LieferscheinPos SET RechnungID=" & Str(ZuRechnungID) & _
    '    " WHERE RechnungID IS NULL AND LieferscheinID=" & Str(LieferscheinID)
    'CurrentDb.Execute Sql

    If Buchen Then
        mUW_RechnungBuchen ZuRechnungID
    End If

mUW_LeistungenVerrechnenExit:
    Exit Function

mUW_LeistungenVerrechnenErr:
    MsgBox "Fehler in der Prozedur mUW_LeistungenVerrechnen: " & Err.Description
    Resume mUW_LeistungenVerrechnenExit
End Function

'** Leistungen in Rechnungen umwandeln **
'** Rückgabe der neuen ID **
Function mUW_VerrechnungZuRechnungen(VerrechnungID As Long, Optional ID As Long, Optional Buchen As Boolean) As Long
    Dim Sql As String, MaxID As Long, NewID As Long, AdressenID As Long
    Dim RS As Recordset, Kundenname As String, ProjektGruppierung As Boolean

    On Error GoTo mUW_VerrechnungZuRechnungenErr

    Sql = "SELECT * FROM T_VerrechnungKunden WHERE VerrechnungID=" & Str(VerrechnungID)
    If ID <> 0 Then Sql = Sql & " AND ID=" & Str(ID)
    Set RS = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)

    With RS
        Do While Not .EOF
            MaxID = DMax("ID", "T_Rechnungen")
            AdressenID = Nz(DMax("ID", "T_Adressen", "StandardFürRechnungen=True AND KarteiID=" & Str(!KarteiID)), 0)
            '** Rechnung erstellen **
            If Buchen Then NewNr = mNr_NewNr("RechnungNr", Date)
            Kundenname = mUW_Kunde(!KarteiID)
            Sql = "INSERT INTO T_Rechnungen (KarteiID, AdressenID, Nr, Datum, VerrechnungID, ZahlartID, Kunde, ZuHandenKontPersID, ZuHanden, AngelegtVon)" & _
                  " VALUES(" & Str(!KarteiID) & _
                  "," & IIf(AdressenID <> 0, Str(AdressenID), "NULL") & _
                  "," & IIf(Buchen, Str(NewNr), "NULL") & _
                  "," & IIf(Buchen, CDate2SQL(Date), "NULL") & _
                  "," & Str(VerrechnungID) & _
                  "," & Str(Nz(DMin("ID", "T_Zahlungsarten", "Standard=True"), 0)) & _
                  ",'" & Kundenname & "'" & _
                  "," & mUW_ZuHandenKontPersID(!KarteiID) & _
                  ",'" & mUW_ZuHanden(!KarteiID) & "'" & _
                  ",'" & mWin_UserName() & "')"

            CurrentDb.Execute Sql
            NewID = DMax("ID", "T_Rechnungen")

            If NewID = MaxID Then
                MsgBox "Fehler beim Anlegen eine neuen Rechnung!", , "mUW_VerrechnungZuRechnungen"
                Exit Function
            End If

            mUW_VerrechnungZuRechnungen = NewID

            '** Rechnungsadresse eintragen **
            mUW_RechnungSetAdresse NewID

            If !PositionenGruppieren Then
                If MsgBox(Kundenname & ":, mit Projektgruppierung?", vbYesNo, "Projektgruppierung") = vbYes Then ProjektGruppierung = True


                '** Positionen übernehmen **
                Sql = "INSERT INTO T_RechnungsPos ( RechnungID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
                      " SELECT " & Str(NewID) & _
                      ", T_Leistungen.ProduktID" & _
                      ", T_Produkte.Beschreibung" & _
                      ", Sum(Menge)" & _
                      ", T_Leistungen.Einheit" & _
                      ", T_Leistungen.VKexkl" & _
                      ", T_Leistungen.UstProz" & _
                      IIf(ProjektGruppierung, ", T_Leistungen.ProjektName", ", NULL") & _
                      ",'" & mWin_UserName() & "'" & _
                      " FROM T_Leistungen LEFT JOIN T_Produkte ON T_Produkte.ID=T_Leistungen.ProduktID" & _
                      " WHERE VerrKndID=" & Str(!ID) & _
                      " GROUP BY T_Leistungen.ProduktID, T_Produkte.Beschreibung, T_Leistungen.Einheit, T_Leistungen.VKexkl, T_Leistungen.UstProz"
                If ProjektGruppierung Then Sql = Sql & ", T_Leistungen.ProjektName"

            Else

                '** Positionen übernehmen **
                Sql = "INSERT INTO T_RechnungsPos ( RechnungID, ProduktID, Beschreibung, Menge, Einheit, VKexkl, UstProz, Anmerkungen, AngelegtVon )" & _
                      " SELECT " & Str(NewID) & _
                      ", T_Leistungen.ProduktID" & _
                      ", T_Produkte.Beschreibung" & _
                      ", T_Leistungen.Menge" & _
                      ", T_Leistungen.Einheit" & _
                      ", T_Leistungen.VKexkl" & _
                      ", T_Leistungen.UstProz" & _
                      ", A_LeistungenPostext.Postext" & _
                      ",'" & mWin_UserName() & "'" & _
                      " FROM (T_Leistungen INNER JOIN A_LeistungenPostext ON T_Leistungen.ID = A_LeistungenPostext.LeistungID)" & _
                      " LEFT JOIN T_Produkte ON T_Produkte.ID=T_Leistungen.ProduktID WHERE VerrKndID=" & Str(!ID)
            End If

            'Debug.Print Sql

            CurrentDb.Execute Sql

            '** RechnungID in Leistungen eintragen **
            Sql = "UPDATE T_Leistungen SET AufRechnungID=" & Str(NewID) & _
                  " WHERE VerrKndID=" & Str(!ID)

            CurrentDb.Execute Sql

            If Buchen Then
                mUW_RechnungBuchen NewID
            End If

            .MoveNext
        Loop
    End With

mUW_VerrechnungZuRechnungenExit:
    Exit Function

mUW_VerrechnungZuRechnungenErr:
    MsgBox "Fehler in der Prozedur mUW_VerrechnungZuRechnungen: " & Err.Description
    Resume mUW_VerrechnungZuRechnungenExit
End Function

Sub mUW_RechnungSetAdresse(RechnungID As Long)
    Dim KarteiID As Long, AdressenID As Long, Sql As String

    KarteiID = Nz(DLookup("KarteiID", "T_Rechnungen", "ID=" & Str(RechnungID)), 0)
    AdressenID = Nz(DLookup("AdressenID", "T_Rechnungen", "ID=" & Str(RechnungID)), 0)

    If AdressenID <> 0 Then
        Sql = "UPDATE T_Rechnungen INNER JOIN T_Adressen ON T_Rechnungen.AdressenID = T_Adressen.ID" & _
              " SET T_Rechnungen.Anschriftname = '" & mDF_AdressAnschriftName(Nz(AdressenID, 0)) & "'" & _
              ", T_Rechnungen.Adresse = T_Adressen.Adresse" & _
              ", T_Rechnungen.LandCode = T_Adressen.LandCode" & _
              ", T_Rechnungen.PLZ = T_Adressen.PLZ" & _
              ", T_Rechnungen.Ort = T_Adressen.Ort" & _
              " WHERE T_Rechnungen.ID=" & Str(RechnungID)
    Else
        Sql = "UPDATE T_Rechnungen INNER JOIN T_Kartei ON T_Kartei.ID=T_Rechnungen.KarteiID" & _
              " SET T_Rechnungen.Anschriftname = '" & mDF_KarteiAnschriftName(Nz(KarteiID, 0)) & "'" & _
              ", T_Rechnungen.Adresse = T_Kartei.Adresse" & _
              ", T_Rechnungen.LandCode = T_Kartei.LandCode" & _
              ", T_Rechnungen.PLZ = T_Kartei.PLZ" & _
              ", T_Rechnungen.Ort = T_Kartei.Ort" & _
              " WHERE T_Rechnungen.ID=" & Str(RechnungID)
    End If
    Debug.Print Sql
    CurrentDb.Execute Sql
End Sub

Function mUW_Kunde(KarteiID As Long) As String
    mUW_Kunde = Left(DMin("Text", "A_AuswahlKartei", "ID=" & Str(KarteiID)), 100)
End Function

Function mUW_ZuHandenKontPersID(KarteiID As Long) As Long
    mUW_ZuHandenKontPersID = Nz(DMin("ID", "T_Ansprechpartner", "BelegeZuHanden=True AND KarteiID=" & Str(KarteiID)), 0)
End Function

Function mUW_ZuHanden(KarteiID As Long) As String
    mUW_ZuHanden = Nz(DMin("Name1", "A_Ansprechpartner", "BelegeZuHanden=True AND KarteiID=" & Str(KarteiID)), "")
End Function

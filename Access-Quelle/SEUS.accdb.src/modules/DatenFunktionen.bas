Attribute VB_Name = "DatenFunktionen"
Option Compare Database
Option Explicit

Function mDF_KundenProduktpreiseLastID(KarteiID As Long, ProduktID As Long) As Long
    mDF_KundenProduktpreiseLastID = Nz(DMax("ID", "T_KundenProduktpreise", "KarteiID=" & Str(KarteiID) & " AND ProduktID=" & Str(ProduktID)), 0)
End Function

Function mDF_ProjektLeistungenStd(ProjektID As Long) As Currency
    mDF_ProjektLeistungenStd = Nz(DSum("Menge", "T_Leistungen", "LEFT(Einheit,3)='Std' AND ProjektID=" & Str(ProjektID)), 0)
End Function

Function mDF_ProduktMengeAusStunden(ProduktID As Long) As Boolean
    mDF_ProduktMengeAusStunden = Nz(DLookup("MengeAusStunden", "T_Produkte", "ID=" & Str(ProduktID)), 0)
End Function

Function mDF_ProjektLeistungenStdVerrechnet(ProjektID As Long) As Currency
    mDF_ProjektLeistungenStdVerrechnet = Nz(DSum("Menge", "T_Leistungen", "LEFT(Einheit,3)='Std' AND ProjektID=" & Str(ProjektID) & " AND NOT AufRechnungID IS NULL"), 0)
End Function

Function mDF_ProjektLeistungenEuro(ProjektID As Long) As Currency
    mDF_ProjektLeistungenEuro = Nz(DSum("Menge*VKExkl", "T_Leistungen", "ProjektID=" & Str(ProjektID)), 0)
End Function

Function mDF_ProjektLeistungenEuroVerrechnet(ProjektID As Long) As Currency
    mDF_ProjektLeistungenEuroVerrechnet = Nz(DSum("Menge*VKExkl", "T_Leistungen", "ProjektID=" & Str(ProjektID) & " AND NOT AufRechnungID IS NULL"), 0)
End Function

Function mDF_ReturnTextFromSQL(Sql As String) As String
    Dim RS As DAO.Recordset

    Set RS = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
    If Not RS.EOF Then mDF_ReturnTextFromSQL = Nz(RS.Fields(0), "")
End Function

Function mDF_ReturnLongFromSQL(Sql As String) As Long
    Dim RS As DAO.Recordset

    Set RS = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)
    If Not RS.EOF Then mDF_ReturnLongFromSQL = Nz(RS.Fields(0), 0)
End Function

Function mDF_ZE_RechungenBetrag(ZEID As Long) As Currency
    mDF_ZE_RechungenBetrag = Nz(DSum("Betrag", "T_Rechnungen_ZE", "ZEID=" & Str(ZEID)), 0)
End Function

Function mDF_Rechungen_ZEBetrag(RechnungID As Long) As Currency
    mDF_Rechungen_ZEBetrag = Nz(DSum("Betrag", "T_Rechnungen_ZE", "RechnungID=" & Str(RechnungID)), 0)
End Function

Function mDF_ER_BetragBezahlt(ERID As Long) As Currency
    mDF_ER_BetragBezahlt = Nz(DSum("Betrag", "T_EingangsRechnungen_ZA", "ERID=" & Str(ERID)), 0)
End Function

Function mDF_ER_BezahltAm(ERID As Long) As Date

    mDF_ER_BezahltAm = Nz(DMax("Datum", "A_ER_ZA", "ERID=" & Str(ERID)), 0)
End Function

Function mDF_ZEBetragZugeordnet(ZEID As Long) As Currency
    mDF_ZEBetragZugeordnet = Nz(DSum("Betrag", "T_Rechnungen_ZE", "ZEID=" & Str(ZEID)), 0)
End Function

Function mDF_ZABetragZugeordnet(ZAID As Long) As Currency
    mDF_ZABetragZugeordnet = Nz(DSum("Betrag", "T_EingangsRechnungen_ZA", "ZAID=" & Str(ZAID)), 0)
End Function

Function mDF_ZABetrag(ZAID As Long) As Currency
    mDF_ZABetrag = Nz(DSum("Betrag", "T_Zahlungsausgänge", "ID=" & Str(ZAID)), 0)
End Function

Function mDF_ZE_RechnungsNummern(ZEID As Long) As String
    Dim Sql As String, RS As DAO.Recordset, x As String

    Sql = "SELECT T_Rechnungen.Nr FROM T_Rechnungen INNER JOIN T_Rechnungen_ZE ON T_Rechnungen.ID = T_Rechnungen_ZE.RechnungID" & _
          " WHERE T_Rechnungen_ZE.ZEID=" & Str(ZEID)

    Set RS = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)

    With RS
        Do While Not .EOF
            x = x & !Nr & ", "
            .MoveNext
        Loop
    End With

    If Len(x) > 2 Then x = Left(x, Len(x) - 2)
    mDF_ZE_RechnungsNummern = x
End Function

Function mDF_KarteiAnschriftName(KarteiID As Long)
    mDF_KarteiAnschriftName = mDF_ReturnTextFromSQL("SELECT Nz([Titel] + ' ', '') & [Name1] & ' ' & (Chr(13) + Chr(10) + [Name2]) FROM T_Kartei WHERE ID=" & Str(KarteiID))
End Function

Function mDF_AdressAnschriftName(AdresseID As Long)
    mDF_AdressAnschriftName = mDF_ReturnTextFromSQL("SELECT Nz([Titel] + ' ', '') & [Name1] & ' ' & (Chr(13) + Chr(10) + [Name2]) FROM T_Adressen WHERE ID=" & Str(AdresseID))
End Function

Function mDF_RechnungAnschriftName(RechnungID As Long) As String
    Dim AdressenID As Long, KarteiID As Long

    AdressenID = Nz(DLookup("AdressenID", "T_Rechnungen", "ID=" & Str(RechnungID)), 0)

    If AdressenID <> 0 Then
        mDF_RechnungAnschriftName = mDF_AdressAnschriftName(AdressenID)
    Else
        mDF_RechnungAnschriftName = mDF_KarteiAnschriftName(Nz(DLookup("KarteiID", "T_Rechnungen", "ID=" & Str(RechnungID)), 0))
    End If
End Function

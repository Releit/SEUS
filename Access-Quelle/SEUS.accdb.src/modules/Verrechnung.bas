Attribute VB_Name = "Verrechnung"
Option Compare Database
Option Explicit

Function mVR_FillVerrechnungKunden(sFilter As String, Optional ProjektGruppierung As Boolean) As Long
    Dim Sql As String, NewID As Long, Filter  As String

    sFilter = sFilter & " AND NZ(Menge,0) <>0 AND NZ(VKexkl,0)<>0 AND AufRechnungID IS NULL"

    '** Leistungen die zu gelöschten Rechnungen zugeordnet sind wieder freigeben **
    Sql = "UPDATE T_Leistungen LEFT JOIN T_Rechnungen ON T_Leistungen.AufRechnungID = T_Rechnungen.ID SET T_Leistungen.AufRechnungID = Null" & _
          " WHERE (((T_Leistungen.AufRechnungID) Is Not Null) AND ((T_Rechnungen.ID) Is Null));"
    CurrentDb.Execute Sql

    If DCount("ID", "T_Leistungen", sFilter) > 0 Then

        Sql = "INSERT INTO T_Verrechnung (DatumVon, DatumBis, AngelegtVon)" & _
              " SELECT Min(T_Leistungen.Datum) AS MinvonDatum, Max(T_Leistungen.Datum) AS MaxvonDatum" & _
              " ,'" & mWin_UserName() & "'" & _
              " FROM T_Leistungen LEFT JOIN T_Projekte ON T_Projekte.ID=T_Leistungen.ProjektID" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter

        CurrentDb.Execute Sql
        NewID = DMax("ID", "T_Verrechnung")

        '** Aufräumen **
        Sql = "UPDATE T_Leistungen SET AufRechnungID = Null WHERE AufRechnungID NOT IN (SELECT ID FROM T_Rechnungen);"
        CurrentDb.Execute Sql

        Sql = "UPDATE T_Leistungen SET VerrKndID = Null WHERE VerrKndID NOT IN (SELECT ID FROM T_VerrechnungKunden);"
        CurrentDb.Execute Sql

        Sql = "DELETE * FROM T_Verrechnung WHERE ID NOT IN (SELECT VerrechnungID FROM T_Rechnungen)"
        CurrentDb.Execute Sql


        '** Datensätze in T_VerrechnungKunden eintragen **
        Sql = "INSERT INTO T_VerrechnungKunden ( VerrechnungID, KarteiID, ProjektID, DatumVon, DatumBis, Kunde, AnzahlPos, Betrag, AngelegtVon)" & _
              " SELECT " & Str(NewID) & ",T_Leistungen.KarteiID" & IIf(ProjektGruppierung, ",T_Leistungen.ProjektID", ",NULL") & _
              ", Min(T_Leistungen.Datum), Max(T_Leistungen.Datum)" & _
              " ,[Name1] & ' ' & [Name2], Count(T_Leistungen.ID), Sum([Menge]*[VKexkl]) " & _
              " ,'" & mWin_UserName() & "'" & _
              " FROM  (T_Kartei INNER JOIN T_Leistungen ON T_Kartei.ID = T_Leistungen.KarteiID)" & _
              " LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter
        Sql = Sql & " GROUP BY T_Leistungen.KarteiID, [Name1] & ' ' & [Name2]"

        If ProjektGruppierung Then Sql = Sql & ", T_Leistungen.ProjektID"

        'Debug.Print Sql

        CurrentDb.Execute Sql

        '** Leistungsdatensätze der richtigen ID in T_VerrechnungKunden zuordnen (zur Erstellung der Rechnung) **
        Sql = "UPDATE (T_Leistungen INNER JOIN T_VerrechnungKunden ON T_Leistungen.KarteiID = T_VerrechnungKunden.KarteiID" & _
              IIf(ProjektGruppierung, " AND T_Leistungen.ProjektID = T_VerrechnungKunden.ProjektID)", ")") & _
              " LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID" & _
              " SET T_Leistungen.VerrKndID = [T_VerrechnungKunden].[ID]" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter & " AND T_VerrechnungKunden.VerrechnungID=" & Str(NewID)

        'Debug.Print Sql

        CurrentDb.Execute Sql

        mVR_FillVerrechnungKunden = NewID
    Else
        MsgBox "Keine zu verrechnenden Leistungen vorhanden", , "Leistungen verrechnen"
    End If

End Function

Function mVR_FillVerrechnungMitarbeiter(sFilter As String, Optional ProjektGruppierung As Boolean) As Long
    Dim Sql As String, NewID As Long, Filter  As String, SeusKarteiID As Long

    SeusKarteiID = 796

    sFilter = sFilter & " AND NZ(Menge,0) <>0 AND NZ(VKexkl,0)<>0 AND AufHonorarRechnungID IS NULL"

    '** Honorare die zu gelöschten Rechnungen zugeordnet sind wieder freigeben **
    Sql = "UPDATE T_Leistungen LEFT JOIN T_Rechnungen ON T_Leistungen.AufHonorarRechnungID = T_Rechnungen.ID SET T_Leistungen.AufHonorarRechnungID = Null" & _
          " WHERE (((T_Leistungen.AufHonorarRechnungID) Is Not Null) AND ((T_Rechnungen.ID) Is Null));"
    CurrentDb.Execute Sql

    If DCount("ID", "T_Leistungen", sFilter) > 0 Then

        Sql = "INSERT INTO T_Verrechnung (DatumVon, DatumBis, HonorarVerrechnung, AngelegtVon)" & _
              " SELECT Min(T_Leistungen.Datum) AS MinvonDatum, Max(T_Leistungen.Datum) AS MaxvonDatum" & _
              " , True, '" & mWin_UserName() & "'" & _
              " FROM T_Leistungen LEFT JOIN T_Projekte ON T_Projekte.ID=T_Leistungen.ProjektID" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter

        CurrentDb.Execute Sql
        NewID = DMax("ID", "T_Verrechnung")

        '** Aufräumen **
        Sql = "UPDATE T_Leistungen SET AufHonorarRechnungID = Null WHERE AufHonorarRechnungID NOT IN (SELECT ID FROM T_Rechnungen);"
        CurrentDb.Execute Sql

        Sql = "UPDATE T_Leistungen SET VerrHonorarID = Null WHERE VerrHonorarID NOT IN (SELECT ID FROM T_VerrechnungKunden);"
        CurrentDb.Execute Sql

        Sql = "DELETE * FROM T_Verrechnung WHERE ID NOT IN (SELECT VerrechnungID FROM T_Rechnungen)"
        CurrentDb.Execute Sql


        '** Datensätze in T_VerrechnungKunden eintragen **
        Sql = "INSERT INTO T_VerrechnungKunden ( VerrechnungID, KarteiID, ProjektID, DatumVon, DatumBis, Kunde, AnzahlPos, Betrag, AngelegtVon)" & _
              " SELECT " & Str(NewID) & "," & Str(SeusKarteiID) & IIf(ProjektGruppierung, ",T_Leistungen.ProjektID", ",NULL") & _
              ", Min(T_Leistungen.Datum), Max(T_Leistungen.Datum)" & _
              " ,[Name1] & ' ' & [Name2], Count(T_Leistungen.ID), Sum([Menge]*[VKexkl]) " & _
              " ,'" & mWin_UserName() & "'" & _
              " FROM (T_Kartei INNER JOIN T_Leistungen ON T_Kartei.ID = " & Str(SeusKarteiID) & ")" & _
              " LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter
        Sql = Sql & " GROUP BY T_Leistungen.KarteiID, [Name1] & ' ' & [Name2]"

        If ProjektGruppierung Then Sql = Sql & ", T_Leistungen.ProjektID"

        'Debug.Print Sql

        CurrentDb.Execute Sql

        '** Leistungsdatensätze der richtigen ID in T_VerrechnungKunden zuordnen (zur Erstellung der Rechnung) **
        Sql = "UPDATE (T_Leistungen INNER JOIN T_VerrechnungKunden ON T_Leistungen.KarteiID = T_VerrechnungKunden.KarteiID" & _
              IIf(ProjektGruppierung, " AND T_Leistungen.ProjektID = T_VerrechnungKunden.ProjektID)", ")") & _
              " LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID" & _
              " SET T_Leistungen.VerrKndID = [T_VerrechnungKunden].[ID]" & _
              " WHERE NZ(T_Projekte.Pauschalverrechnung,0)=0 AND " & sFilter & " AND T_VerrechnungKunden.VerrechnungID=" & Str(NewID)

        'Debug.Print Sql

        CurrentDb.Execute Sql

        mVR_FillVerrechnungMitarbeiter = NewID
    Else
        MsgBox "Keine zu verrechnenden Leistungen vorhanden", , "Leistungen verrechnen"
    End If

End Function

Attribute VB_Name = "RechnungenDaten"
Option Compare Database
Option Explicit

Function mRD_BetragInkl(RechnungID As Long) As Currency
    mRD_BetragInkl = Nz(DSum("BetragInkl", "V_Rechnungen", "RechnungID=" & Str(RechnungID)), 0)
End Function

Function mRD_RechnungsDatum(ZEID As Long) As String
    Dim Sql As String, RS As DAO.Recordset, x As String

    Sql = "SELECT T_Rechnungen.Datum FROM T_Rechnungen INNER JOIN T_Rechnungen_ZE ON T_Rechnungen.ID = T_Rechnungen_ZE.RechnungID" & _
          " WHERE T_Rechnungen_ZE.ZEID=" & Str(ZEID)

    Set RS = CurrentDb.OpenRecordset(Sql, dbOpenSnapshot)

    With RS
        Do While Not .EOF
            x = x & Format(!Datum, "DD.MM.YYYY") & ", "
            .MoveNext
        Loop
    End With

    If Len(x) > 2 Then x = Left(x, Len(x) - 2)
    mRD_RechnungsDatum = x
End Function

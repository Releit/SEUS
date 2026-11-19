SELECT
  CLng(
    Nz(
      [T_Buchungen].[GegenRechnungID],
      [T_Buchungen].[RechnungID]
    )
  ) AS RechnungID,
  Sum(T_Buchungen.Betrag) AS Saldo
FROM
  T_Buchungen
GROUP BY
  CLng(
    Nz(
      [T_Buchungen].[GegenRechnungID],
      [T_Buchungen].[RechnungID]
    )
  )
ORDER BY
  CLng(
    Nz(
      [T_Buchungen].[GegenRechnungID],
      [T_Buchungen].[RechnungID]
    )
  );

SELECT
  T_Buchungen.RechnungID,
  Sum(T_Buchungen.Betrag) AS SummevonBetrag
FROM
  T_Buchungen
GROUP BY
  T_Buchungen.RechnungID;

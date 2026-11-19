SELECT
  A_RechnungenMitBetrag.RechnungID,
  A_RechnungenMitBetrag.Datum,
  A_RechnungenMitBetrag.Nr,
  A_RechnungenMitBetrag.Kunde,
  A_RechnungenMitBetrag.PLZ,
  A_RechnungenMitBetrag.Ort,
  A_RechnungenMitBetrag.Adresse,
  Round(
    [A_RechnungenMitBetrag].[RechnungsbetragInkl],
    2
  ) AS BetragRechnungInkl,
  Round([SummevonBetrag], 2) AS BetragGebucht
FROM
  A_BuchungenNachRechnungID
  RIGHT JOIN A_RechnungenMitBetrag ON A_BuchungenNachRechnungID.RechnungID = A_RechnungenMitBetrag.RechnungID;

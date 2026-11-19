SELECT
  T_Rechnungen.ID,
  T_Rechnungen.Nr,
  T_Rechnungen.Datum,
  Nz([anschrift], [Kunde]) AS NameAnschrift,
  T_Rechnungen.Anmerkungen,
  Sum([menge] * [vkexkl]) AS Netto,
  Sum(
    [menge] * [vkexkl] *(1 + [Ustproz])
  ) AS Brutto
FROM
  T_Rechnungen
  INNER JOIN T_RechnungsPos ON T_Rechnungen.ID = T_RechnungsPos.RechnungID
GROUP BY
  T_Rechnungen.ID,
  T_Rechnungen.Nr,
  T_Rechnungen.Datum,
  Nz([anschrift], [Kunde]),
  T_Rechnungen.Anmerkungen;

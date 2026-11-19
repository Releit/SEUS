SELECT
  CStr(
    Year(T_Leistungen.Datum)
  )& ' / ' & Month(T_Leistungen.Datum) AS Datum,
  T_Produkte.Beschreibung,
  Sum(T_Leistungen.Menge) AS Mengen,
  T_Leistungen.Einheit
FROM
  T_Leistungen
  INNER JOIN T_Produkte ON T_Produkte.ID = T_Leistungen.ProduktID
GROUP BY
  T_Produkte.Beschreibung,
  T_Leistungen.Einheit,
  T_Leistungen.MitarbeiterID,
  Month(T_Leistungen.Datum),
  Year(T_Leistungen.Datum)
HAVING
  T_Leistungen.MitarbeiterID = 2
  AND Month(T_Leistungen.Datum) = 9
  AND Year(T_Leistungen.Datum) = 2025;

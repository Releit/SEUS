SELECT
  T_Leistungen.Datum,
  T_Produkte.Beschreibung,
  Sum(T_Leistungen.Menge) AS SummevonMenge,
  T_Leistungen.Einheit,
  T_Leistungen.VKExkl,
  T_Leistungen.AufRechnungID
FROM
  T_Produkte
  INNER JOIN T_Leistungen ON T_Produkte.ID = T_Leistungen.ProduktID
WHERE
  (
    (
      (T_Leistungen.ProjektID) = mBR_GetID()
    )
  )
GROUP BY
  T_Leistungen.Datum,
  T_Produkte.Beschreibung,
  T_Leistungen.Einheit,
  T_Leistungen.VKExkl,
  T_Leistungen.AufRechnungID;

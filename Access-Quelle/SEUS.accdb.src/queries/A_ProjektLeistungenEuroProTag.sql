SELECT
  T_Leistungen.Datum,
  Sum(
    [Menge] * [T_Leistungen].[VKexkl]
  ) AS Gesamt
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
  T_Leistungen.Datum
ORDER BY
  T_Leistungen.Datum;

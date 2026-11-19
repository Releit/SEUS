SELECT
  T_Zahlungseingänge.Datum,
  [Betrag] / 1.2 AS Netto,
  T_Zahlungseingänge.Betrag,
  T_Kartei.Name1,
  T_Kartei.Name2
FROM
  T_Kartei
  INNER JOIN T_Zahlungseingänge ON (
    T_Kartei.ID = T_Zahlungseingänge.KarteiID
  )
  AND (
    T_Kartei.ID = T_Zahlungseingänge.KarteiID
  )
WHERE
  (
    (
      (T_Zahlungseingänge.Betrag) <> 0
    )
    AND (
      (
        Year([Datum]) * 100 + Month([Datum])
      ) >= 201901
    )
  )
ORDER BY
  T_Zahlungseingänge.Datum;

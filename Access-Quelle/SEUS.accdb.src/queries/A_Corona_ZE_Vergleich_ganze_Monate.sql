SELECT
  Year([Datum]) * 100 + Month([Datum]) AS JahrMonat,
  Sum([Betrag] / 1.2) AS Monatssumme,
  Min(T_Zahlungseingänge.Datum) AS MinvonDatum,
  Max(T_Zahlungseingänge.Datum) AS MaxvonDatum
FROM
  T_Zahlungseingänge
GROUP BY
  Year([Datum]) * 100 + Month([Datum])
HAVING
  (
    (
      (
        Year([Datum]) * 100 + Month([Datum])
      ) >= 201901
    )
  );

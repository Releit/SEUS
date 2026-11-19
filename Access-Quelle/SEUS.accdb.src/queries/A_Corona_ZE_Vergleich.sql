SELECT
  Year([Datum] - 15) * 100 + Month([Datum] - 15) AS JahrMonat,
  Sum([Betrag] / 1.2) AS Monatssumme,
  Min(T_Zahlungseingänge.Datum) AS MinvonDatum,
  Max(T_Zahlungseingänge.Datum) AS MaxvonDatum
FROM
  T_Zahlungseingänge
GROUP BY
  Year([Datum] - 15) * 100 + Month([Datum] - 15)
HAVING
  (
    (
      (
        Year([Datum] - 15) * 100 + Month([Datum] - 15)
      ) >= 201901
    )
  );

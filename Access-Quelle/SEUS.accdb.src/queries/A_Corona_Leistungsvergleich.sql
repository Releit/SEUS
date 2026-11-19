SELECT
  Year([Datum] - 15) * 100 + Month([Datum] - 15) AS JahrMonat,
  Sum([Menge] * [VKexkl]) AS Monatssumme,
  Min(T_Leistungen.datum) AS Minvondatum,
  Max(T_Leistungen.datum) AS Maxvondatum
FROM
  T_Leistungen
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

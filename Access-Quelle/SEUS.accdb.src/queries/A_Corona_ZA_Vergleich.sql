SELECT
  Year(
    [T_Zahlungsausgänge].[Datum] - 15
  ) * 100 + Month(
    [T_Zahlungsausgänge].[Datum] - 15
  ) AS JahrMonat,
  T_AusgabenArten.Ausgabeart,
  Sum(T_Zahlungsausgänge.Betrag) AS Monatssumme,
  Min(T_Zahlungsausgänge.Datum) AS MinvonDatum,
  Max(T_Zahlungsausgänge.Datum) AS MaxvonDatum
FROM
  (
    (
      T_EingangsRechnungen
      INNER JOIN T_EingangsRechnungen_ZA ON T_EingangsRechnungen.ID = T_EingangsRechnungen_ZA.ERID
    )
    INNER JOIN T_Zahlungsausgänge ON T_EingangsRechnungen_ZA.ZAID = T_Zahlungsausgänge.ID
  )
  INNER JOIN T_AusgabenArten ON (
    T_EingangsRechnungen.AusgabenArtID = T_AusgabenArten.ID
  )
  AND (
    T_AusgabenArten.ID = T_EingangsRechnungen.AusgabenArtID
  )
WHERE
  (
    (
      (T_AusgabenArten.CoronaRelevant) = Yes
    )
  )
GROUP BY
  Year(
    [T_Zahlungsausgänge].[Datum] - 15
  ) * 100 + Month(
    [T_Zahlungsausgänge].[Datum] - 15
  ),
  T_AusgabenArten.Ausgabeart
HAVING
  (
    (
      (
        Year(
          [T_Zahlungsausgänge].[Datum] - 15
        ) * 100 + Month(
          [T_Zahlungsausgänge].[Datum] - 15
        )
      ) >= 201901
    )
  );

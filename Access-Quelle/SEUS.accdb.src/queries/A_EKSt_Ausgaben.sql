SELECT
  T_Zahlungsausgänge.LfdNr,
  T_Zahlungsausgänge.Datum,
  T_Zahlungsarten.Zahlungsart,
  T_Zahlungsausgänge.Betrag AS ZABetrag,
  T_EingangsRechnungen_ZA.Betrag AS BetragZuR,
  [Name1] & " " & [Name2] AS Lieferant,
  T_EingangsRechnungen.Datum AS R_Datum,
  T_EingangsRechnungen.Brutto,
  T_EingangsRechnungen.Ust10Proz,
  T_EingangsRechnungen.Ust20Proz,
  T_EingangsRechnungen.Skonto,
  T_AusgabenArten.Ausgabeart,
  T_EingangsRechnungen.InnerGemErwerb AS IGE,
  T_EingangsRechnungen.PfadDatei
FROM
  (
    (
      (
        (
          T_Zahlungsausgänge
          LEFT JOIN T_EingangsRechnungen_ZA ON T_Zahlungsausgänge.ID = T_EingangsRechnungen_ZA.ZAID
        )
        LEFT JOIN T_EingangsRechnungen ON T_EingangsRechnungen_ZA.ERID = T_EingangsRechnungen.ID
      )
      LEFT JOIN T_Zahlungsarten ON T_Zahlungsausgänge.ZahlungsartID = T_Zahlungsarten.ID
    )
    LEFT JOIN T_AusgabenArten ON T_EingangsRechnungen.AusgabenArtID = T_AusgabenArten.ID
  )
  LEFT JOIN T_Kartei ON T_EingangsRechnungen.LiefKarteiID = T_Kartei.ID
WHERE
  (
    (
      (
        Year([T_Zahlungsausgänge].[Datum])
      ) = [Jahr eingeben]
    )
  )
ORDER BY
  T_Zahlungsausgänge.Datum;

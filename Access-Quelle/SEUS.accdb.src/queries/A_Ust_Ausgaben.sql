SELECT
  T_EingangsRechnungen.BezahltAm,
  T_Kartei.Name1 AS Lieferant,
  CDbl(
    Nz(
      [BetragBezahlt], [T_EingangsRechnungen].[Brutto]
    )
  ) AS Betrag,
  Round(
    Nz([Ust10Proz], 0) + Nz([Ust20Proz], 0),
    2
  ) AS VSt,
  T_EingangsRechnungen.Nr,
  CDbl(
    IIf([Ausgabeart] = "KFZ", [VSt], 0)
  ) AS KFZ,
  CDbl(
    IIf([InnerGemErwerb], [Brutto], 0)
  ) AS IGE
FROM
  (
    T_EingangsRechnungen
    INNER JOIN T_Kartei ON T_EingangsRechnungen.LiefKarteiID = T_Kartei.ID
  )
  LEFT JOIN T_AusgabenArten ON T_EingangsRechnungen.AusgabenArtID = T_AusgabenArten.ID
WHERE
  (
    (
      (T_EingangsRechnungen.BezahltAm) Between [von]
      And [bis]
    )
  )
ORDER BY
  T_EingangsRechnungen.BezahltAm;

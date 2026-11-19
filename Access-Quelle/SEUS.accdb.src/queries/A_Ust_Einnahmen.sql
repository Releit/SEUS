SELECT
  T_Zahlungseingänge.Datum,
  T_Kartei.Name1 AS Kunde,
  T_Rechnungen.Nr AS RechnungNr,
  T_Rechnungen.Datum AS [R-Datum],
  T_Rechnungen_ZE.Betrag,
  T_Rechnungen.InnerGemErwerb,
  Round(
    [T_Rechnungen_ZE].[Betrag] /(
      1 + Nz(
        [T_UstRegelungen].[UstProz], 0.2
      )
    ) * Nz(
      [T_UstRegelungen].[UstProz], 0.2
    ),
    2
  ) AS USt
FROM
  (
    T_Kartei
    INNER JOIN T_Zahlungseingänge ON T_Kartei.ID = T_Zahlungseingänge.KarteiID
  )
  INNER JOIN (
    (
      T_Rechnungen
      LEFT JOIN T_UstRegelungen ON T_Rechnungen.UStRegelungID = T_UstRegelungen.ID
    )
    INNER JOIN T_Rechnungen_ZE ON T_Rechnungen.ID = T_Rechnungen_ZE.RechnungID
  ) ON T_Zahlungseingänge.ID = T_Rechnungen_ZE.ZEID
WHERE
  (
    (
      (T_Zahlungseingänge.Datum) Between [von]
      And [bis]
    )
  );

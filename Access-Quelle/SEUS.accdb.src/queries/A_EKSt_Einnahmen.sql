SELECT
  T_Zahlungseingänge.LfdNr,
  T_Zahlungseingänge.Datum,
  [Name1] & " " & [Name2] AS Kunde,
  T_Zahlungseingänge.Betrag,
  mDF_ZE_RechnungsNummern([T_Zahlungseingänge].[ID]) AS RNr,
  mRD_RechnungsDatum([T_Zahlungseingänge].[ID]) AS RDatum
FROM
  T_Kartei
  INNER JOIN T_Zahlungseingänge ON T_Kartei.ID = T_Zahlungseingänge.KarteiID
WHERE
  (
    (
      (
        Year([T_Zahlungseingänge].[Datum])
      ) = [Jahr eingeben]
    )
  )
ORDER BY
  T_Zahlungseingänge.Datum;

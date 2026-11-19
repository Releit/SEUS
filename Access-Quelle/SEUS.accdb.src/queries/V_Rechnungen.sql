SELECT
  T_Rechnungen.ID AS RechnungID,
  T_Rechnungen.GegenRechnungID,
  T_Rechnungen.VerrechnungID,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  T_Rechnungen.KarteiID,
  T_Kartei.Nr AS KundenNr,
  [T_Kartei].[Name1] & ' ' & [T_Kartei].[Name2] AS KundenName,
  Nz(
    [T_Adressen].[Name1], [T_Kartei].[Name1]
  )& ' ' & Nz(
    [T_Adressen].[Name2], [T_Kartei].[Name2]
  ) AS Anschriftname,
  Nz(
    [T_Adressen].[Adresse], [T_Kartei].[Adresse]
  ) AS Adresse,
  Nz(
    [T_Adressen].[Anrede], [T_Kartei].[Anrede]
  ) AS Anrede,
  Nz(
    [T_Adressen].[PLZ], [T_Kartei].[PLZ]
  ) AS PLZ,
  Nz(
    [T_Adressen].[Ort], [T_Kartei].[Ort]
  ) AS Ort,
  Round(
    Sum(
      Round(
        [Menge] * [VKExkl] *(1 - [RabattProz]),
        4
      )
    ),
    2
  ) AS BetragExkl,
  Round(
    Sum(
      Round(
        [UstProz] * [VKexkl] * [menge] *(1 - [RabattProz]),
        4
      )
    ),
    2
  ) AS USt,
  [BetragExkl] + [USt] AS BetragInkl,
  T_Rechnungen.DruckText,
  T_Rechnungen.MahnStufe,
  T_Rechnungen.Mahnung1Am,
  T_Rechnungen.Mahnung2Am,
  T_Rechnungen.Mahnung3Am,
  T_Rechnungen.RechtsanwaltAm,
  T_Rechnungen.ZuHanden,
  Sum(
    Round(
      [Menge] * [VKExkl] *([RabattProz]),
      4
    )
  ) AS RabattExkl,
  T_Kartei.Name1,
  T_Kartei.Name2,
  Year(
    Nz(
      [T_Rechnungen].[Datum],
      Date()
    )
  ) AS Jahr,
  mDF_Rechungen_ZEBetrag([T_Rechnungen].[ID]) AS Bezahlt,
  T_Rechnungen.InnerGemErwerb
FROM
  (
    T_Kartei
    INNER JOIN (
      T_Adressen
      RIGHT JOIN T_Rechnungen ON T_Adressen.ID = T_Rechnungen.AdressenID
    ) ON T_Kartei.ID = T_Rechnungen.KarteiID
  )
  LEFT JOIN T_RechnungsPos ON T_Rechnungen.ID = T_RechnungsPos.RechnungID
GROUP BY
  T_Rechnungen.ID,
  T_Rechnungen.GegenRechnungID,
  T_Rechnungen.VerrechnungID,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  T_Rechnungen.KarteiID,
  T_Kartei.Nr,
  [T_Kartei].[Name1] & ' ' & [T_Kartei].[Name2],
  Nz(
    [T_Adressen].[Name1], [T_Kartei].[Name1]
  )& ' ' & Nz(
    [T_Adressen].[Name2], [T_Kartei].[Name2]
  ),
  Nz(
    [T_Adressen].[Adresse], [T_Kartei].[Adresse]
  ),
  Nz(
    [T_Adressen].[Anrede], [T_Kartei].[Anrede]
  ),
  Nz(
    [T_Adressen].[PLZ], [T_Kartei].[PLZ]
  ),
  Nz(
    [T_Adressen].[Ort], [T_Kartei].[Ort]
  ),
  T_Rechnungen.DruckText,
  T_Rechnungen.MahnStufe,
  T_Rechnungen.Mahnung1Am,
  T_Rechnungen.Mahnung2Am,
  T_Rechnungen.Mahnung3Am,
  T_Rechnungen.RechtsanwaltAm,
  T_Rechnungen.ZuHanden,
  T_Kartei.Name1,
  T_Kartei.Name2,
  Year(
    Nz(
      [T_Rechnungen].[Datum],
      Date()
    )
  ),
  mDF_Rechungen_ZEBetrag([T_Rechnungen].[ID]),
  T_Rechnungen.InnerGemErwerb;

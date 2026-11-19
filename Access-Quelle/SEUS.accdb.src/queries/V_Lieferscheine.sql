SELECT
  T_Lieferscheine.ID AS LieferscheinID,
  T_Lieferscheine.Datum,
  T_Lieferscheine.Nr,
  T_Kartei.Nr AS KNr,
  T_Lieferscheine.KarteiID,
  Nz(
    [T_Adressen].[Anrede], [T_Kartei].[Anrede]
  ) AS Anrede,
  T_Lieferscheine.ZuHanden,
  [T_Kartei].[Name1] & ' ' & [T_Kartei].[Name2] AS Kundenname,
  Nz(
    [T_Adressen].[Name1], [T_Kartei].[Name1]
  )& ' ' & Nz(
    [T_Adressen].[Name2], [T_Kartei].[Name2]
  ) AS Anschriftname,
  Nz(
    [T_Adressen].[Adresse], [T_Kartei].[Adresse]
  ) AS Adresse,
  Nz(
    [T_Adressen].[PLZ], [T_Kartei].[PLZ]
  ) AS PLZ,
  Nz(
    [T_Adressen].[Ort], [T_Kartei].[Ort]
  ) AS Ort,
  Sum(
    Round(
      [Menge] * [VKExkl] *(1 - [RabattProz]),
      4
    )
  ) AS BetragExkl,
  Sum(
    Round(
      [Menge] * [VKExkl] *(1 - [RabattProz]) *(1 + [UstProz]),
      4
    )
  ) AS BetragInkl,
  Sum(
    - IsNull([RechnungID])
  ) AS Unverrechnet
FROM
  (
    T_Kartei
    INNER JOIN (
      T_Adressen
      RIGHT JOIN T_Lieferscheine ON T_Adressen.ID = T_Lieferscheine.AdressenID
    ) ON T_Kartei.ID = T_Lieferscheine.KarteiID
  )
  LEFT JOIN T_LieferscheinPos ON T_Lieferscheine.ID = T_LieferscheinPos.LieferscheinID
GROUP BY
  T_Lieferscheine.ID,
  T_Lieferscheine.Datum,
  T_Lieferscheine.Nr,
  T_Kartei.Nr,
  T_Lieferscheine.KarteiID,
  Nz(
    [T_Adressen].[Anrede], [T_Kartei].[Anrede]
  ),
  T_Lieferscheine.ZuHanden,
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
    [T_Adressen].[PLZ], [T_Kartei].[PLZ]
  ),
  Nz(
    [T_Adressen].[Ort], [T_Kartei].[Ort]
  );

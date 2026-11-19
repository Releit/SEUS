SELECT
  T_RechnungsPos.RechnungID,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  [Name1] & ' ' & [Name2] AS Kunde,
  T_Kartei.PLZ,
  T_Kartei.Ort,
  T_Kartei.Adresse,
  Sum(
    ([Menge] * [VKexkl]) *(1 - [RabattProz]) *(1 + [UStProz])
  ) AS RechnungsbetragInkl,
  T_Rechnungen.MahnStufe,
  T_Rechnungen.RechtsanwaltAm
FROM
  (
    T_Kartei
    INNER JOIN T_Rechnungen ON T_Kartei.ID = T_Rechnungen.KarteiID
  )
  INNER JOIN T_RechnungsPos ON T_Rechnungen.ID = T_RechnungsPos.RechnungID
GROUP BY
  T_RechnungsPos.RechnungID,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  [Name1] & ' ' & [Name2],
  T_Kartei.PLZ,
  T_Kartei.Ort,
  T_Kartei.Adresse,
  T_Rechnungen.MahnStufe,
  T_Rechnungen.RechtsanwaltAm;

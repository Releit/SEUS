SELECT
  T_Zahlungseingänge.*,
  Year([Datum]) AS Jahr,
  [Name1] & " " & [Name2] AS Kundenname
FROM
  T_Kartei
  INNER JOIN T_Zahlungseingänge ON T_Kartei.ID = T_Zahlungseingänge.KarteiID;

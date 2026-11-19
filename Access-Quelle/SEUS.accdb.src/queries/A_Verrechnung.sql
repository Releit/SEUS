SELECT
  T_Leistungen.KarteiID,
  T_Kartei.Name1
FROM
  T_Produkte
  INNER JOIN (
    T_Kartei
    INNER JOIN T_Leistungen ON T_Kartei.ID = T_Leistungen.KarteiID
  ) ON T_Produkte.ID = T_Leistungen.ProduktID
GROUP BY
  T_Leistungen.KarteiID,
  T_Kartei.Name1;

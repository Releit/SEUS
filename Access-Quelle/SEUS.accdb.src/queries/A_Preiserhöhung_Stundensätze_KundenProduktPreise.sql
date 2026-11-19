PARAMETERS [Erhöhung in Prozent] Single,
[Gültig ab] DateTime; INSERT INTO T_KundenProduktpreise (
  VKexkl, GültigAb, KarteiID, ProduktID,
  AngelegtVon
)
SELECT DISTINCT
  Round(
    [T_KundenProduktpreise].[VKexkl] *(1 + [Erhöhung in Prozent] / 100),
    0
  ) AS Ausdr1,
  [Gültig ab] AS Ausdr2,
  T_KundenProduktpreise.KarteiID,
  T_KundenProduktpreise.ProduktID,
  mwin_Username() AS Ausdr3
FROM
  T_Produkte
  INNER JOIN T_KundenProduktpreise ON T_Produkte.ID = T_KundenProduktpreise.ProduktID
WHERE
  (
    (
      (T_Produkte.MengeAusStunden) = Yes
    )
    AND (
      (T_KundenProduktpreise.ID) = mDF_KundenProduktpreiseLastID([KarteiID], [ProduktID])
    )
  );

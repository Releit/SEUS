PARAMETERS [Erhöhung in Prozent] Single;
UPDATE
  T_Produkte
SET
  T_Produkte.VKexkl = Round(
    [T_Produkte].[VKexkl] *(1 + [Erhöhung in Prozent] / 100),
    0
  ),
  T_Produkte.BearbeitetAm = Date(),
  T_Produkte.BearbeitetVon = mWin_UserName()
WHERE
  (
    (
      (T_Produkte.MengeAusStunden) = Yes
    )
  );

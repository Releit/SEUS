INSERT INTO Rechnungen (
  KdLfNr, SEUSID, Datum, RNr, Beschreibung,
  PrintTxt, ZuHanden
)
SELECT
  T_Kartei.KdLfNr,
  T_Rechnungen.ID,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  T_Rechnungen.Anmerkungen,
  T_Rechnungen.DruckText,
  T_Rechnungen.zuhanden
FROM
  (
    T_Kartei
    INNER JOIN T_Rechnungen ON T_Kartei.ID = T_Rechnungen.KarteiID
  )
  LEFT JOIN Rechnungen ON T_Rechnungen.ID = Rechnungen.SEUSID
WHERE
  (
    (
      (T_Rechnungen.Nr) Is Not Null
    )
    AND (
      (Rechnungen.SEUSID) Is Null
    )
  );

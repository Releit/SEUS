SELECT
  T_Kartei.Nr,
  [Name1] & " " & [Name2] AS Kundenname,
  T_Kartei.PLZ,
  T_Kartei.Ort,
  T_Kartei.Adresse,
  T_Kartei.Anrede,
  T_Kartei.Titel,
  T_Kartei.Tel,
  T_Kartei.Fax,
  A_LetzteLieferung.MaxvonDatum AS LetzteLieferung
FROM
  T_Kartei
  LEFT JOIN A_LetzteLieferung ON T_Kartei.ID = A_LetzteLieferung.KarteiID
WHERE
  (
    (
      ([KundenName]) Like 'al*'
    )
  );

SELECT
  T_Adressen.ID,
  [Name1] & " " & [Name2] & (" " + [Titel])& ", " & [PLZ] & " " & [Ort] & ", " & [adresse] AS Ausdr1,
  [Name1] & " " & [Name2] AS Kundenname,
  T_Adressen.KarteiID
FROM
  T_Adressen
ORDER BY
  [Name1] & " " & [Name2] & (" " + [Titel])& ", " & [PLZ] & " " & [Ort] & ", " & [adresse],
  [Name1] & " " & [Name2];

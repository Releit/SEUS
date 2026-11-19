SELECT
  T_Kartei.ID,
  [Name1] & " " & [Name2] & (" " + [Titel])& ", " & [PLZ] & " " & [Ort] & ", " & [adresse] AS [Text]
FROM
  T_Kartei
ORDER BY
  [Name1] & " " & [Name2] & (", " + [Titel])& ", " & [PLZ] & " " & [Ort] & ", " & [adresse];

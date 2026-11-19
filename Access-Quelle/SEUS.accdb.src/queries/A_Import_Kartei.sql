INSERT INTO T_Kartei (
  Name1, Name2, Branche, Adresse, PLZ,
  Ort, LandCode, Tel, FAX, Kunde, Homepage,
  UnsereKdNr, AngelegtVon, BearbeitetVon,
  AngelegtAm, BearbeitetAm, KdLfNr,
  Anmerkungen
)
SELECT DISTINCT
  Kund_Lief_Mit.Name,
  Kund_Lief_Mit.Name2,
  Kund_Lief_Mit.Branche,
  Kund_Lief_Mit.Adresse,
  Kund_Lief_Mit.PLZ,
  Kund_Lief_Mit.Ort,
  "A" AS Ausdr1,
  Kund_Lief_Mit.Tel,
  Kund_Lief_Mit.FAX,
  Yes AS Ausdr2,
  Kund_Lief_Mit.Homepage,
  Kund_Lief_Mit.unsere_KDNr,
  Kund_Lief_Mit.Angelegt_von,
  Kund_Lief_Mit.Bearbeitet_von,
  Kund_Lief_Mit.Angelegt_am,
  Kund_Lief_Mit.Bearbeitet_am,
  Kund_Lief_Mit.KdLfNr,
  [Memo] & Chr(13)& Chr(10)& [Sondervorgaben] AS Ausdr3
FROM
  Kund_Lief_Mit
  INNER JOIN Rechnungen ON Kund_Lief_Mit.KdLfNr = Rechnungen.KdLfNr;

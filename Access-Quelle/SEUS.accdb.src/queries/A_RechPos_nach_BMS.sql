INSERT INTO Rech_pos (
  RNr, SEUSID, ProdNr, Menge, Bemerkung,
  Preis_Exkl, Preis_Inkl
)
SELECT
  Rechnungen.RNr,
  T_RechnungsPos.ID,
  T_RechnungsPos.ProduktID,
  T_RechnungsPos.Menge,
  T_RechnungsPos.Anmerkungen,
  T_RechnungsPos.VKexkl,
  [VKexkl] * 1.2 AS Ausdr1
FROM
  Rechnungen
  INNER JOIN (
    T_RechnungsPos
    LEFT JOIN Rech_pos ON T_RechnungsPos.ID = Rech_pos.SEUSID
  ) ON Rechnungen.SEUSID = T_RechnungsPos.RechnungID
WHERE
  (
    (
      (Rech_pos.SEUSID) Is Null
    )
  );

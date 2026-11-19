UPDATE
  T_Leistungen
SET
  T_Leistungen.AufRechnungID = Null
WHERE
  (
    (
      (T_Leistungen.AufRechnungID) Not In (
      SELECT
        ID
      FROM
        T_Rechnungen)
    )
  );

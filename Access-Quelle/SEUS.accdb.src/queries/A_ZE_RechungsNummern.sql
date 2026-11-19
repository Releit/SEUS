SELECT
  T_Rechnungen_ZE.ZEID,
  T_Rechnungen.Nr
FROM
  T_Rechnungen
  INNER JOIN T_Rechnungen_ZE ON T_Rechnungen.ID = T_Rechnungen_ZE.RechnungID;

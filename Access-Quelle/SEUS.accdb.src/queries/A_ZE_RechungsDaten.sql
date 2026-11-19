SELECT
  T_Rechnungen_ZE.ZEID,
  V_Rechnungen.Nr,
  T_Rechnungen_ZE.RechnungID,
  V_Rechnungen.BetragInkl,
  V_Rechnungen.USt
FROM
  V_Rechnungen
  INNER JOIN T_Rechnungen_ZE ON V_Rechnungen.RechnungID = T_Rechnungen_ZE.RechnungID;

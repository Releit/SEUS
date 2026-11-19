SELECT
  T_EingangsRechnungen_ZA.ZAID,
  T_EingangsRechnungen_ZA.ERID,
  T_EingangsRechnungen_ZA.Betrag,
  T_Zahlungsausgänge.Datum,
  T_Zahlungsausgänge.Betrag AS BetragZA
FROM
  T_Zahlungsausgänge
  INNER JOIN T_EingangsRechnungen_ZA ON T_Zahlungsausgänge.ID = T_EingangsRechnungen_ZA.ZAID;

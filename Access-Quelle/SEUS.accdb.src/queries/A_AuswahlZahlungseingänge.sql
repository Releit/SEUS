SELECT
  T_Zahlungseingänge.ID,
  T_Zahlungseingänge.Datum,
  A_AuswahlKartei.Text AS Kunde,
  T_Zahlungseingänge.Betrag
FROM
  T_Zahlungseingänge
  LEFT JOIN A_AuswahlKartei ON T_Zahlungseingänge.ID = A_AuswahlKartei.ID
ORDER BY
  T_Zahlungseingänge.Datum;

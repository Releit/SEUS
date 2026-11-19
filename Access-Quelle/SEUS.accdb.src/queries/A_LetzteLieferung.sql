SELECT
  T_Lieferscheine.KarteiID,
  Max(T_Lieferscheine.Datum) AS MaxvonDatum
FROM
  T_Lieferscheine
GROUP BY
  T_Lieferscheine.KarteiID;

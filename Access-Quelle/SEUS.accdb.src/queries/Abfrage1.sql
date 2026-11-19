UPDATE
  T_Rechnungen
  INNER JOIN T_Lieferscheine ON T_Rechnungen.Nr = T_Lieferscheine.Nr
SET
  T_Rechnungen.AdressenID = [T_Lieferscheine].[AdressenID]
WHERE
  (
    (
      (T_Rechnungen.AngelegtAm) = #3/16/2011#
    )
  );

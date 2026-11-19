SELECT
  T_Kartei.Lieferant
FROM
  T_Kartei
WHERE
  (
    (
      (T_Kartei.KdLfNr) In (
        SELECT
          KdLfNr
        FROM
          [LiefRechnungen]
      )
    )
  );

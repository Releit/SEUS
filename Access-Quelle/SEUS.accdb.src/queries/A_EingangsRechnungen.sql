SELECT
  T_EingangsRechnungen.*,
  Year(
    Nz(
      [Datum],
      Date()
    )
  ) AS Jahr,
  DatePart(
    "q",
    Nz(
      [datum],
      Date()
    )
  ) AS Qt,
  Round(
    [Brutto] - Nz([BetragBezahlt], 0) - Nz([skonto], 0),
    3
  ) AS offen
FROM
  T_EingangsRechnungen
ORDER BY
  T_EingangsRechnungen.Datum;

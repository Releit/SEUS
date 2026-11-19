SELECT
  T_RechnungsPos.RechnungID,
  Sum(
    Round(
      [Menge] * [VKExkl] *([RabattProz]),
      4
    )
  ) AS RabattExkl
FROM
  T_RechnungsPos
GROUP BY
  T_RechnungsPos.RechnungID;

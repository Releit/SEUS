SELECT
  T_Leistungen.ID AS LeistungID,
  Format([Datum], "dd/mm/yyyy")& " " & Format([von], "Short Time")& "-" & Format([bis], "Short Time")& " " & ("[" + [Projektname] + "] ")& [Anmerkungen] AS PosText
FROM
  T_Leistungen;

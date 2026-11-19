SELECT
  T_Ansprechpartner.ID,
  T_Ansprechpartner.KarteiID,
  (
    Trim([Titel]) + ' '
  )& ([Vorname] + ' ')& [Nachname] AS Name1,
  (
    UCase([Nachname]) + ' '
  )& [Vorname] & (", " + [Titel]) AS Name2,
  T_Ansprechpartner.BelegeZuHanden
FROM
  T_Ansprechpartner;

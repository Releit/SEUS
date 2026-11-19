SELECT
  T_RechnungsPos.RechnungID,
  Nz([Titel] + " ", "")& [Name1] & " " & [Name2] & Nz(
    Chr(13)& Chr(10)& [zuhanden],
    ""
  ) AS Kundenname,
  T_Kartei.Adresse,
  [PLZ] & " " & [Ort] AS PLZOrt,
  T_Rechnungen.Datum,
  T_Rechnungen.Nr,
  T_RechnungsPos.Datum,
  [Menge] & " " & [Einheit] AS MengeEinheit,
  IIf(
    IsNull([Beschreibung]),
    "",
    [Beschreibung] + Chr(13) + Chr(10)
  )& [T_RechnungsPos].[anmerkungen] AS PosText,
  T_Einstellungen.RechnungsFussText,
  T_RechnungsPos.VKexkl,
  T_RechnungsPos.UStProz,
  T_RechnungsPos.Menge,
  T_Zahlungsarten.RechnungsText,
  [kurzzeichen] & "-" & Format([T_Rechnungen].[Nr]) AS RNr,
  T_Kartei.UID,
  IIf(
    [UstProz] = 0.1,
    Round([UstProz] * [VKexkl] * [menge], 2),
    Null
  ) AS Ust10Proz,
  IIf(
    [UstProz] = 0.12,
    Round([UstProz] * [VKexkl] * [menge], 2),
    Null
  ) AS Ust12Proz,
  IIf(
    [UstProz] = 0.2,
    Round([UstProz] * [VKexkl] * [menge], 2),
    Null
  ) AS Ust20Proz,
  [VKexkl] * [Menge] AS Gesamt,
  IIf(
    IsNull([T_Kartei].[uid]),
    Null,
    "Ihre UID:"
  ) AS UIDNrTxt,
  "#_ EMailto " & [eMail] & " _# #_ EMailsubject " & [eMailSubject] & " _#" AS PdfParameter,
  "#_ EMailbody " & [eMailBody1] AS PdfEMailBody1,
  [emailbody2] & " _# #_ NextAction 4 _#" AS pdfEMailBody,
  T_Rechnungen.DruckText
FROM
  (
    (
      T_Kartei
      INNER JOIN (
        T_Rechnungen
        INNER JOIN T_Zahlungsarten ON T_Rechnungen.ZahlartID = T_Zahlungsarten.ID
      ) ON T_Kartei.ID = T_Rechnungen.KarteiID
    )
    INNER JOIN T_RechnungsPos ON T_Rechnungen.ID = T_RechnungsPos.RechnungID
  ),
  T_Einstellungen
WHERE
  (
    (
      (T_RechnungsPos.RechnungID) = mBR_GetID()
    )
  );

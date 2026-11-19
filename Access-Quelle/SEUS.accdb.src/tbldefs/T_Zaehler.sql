CREATE TABLE [T_Zaehler] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Jahr] LONG,
  [AngebotNr] LONG,
  [AuftragNr] LONG,
  [LieferscheinNr] LONG,
  [RechnungNr] LONG,
  [ZahlungseingangNr] LONG,
  [BuchungNr] LONG
)

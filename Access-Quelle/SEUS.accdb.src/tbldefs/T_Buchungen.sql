CREATE TABLE [T_Buchungen] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ZahlungseingangID] LONG,
  [KarteiID] LONG,
  [RechnungID] LONG,
  [GegenRechnungID] LONG,
  [Datum] DATETIME,
  [Betrag] CURRENCY,
  [Buchungsart] VARCHAR (30),
  [Anmerkung] VARCHAR (200),
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30)
)

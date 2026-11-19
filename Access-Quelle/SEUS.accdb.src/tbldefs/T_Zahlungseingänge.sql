CREATE TABLE [T_Zahlungseingänge] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ZahlungsartID] LONG,
  [KarteiID] LONG,
  [LfdNr] LONG,
  [Datum] DATETIME,
  [Betrag] CURRENCY,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30)
)

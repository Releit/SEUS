CREATE TABLE [T_Zahlungsausgänge] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ZahlungsartID] LONG,
  [KarteiID] LONG,
  [LfdNr] LONG,
  [Datum] DATETIME,
  [Betrag] CURRENCY,
  [Anmerkung] VARCHAR (50),
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30)
)

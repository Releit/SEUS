CREATE TABLE [T_Rechnungen_ZE] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ZEID] LONG,
  [RechnungID] LONG,
  [Betrag] CURRENCY,
  [Anmerkungen] VARCHAR (100),
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30)
)

CREATE TABLE [T_EingangsRechnungen_ZA] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ZAID] LONG,
  [ERID] LONG,
  [Betrag] CURRENCY,
  [Anmerkungen] VARCHAR (100),
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30)
)

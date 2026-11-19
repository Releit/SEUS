CREATE TABLE [T_Lieferscheine] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [KarteiID] LONG,
  [ZahlartID] LONG,
  [AdressenID] LONG,
  [Datum] DATETIME,
  [Nr] LONG CONSTRAINT [Nr] UNIQUE,
  [Anmerkungen] LONGTEXT,
  [ZuHanden] VARCHAR (50),
  [DruckText] LONGTEXT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30),
  [Auftragnummer] VARCHAR (16)
)

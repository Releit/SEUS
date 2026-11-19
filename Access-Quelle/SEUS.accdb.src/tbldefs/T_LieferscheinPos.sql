CREATE TABLE [T_LieferscheinPos] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [LieferscheinID] LONG,
  [RechnungID] LONG,
  [ProduktID] LONG,
  [Datum] DATETIME,
  [Menge] SINGLE,
  [Einheit] VARCHAR (10),
  [Beschreibung] VARCHAR (100),
  [VKexkl] CURRENCY,
  [RabattProz] SINGLE,
  [UStProz] SINGLE,
  [Anmerkungen] LONGTEXT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30)
)

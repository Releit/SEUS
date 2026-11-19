CREATE TABLE [T_KundenProduktpreise_alt] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProduktID] LONG,
  [KarteiID] LONG,
  [VKExkl] CURRENCY,
  [GültigAb] DATETIME,
  [Anmerkungen] LONGTEXT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30)
)

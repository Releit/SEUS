CREATE TABLE [T_Ansprechpartner] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [KarteiID] LONG,
  [Anrede] VARCHAR (30),
  [Titel] VARCHAR (20),
  [Vorname] VARCHAR (50),
  [Nachname] VARCHAR (50),
  [Tel] VARCHAR (30),
  [Mobil] VARCHAR (30),
  [eMail] VARCHAR (50),
  [Fax] VARCHAR (30),
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30),
  [BelegeZuHanden] BIT,
  [Anmerkungen] VARCHAR (100)
)

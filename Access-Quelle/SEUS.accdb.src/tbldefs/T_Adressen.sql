CREATE TABLE [T_Adressen] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [KarteiID] LONG,
  [Anrede] VARCHAR (30),
  [Titel] VARCHAR (20),
  [Name1] VARCHAR (50),
  [Name2] VARCHAR (50),
  [Adresse] VARCHAR (50),
  [PLZ] VARCHAR (10),
  [Ort] VARCHAR (50),
  [LandCode] VARCHAR (4),
  [StandardFürRechnungen] BIT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30)
)

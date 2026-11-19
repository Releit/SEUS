CREATE TABLE [T_Projekte_alt] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [KarteiID] LONG,
  [ProjektName] VARCHAR (50),
  [Beschreibung] LONGTEXT,
  [StartDatum] DATETIME,
  [EndDatum] DATETIME,
  [obsolet] BIT,
  [Pauschalverrechnung] BIT,
  [Pauschalbetrag] CURRENCY,
  [PauschalStd] SINGLE,
  [ProjektStdSatz] CURRENCY
)

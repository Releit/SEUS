CREATE TABLE [T_Produkte] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ProdGrpID] LONG,
  [Beschreibung] VARCHAR (100),
  [VKexkl] CURRENCY,
  [UstProz] SINGLE,
  [Einheit] VARCHAR (10),
  [obsolet] BIT,
  [StandardMenge] SINGLE,
  [MengeAusStunden] BIT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [BearbeitetAm] DATETIME,
  [BearbeitetVon] VARCHAR (30),
  [ProduktNrAlt] LONG,
  [BeschreibungSp1] VARCHAR (100),
  [BeschreibungSp2] VARCHAR (100),
  [HonorarExkl] CURRENCY
)

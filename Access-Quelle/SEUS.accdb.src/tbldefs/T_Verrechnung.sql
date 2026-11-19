CREATE TABLE [T_Verrechnung] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [DatumVon] DATETIME,
  [DatumBis] DATETIME,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [HonorarVerrechnung] BIT
)

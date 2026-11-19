CREATE TABLE [T_VerrechnungKunden] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [VerrechnungID] LONG,
  [KarteiID] LONG,
  [DatumVon] DATETIME,
  [DatumBis] DATETIME,
  [Kunde] VARCHAR (100),
  [AnzahlPos] LONG,
  [Betrag] CURRENCY,
  [PositionenGruppieren] BIT,
  [Verrechnet] BIT,
  [AngelegtAm] DATETIME,
  [AngelegtVon] VARCHAR (30),
  [ProjektID] LONG
)

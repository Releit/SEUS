CREATE TABLE [T_BikeChef_PartnerImports] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [GestartetAm] DATETIME,
  [GestartetVon] VARCHAR (50),
  [AnzahlImports] LONG
)

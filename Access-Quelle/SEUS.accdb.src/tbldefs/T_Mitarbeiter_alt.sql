CREATE TABLE [T_Mitarbeiter_alt] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Vorname] VARCHAR (50),
  [Nachname] VARCHAR (50),
  [Titel] VARCHAR (50),
  [TitelHinten] VARCHAR (50)
)

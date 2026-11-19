CREATE TABLE [T_PLZ] (
  [PLZ] LONG CONSTRAINT [PLZ] PRIMARY KEY UNIQUE NOT NULL,
  [Ort] VARCHAR (100),
  [Bundesland] VARCHAR (10),
  [gültigAb] DATETIME,
  [gültigBis] DATETIME,
  [NamePLZTyp] VARCHAR (50),
  [adressierbar] VARCHAR (10),
  [Postfach] VARCHAR (10)
)

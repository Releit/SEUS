CREATE TABLE [T_VersionsNummern_Fügetechnik] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [VersionsNummer] LONG,
  [Datum] DATETIME,
  [AngelegtVon] VARCHAR (255),
  [Anmerkung] LONGTEXT
)

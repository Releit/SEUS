CREATE TABLE [T_Zahlungsarten] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Zahlungsart] VARCHAR (50),
  [RechnungsText] LONGTEXT,
  [Kurzzeichen] VARCHAR (2),
  [Standard] BIT
)

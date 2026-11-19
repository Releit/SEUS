CREATE TABLE [T_AusgabenArten] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Ausgabeart] VARCHAR (50),
  [CoronaRelevant] BIT
)

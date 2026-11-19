CREATE TABLE [T_UstRegelungen] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Titel] VARCHAR (50),
  [BelegText] VARCHAR (100),
  [UIDErforderlich] BIT,
  [Anmerkungen] LONGTEXT,
  [UstProz] SINGLE,
  [BelegTextSP1] VARCHAR (100),
  [BelegTextSP2] VARCHAR (100)
)

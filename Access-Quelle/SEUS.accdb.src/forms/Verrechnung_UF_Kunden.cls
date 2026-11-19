Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdMakeRng_Click()
    Dim NewID As Long

    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    NewID = mUW_VerrechnungZuRechnungen(Me!VerrechnungID, Me!ID)

    m_GoToID "Rechnungen", NewID

    Me!Verrechnet = True
    cmdMakeRng.Enabled = False
End Sub

Private Sub Form_Current()
    Dim Sql As String, sID As String, sVon As String, sBis As String

    sID = Str(Nz(Me![ID], 0))
    sVon = CDate2SQL(Nz(Me!DatumVon, 0))
    sBis = CDate2SQL(Nz(Me!DatumBis, 100000))

    If PositionenGruppieren Then
        Sql = "SELECT T_Leistungen.VerrKndID, T_Leistungen.KarteiID, Count(T_Leistungen.ID) AS AnzahlvonID, Sum(T_Leistungen.Menge) AS SummevonMenge" & _
              ",T_Produkte.Einheit, T_Produkte.Beschreibung AS Produkt, '' AS Anmerkungen, T_Projekte.ProjektName" & _
              " FROM T_Produkte INNER JOIN (T_Leistungen LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID) ON T_Produkte.ID = T_Leistungen.ProduktID" & _
              " WHERE (((T_Leistungen.Menge) <> 0) And ((T_Leistungen.VKexkl) <> 0))" & _
              " AND T_Leistungen.Datum>=" & sVon & " AND T_Leistungen.Datum<=" & sBis
        Sql = Sql & " GROUP BY T_Leistungen.VerrKndID, T_Leistungen.KarteiID, T_Produkte.Einheit, T_Produkte.Beschreibung, T_Projekte.ProjektName" & _
              " HAVING (((T_Leistungen.VerrKndID) = [Formulare]![Verrechnung]![UF_Kunden].[Formular]![ID]))"
    Else
        Sql = "SELECT T_Leistungen.VerrKndID, T_Leistungen.KarteiID, 1 AS AnzahlvonID, T_Leistungen.Menge AS SummevonMenge" & _
              ",T_Produkte.Einheit, T_Produkte.Beschreibung AS Produkt, T_Leistungen.Anmerkungen, T_Projekte.ProjektName" & _
              " FROM T_Produkte INNER JOIN (T_Leistungen LEFT JOIN T_Projekte ON T_Leistungen.ProjektID = T_Projekte.ID) ON T_Produkte.ID = T_Leistungen.ProduktID" & _
              " WHERE (((T_Leistungen.Menge) <> 0) And ((T_Leistungen.VKexkl) <> 0))" & _
              " AND T_Leistungen.Datum>=" & sVon & " AND T_Leistungen.Datum<=" & sBis

        Sql = Sql & " AND (((T_Leistungen.VerrKndID) =" & sID & "))"
    End If

    Sql = Sql & " ORDER BY T_Produkte.Beschreibung;"

    Me.Parent!UF_Positionen.Form.RecordSource = Sql

    cmdMakeRng.Enabled = Not Me!Verrechnet
End Sub

Private Sub PositionenGruppieren_AfterUpdate()
    Form_Current
End Sub

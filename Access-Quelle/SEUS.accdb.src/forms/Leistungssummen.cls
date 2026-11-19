Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmb_MitarbeiterID_AfterUpdate()
    updateSQL
End Sub


Private Sub updateSQL()
    CurrentDb.QueryDefs("A_Leistungssummen").Sql = "SELECT  CStr(Year(T_Leistungen.Datum)) & ' / ' & Month(T_Leistungen.Datum) AS Datum, T_Produkte.Beschreibung, Sum(T_Leistungen.Menge) as Mengen,  T_Leistungen.Einheit" & _
                                                    " FROM T_Leistungen INNER JOIN T_Produkte ON T_Produkte.ID = T_Leistungen.ProduktID" & _
                                                    " GROUP BY  T_Produkte.Beschreibung, T_Leistungen.Einheit, T_Leistungen.MitarbeiterID, Month(T_Leistungen.Datum), Year(T_Leistungen.Datum)" & _
                                                    " HAVING T_Leistungen.MitarbeiterID = " & Str(Me!cmb_MitarbeiterID) & " AND Month(T_Leistungen.Datum) = " & Str(Me!cmb_Monat) & " AND Year(T_Leistungen.Datum) = " & Str(Me!cmb_Year)
    Me.Leistungssummen_UF.Form.RecordSource = "A_Leistungssummen"
    Me.Leistungssummen_UF_Einzel.Form.RecordSource = "SELECT T_Leistungen.Datum, T_Produkte.Beschreibung, T_Leistungen.Menge as Mengen,  T_Leistungen.Einheit" & _
                                                    " FROM T_Leistungen INNER JOIN T_Produkte ON T_Produkte.ID = T_Leistungen.ProduktID" & _
                                                    " WHERE T_Leistungen.MitarbeiterID = " & Str(Me!cmb_MitarbeiterID) & " AND Month(T_Leistungen.Datum) = " & Str(Me!cmb_Monat) & " AND Year(T_Leistungen.Datum) = " & Str(Me!cmb_Year) & " Order by Datum"
    CurrentDb.QueryDefs("A_Leistungssummen").OpenRecordset
    Me.Requery
End Sub

Private Sub cmb_Monat_AfterUpdate()
    updateSQL
End Sub

Private Sub cmd_close_Click()
    DoCmd.Close
End Sub

Private Sub Form_Current()
    Me!cmb_Monat = IIf(Month(Now()) - 1 = 0, 1, Month(Now()) - 1)
    Me!cmb_MitarbeiterID = 2
    Me!cmb_Year = Year(Now())
End Sub

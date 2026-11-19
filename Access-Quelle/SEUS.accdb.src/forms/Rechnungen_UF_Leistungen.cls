Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    If Not IsNull(Me!ID) Then m_GoToID "Rechnungen", Me!ID
End Function

Private Sub Bis_AfterUpdate()
    pv_Menge
End Sub

Private Sub Bis_DblClick(Cancel As Integer)
    If IsNull(Bis) Then
        '** aktuelle Uhrzeit auf eine Viertelstd. gerundet
        Bis = Round((Now() - Date) * 96, 0) / 96
    End If
End Sub

Private Sub Datum_DblClick(Cancel As Integer)
    If IsNull(Datum) Then Datum = Date
End Sub

Private Sub Form_Current()
    Projekt.Requery
End Sub

Private Sub Form_Load()
    Projekt.Visible = mED_GetBoolean("Projekte")
End Sub

Private Sub KarteiID_AfterUpdate()
    ProjektID = Null
    Projekt.Requery
End Sub

Private Sub KarteiID_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", KarteiID
End Sub

Private Sub ProduktID_AfterUpdate()
    Me!Menge = Nz(ProduktID.Column(5), 1)
    Me!MengeAusStunden = Nz(ProduktID.Column(6), False)
    Me!Einheit = ProduktID.Column(3)
    pv_Menge
End Sub

Sub pv_Menge()
    If Me!MengeAusStunden Then
        If Not IsNull(Von) And Not IsNull(Bis) Then
            Menge = (Bis - Von) * 24
        Else
            Menge = Null
        End If
    End If
End Sub

Private Sub ProduktID_DblClick(Cancel As Integer)
    m_GoToID "Produkte", Nz(ProduktID, 0)
End Sub

Private Sub Projekt_AfterUpdate()
    Me!ProjektID = Projekt.Column(1)
End Sub

Private Sub Projekt_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", KarteiID
End Sub

Private Sub Von_AfterUpdate()
    pv_Menge
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!BearbeitetAm = Date
    Me!BearbeitetVon = mWin_UserName
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!AngelegtVon = mWin_UserName
    Me!KarteiID = Me.Parent!KarteiID
End Sub

Private Sub Von_DblClick(Cancel As Integer)
    If IsNull(Von) Then
        '** aktuelle Uhrzeit auf eine Viertelstd. gerundet
        Von = Round((Now() - Date) * 96, 0) / 96
    End If
End Sub

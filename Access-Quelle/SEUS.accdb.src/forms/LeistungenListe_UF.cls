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
        pv_Menge
    End If
End Sub

Private Sub Datum_AfterUpdate()
    Projekt.Requery
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
    pv_ReadVKexkl
End Sub

Private Sub KarteiID_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", KarteiID
End Sub

Public Sub ProduktID_AfterUpdate()

    If Me!Menge <> Nz(ProduktID.Column(5), 1) Then Me!Menge = Nz(ProduktID.Column(5), 1)
    If Me!MengeAusStunden <> Nz(ProduktID.Column(6), False) Then Me!MengeAusStunden = Nz(ProduktID.Column(6), False)
    If Me!Einheit <> ProduktID.Column(3) Then Me!Einheit = ProduktID.Column(3)
    pv_ReadVKexkl
    pv_Menge
End Sub

Sub pv_ReadVKexkl()
    Dim vk As Currency, ProjektStdSatz As Currency

    If Not IsNull(ProduktID) And Not IsNull(KarteiID) Then
        vk = mKD_VKexkl(ProduktID, KarteiID)
        ProjektStdSatz = Val(Nz(Projekt.Column(2), "0"))
        If vk <> -1 Then
            If ProjektStdSatz > 0 Then
                If vk > ProjektStdSatz Then
                    VKExkl = ProjektStdSatz
                Else
                    VKExkl = vk
                End If
            Else
                VKExkl = vk
            End If
        Else
            VKExkl = Null
        End If
    Else
        VKExkl = Null
    End If
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

Public Sub Projekt_AfterUpdate()
    Me!ProjektName = Projekt.Column(1)
    '** eventuellen Projektstundensatz eintragen **
    If Val(Nz(Projekt.Column(2), "0")) <> 0 Then
        VKExkl = Projekt.Column(2)
    Else
        ProduktID_AfterUpdate
    End If
End Sub

Private Sub Projekt_DblClick(Cancel As Integer)
    If Not IsNull(KarteiID) Then m_GoToID "Kartei", KarteiID
End Sub

Private Sub Von_AfterUpdate()
    pv_Menge
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsNull(ProduktID) Then
        MsgBox "Bitte Produkt auswählen", , "Fehlende Eingabe"
        Cancel = True
    End If

    If IsNull(MitarbeiterID) And mDF_ProduktMengeAusStunden(Nz(ProduktID, 0)) Then
        MsgBox "Bitte Mitarbeiter auswählen", , "Fehlende Eingabe"
        Cancel = True
    End If

    Me!BearbeitetAm = Date
    Me!BearbeitetVon = mWin_UserName
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!AngelegtVon = mWin_UserName
End Sub

Private Sub Von_DblClick(Cancel As Integer)
    If IsNull(Von) Then
        '** aktuelle Uhrzeit auf eine Viertelstd. gerundet
        Von = Round((Now() - Date) * 96, 0) / 96
        pv_Menge
    End If
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterDelConfirm(Status As Integer)
    If Status = 2 Then pv_Summen
End Sub

Private Sub Form_AfterUpdate()
    pv_Summen
End Sub

Sub pv_Summen()
    Me.Parent.pv_Summen
End Sub

Private Sub ProduktID_AfterUpdate()
    If Not IsNull(ProduktID) Then
        Me!Beschreibung = ProduktID.Column(1)
        Me!VKExkl = ProduktID.Column(2)
        Me!Einheit = ProduktID.Column(3)
        Me!UstProz = ProduktID.Column(4)
        If Not IsNull(ProduktID.Column(5)) And IsNull(Me!Menge) Then Me!Menge = ProduktID.Column(5)
    Else
        Me!Beschreibung = Null
        Me!VKExkl = Null
        Me!Einheit = Null
    End If
End Sub

Private Sub ProduktID_DblClick(Cancel As Integer)
    If Not IsNull(ProduktID) Then m_GoToID "Produkte", ProduktID
End Sub

Private Sub ProduktID_NotInList(NewData As String, Response As Integer)
    Dim ctl As Control

    Set ctl = Me!ProduktID

    If MsgBox("Soll ihre Eingabe als neues Produkt angelegt werden?", vbYesNo, "Produkt nicht vorhanden") = vbYes Then
        '** neues Produkt anlegen **
        CurrentDb.Execute "INSERT INTO T_Produkte(Beschreibung) VALUES('" & NewData & "')"
        m_GoToID "Produkte", DMax("ID", "T_Produkte")
        Response = acDataErrAdded
    Else
        Response = acDataErrContinue
        ctl.Undo
    End If
End Sub

Private Sub RechnungNr_DblClick(Cancel As Integer)
    If Not IsNull(Me!RechnungID) Then
        m_GoToID "Rechnungen", Me!RechnungID
    End If
End Sub

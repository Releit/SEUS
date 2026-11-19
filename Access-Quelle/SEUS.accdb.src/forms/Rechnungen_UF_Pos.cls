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

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    pv_Summen
End Sub

Private Sub LfsDatum_DblClick(Cancel As Integer)
    If Not IsNull(Me!LieferscheinID) Then
        m_GoToID "Lieferscheine", Me!LieferscheinID
    End If
End Sub

Private Sub LieferscheinNr_DblClick(Cancel As Integer)
    If Not IsNull(Me!LieferscheinID) Then
        m_GoToID "Lieferscheine", Me!LieferscheinID
    End If
End Sub

Private Sub ProduktID_AfterUpdate()
    If Not IsNull(ProduktID) Then
        Me!Beschreibung = ProduktID.Column(1)
        Me!Menge = Nz(ProduktID.Column(5), 1)
        Me!Einheit = ProduktID.Column(3)
        Me!UstProz = ProduktID.Column(4)
        pv_ReadVKexkl
    Else
        Me!Beschreibung = Null
        Me!VKExkl = Null
        Me!Einheit = Null
    End If
End Sub

Sub pv_ReadVKexkl()
    Dim vk As Currency

    If Not IsNull(ProduktID) And Not IsNull(Me.Parent!KarteiID) Then
        vk = mKD_VKexkl(ProduktID, Me.Parent!KarteiID)
        If vk <> -1 Then VKExkl = vk Else VKExkl = Null
    Else
        VKExkl = Null
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

Private Sub UStProz_AfterUpdate()
    If Not IsNull(VKExkl) And Not IsNull(UstProz) Then
        VKInkl = VKExkl * (1 + UstProz)
    End If
End Sub

Private Sub VKexkl_AfterUpdate()
    If Not IsNull(VKExkl) And Not IsNull(UstProz) Then
        VKInkl = VKExkl * (1 + UstProz)
    End If
End Sub

Private Sub VKinkl_AfterUpdate()
    If Not IsNull(VKInkl) And Not IsNull(UstProz) Then
        VKExkl = VKInkl / (1 + UstProz)
    End If
End Sub

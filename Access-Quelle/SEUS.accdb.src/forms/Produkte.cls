Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdClose_Click()
    '** Bei Produktneuanlage aus Rechungspositionen VKExkl und Einheit übernehmen **
    If Me!ID = DMax("ID", "T_Produkte") Then
        On Error Resume Next
        With Forms.Rechnungen.Form.UF_Pos
            !VKExkl = Me!VKExkl
            !Einheit = Me!Einheit
        End With
        On Error GoTo 0
    End If

    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!BearbeitetAm = Date
    Me!BearbeitetVon = mWin_UserName
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!AngelegtVon = mWin_UserName
End Sub

Private Sub Form_Current()
    VKinkl_AfterUpdate
End Sub

Private Sub Form_Load()
    Dim txt As String

    txt = Nz(DLookup("Sprache", "T_Sprachen", "ID=1"), "")
    If txt <> "" Then
        lbBeschreibungSp1.Caption = txt
        BeschreibungSp1.Enabled = True
    Else
        BeschreibungSp1.Enabled = False
    End If

    txt = Nz(DLookup("Sprache", "T_Sprachen", "ID=2"), "")
    If txt <> "" Then
        lbBeschreibungSp2.Caption = txt
        BeschreibungSp2.Enabled = True
    Else
        BeschreibungSp2.Enabled = False
    End If
End Sub

Private Sub VKexkl_AfterUpdate()
    VKInkl = VKExkl * (1 + Nz(UstProz, 0.2))
End Sub

Private Sub VKinkl_AfterUpdate()
    VKExkl = VKInkl / (1 + Nz(UstProz, 0.2))
End Sub

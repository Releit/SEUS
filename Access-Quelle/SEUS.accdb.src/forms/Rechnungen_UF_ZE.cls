Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ZEID_AfterUpdate()
    Dim offen As Currency

    If Not IsNull(ZEID) Then
        offen = ZEID.Column(2)
        If offen > Me.Parent!BetragInkl Then
            Betrag = Me.Parent!BetragInkl
        Else
            Betrag = offen
        End If

    End If
End Sub

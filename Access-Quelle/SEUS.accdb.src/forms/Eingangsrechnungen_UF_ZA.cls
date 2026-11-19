Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Datum_DblClick(Cancel As Integer)
    If Not IsNull(Me!ERID) Then DoCmd.OpenForm "Eingangsrechnungen", , , "ID=" & Str(Me!ERID)
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    'Me.Parent!BetragBezahlt = mDF_ER_BetragBezahlt(Me.Parent!ID)
End Sub

Private Sub Form_AfterUpdate()
    'Me.Parent!BetragBezahlt = mDF_ER_BetragBezahlt(Me.Parent!ID)
End Sub

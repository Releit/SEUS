Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Nr_Click()
    If Not IsNull(Me!RechnungID) Then
        DoCmd.OpenForm "Rechnungen", , , "ID=" & Str(Me!RechnungID)
    End If
End Sub

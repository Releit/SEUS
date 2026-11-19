Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    If Not IsNull(Me!ID) Then m_GoToID "Rechnungen", Me!ID
End Function

Private Sub EuroProTag_Click()
    mBR_SetID Me!ID
    DoCmd.OpenQuery "A_ProjektLeistungenEuroProTag"
End Sub

Private Sub Leistungen_Click()
    mBR_SetID Me!ID
    DoCmd.OpenReport "Leistungsverzeichnis_Projekt", acViewPreview
End Sub

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Load()
    FirmenPapier.Visible = mBR_GetShowLogo()
    BelegKopf.Visible = mBR_GetShowLogo()
End Sub

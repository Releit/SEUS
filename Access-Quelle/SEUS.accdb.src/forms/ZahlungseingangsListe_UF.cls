Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    m_GoToID "ZahlungsEingang", Me!ID
End Function

Private Sub Datum_DblClick(Cancel As Integer)
    pv_GoToID
End Sub

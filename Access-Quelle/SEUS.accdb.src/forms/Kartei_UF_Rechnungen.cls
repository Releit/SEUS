Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    If Not IsNull(Me!ID) Then m_GoToID "Rechnungen", Me!ID
End Function

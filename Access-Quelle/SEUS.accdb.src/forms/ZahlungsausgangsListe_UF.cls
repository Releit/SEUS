Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    DoCmd.OpenForm "Zahlungsausgang", , , "ID=" & Str(Nz(Me!ID, 0))
End Function

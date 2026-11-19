Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    Dim FormName As String

    FormName = "Rechnungen"

    DoCmd.OpenForm FormName, , , "ID=" & Str(Me!ID)

End Function

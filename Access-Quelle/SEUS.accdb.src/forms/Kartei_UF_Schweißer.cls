Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()
    If Not IsNull(Me!ID) Then m_GoToID "Rechnungen", Me!ID
End Function

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me!AngelegtVon = mWin_UserName
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!BearbeitetAm = Date
    Me!BearbeitetVon = mWin_UserName
End Sub

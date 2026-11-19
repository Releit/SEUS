Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOK_Click()
    MemoTextCtrl = MemoText
    DoCmd.Close acForm, Me.Name
End Sub

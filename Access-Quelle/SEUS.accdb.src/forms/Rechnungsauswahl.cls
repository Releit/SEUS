Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAdd2Rechnung_Click()
    Dim NewID As Long

    If IsNull(UF_Rechnungen.Form!ID) Then
        MsgBox "Bitte eine Rechnung aus der Liste auswählen"
    Else
        With Forms.Lieferscheine
            NewID = mUW_LiefersZuRechnung(!ID, !KarteiID, UF_Rechnungen.Form!ID)
            If NewID <> 0 Then
                !UF_Pos.Requery
                m_GoToID "Rechnungen", NewID
            End If
        End With
    End If

    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdMakeNewRechnung_Click()
    Dim NewID As Long

    With Forms.Lieferscheine
        NewID = mUW_LiefersZuRechnung(!ID, !KarteiID, 0)
        If NewID <> 0 Then
            !UF_Pos.Requery
            m_GoToID "Rechnungen", NewID
        End If
    End With

    DoCmd.Close acForm, Me.Name
End Sub

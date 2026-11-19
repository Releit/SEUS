Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function pv_GoToID()



    Dim FormName As String

    FormName = "Kartei"

    If mSU_KarteiSuche Then
        Forms(mSU_StartForm).KarteiID.Requery
        Forms(mSU_StartForm).KarteiID = mSU_KarteiID
        mSU_KarteiSuche = False
        DoCmd.Close acForm, Me.Parent.Name
    Else
        DoCmd.OpenForm FormName, , , , , acHidden
        With Forms(FormName)
            .RecordsetClone.FindFirst "ID=" & Str(Me!ID)
            If Not .RecordsetClone.EOF Then
                .Recordset.Bookmark = .RecordsetClone.Bookmark
                .Visible = True
            Else
                MsgBox "Datensatz nicht gefunden!", , FormName
                DoCmd.Close acForm, FormName
            End If
        End With
    End If
End Function

Private Sub Form_Current()
    mSU_KarteiID = Nz(Me!ID)
End Sub

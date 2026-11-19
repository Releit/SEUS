Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim sPfad As String

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name

End Sub

Function LoopThroughFiles(inputDirectoryToScanForFile, filenameCriteria) As String

    Dim StrFile As String
    'Debug.Print "in LoopThroughFiles. inputDirectoryToScanForFile: ", inputDirectoryToScanForFile

    StrFile = Dir(inputDirectoryToScanForFile & "\*" & filenameCriteria)
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        StrFile = Dir

    Loop

End Function

Private Sub cmdDelFile_Click()
    If Not IsNull(LstFiles) Then
        If MsgBox("Datei löschen?", vbYesNo) = vbYes Then
            mFM_DeleteFile sPfad & "\" & LstFiles
            LstFiles.Requery
        End If
    End If
End Sub

Private Sub cmdRequery_Click()
    pv_ReadFolder SearchFolder

End Sub

Public Sub pv_ReadFolder(Pfad As String)

    sPfad = Pfad

    mFM_FillTempDateien Pfad
    If DateiAuswahlStartForm.Name <> "" Then
        mFM_TempDateienSetZugeordnetER Nz(DateiAuswahlStartForm!Datum, Date)
    Else
        mFM_TempDateienSetZugeordnetER Date
    End If
    LstFiles.Requery
End Sub

Private Sub cmdShowFile_Click()
    If Not IsNull(LstFiles) Then mFM_OpenFile sPfad & "\" & LstFiles
End Sub

Public Sub FilterTxt_AfterUpdate()
    LstFiles.RowSource = "SELECT TempDateien.*, TempDateien.ZugeordnetER FROM TempDateien WHERE TempDateien.ZugeordnetER=False" & _
                         IIf(FilterTxt <> "", " AND Filename LIKE '*" & FilterTxt & "*'", "")
End Sub

Private Sub LstFiles_DblClick(Cancel As Integer)
    If Not IsNull(LstFiles) Then
        If DateiAuswahlStartForm.Name <> "" Then
            DateiAuswahlStartForm!PfadDatei = LstFiles
            'DoCmd.Close acForm, Me.Name
            LstFiles.Requery
        End If
    End If
End Sub

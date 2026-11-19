Attribute VB_Name = "FileManagement"
Option Compare Database
Option Explicit

Const msoFileDialogFolderPicker = 4
Const msoFileDialogFilePicker = 3

Sub testmFM_WriteUTF8TextFile()
    mFM_WriteUTF8TextFile "Förderung in Österreich über Äpfel", "c:\temp\testutf8.txt"
End Sub

Sub mFM_WriteUTF8TextFile(txt As String, sFileName As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")

    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object

    fsT.WriteText txt
    fsT.SaveToFile sFileName, 2 'Save binary data To disk

End Sub

Sub mFM_MoveFile(SrcFileName As String, DestFileName As String)
    Dim fso As Object
    Dim DestFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    DestFolder = LeftOfLast((DestFileName), "\")

    ' Falls Zielordner nicht existiert ? erstellen
    If Not fso.FolderExists(DestFolder) Then
        fso.CreateFolder DestFolder
    End If

    ' Datei verschieben
    fso.MoveFile SrcFileName, DestFileName
End Sub


Function mFM_GetFolderName(Optional OpenAt As String) As String
    Dim lCount As Long

    mFM_GetFolderName = vbNullString

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            mFM_GetFolderName = .SelectedItems(lCount)
        Next lCount
    End With
End Function

Function mFM_GetFileName(Optional OpenAt As String) As String
    Dim lCount As Long

    mFM_GetFileName = vbNullString

    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = OpenAt
        .Show
        For lCount = 1 To .SelectedItems.Count
            mFM_GetFileName = .SelectedItems(lCount)
        Next lCount
    End With
End Function

Sub mFM_OpenFile(FileName As String)
    Dim x As Shell32.Shell

    On Error GoTo OpenFileErr

    If Left(FileName, 4) = "ftp:" Or mFM_FileExists(FileName) Or mFM_FolderExists(FileName) Then
        Set x = New Shell32.Shell
        x.Open FileName
    Else
        MsgBox "Datei oder Ordner:'" & FileName & "' nicht vorhanden!", , "Datei/Ordner öffnen"
    End If

OpenFileExit:
    Exit Sub

OpenFileErr:
    MsgBox Err.Description
    Resume OpenFileExit
End Sub

Sub testsdfsdf()
    'mFM_OpenFile "c:\temp\test.pptx"
    Application.FollowHyperlink "c:\test.jpg"
End Sub

Sub mFM_CreateFolder(NewFolderName As String, CreateInFolder As String)
    Dim fs As Object

    On Error GoTo MyErr

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateFolder CreateInFolder & IIf(Right(CreateInFolder, 1) <> "\", "\", "") & NewFolderName

MyExit:
    Exit Sub

MyErr:
    MsgBox Err.Description
    Resume MyExit
End Sub

Sub mFM_DeleteFolder(FolderName As String)
    Dim fs As Object

    On Error GoTo MyErr

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.deletefolder FolderName

MyExit:
    Exit Sub

MyErr:
    MsgBox Err.Description
    Resume MyExit
End Sub

Function mFM_FolderExists(FolderName As String, Optional NoMsg As Boolean) As Boolean
    Dim fs As Object

    On Error GoTo MyErr

    Set fs = CreateObject("Scripting.FileSystemObject")
    mFM_FolderExists = fs.FolderExists(FolderName)

MyExit:
    Exit Function

MyErr:
    If Not NoMsg Then MsgBox Err.Description
    Resume MyExit
End Function

Function mFM_FileExists(FileName As String, Optional NoMsg As Boolean) As Boolean
    Dim fs As Object

    On Error GoTo MyErr

    Set fs = CreateObject("Scripting.FileSystemObject")
    mFM_FileExists = fs.FileExists(FileName)

MyExit:
    Exit Function

MyErr:
    If Not NoMsg Then MsgBox Err.Description
    Resume MyExit
End Function

Function mFM_DeleteFile(FileName As String) As Boolean
    Dim fs As Object

    On Error GoTo MyErr

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.deletefile FileName
    mFM_DeleteFile = True

MyExit:
    Exit Function

MyErr:
    MsgBox Err.Description
    Resume MyExit
End Function

'** Dateien aus Ordner lesen und Rückgabe des ersten Dateinamen
Function mFM_ReadFolder(Pfad As String, Optional ZielControl As Access.Control, _
                        Optional ZielCombo As Access.Control, Optional SearchSubFolders As Boolean) As String
    Dim txt As String
    Dim cmbTxt As String                         '** für Kombinationsfeld mit zwei Spalten, Datei und Datei mit Pfad
    Dim i As Long, Anzahl As Integer

    Dim RSTempDateien As DAO.Recordset

    '** Tabelle TempDateien anlegen oder leeren **
    mFM_CreateTempDateien

    If Pfad = "" Then Exit Function

    DoCmd.Hourglass True

    Set RSTempDateien = CurrentDb.OpenRecordset("TempDateien", dbOpenDynaset)

    If Right(Trim(Pfad), 1) <> "\" Then Pfad = Pfad + "\"

    With Application.FileSearch
        '.NewSearch
        .LookIn = Pfad
        .TextOrProperty = "*.*"
        .FileType = 1                            'msoFileTypeAllFiles
        '.MatchTextExactly = True
        .SearchSubFolders = SearchSubFolders

        If .Execute(SortBy:=4, _
                    SortOrder:=1) > 0 Then
            '4=msoSortByLastModified, 1=msoSortOrderAscending
            For i = 1 To .FoundFiles.Count
                If Left(.FoundFiles(i), Len(Pfad)) = Pfad Then '** Datei wirklich im Pfad gefunden **
                    RSTempDateien.AddNew
                    If LeftOfLast(.FoundFiles(i), "\") & "\" = Pfad Then
                        '** ersten Dateinnamen zurückgeben
                        If mFM_ReadFolder = "" Then mFM_ReadFolder = .FoundFiles(i)
                    End If
                    RSTempDateien!FileName = Right(.FoundFiles(i), Len(.FoundFiles(i)) - Len(Pfad))
                    RSTempDateien!DateCreated = mFM_FileDateCreated(.FoundFiles(i))
                    RSTempDateien!DateLastModified = mFM_FileDateLastModified(.FoundFiles(i))
                    RSTempDateien.Update
                    Anzahl = Anzahl + 1
                End If
            Next

            'On Error Resume Next 'falls keine ZielControls verwendet werden
            If Anzahl > 0 Then
                'ZielControl = Format(Anzahl) & " Datei(en) gefunden !" & Chr(13) & Chr(10) & Txt
                'ZielCombo.RowSource = cmbTxt
            Else
                'ZielControl = "Ordner nicht gefunden !" & Chr(13) & Chr(10)
                'ZielCombo.RowSource = "'Ordner nicht gefunden';''"
            End If
        Else
            'On Error Resume Next
            'ZielControl = "0 Dateien gefunden !" & Chr(13) & Chr(10)
            'ZielCombo.RowSource = "'leer';''"
        End If
    End With

    DoCmd.Hourglass False
End Function

Sub testLoopThroughFiles()
    mFM_FillTempDateien "C:\Temp"
    DoCmd.OpenTable "TempDateien"
End Sub

Function mFM_FillTempDateien(SearchFolder As String, Optional filenameCriteria As String) As Long

    Dim FileName As String
    Dim rs As Recordset, i As Long

    mFM_CreateTempDateien

    Set rs = CurrentDb.OpenRecordset("TempDateien", dbOpenDynaset)

    FileName = Dir(SearchFolder & "\*" & filenameCriteria)
    Do While Len(FileName) > 0
        With rs
            .AddNew
            !FileName = FileName
            !DateCreated = mFM_FileDateCreated(SearchFolder & "\" & FileName)
            !DateLastModified = mFM_FileDateLastModified(SearchFolder & "\" & FileName)
            i = i + 1
            .Update
        End With

        FileName = Dir
    Loop
    mFM_FillTempDateien = i
End Function

Sub mFM_TempDateienSetZugeordnetER(Datum As Date)
    Dim rs As Recordset, sPath As String
    sPath = mED_ER_Pfad(Datum)


    Set rs = CurrentDb.OpenRecordset("TempDateien", dbOpenDynaset)
    With rs
        Do While Not .EOF()
            If DCount("ID", "T_Eingangsrechnungen", "PfadDatei='" & !FileName & "'") Then
                .Edit
                !ZugeordnetER = True
                .Update
            End If
            .MoveNext
        Loop
    End With
End Sub

'** Tabelle TempDateien anlegen oder Daten löschen **
'** diese Tabelle dienst zum zwischenspeichern der Inhaltes einen Verzeichisses **
Sub mFM_CreateTempDateien()
    Dim tempDateien As TableDef
    Dim Sql As String

    '** temporäre Zwischentabelle löschen
    On Error Resume Next
    Err = 0
    CurrentDb.TableDefs.Delete "TempDateien"
    'MsgBox Err
    'On Error GoTo 0

    If Err = 3211 Then                           '** Tabelle in Verwendung
        Sql = "Delete * From TempDateien"
        CurrentDb.Execute Sql

    Else
        '** Tabledef erzeugen
        Set tempDateien = CurrentDb.CreateTableDef("TempDateien")

        '** Felder in Zwischentabelle einfügen
        With tempDateien
            .Fields.Append .CreateField("FileName", dbText, 255)
            .Fields.Append .CreateField("DateCreated", dbDate)
            .Fields.Append .CreateField("DateLastModified", dbDate)
            .Fields.Append .CreateField("ZugeordnetER", dbBoolean)
        End With

        '** Tabelle erzeugen
        CurrentDb.TableDefs.Append tempDateien
    End If
End Sub

Function mFM_FileDateCreated(FileSpec) As Date
    Dim fs As Object
    Dim f As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileSpec)
    mFM_FileDateCreated = f.DateCreated
End Function

Function mFM_FileDateLastModified(FileSpec) As Date
    Dim fs As Object
    Dim f As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileSpec)
    mFM_FileDateLastModified = f.DateLastModified
End Function

Function mFM_FileDateLastAccessed(FileSpec) As Date
    Dim fs As Object
    Dim f As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileSpec)
    mFM_FileDateLastAccessed = f.DateLastAccessed
End Function

Function mFM_NewWordDoc(NewDocName As String, DocsFolder As String, Optional TemplateName As String) As Boolean

    Dim oApp As Object

    On Error GoTo mFM_NewWordDocErr

    If Len(NewDocName) > 0 Then
        If Not Right(NewDocName, 4) = ".doc" Then
            NewDocName = NewDocName & ".doc"
        End If
        Set oApp = CreateObject("Word.Application")
        oApp.Visible = True
        oApp.ChangeFileOpenDirectory DocsFolder & "\"

        If TemplateName <> "" Then
            oApp.Documents.Add Template:=TemplateName _
                                          , NewTemplate:=False, DocumentType:=0
        Else
            oApp.Documents.Add DocumentType:=0
        End If

        oApp.ActiveDocument.SaveAs FileName:=NewDocName, FileFormat:=0, _
                                   LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
                                   :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                                   SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False

        mFM_NewWordDoc = True
    Else
        MsgBox "ungültiger Dateiname!"
        mFM_NewWordDoc = False
    End If

mFM_NewWordDocExit:
    Exit Function

mFM_NewWordDocErr:
    MsgBox Err.Description
    mFM_NewWordDoc = False
    Resume mFM_NewWordDocExit
End Function

'** Dateien aus Ordner lesen und Rückgabe des ersten Dateinamen
Function mFM_FindFile(Pfad As String, FilterTxt As String, Optional SearchSubFolders As Boolean) As String
    Dim txt As String
    Dim cmbTxt As String                         '** für Kombinationsfeld mit zwei Spalten, Datei und Datei mit Pfad
    Dim i As Long, Anzahl As Integer

    '** Tabelle TempDateien anlegen oder leeren **
    mFM_CreateTempDateien

    If Pfad = "" Then Exit Function

    DoCmd.Hourglass True

    If Right(Trim(Pfad), 1) <> "\" Then Pfad = Pfad + "\"

    With Application.FileSearch
        '.NewSearch
        .LookIn = Pfad
        .TextOrProperty = "*.*"
        .FileType = 1                            'msoFileTypeAllFiles
        '.MatchTextExactly = True
        .SearchSubFolders = SearchSubFolders

        If .Execute(SortBy:=4, _
                    SortOrder:=1) > 0 Then
            '4=msoSortByLastModified, 1=msoSortOrderAscending
            For i = 1 To .FoundFiles.Count
                If Left(.FoundFiles(i), Len(Pfad)) = Pfad Then '** Datei wirklich im Pfad gefunden **

                    If LeftOfLast(.FoundFiles(i), "\") & "\" = Pfad Then
                        '** ersten Dateinnamen zurückgeben
                        If mFM_FindFile = "" Then mFM_FindFile = .FoundFiles(i)
                    End If
                    Anzahl = Anzahl + 1
                End If
            Next
        End If
    End With

    DoCmd.Hourglass False
End Function

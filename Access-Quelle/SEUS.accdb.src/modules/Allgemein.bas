Attribute VB_Name = "Allgemein"
Option Compare Database
Option Explicit
'@Folder main


'** Zeitraumtext für Berichte
Public mBR_Zeitraum As String

Public MemoTextCtrl As Control

Public DateiAuswahlStartForm As Form

Sub asfsdfsdf()
    MsgBox Len("Technische und rechtliche Spezialausbildung (z.B. Gefahrstoffe, Gefahrguttransport, auditieren von rechtlichen Verpflichtungen, Betreiben von Kläranlagen, 'Carbon Footprin', 'Energietechnik', Kreislaufwirtschaft usw.)")
End Sub



'CDate2SQL **************************************************************
'***********
'Zweck: Wandelt ein Datum in die SQL-konforme Schreibweise mm/dd/yyyy um
'
'Beispiel:
'          sqldat=CDate2SQL(Date)
'
'Rückgabewerte:      Der übergebene Datum umgewandelt
'                    Chr$(0) bei <Fehler>
'
'letztes Update : Reinhard H. Leitner 19/08/2004
'
'************************************************************************
Public Function CTime2SQL(Datum As Date) As String
    On Error Resume Next
    CTime2SQL = DatePart("h", Datum) & ":" & DatePart("n", Datum) & ":" & DatePart("s", Datum)
    If Err.Number <> 0 Then CTime2SQL = Chr$(0)
End Function

Public Function CTime2TSQL(Datum As Date) As String
    On Error Resume Next
    CTime2TSQL = "'" & DatePart("h", Datum) & ":" & DatePart("n", Datum) & ":" & DatePart("s", Datum) & "'"
    If Err.Number <> 0 Then CTime2TSQL = Chr$(0)
End Function

Public Function CDate2SQL(Datum As Date) As String
    On Error Resume Next
    CDate2SQL = "#" & DatePart("m", Datum) & "/" & DatePart("d", Datum) & "/" & DatePart("yyyy", Datum) & "#"
    If Err.Number <> 0 Then CDate2SQL = Chr$(0)
End Function

Public Function CDateTime2SQL(Datum As Date) As String
    Err = 0
    'On Error Resume Next
    CDateTime2SQL = DatePart("m", Datum) & "/" & DatePart("d", Datum) & "/" & DatePart("yyyy", Datum) & _
                    " " & DatePart("h", Datum) & ":" & DatePart("n", Datum) & ":" & DatePart("s", Datum)
    If Err.Number <> 0 Then CDateTime2SQL = Chr$(0)
End Function

Public Function CDateTime2TSQL(Datum As Date) As String
    Err = 0
    'On Error Resume Next
    CDateTime2TSQL = "'" & DatePart("d", Datum) & "." & DatePart("m", Datum) & "." & DatePart("yyyy", Datum) & _
                     " " & DatePart("h", Datum) & ":" & DatePart("n", Datum) & ":" & DatePart("s", Datum) & "'"
    If Err.Number <> 0 Then CDateTime2TSQL = Chr$(0)
End Function

Public Function CDate2TSQL(Datum As Date) As String
    On Error Resume Next
    CDate2TSQL = "'" & DatePart("d", Datum) & "." & DatePart("m", Datum) & "." & DatePart("yyyy", Datum) & "'"
    If Err.Number <> 0 Then CDate2TSQL = Chr$(0)
End Function

'** Linken Teil eines Strings Txt nach dem letzten Vorkommen von RightOff
Function LeftOfLast(txt As String, LeftOf As String) As String
    Dim Pos As Integer
    Dim txtOrg As String

    txtOrg = txt

    Do While InStr(txt, LeftOf) > 0
        Pos = Pos + InStr(txt, LeftOf)
        txt = Right(txt, Len(txt) - InStr(txt, LeftOf))
    Loop

    If Pos > 0 Then
        LeftOfLast = Left(txtOrg, Pos - 1)
    Else
        LeftOfLast = txtOrg
    End If
End Function

'** Rechter Teil eines Strings Txt nach dem letzten Vorkommen von RightOff
Function RightOfLast(ByVal txt As String, RightOf As String) As String
    Do While InStr(txt, RightOf) > 0
        txt = Right(txt, Len(txt) - InStr(txt, RightOf) - Len(RightOf) + 1)
    Loop
    RightOfLast = txt
End Function

'** Linken Teil eines Strings Txt nach dem ersten Vorkommen von LeftOF
Function LeftOfFirst(txt As String, LeftOf As String) As String
    LeftOfFirst = Left(txt, InStr(1, txt, LeftOf))
End Function

Sub ShowConnection()
    Dim Tb As TableDef
    For Each Tb In CurrentDb.TableDefs
        If Tb.Connect <> "" Then Debug.Print Tb.Name & ": " & Tb.Connect
    Next

    ';DATABASE=D:\Firma\Access\RechnungsDB\Daten.accdb
End Sub

Sub SetConnection()
    Dim Tb As TableDef, db As Database

    Set db = CurrentDb
    For Each Tb In db.TableDefs
        If Tb.Connect <> "" Then Tb.Connect = ";DATABASE=D:\Firma\Access\RechnungsDB\Daten.accdb"
    Next
End Sub

Sub SetConnectionLeitner()
    Dim Tb As TableDef, db As Database

    Set db = CurrentDb
    For Each Tb In db.TableDefs
        If Tb.Connect <> "" Then
            Tb.Connect = ";DATABASE=D:\Daten\Datenbanken\Schaffer Arnold\Daten.accdb"
        End If
    Next
End Sub

Sub OpenMemoForm(StartCtrl As Control)

    DoCmd.OpenForm "MemoForm", , , , , acHidden
    With Forms("MemoForm").Form
        !lbTitel.Caption = StartCtrl.Name
        Set MemoTextCtrl = StartCtrl
        Set !MemoText = StartCtrl '** Übergabe des Inhalts
        .Visible = True
    End With
End Sub

Sub test()
    Dim frm As Form
    Set frm = Forms.EingangsRechnungenListe!UF.Form
    MsgBox frm!Datum
End Sub

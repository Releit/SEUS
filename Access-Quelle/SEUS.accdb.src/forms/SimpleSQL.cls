Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdRunSQL_Click()
    If Len(Nz(SQLtext, "")) < 4 Then Exit Sub
    If Left(SQLtext, 6) = "SELECT" Then
        CurrentDb.QueryDefs("tmp_Abfrage").Sql = SQLtext
        DoCmd.OpenQuery "tmp_Abfrage", acViewNormal


    Else
        If InStr(1, SQLtext, vbCrLf) Then
            Do While SQLtext <> "" And Len(SQLtext) > 4
                If LeftOfFirst(SQLtext, vbCrLf) = "" Then
                    '** Kein vbcrlf mehr da -> letzter Befehl
                    CurrentDb.Execute SQLtext
                    SQLtext = ""
                Else
                    CurrentDb.Execute LeftOfFirst(SQLtext, vbCrLf)
                    SQLtext = Mid(SQLtext, InStr(1, SQLtext, vbCrLf) + 2)
                End If
            Loop
        Else
            CurrentDb.Execute SQLtext
        End If
    End If
End Sub

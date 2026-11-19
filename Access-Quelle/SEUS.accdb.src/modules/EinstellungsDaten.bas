Attribute VB_Name = "EinstellungsDaten"
Option Compare Database
Option Explicit

Function mED_GetLong(Feldname As String) As Long
    mED_GetLong = Nz(DMin(Feldname, "T_Einstellungen"), 0)
End Function

Function mED_GetText(Feldname As String) As String
    mED_GetText = Nz(DMin(Feldname, "T_Einstellungen"), 0)
End Function

Sub testmED_GetMemo()
    MsgBox mED_GetMemo("MahnFussText")
End Sub

Function mED_GetMemo(Feldname As String) As String
    Dim RS As Object
    Set RS = CurrentDb.OpenRecordset("SELECT " & Feldname & " FROM T_Einstellungen", dbOpenSnapshot)
    mED_GetMemo = Nz(RS.Fields(0), "")
End Function

Function mED_LetztesMahnDatum(RechnungID As Long) As Date
    mED_LetztesMahnDatum = Nz(DMin("Mahnung" & Format(mED_GetLong("MahnAnzahl")) & "am", "T_Rechnungen", "ID=" & Str(RechnungID)), 100000)
End Function

Sub TestmED_LetztesMahnDatum()
    MsgBox mED_LetztesMahnDatum(11)

End Sub

Function mED_GetBoolean(Feldname As String) As Boolean
    mED_GetBoolean = Nz(DMin(Feldname, "T_Einstellungen"), 0)
End Function

Function mED_GetDate(Feldname As String) As Date
    mED_GetDate = Nz(DMin(Feldname, "T_Einstellungen"), 0)
End Function

Sub testmED_ER_Pfad()
    MsgBox mED_ER_Pfad(Date)
End Sub

Function mED_PfadReinhardOrPrivat(sPath As String) As String
    If Not mFM_FolderExists(sPath) Then
        If InStr(sPath, "Privat") Then
            sPath = Replace(sPath, "Privat", "Reinhard")
        ElseIf InStr(sPath, "Reinhard") Then
            sPath = Replace(sPath, "Reinhard", "Privat")
        End If
    End If

    mED_PfadReinhardOrPrivat = sPath
End Function

Function mED_ER_Pfad(Datum As Date) As String
    Dim sPath As String

    sPath = mED_PfadReinhardOrPrivat(Nz(DLookup("BelegBasisOrdner", "T_Einstellungen"), ""))

    If sPath <> "" Then
        If Datum > 0 Then
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
            If mFM_FolderExists(sPath & Format(Year(Datum)), True) Then sPath = sPath & Format(Year(Datum))
            If mFM_FolderExists(sPath & Format(Year(Datum)) & "\ER", True) Then sPath = sPath & Format(Year(Datum))
            If mFM_FolderExists(sPath & "\ER", True) Then sPath = sPath & "\ER"
            If mFM_FolderExists(sPath & "\Eingangsrechnungen", True) Then sPath = sPath & "\Eingangsrechnungen"
        End If
    End If

    mED_ER_Pfad = sPath
End Function

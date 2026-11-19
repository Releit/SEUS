Attribute VB_Name = "ToDo"
Option Compare Database
Option Explicit

Sub mTD_ShowToDo()
    If Nz(DLookup("ToDoListeBeimStart", "T_Einstellungen"), False) Then
        If DCount("ID", "T_ToDoListe", "erledigt=0 and obsolet=0") Then
            DoCmd.OpenForm "ToDoListe", , , , , acDialog

        End If
    End If
End Sub

Sub dfgdfgd()
    DoCmd.RunCommand acCmdWorkgroupAdministrator

End Sub

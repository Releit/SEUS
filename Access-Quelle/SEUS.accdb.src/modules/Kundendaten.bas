Attribute VB_Name = "Kundendaten"
Option Compare Database
Option Explicit

Function mKD_VKexkl(ProduktID As Long, KarteiID As Long) As Currency
    Dim rs As Recordset

    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 VKexkl FROM T_KundenProduktpreise WHERE ProduktID=" & Str(ProduktID) & _
                                     " AND KarteiID=" & Str(KarteiID) & _
                                     " ORDER BY GültigAb DESC", dbOpenSnapshot)

    If Not rs.EOF Then
        mKD_VKexkl = Nz(rs.Fields(0), -1)
    Else
        mKD_VKexkl = Nz(DMin("VKexkl", "T_Produkte", "ID=" & Str(ProduktID)), -1)
    End If
End Function

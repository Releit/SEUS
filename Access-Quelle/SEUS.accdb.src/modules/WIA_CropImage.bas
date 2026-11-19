Attribute VB_Name = "WIA_CropImage"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : WIA_CropImage
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Crop an image
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Late Binding version  -> None required
'             Early Binding version -> Microsoft Windows Image Acquisition Library vX.X
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sInitialImage     : Fully qualified path and filename of the image file to gcrop
' lLeft             : Distance for the Left most side of the image to start cropping
' lTop              : Distance for the Top most side of the image to start cropping
' lRight            : Distance for the Right most side of the image to start cropping
' lBottom           : Distance for the Bottom most side of the image to start cropping
' sDestiationImage  : Optional fully qualified path and filename of the cropped image file
'                       If omitted, it overwrites the original sInitialImage image file
' bOpenCropFile     : Whether the final cropped image should be automatically opened
'
' Usage:
' ~~~~~~
' Cropped and create a new image
' WIA_CropImage "C:\Temp\Img01.jpg", 0, 0, 1000, 1000, "C:\Temp\Img01_Cropped.jpg", False
' Cropped and overwrite the original source image
' WIA_CropImage "C:\Temp\Img01.jpg", 0, 0, 1000, 1000
'
' Revision History:
' Rev       Date(yyyy-mm-dd)        Description
' **************************************************************************************
' 1         2022-01-30              Initial Public Release
'---------------------------------------------------------------------------------------
Sub mWIA_CropImage(sInitialImage As String, _
                   lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, _
                   Optional sDestiationImage As String, _
                   Optional bOpenCropFile As Boolean = True)
    On Error GoTo Error_Handler
    #Const WIA_EarlyBind = False                 'True => Early Binding / False => Late Binding
    #If WIA_EarlyBind = True Then
        Dim oIF               As WIA.ImageFile
        Dim oIP               As WIA.ImageProcess

        Set oIF = New WIA.ImageFile
        Set oIP = New WIA.ImageProcess
    #Else
        Dim oIF               As Object
        Dim oIP               As Object

        Set oIF = CreateObject("WIA.ImageFile")
        Set oIP = CreateObject("WIA.ImageProcess")
    #End If
    Dim sCropImage            As String

    'Load and crop
    oIF.LoadFile sInitialImage
    oIP.Filters.Add oIP.FilterInfos("Crop").FilterID
    With oIP.Filters(1)
        .Properties("Left") = lLeft
        .Properties("Top") = lTop
        .Properties("Right") = lRight
        .Properties("Bottom") = lBottom
    End With
    Set oIF = oIP.Apply(oIF)

    'Save the cropped image
    If sDestiationImage <> "" Then
        'New file specified
        sCropImage = sDestiationImage
    Else
        'Overwrite original if no Dest. file specified
        sCropImage = sInitialImage
    End If
    If Len(Dir(sCropImage)) > 0 Then Kill sCropImage
    oIF.SaveFile sCropImage

    'Open the cropped image
    If bOpenCropFile = True Then Application.FollowHyperlink sCropImage 'Only works in Access.

Error_Handler_Exit:
    On Error Resume Next
    If Not oIP Is Nothing Then Set oIP = Nothing
    If Not oIF Is Nothing Then Set oIF = Nothing
    Exit Sub

Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: WIA_CropImage" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub

Sub testWIA_CropImage()
    mWIA_CropImage "C:\temp\test.png", 15, 15, 16, 16, "C:\temp\test1516.png"

End Sub

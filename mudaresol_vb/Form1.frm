VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim resp As String
resp = InputBox("Qual a resolucao?" & vbCrLf & vbCrLf & _
    "    1. 640x480" & vbCrLf & _
    "    2. 800x600" & vbCrLf & _
    "    3. 1024x768" & vbCrLf & _
    "    4. 1152x864" & vbCrLf & _
    "    5. 1280x1024" & vbCrLf & _
    "Digite o numero referente a resolucao que se deseja configurar.", "mudaresol", "3")
    If resp = "1" Then
        Call ChangeScreen_Resol(640, 480)
    ElseIf resp = "2" Then
        Call ChangeScreen_Resol(800, 600)
    ElseIf resp = "3" Then
        Call ChangeScreen_Resol(1024, 768)
    ElseIf resp = "4" Then
        Call ChangeScreen_Resol(1152, 864)
    ElseIf resp = "5" Then
        Call ChangeScreen_Resol(1280, 1024)
    Else
        Call MsgBox("Opcao invalida. Tente novamente.", vbExclamation)
    End If
End Sub


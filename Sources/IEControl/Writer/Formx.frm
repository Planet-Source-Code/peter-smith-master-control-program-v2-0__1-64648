VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INTERNET CONTROLLER - ADMIN"
   ClientHeight    =   1710
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Formx.frx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Disable"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path & "\IEControl.db" For Output As #1
Print #1, "ENABLED"
Close #1

End Sub

Private Sub Command2_Click()
Open App.Path & "\IEControl.db" For Output As #1
Print #1, "DISABLED"
Close #1
End Sub

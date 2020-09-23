VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Username to Message"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send >"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
Dim strUsername As String

strUsername = Text1.Text
Text2.Visible = False
Text3.Visible = True
Form1.Width = 5985
Command1.Visible = False
Form1.Caption = "Enter message to send"
Command2.Visible = True

End Sub

Public Sub Command2_Click()
Dim strMessage As String
strMessage = Text2.Text
Call Send_Message

End Sub

Private Sub Command3_Click()

End Sub


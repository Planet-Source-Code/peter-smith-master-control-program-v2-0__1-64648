VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoban"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Letter"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   120
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Ban"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strName = Text1.Text

strUsername = Text3.Text
strreason = Text4.Text
strTeacher = Text6.Text

strDate = Now

strmess4 = "Name: " & strName
strmess5 = "Username: " & strUsername
strmess6 = "Date and Time: " & strDate
strmess8 = "Reason: " & strreason
strmess9 = "Reporters Name: " & strTeacher
strLine = "-----------------------------------------------------------------"


Open App.Path & "\LetterTemp.txt" For Append As #1
Print #1, strLine
Print #1, strmess4
Print #1, strmess5
Print #1, strmess6
Print #1, strmess7
Print #1, strmess8
Print #1, strmess9
Print #1, strLine
Close #1

MsgBox ("Report Created")
End
End Sub


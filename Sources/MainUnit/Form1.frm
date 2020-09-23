VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Control"
   ClientHeight    =   5955
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Lock Machine"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show KEYLOGS"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Logons"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Report User"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Password Lookup"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Messenger"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Internet Control"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1800
      X2              =   4320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEWS"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   3255
      Left            =   1800
      Top             =   1680
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell App.Path & "\IEControl\Writer.exe", vbNormalFocus

End Sub

Private Sub Command2_Click()
Shell App.Path & "\MailBox\SendThis.exe", vbNormalFocus

End Sub

Private Sub Command3_Click()
strfilenamex = App.Path & "\Whologon\whologon.txt"
Call ShellExecute(Me.hwnd, "Open", strfilenamex, vbNullString, "c:\", 1)
End Sub

Private Sub Command4_Click()
Shell App.Path & "\PassFind\PassFind.exe", vbNormalFocus

End Sub

Private Sub Command5_Click()
Shell App.Path & "\Keylogger\KeyFind.exe", vbNormalFocus
End Sub

Private Sub Command6_Click()
Shell App.Path & "\Secure\VTest.exe", vbNormalFocus

End Sub

Private Sub Command7_Click()
Shell App.Path & "\BanRequest\BanRequest.exe", vbNormalFocus

End Sub

Private Sub Form_Load()
Dim strDate As String
Dim strmessage As String
Dim strNews As String

Open App.Path & "\MAILBOX\news.txt" For Input As #1
Input #1, strNews
Close #1

strDate = Split(strNews, ":")(0)
strmessage = Split(strNews, ":")(1)
Label2.Caption = strDate
Text1.Text = strmessage
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Zodiac Studios NetMessage - Message Received"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Timer Scanner 
      Interval        =   500
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MSG"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set objnet = CreateObject("WScript.NetWork")
  strmyusername = objnet.UserName

Open App.Path & "\MsgPath.txt" For Input As #1
Input #1, datpath
Close #1

mypath = datpath & strmyusername & ".txt"

Kill mypath
Form1.Hide

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
End
End If

Form1.Hide

End Sub

Private Sub Scanner_Timer()
On Error Resume Next

Set objnet = CreateObject("WScript.NetWork")
  strmyusername = objnet.UserName

Open App.Path & "\MsgPath.txt" For Input As #1
Input #1, datpath
Close #1

MkDir datpath
mypath = datpath & strmyusername & ".txt"

On Error Resume Next
Open mypath For Input As #1
Input #1, message
Close #1

If message = "" Then
Form1.Hide
GoTo none
End If

Label1.Caption = message
Form1.Show

none:
    




End Sub

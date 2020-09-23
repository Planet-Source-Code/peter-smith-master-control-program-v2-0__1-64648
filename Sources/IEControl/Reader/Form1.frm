VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Controller"
   ClientHeight    =   510
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label labelstat 
      Caption         =   "Checking Availability...."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub



Private Sub Timer1_Timer()
If ProgressBar1.Value < ProgressBar1.Max Then
ProgressBar1.Value = ProgressBar1.Value + 1
ElseIf ProgressBar1.Value = ProgressBar1.Value Then
On Error Resume Next
Open App.Path & "\IEControl.db" For Input As #1
Input #1, strEnabled
Close #1

If strEnabled = "ENABLED" Then
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE"), vbNormalFocus
End
ElseIf strEnabled = "DISABLED" Then
labelstat.Caption = "INTERNET DISABLED/OFFLINE"
Form1.Height = 1500
Else
MsgBox ("Error Getting Status. Please RETRY")
End
End If
End If
End Sub

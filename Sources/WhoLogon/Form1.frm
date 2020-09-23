VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Registering User"
   ClientHeight    =   375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   3360
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set objnet = CreateObject("WScript.NetWork")
  struser = objnet.UserName
  strcompa = objnet.computername
stroutput = (struser & " on " & strcompa & " at " & Now)
Open App.Path & "\whologon.txt" For Append As #1
Print #1, stroutput
Close #1


End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set objnet = CreateObject("WScript.NetWork")
  struser = objnet.UserName
  strcompa = objnet.computername
stroutput = (struser & " on " & strcompa & " at " & Now)
Open App.Path & "\whologoff.txt" For Append As #1
Print #1, stroutput
Close #1
End Sub

Private Sub Form_Terminate()
Set objnet = CreateObject("WScript.NetWork")
  struser = objnet.UserName
  strcompa = objnet.computername
stroutput = (struser & " on " & strcompa & " at " & Now)
Open App.Path & "\whologoff.txt" For Append As #1
Print #1, stroutput
Close #1
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set objnet = CreateObject("WScript.NetWork")
  struser = objnet.UserName
  strcompa = objnet.computername
stroutput = (struser & " on " & strcompa & " at " & Now)
Open App.Path & "\whologoff.txt" For Append As #1
Print #1, stroutput
Close #1
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value < ProgressBar1.Max Then
ProgressBar1.Value = ProgressBar1.Value + 1

ElseIf ProgressBar1.Value >= ProgressBar1.Max Then
Form1.Hide
End If

End Sub

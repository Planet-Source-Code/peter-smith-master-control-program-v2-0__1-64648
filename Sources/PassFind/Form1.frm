VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Username"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.ListBox lstUsername 
      Height          =   6105
      Left            =   7560
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Details:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1335
      Left            =   120
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strListNum, strListCount As String

strListNum = lstUsername.ListCount

For i = 0 To strListNum
If i >= lstUsername.ListCount Then GoTo err2
    lstUsername.ListIndex = i
    strcontents = lstUsername.Text
    struser = Split(strcontents, ":")(0)


    If struser = Text1.Text Then
        strpass = Split(strcontents, ":")(1)
        Label2.Caption = Text1.Text
        Label3.Caption = strpass
        GoTo errmain
    End If



Next i

Exit Sub
errmain:
    Command1.Value = False
    GoTo err3
    
err2:

    MsgBox ("Finished Searching, NO RESULTS")
    Command1.Value = False
    
err3:
    Command1.Value = False
    
End Sub

Private Sub Form_Load()
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long

nFileNum = FreeFile

Open App.Path & "\details.txt" For Input As nFileNum
lLineCount = 1
' Read the contents of the file
Do While Not EOF(nFileNum)
    Line Input #nFileNum, sNextLine
    lstUsername.AddItem sNextLine
    'add line numbers to it, in this case!
    sNextLine = lLineCount & " " & sNextLine & vbCrLf
    sText = sText & sNextLine
    lLineCount = lLineCount + 1
Loop

Close nFileNum


End Sub

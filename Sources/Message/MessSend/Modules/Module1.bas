Attribute VB_Name = "Module1"
Public Sub Send_Message()
On Error Resume Next

Set objnet = CreateObject("WScript.NetWork")
  strmyusername = objnet.UserName
  strmycomputer = objnet.computername
  
  
strUsername = Form1.Text2.Text
strMessage = Form1.Text3.Text




   If strUsername = "PathChange" Then
        Path = InputBox("Input new path:")
        Open App.Path & "\MsgPath.txt" For Output As #1
        Print #1, Path
        Close #1
        Call reset_all

End If

If strUsername = "PathChange" Then
Call reset_all
GoTo none
End If

    

Open App.Path & "\MsgPath.txt" For Input As #1
Input #1, datpath
Close #1

Open (datpath & strUsername & ".txt") For Output As #2
Print #2, strmyusername & " says: " & strMessage
Close #2

x = MsgBox("Message sent to " & strUsername, , "Zodiac Messenger")


Open (datpath & "LOG.txt") For Append As #2
Print #2, strmyusername & " on " & strmycomputer & " at " & Now & " sent: " & strMessage & " to " & strUsername
Close #2

Call reset_all
 Exit Sub
 
none:

End Sub

Public Sub reset_all()
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text2.Visible = False
Form1.Text3.Visible = True
Form1.Width = 3375
Form1.Command1.Visible = True

End Sub

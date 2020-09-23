Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long





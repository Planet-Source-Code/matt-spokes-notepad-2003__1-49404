Attribute VB_Name = "File1"
Public Sub NewFile(Text As TextBox, MainForm As Form)
On Error Resume Next
YesNo = MsgBox("Are you sure you want to start a new file?", vbQuestion + vbYesNo, "Notepad 2003")
If YesNo = vbYes Then
Text.Tag = ""
Text.Text = ""
MainForm.Caption = "Notepad 2003 - Untitled"
End If
End Sub

Public Sub LoadFile(Text As TextBox, CD As CommonDialog, MainForm As Form)
Text.Text = ""
With CD
.DialogTitle = "Open File"
.InitDir = Mid(App.Path, 1, 3) & "My Documents"
.Filter = "Text Files|*.txt|HTML Files|*.html;*.htm|ASP Files|*.asp|CSS Files|*.css|INI Files|*.ini|All Files|*.*"
.ShowOpen
End With
Open CD.FileName For Input As #1
Do
Line Input #1, i
Text.Text = Text.Text & i & vbNewLine
Loop Until EOF(1)
Close #1
Text.Tag = CD.FileName
MainForm.Caption = "Notepad 2003 - " & CD.FileTitle
Text.Text = Mid(Text.Text, 1, Len(Text.Text) - 2)
Text.SelStart = Len(Text.Text)
End Sub

Public Sub SaveFile(Text As TextBox, CD As CommonDialog, MainForm As Form)
On Error Resume Next
With CD
.DialogTitle = "Save File"
.InitDir = Mid(App.Path, 1, 3) & "My Documents"
.Filter = "Text Files|*.txt|HTML Files|*.htm|ASP Files|*.asp|CSS Files|*.css|INI Files|*.ini|All Files|*.*"
.ShowSave
End With
Open CD.FileName For Output As #1
Print #1, Text.Text
Close #1
MainForm.Caption = "Notepad 2003 - " & CD.FileTitle
Text.Tag = CD.FileName
End Sub

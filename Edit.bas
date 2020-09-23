Attribute VB_Name = "Edit1"
Dim IsCut As Boolean

Public Sub Copy(Text As TextBox)
On Error Resume Next
Clipboard.SetText (Text.SelText)
IsCut = False
End Sub

Public Sub Cut(Text As TextBox)
On Error Resume Next
Clipboard.SetText (Text.SelText)
Text.Text = Replace(Text.Text, Text.SelText, "", 1, 1)
IsCut = True
End Sub

Public Sub DeleteText(Text As TextBox)
On Error Resume Next
Text.Text = Replace(Text.Text, Text.SelText, "")
End Sub

Public Sub Paste(Text As TextBox)
On Error Resume Next
If IsCut = True Then
Text.SelText = Clipboard.GetText
Clipboard.Clear
Else
Text.SelText = Clipboard.GetText
End If
End Sub

Public Sub SelectAll(Text As TextBox)
Text.SelStart = 0
Text.SelLength = Len(Text.Text)
End Sub

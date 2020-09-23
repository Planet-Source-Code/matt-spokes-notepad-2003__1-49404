Attribute VB_Name = "Registry"
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Const REG_SZ = 1
Private Const LOCALMACHINE = &H80000002

Public Sub ChangeRegisteredOwner(strName As String)
On Error Resume Next
Dim nBufferKey, nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows", nBufferKey
RegOpenKey nBufferKey, "CurrentVersion", nBufferSubKey
RegSetValueEx nBufferSubKey, "RegisteredOwner", 0, REG_SZ, strName, Len(nBufferSubKey)
End Sub

Public Sub ChangeComputerName(strName As String)
On Error Resume Next
Dim nBufferKey, nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "System\CurrentControlSet\Control\ComputerName", nBufferKey
RegOpenKey nBufferKey, "ComputerName", nBufferSubKey
RegSetValueEx nBufferSubKey, "ComputerName", 0, REG_SZ, strName, Len(nBufferSubKey)
End Sub

Public Sub GetNotepadSettings(Text As TextBox)
On Error Resume Next
Dim nBufferKey As Long
RegCreateKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
Dim nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegCreateKey nBufferKey, "Fonts", nBufferSubKey
Dim nBufferDATA As String
nBufferDATA = Space(256)
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey nBufferKey, "Fonts", nBufferSubKey
RegQueryValueEx nBufferSubKey, "FontName", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontName = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontSize", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontSize = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontUnderline", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontUnderline = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontBold", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontBold = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontItalic", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontItalic = nBufferDATA
RegOpenKey nBufferKey, "Colors", nBufferSubKey
RegQueryValueEx nBufferSubKey, "Back", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.BackColor = nBufferDATA
RegQueryValueEx nBufferSubKey, "Text", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.ForeColor = nBufferDATA
End Sub

Public Sub SaveFonts(Text As TextBox)
On Error Resume Next
Dim nBufferKey As Long, nBufferSubKey As Long
RegCreateKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegCreateKey nBufferKey, "Fonts", nBufferSubKey
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey nBufferKey, "Fonts", nBufferSubKey
RegSetValueEx nBufferSubKey, "FontName", 0, REG_SZ, Text.FontName, Len(nBufferSubKey)
RegSetValueEx nBufferSubKey, "FontSize", 0, REG_SZ, Text.FontSize, Len(nBufferSubKey)
RegSetValueEx nBufferSubKey, "FontItalic", 0, REG_SZ, Text.FontItalic, Len(nBufferSubKey)
RegSetValueEx nBufferSubKey, "FontBold", 0, REG_SZ, Text.FontBold, Len(nBufferSubKey)
RegSetValueEx nBufferSubKey, "FontUnderline", 0, REG_SZ, Text.FontUnderline, Len(nBufferSubKey)
End Sub

Public Sub SaveColors(TextColor As PictureBox, BackColor As PictureBox)
On Error Resume Next
Dim nBufferKey As Long, nBufferSubKey As Long
RegCreateKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegCreateKey nBufferKey, "Colors", nBufferSubKey
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey nBufferKey, "Colors", nBufferSubKey
RegSetValueEx nBufferSubKey, "Back", 0, REG_SZ, BackColor.BackColor, Len(nBufferSubKey)
RegSetValueEx nBufferSubKey, "Text", 0, REG_SZ, TextColor.BackColor, Len(nBufferSubKey)
End Sub

Public Sub GetFontSettings(Text As TextBox, Check As CheckBox)
On Error Resume Next
Dim nBufferKey As Long
RegCreateKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
Dim nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegCreateKey nBufferKey, "Fonts", nBufferSubKey
Dim nBufferDATA As String
nBufferDATA = Space(256)
RegOpenKey LOCALMACHINE, "SOFTWARE\MattSpokes", nBufferKey
RegOpenKey nBufferKey, "Fonts", nBufferSubKey
RegQueryValueEx nBufferSubKey, "FontName", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontName = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontSize", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontSize = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontUnderline", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontUnderline = nBufferDATA
If Text.FontUnderline = True Then Check.Value = 1
RegQueryValueEx nBufferSubKey, "FontBold", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontBold = nBufferDATA
RegQueryValueEx nBufferSubKey, "FontItalic", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.FontItalic = nBufferDATA
End Sub

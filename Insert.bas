Attribute VB_Name = "Insert1"
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Const REG_SZ = 1
Private Const LOCALMACHINE = &H80000002

Public Sub TimeDate(Text As TextBox)
Text.SelText = Mid(Time, 1, 5) & "  " & Date
End Sub

Public Sub NotepadVersion(Text As TextBox)
Text.SelText = App.Title & "  -  " & App.Major & "." & App.Minor
End Sub

Public Sub CurrentFile(Text As TextBox, MainForm As Form)
Text.SelText = Mid(MainForm.Caption, 16, 255)
End Sub

Public Sub RegisteredOwner(Text As TextBox)
On Error Resume Next
Dim nBufferKey, nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows", nBufferKey
RegOpenKey nBufferKey, "CurrentVersion", nBufferSubKey
Dim nBufferDATA As String
nBufferDATA = Space(256)
RegQueryValueEx nBufferSubKey, "RegisteredOwner", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.SelText = nBufferDATA
End Sub

Public Sub WindowsVersion(Text As TextBox)
On Error Resume Next
Dim nBufferKey, nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows", nBufferKey
RegOpenKey nBufferKey, "CurrentVersion", nBufferSubKey
Dim nBufferDATA As String
nBufferDATA = Space(256)
RegQueryValueEx nBufferSubKey, "Version", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.SelText = nBufferDATA
End Sub

Public Sub ComputerName(Text As TextBox)
On Error Resume Next
Dim nBufferKey, nBufferSubKey As Long
RegOpenKey LOCALMACHINE, "System\CurrentControlSet\Control\ComputerName", nBufferKey
RegOpenKey nBufferKey, "ComputerName", nBufferSubKey
Dim nBufferDATA As String
nBufferDATA = Space(256)
RegQueryValueEx nBufferSubKey, "ComputerName", 0, REG_SZ, nBufferDATA, Len(nBufferDATA)
Text.SelText = nBufferDATA
End Sub

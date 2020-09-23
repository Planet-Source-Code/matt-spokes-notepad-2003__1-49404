VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Notepad 2003 - Untitled"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8805
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Notepad 
      Height          =   6015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1680
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu PrintSetup 
         Caption         =   "Print Setup"
         Shortcut        =   ^Q
      End
      Begin VB.Menu PrintText 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut1 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy1 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste1 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu SelectAll1 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ChangeFonts 
         Caption         =   "Change Fonts"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Insert 
      Caption         =   "Insert"
      Begin VB.Menu TimeDate1 
         Caption         =   "Time And Date"
         Shortcut        =   {F1}
      End
      Begin VB.Menu YourName 
         Caption         =   "Your Name"
         Shortcut        =   {F2}
      End
      Begin VB.Menu WindowsVer 
         Caption         =   "Windows Version"
         Shortcut        =   {F3}
      End
      Begin VB.Menu NotepadVer 
         Caption         =   "Notepad Version"
         Shortcut        =   {F4}
      End
      Begin VB.Menu CurrentFile1 
         Caption         =   "Current File"
         Shortcut        =   {F5}
      End
      Begin VB.Menu CompName 
         Caption         =   "Computer Name"
         Shortcut        =   {F6}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu edityourname 
         Caption         =   "Edit Your Name"
         Shortcut        =   {F11}
      End
      Begin VB.Menu editcompname 
         Caption         =   "Edit Computer Name"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu findreplace 
      Caption         =   "Find And Replace"
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsSaved As Boolean
Dim CommandLine As String

Private Sub about_Click()
Form3.Show
End Sub

Private Sub ChangeFonts_Click()
Form2.Show
End Sub

Private Sub CompName_Click()
Call ComputerName(Notepad)
End Sub

Private Sub Copy1_Click()
Call Copy(Notepad)
End Sub

Private Sub CurrentFile1_Click()
Call CurrentFile(Notepad, Form1)
End Sub

Private Sub Cut1_Click()
Call Cut(Notepad)
End Sub

Private Sub Delete_Click()
Call DeleteText(Notepad)
End Sub

Private Sub Edit_Click()
If Notepad.SelLength = 0 Then
Cut1.Enabled = False
Copy1.Enabled = False
Delete.Enabled = False
Else
Cut1.Enabled = True
Copy1.Enabled = True
Delete.Enabled = True
End If
If Clipboard.GetText = "" Then
Paste1.Enabled = False
Else
Paste1.Enabled = True
End If
End Sub

Private Sub editcompname_Click()
Dim ComputerName As String
ComputerName = InputBox("Insert your new computer name", "Change Computer Name", "New Computer Name Here")
Call ChangeComputerName(ComputerName)
End Sub

Private Sub edityourname_Click()
Dim UserName As String
UserName = InputBox("Insert your new User Name", "Change User Name", "Your Name Here")
Call ChangeRegisteredOwner(UserName)
End Sub

Private Sub findreplace_Click()
Form4.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Call GetNotepadSettings(Notepad)
If Command() <> "" Then
CommandLine = Mid(Command(), 2, Len(Command()) - 2)
Call CommandFile(CommandLine)
End If
End Sub

Public Sub CommandFile(Filepath As String)
Dim ShortFile As String
Open Filepath For Input As #1
Do
Line Input #1, i
Notepad.Text = Notepad.Text & i & vbNewLine
Loop Until EOF(1)
Close #1
Notepad.Tag = Filepath
For i = 1 To Len(Filepath)
If Mid(Filepath, i, 1) = "\" Then x = i
Next i
ShortFile = Mid(Filepath, x + 1)
Form1.Caption = "Notepad 2003 - " & ShortFile
Notepad.Text = Mid(Notepad.Text, 1, Len(Notepad.Text) - 2)
Notepad.SelStart = Len(Notepad.Text)
IsSaved = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not Notepad.Text = "" Then
it = 1
Else
it = 0
End If
If it = 1 Then
YesNo = MsgBox("Do you want to save changes to your current file?", vbExclamation + vbYesNo, "Notepad 2003")
If YesNo = vbYes Then
Call Save_Click
ElseIf YesNo = vbNo Then
Unload Me
End If
End If
Unload Form2
Unload Form3
Unload Form4
End Sub

Private Sub Form_Resize()
On Error Resume Next
Notepad.Height = Me.Height - 675
Notepad.Width = Me.Width - 105
End Sub

Private Sub New_Click()
On Error Resume Next
Call NewFile(Notepad, Form1)
IsSaved = False
End Sub

Private Sub NotepadVer_Click()
Call NotepadVersion(Notepad)
End Sub

Private Sub Open_Click()
On Error Resume Next
Call LoadFile(Notepad, CD, Form1)
IsSaved = True
End Sub

Private Sub Paste1_Click()
Call Paste(Notepad)
End Sub

Private Sub PrintSetup_Click()
On Error GoTo err
CD.Flags = &H40
CD.CancelError = True
CD.ShowPrinter
err:
End Sub

Private Sub PrintText_Click()
On Error Resume Next
CD.Flags = 0
CD.CancelError = True
CD.ShowPrinter
OldCurrentX = Printer.CurrentX
Printer.Copies = CD.Copies
Printer.Orientation = CD.Orientation
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.CurrentX = Printer.ScaleWidth / 2 - ((Len(Mid(Me.Caption, 16, 255)) * 15 * (Printer.FontSize / 64 * 51)) / 2)
Printer.Print (Mid(Me.Caption, 16, 255) & vbCrLf & vbCrLf)
Printer.FontSize = Notepad.FontSize
Printer.FontBold = Notepad.FontBold
Printer.FontItalic = Notepad.FontItalic
Printer.FontName = Notepad.FontName
Printer.FontUnderline = Notepad.FontUnderline
Printer.CurrentX = OldCurrentX + 2
Dim TextStrW As Long, TempStr As String, LeftStr As String
TextStrW = Printer.TextWidth(Notepad.Text)
LeftStr = Notepad.Text
Do Until Printer.TextWidth(LeftStr) <= Printer.ScaleWidth
TempStr = LeftStr
Do Until Printer.TextWidth(TempStr) <= Printer.ScaleWidth
TempStr = Mid(TempStr, 1, Len(TempStr) - 1)
Loop
Printer.Print TempStr
LeftStr = Mid(LeftStr, Len(TempStr) + 1, Len(LeftStr) - Len(TempStr))
Loop
Printer.Print LeftStr
Printer.EndDoc
End Sub

Private Sub Save_Click()
On Error Resume Next
If IsSaved = True Then
Open Notepad.Tag For Output As #1
Print #1, Notepad.Text
Close #1
Else
Call SaveFile(Notepad, CD, Form1)
End If
IsSaved = True
End Sub

Private Sub SaveAs_Click()
Call SaveFile(Notepad, CD, Form1)
IsSaved = True
End Sub

Private Sub SelectAll1_Click()
Call SelectAll(Notepad)
End Sub

Private Sub TimeDate1_Click()
Call TimeDate(Notepad)
End Sub

Private Sub WindowsVer_Click()
Call WindowsVersion(Notepad)
End Sub

Private Sub YourName_Click()
Call RegisteredOwner(Notepad)
End Sub

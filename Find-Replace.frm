VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find / Replace"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3630
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Replace"
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "Replace"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Replace All"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Replace With :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Find What :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Case Sensitive"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SearchIn As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1

Private Sub Command1_Click()
Dim x As Integer
s = Form1.Notepad.SelStart + Len(Form1.Notepad.SelText) + 1
If Form1.Notepad.SelStart = 0 Then s = 1
If Check1.Value = 1 Then
x = InStr(s, Form1.Notepad.Text, Text1.Text, vbTextCompare)
Else
x = InStr(s, LCase(Form1.Notepad.Text), LCase(Text1.Text), vbTextCompare)
End If
If x <> 0 Then
Form1.Notepad.SelStart = Int(x) - 1
Form1.Notepad.SelLength = Len(Text1.Text)
Form1.Notepad.SetFocus
Else
Form1.Notepad.SelStart = 0
End If
End Sub

Private Sub Command2_Click()
On Error GoTo err
Label3.Caption = ""
With Form1.Notepad
.Text = Replace(.Text, Text2.Text, Text3.Text)
End With
err:
Label3.Caption = Text1.Text & " was not found!"
End Sub

Private Sub Command3_Click()
On Error GoTo err
Label3.Caption = ""
With Form1.Notepad
.Text = Replace(.Text, Text2.Text, Text3.Text, 1, 1)
End With
err:
Label3.Caption = Text1.Text & " was not found!"
End Sub

Private Sub Form_Load()
Dim result As Long
result = SetWindowPos(Form4.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

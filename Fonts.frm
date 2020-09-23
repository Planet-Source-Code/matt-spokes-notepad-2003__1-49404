VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Fonts"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Underline"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Text            =   "Sample"
      Top             =   360
      Width           =   4815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FOnts.frx":0000
      Left            =   3600
      List            =   "FOnts.frx":0010
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   0
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4320
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Back Color :"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Text Color :"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 0 Then
Text1.FontUnderline = False
Else
Text1.FontUnderline = True
End If
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Regular" Then
Text1.FontBold = False
Text1.FontItalic = False
End If
If Combo3.Text = "Bold" Then
Text1.FontBold = True
Text1.FontItalic = False
End If
If Combo3.Text = "Italic" Then
Text1.FontItalic = True
Text1.FontBold = False
End If
If Combo3.Text = "Bold Italic" Then
Text1.FontBold = True
Text1.FontItalic = True
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Combo1.Text < 8 Then Combo1.Text = 8
If Combo1.Text > 72 Then Combo1.Text = 72
Text1.FontSize = Combo1.Text
Form1.Notepad.FontName = Combo2.Text
Form1.Notepad.FontItalic = Text1.FontItalic
Form1.Notepad.FontBold = Text1.FontBold
Form1.Notepad.FontSize = Combo1.Text
Form1.Notepad.FontUnderline = Text1.FontUnderline
Form1.Notepad.ForeColor = Picture1.BackColor
Form1.Notepad.BackColor = Picture2.BackColor
Call SaveColors(Picture1, Picture2)
Call SaveFonts(Form1.Notepad)
Unload Me
End Sub

Private Sub Form_Load()
For i = 1 To Screen.FontCount
Combo2.AddItem Screen.Fonts(i)
Next i
Combo2.RemoveItem 0
For i = 8 To 72 Step 2
Combo1.AddItem i
Next i
Text1.FontName = Form1.Notepad.FontName
Text1.FontSize = Form1.Notepad.FontSize
Text1.FontBold = Form1.Notepad.FontBold
Text1.FontItalic = Form1.Notepad.FontItalic
If Form1.Notepad.FontUnderline = True Then
Check1.Value = 1
Else
Check1.Value = False
End If
Combo2.Text = Text1.FontName
Combo1.Text = Text1.FontSize
If Text1.FontBold = True And Text1.FontItalic = True Then
Combo3.Text = "Bold Italic"
ElseIf Text1.FontBold = True Then
Combo3.Text = "Bold"
ElseIf Text1.FontItalic = True Then
Combo3.Text = "Italic"
Else
Combo3.Text = "Regular"
End If
Text1.FontUnderline = Form1.Notepad.FontUnderline
Text1.ForeColor = Form1.Notepad.ForeColor
Text1.BackColor = Form1.Notepad.BackColor
Picture1.BackColor = Text1.ForeColor
Picture2.BackColor = Text1.BackColor
Call GetFontSettings(Text1, Check1)
Call SelectAll(Text1)
End Sub

Private Sub Picture1_Click()
With CD
.DialogTitle = "Choose Text Color"
.ShowColor
End With
Picture1.BackColor = CD.Color
Text1.ForeColor = CD.Color
End Sub

Private Sub Picture2_Click()
With CD
.DialogTitle = "Choose Back Color"
.ShowColor
End With
Picture2.BackColor = CD.Color
Text1.BackColor = CD.Color
End Sub

Private Sub combo2_Click()
Text1.FontName = Combo2.Text
End Sub

Private Sub combo1_Click()
Text1.FontSize = Combo1.Text
End Sub

VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "By Matt Spokes 25/10/03"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Ver"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   0
      Picture         =   "About.frx":0000
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = App.Title
Label2.Caption = "Version " & App.Major & "." & App.Minor
End Sub

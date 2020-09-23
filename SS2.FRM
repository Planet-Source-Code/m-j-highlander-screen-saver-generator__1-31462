VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SS-Gen"
   ClientHeight    =   2295
   ClientLeft      =   2250
   ClientTop       =   2460
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   4425
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2175
      TabIndex        =   3
      Top             =   1665
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   480
      Left            =   480
      TabIndex        =   2
      Top             =   1650
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   345
      MaxLength       =   6
      TabIndex        =   1
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enter a name for the new Screen Saver       (6 characters max)"
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   255
      TabIndex        =   0
      Top             =   135
      Width           =   3960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
gname = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = ""
gname = ""
Unload Me
End Sub


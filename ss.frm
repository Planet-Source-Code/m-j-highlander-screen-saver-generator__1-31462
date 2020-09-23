VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Generator"
   ClientHeight    =   3615
   ClientLeft      =   1800
   ClientTop       =   1890
   ClientWidth     =   5550
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
   Icon            =   "ss.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   5550
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "&Help"
      Height          =   495
      Left            =   400
      TabIndex        =   3
      Top             =   2880
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   400
      TabIndex        =   2
      Top             =   2295
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   500
      Left            =   400
      TabIndex        =   1
      Top             =   1695
      Width           =   2000
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Columns         =   5
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1590
      Left            =   2835
      TabIndex        =   4
      Top             =   1845
      Visible         =   0   'False
      Width           =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim x As String * 1
Open App.Path & "\ssw.exe" For Binary As #1
x = Left(Label1.Caption, 1)
Put #1, 11987, x
Close
S = Shell(App.Path & "\ssw.exe /s")
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Static a$
'Static i
On Error GoTo e
If a$ = "" Then
coco:
    a$ = InputBox$("Please enter your Windows directory", "SS-Gen", "C:\WINDOWS")
    If a$ = "" Then Exit Sub
    a$ = Trim(a$)
    D$ = Dir(a$, 16)
    If D$ = "" Then MsgBox "Directory does not exist" + Chr(10) + Chr(13) + "Try again.", 16, "SS-Gen": GoTo coco
    If Right(a$, 1) <> "\" Then a$ = a$ + "\"
End If

Form2.Show 1
If gname = "" Then Exit Sub
Open "ssw.exe" For Binary As #2
Dim cx As String * 7
LSet cx = gname
Put #2, 1588, cx
Close #2
FileCopy "ssw.exe", a$ + "ss" + Trim(gname) + ".scr"
Command2.Enabled = False
Exit Sub
e:
MsgBox Error$, 26, "SS-Gen"
Exit Sub
End Sub

Private Sub Command3_Click()
nl = Chr(10) + Chr(13)
msg = "SS-Gen: The Screen Saver Generator" + nl
msg = msg + "by Muhammad S." + nl + nl + nl
msg = msg + "Using SS-Gen is very simple :" + nl + nl
msg = msg + "1. Click the symbol you want." + nl
msg = msg + "2. Click [Go] to view the screen saver." + nl
msg = msg + "3. To save the screen saver click [Save] then enter a name for it." + nl
msg = msg + "4. Finaly run Control Panel to set the new screen saver."

MsgBox msg, , "SS-Gen Help"
End Sub

Private Sub Form_Load()
ChDir App.Path
For i = 33 To 255
List1.AddItem Chr(i)
Next i

Label1.Font.Name = "Windings"
End Sub

Private Sub list1_Click()
Label1.Caption = List1.List(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
Command1.Value = 1
End Sub


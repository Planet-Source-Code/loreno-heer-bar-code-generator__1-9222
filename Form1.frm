VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "BCode"
   ClientHeight    =   1215
   ClientLeft      =   7245
   ClientTop       =   5220
   ClientWidth     =   1875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   125
   Begin VB.CommandButton Command1 
      Appearance      =   0  '2D
      Caption         =   "Paint"
      Default         =   -1  'True
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   195
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "002482"
      Top             =   315
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   195
      Left            =   360
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "610800"
      Top             =   315
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   195
      Left            =   0
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "7"
      Top             =   315
      Width           =   135
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '2D
      Caption         =   "Print"
      Height          =   255
      Left            =   780
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Left            =   1410
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Call the PaintCode Funktion Me<--Form to paint to,1 Number from Left,6 mid. Numbers , Last 6 Numbers
PaintCode Me, Text1.Text, Text2.Text, Text3.Text
End Sub

Private Sub Command2_Click()
'Set the Buttons to Invisible (now they will not by printed)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Me.PrintForm 'Print the Form
Command1.Visible = True 'Set the Buttons back to visible
Command2.Visible = True
Command3.Visible = True
End Sub

Private Sub Command3_Click()
'Show Form2 With Information
Form2.Show
End Sub

Private Sub Form_Load()
'Scale Mode = 3 - Pixel
Text1.Left = 0 'Set the edit boxes to the right position
Text1.Top = 21
Text2.Top = 21
Text3.Top = 21
Text2.Left = 15
Text3.Left = Text2.Left + Text2.Width + 6
PaintCode Me, Text1.Text, Text2.Text, Text3.Text 'PaintCode
End Sub

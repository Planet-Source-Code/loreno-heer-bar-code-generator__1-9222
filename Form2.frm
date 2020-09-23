VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Help"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      MaxLength       =   13
      TabIndex        =   5
      Text            =   "7610800002482"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "or"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "mailto:fat_fish@bluewin.ch"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "mailto:borg@bluewin.ch"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   $"Form2.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Number Check:"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   5520
      X2              =   5520
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   3960
      X2              =   5520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   5520
      X2              =   3960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3960
      X2              =   3960
      Y1              =   1560
      Y2              =   2880
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "(D)Check number"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "(C) Product Code"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "(B) Manufacturer"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   $"Form2.frx":00C0
      Height          =   3195
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "(A)Country code:"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   0
      Picture         =   "Form2.frx":022E
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Test the Code
Dim a
Dim b
Dim c
b = 1
If Len(Text1.Text) = 13 Then
For a = 1 To 12
    If b = 1 Then
        c = c + Mid(Text1.Text, a, 1)
        b = 0
    Else
        c = c + (Mid(Text1.Text, a, 1) * 3)
        b = 1
    End If
Next
If (c + Mid(Text1.Text, 13, 1)) Mod 10 = 0 Then
    Label8.Caption = "Number OK"
Else
    Label8.Caption = "Number Incorrect"
End If
Else
    Label8.Caption = "Number Incorrect"
End If
'e.g:
'Code:   4  0  1  2  3  4  5  0  6  7  8  9  7
'        *1|*3|*1|*3|*1|*3|*1|*3|*1|*3|*1|*3|*1
'Result: 4+ 0+ 1+ 6+ 3+ 12+5+ 0+ 6+ 21+8+ 27 +7 = 100  || 100 Mod 10 = 0 Code is Correct
End Sub


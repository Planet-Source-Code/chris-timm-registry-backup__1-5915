VERSION 5.00
Begin VB.Form frmReadme 
   Caption         =   "Read Me"
   ClientHeight    =   5940
   ClientLeft      =   1155
   ClientTop       =   1755
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Readme.frx":0000
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()


Unload Me
frmMain.Show

End Sub

Private Sub Form_Load()

Dim Readme As String
Dim ReadString As String
Dim Endoffile As String

Readme = App.Path & "\readme.txt"

If Dir(Readme) <> "" Then


Open Readme For Input As #1
    Do While Not EOF(1)
     Line Input #1, ReadString
      
     'MsgBox ReadString
     Endoffile = Endoffile & ReadString & Chr$(13) + Chr$(10)
      
      Text1.Text = Endoffile
Loop

Close #1

Else
    MsgBox "Readme file not found in Windows", vbOKOnly, "Error"
    Exit Sub
End If

End Sub



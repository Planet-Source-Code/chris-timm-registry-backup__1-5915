VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Password Form"
   ClientHeight    =   2325
   ClientLeft      =   3570
   ClientTop       =   2865
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   3510
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the password"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim RunRegEdit


If txtpassword.Text = "password" Then
    Unload Me
    RunRegEdit = Shell("c:\windows\regedit.exe", 1)
Else
    MsgBox "Access has been denied, you cannot run Regedit", , "Access Denied"
    txtpassword.Text = ""
    txtpassword.SetFocus
    Exit Sub
End If

    
    
   
End Sub


Private Sub Command2_Click()



Unload Me

End Sub

Private Sub Form_Load()



txtpassword.Text = ""

End Sub



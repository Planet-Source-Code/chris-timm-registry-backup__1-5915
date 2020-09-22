VERSION 5.00
Begin VB.Form frmshowbackup 
   Caption         =   "Show Backed up Files "
   ClientHeight    =   3225
   ClientLeft      =   3735
   ClientTop       =   3315
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3225
   ScaleWidth      =   3435
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete Files"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmshowbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()


Unload Me

frmMain.Show

End Sub

Private Sub Command1_Click()

Dim DeleteDate As String
Dim RestDate As String

DeleteDate = InputBox("Please enter the date of the file you wish to delete", "Restore Date")
If DeleteDate = "" Then Exit Sub



RestDate = Format$(DeleteDate, "ddmmyy")

If Dir("c:\regbackup\*" & RestDate & "*.*") = "" Then
    MsgBox "There are no files to delete for that Date ", vbOKOnly + vbInformation, "No Files to Delete"
    Exit Sub
    
Else
    Kill "c:\regbackup\*" & RestDate & "*.*"
End If



DoEvents


MsgBox "The files from " & DeleteDate & " Were successfully deleted"
cmdOK_Click



End Sub

Private Sub Form_Load()


Dim sFileName As String
Dim sPath As String


sPath = "c:\regbackup\*.*"

sFileName = Dir(sPath)




If Dir(sPath) = "" Then
    MsgBox "There are no files currently Backed Up", vbOKOnly + vbInformation, "No Backed up files"
    
    
    Exit Sub
Else
    List1.AddItem sFileName


Do
    sFileName = Dir
    List1.AddItem sFileName

Loop Until sFileName = ""


End If

End Sub



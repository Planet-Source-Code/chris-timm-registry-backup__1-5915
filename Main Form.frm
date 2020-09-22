VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Windows 95 Files"
   ClientHeight    =   4380
   ClientLeft      =   2040
   ClientTop       =   2895
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   6690
   Begin VB.CheckBox Check1 
      Caption         =   "Backup AUTOEXEC.BAT AND CONFIG.SYS"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Text            =   " "
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   " "
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Backing Up "
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   " "
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Last Backup"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Files"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Original"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnurunreged 
         Caption         =   "&Run Registry Editor"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCustomise 
      Caption         =   "&Customise"
      Begin VB.Menu mnushowbackup 
         Caption         =   "S&how backed up files"
      End
      Begin VB.Menu mnubuabcs 
         Caption         =   "&Backup AUTOEXEC.BAT AND CONFIG.SYS"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnureadme 
         Caption         =   "Read&ME"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim CheckStat As String * 5
'Dim doini As String * 5
'Dim NewDate As String * 8
'Dim DateChange As String
'Dim CopyMe As Integer
'Dim NewDate1 As String


Private Sub Check1_Click()

If Check1.Value = Checked Then
    CheckStat = "Yes"
Else
    CheckStat = "No"
End If

doini = WritePrivateProfileString("BackupFiles", "BackupAutoConfig", CheckStat, "c:\chris\regcheck.ini")

End Sub

Private Sub Form_Load()

'*****************************
'This Program is hard coded to save the files
'into a directory called c:\regbackup
'It does not search for the directory,
'Please create it before running the code
'********************
'Please update me if you make any changes, just
'for interest so that I can see where it going
'**************
'ctimm@primetimetech.com



Dim DefaultDate As String

doini = GetPrivateProfileString("BackupFiles", "BackupAutoConfig", "Yes", CheckStat, 5, "c:\chris\regcheck.ini")


' Check if Autoexec and Config files need to backed
' up

If Left(CheckStat, 3) = "Yes" Then

   Check1.Value = 1

End If


DefaultDate = "01/01/80"

doini = GetPrivateProfileString("CurrentDate", "DateLastBackedUp", DefaultDate, NewDate, 10, "c:\chris\regcheck.ini")

Label4.Caption = NewDate


End Sub

Private Sub mnuAbout_Click()

'Unload Me
'frmAbout.Show
MsgBox "Windows 9x Registry Backup --- Chris Timm" & Chr(10) + Chr(13) & "This Program is FREEWARE, PLEASE FEEL" & Chr(10) + Chr(13) & "FREE TO COPY/UPDATE it", , "Registry Backup"




End Sub


Private Sub mnuBackup_Click()

' Trap todays date
NewDate = Format$(Now, "dd/mm/yy")

' Change the format of todays date
DateChange = Format$(Now, "ddmmyy")

Text1.Text = "C:\WINDOWS\SYSTEM.DAT"
Text2.Text = "C:\REGBACKUP\SYSTEM_" & DateChange & ".DAT"

DoEvents


If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If


ProgressBar1.Visible = True


For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe



FileCopy "c:\windows\system.dat", "C:\REGBACKUP\SYSTEM_" & DateChange & ".DAT"



Text1.Text = "C:\WINDOWS\USER.DAT"
Text2.Text = "C:\REGBACKUP\USER_" & DateChange & ".DAT"

DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If



For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe

FileCopy "c:\windows\user.dat", "C:\REGBACKUP\USER_" & DateChange & ".DAT"

If Check1.Value = Checked Then
    CheckStat = "Yes"
     doini = WritePrivateProfileString("BackupFiles", "BackupAutoConfig", CheckStat, "c:\chris\regcheck.ini")
 
    Text1.Text = "C:\AUTOEXEC.BAT"
    Text2.Text = "C:\REGBACKUP\AUTOEXEC_" & DateChange & ".BAT"

DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If



For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe
FileCopy "c:\Autoexec.bat", "C:\REGBACKUP\AUTOEXEC_" & DateChange & ".BAT"


    Text1.Text = "C:\CONFIG.SYS"
    Text2.Text = "C:\REGBACKUP\CONFIG_" & DateChange & ".SYS"
DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If



For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe
FileCopy "c:\config.sys", "C:\REGBACKUP\CONFIG_" & DateChange & ".SYS"

Else
    CheckStat = "No"
    doini = WritePrivateProfileString("BackupFiles", "BackupAutoConfig", CheckStat, "c:\chris\regcheck.ini")

End If




MsgBox "Backup Successful", vbOKOnly, "Successful"


Text1.Text = ""
Text2.Text = ""

ProgressBar1.Value = 0
ProgressBar1.Visible = False

doini = WritePrivateProfileString("CurrentDate", "DateLastBackedUp", NewDate, "c:\chris\regcheck.ini")



End Sub

Private Sub mnubuabcs_Click()

mnubuabcs.Checked = Not mnubuabcs.Checked

   ' mnubuabcs.Checked = Checked
    
    Check1.Value = Checked

If mnubuabcs.Checked = Unchecked Then
    Check1.Value = Unchecked
End If


End Sub






Private Sub mnuexit_Click()


End

End Sub

Private Sub mnureadme_Click()


Unload Me
frmReadme.Show


End Sub


Private Sub mnuRestore_Click()

Dim RestoreDate As String
Dim RestDate As String

RestoreDate = InputBox("Please enter the date from which you would like the restore to be done.  " & Chr(10) + Chr(13) & "Use the date form of DD/MM/YYYY", "Restore Date")

RestDate = Format$(RestoreDate, "ddmmyy")

If Dir("c:\regbackup\*" & RestDate & "*.*") = "" Then
    MsgBox "There are no files to Restore for that Date ", vbOKOnly + vbInformation, "No Files to Delete"
    Exit Sub
    
Else

ProgressBar1.Visible = True


Text1.Text = "C:\REGBACKUP\CONFIG_" & RestDate & ".sys"
Text2.Text = "C:\TEMP\CONFIG.SYS"

DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If
DoEvents

For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe

FileCopy "C:\REGBACKUP\CONFIG_" & RestDate & ".sys", "c:\temp\CONFIG.SYS"

'**************************************


Text1.Text = "C:\REGBACKUP\SYSTEM_" & DateChange & ".DAT"
Text2.Text = "C:\TEMP\SYSTEM.DAT"

DoEvents


If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If


ProgressBar1.Visible = True


For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe



FileCopy "C:\REGBACKUP\SYSTEM_" & DateChange & ".DAT", "c:\temp\SYSTEM.DAT"





Text1.Text = "C:\REGBACKUP\USER_" & DateChange & ".DAT"
Text2.Text = "C:\TEMP\USER.DAT"

DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If



For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe

FileCopy "C:\REGBACKUP\USER_" & DateChange & ".DAT", "c:\temp\user.dat"

 
    Text1.Text = "C:\REGBACKUP\AUTOEXEC_" & DateChange & ".BAT"
    Text2.Text = "C:\TEMP\AUTOEXEC.BAT"

DoEvents

If ProgressBar1.Value >= 100 Then
    ProgressBar1.Value = 0
End If



For CopyMe = 0 To 99
    ProgressBar1.Value = ProgressBar1.Value + 1
Next CopyMe
FileCopy "C:\REGBACKUP\AUTOEXEC_" & DateChange & ".BAT", "C:\TEMP\AUTOEXEC.BAT"

MsgBox "RESTORE  Successful", vbOKOnly, "Successful"

Text1.Text = ""
Text2.Text = ""

ProgressBar1.Value = 0
ProgressBar1.Visible = False

End If


End Sub


Private Sub mnurunreged_Click()


Dim RunRegEd
Dim Reply As Integer

Reply = MsgBox("You are required to enter a password before continuing , OK to continue", vbYesNo + vbQuestion, "Password Required")
If Reply = vbYes Then
    frmPassword.Show
    
Else
    Exit Sub
End If







'RunRegEd = Shell("C:\WINDOWS\REGEDIT.EXE", 1)








End Sub

Private Sub mnushowbackup_Click()

'Unload Me
frmshowbackup.Show


End Sub



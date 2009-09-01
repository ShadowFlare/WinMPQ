VERSION 4.00
Begin VB.Form frmMpq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPQ Embedder"
   ClientHeight    =   1695
   ClientLeft      =   3045
   ClientTop       =   2730
   ClientWidth     =   2775
   Height          =   2385
   Icon            =   "frmMpq.frx":0000
   Left            =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2775
   Top             =   2100
   Width           =   2895
   Begin VB.CommandButton cmdSaveEXE 
      Caption         =   "Save &EXE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveMPQ 
      Caption         =   "Save &MPQ"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2565
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run EXE"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHReadme 
         Caption         =   "View &Readme..."
      End
      Begin VB.Menu mnuHSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMpq"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim MpqHeader As Long, IsEXE As Boolean, FileDialog As OPENFILENAME
Private Sub cmdAdd_Click()
Dim OldFileName As String, NewMpqHeader As Long, fNum As Long, Text As String, fNum2 As Long, Text2 As String, bNum As Long
FileDialog.Flags = &H1000 Or &H4 Or &H2
FileDialog.Filter = "Mpq Archives (*.mpq;*.exe;*.snp;*.scm;*.scx;*.w3m;*.w3x)|*.mpq;*.exe;*.snp;*.scm;*.scx;*.w3m;*.w3x|All Files (*.*)|*.*"
OldFileName = FileDialog.FileName
FileDialog.hwndOwner = hWnd
If ShowOpen(FileDialog) = False Then GoTo Cancel
NewMpqHeader = FindMpqHeader(FileDialog.FileName)
If NewMpqHeader = -1 Then
    MsgBox "This file does not contain an MPQ archive.", , "MPQ Embedder"
    GoTo Cancel
End If
fNum = FreeFile
Open FileDialog.FileName For Binary As #fNum
fNum2 = FreeFile
Open OldFileName For Binary As #fNum2
If MpqHeader / 512 <> Int(MpqHeader / 512) Then
    bNum = MsgBox("The file you are adding the MPQ archive to" + vbCrLf + "is not the proper size; therefore, most MPQ" + vbCrLf + "archive readers will not be able to read it." + vbCrLf + "Do you want to increase the size of the file," + vbCrLf + "so other programs can read it?", vbQuestion Or vbYesNo Or vbDefaultButton1, "MPQ Embedder")
    If bNum = vbYes Then
        Text = String(512 - (MpqHeader - Int(MpqHeader / 512) * 512), Chr(0))
        Put #fNum2, MpqHeader + 1, Text
        MpqHeader = MpqHeader + Len(Text)
    End If
End If
For bNum = NewMpqHeader + 1 To LOF(fNum) Step 2 ^ 20
    Text = String(2 ^ 20, Chr(0))
    If LOF(fNum) - bNum + 1 >= 2 ^ 20 Then
        Get #fNum, bNum, Text
        Put #fNum2, MpqHeader + bNum - NewMpqHeader, Text
    Else
        Text = String(LOF(fNum) - bNum + 1, Chr(0))
        Get #fNum, bNum, Text
        Put #fNum2, MpqHeader + bNum - NewMpqHeader, Text
    End If
Next bNum
Close #fNum2
Close #fNum
cmdAdd.Enabled = False
cmdRemove.Enabled = True
cmdSaveMPQ.Enabled = True
cmdSaveEXE.Enabled = True
If MpqHeader / 512 = Int(MpqHeader / 512) Then
    Label1.Caption = "This file contains an MPQ archive."
Else
    Label1.Caption = "This file contains an MPQ archive, but other programs may not be able to read it."
End If
Cancel:
FileDialog.FileName = OldFileName
End Sub
Private Sub cmdRemove_Click()
Dim fNum As Long, Text As String, fNum2 As Long, Text2 As String, bNum As Long
bNum = MsgBox("Are you sure you want to permanently" + vbCrLf + "remove the MPQ archive from this file?", vbQuestion Or vbYesNo Or vbDefaultButton2, "MPQ Embedder")
If bNum = vbNo Then Exit Sub
fNum = FreeFile
Open FileDialog.FileName For Binary As #fNum
fNum2 = FreeFile
If Dir(FileDialog.FileName + ".remove") <> "" Then Kill FileDialog.FileName + ".remove"
Open FileDialog.FileName + ".remove" For Binary As #fNum2
For bNum = 1 To MpqHeader Step 2 ^ 20
    Text = String(2 ^ 20, Chr(0))
    If MpqHeader - bNum + 1 >= 2 ^ 20 Then
        Get #fNum, bNum, Text
        Put #fNum2, bNum, Text
    Else
        Text = String(MpqHeader - bNum + 1, Chr(0))
        Get #fNum, bNum, Text
        Put #fNum2, bNum, Text
    End If
Next bNum
Close #fNum2
Close #fNum
Kill FileDialog.FileName
Name FileDialog.FileName + ".remove" As FileDialog.FileName
cmdAdd.Enabled = True
cmdRemove.Enabled = False
cmdSaveMPQ.Enabled = False
cmdSaveEXE.Enabled = True
Label1.Caption = "This file does not contain an MPQ archive."
End Sub
Private Sub cmdSaveEXE_Click()
Dim OldFileName As String, fNum As Long, Text As String, fNum2 As Long, Text2 As String, bNum As Long
FileDialog.Flags = &H1000 Or &H4 Or &H2
FileDialog.Filter = "File (*.*)|*.*"
FileDialog.DefaultExt = ""
OldFileName = FileDialog.FileName
FileDialog.FileName = FileDialog.FileName
FileDialog.hwndOwner = hWnd
If ShowSave(FileDialog) = False Then GoTo Cancel
fNum = FreeFile
Open OldFileName For Binary As #fNum
fNum2 = FreeFile
If Dir(FileDialog.FileName) <> "" Then Kill FileDialog.FileName
Open FileDialog.FileName For Binary As #fNum2
For bNum = 1 To MpqHeader Step 2 ^ 20
    Text = String(2 ^ 20, Chr(0))
    If MpqHeader - bNum + 1 >= 2 ^ 20 Then
        Get #fNum, bNum, Text
        Put #fNum2, bNum, Text
    Else
        Text = String(MpqHeader - bNum + 1, Chr(0))
        Get #fNum, bNum, Text
        Put #fNum2, bNum, Text
    End If
Next bNum
Close #fNum2
Close #fNum
Cancel:
FileDialog.FileName = OldFileName
End Sub
Private Sub cmdSaveMPQ_Click()
Dim OldFileName As String, fNum As Long, Text As String, fNum2 As Long, Text2 As String, bNum As Long
FileDialog.Flags = &H1000 Or &H4 Or &H2
FileDialog.Filter = "MPQ Archive (*.mpq)|*.mpq"
FileDialog.DefaultExt = "mpq"
OldFileName = FileDialog.FileName
FileDialog.FileName = FileDialog.FileName + ".mpq"
FileDialog.hwndOwner = hWnd
If ShowSave(FileDialog) = False Then GoTo Cancel
fNum = FreeFile
Open OldFileName For Binary As #fNum
fNum2 = FreeFile
If Dir(FileDialog.FileName) <> "" Then Kill FileDialog.FileName
Open FileDialog.FileName For Binary As #fNum2
For bNum = MpqHeader + 1 To LOF(fNum) Step 2 ^ 20
    Text = String(2 ^ 20, Chr(0))
    If LOF(fNum) - bNum + 1 >= 2 ^ 20 Then
        Get #fNum, bNum, Text
        Put #fNum2, bNum - MpqHeader, Text
    Else
        Text = String(LOF(fNum) - bNum + 1, Chr(0))
        Get #fNum, bNum, Text
        Put #fNum2, bNum - MpqHeader, Text
    End If
Next bNum
Close #fNum2
Close #fNum
Cancel:
FileDialog.FileName = OldFileName
End Sub

Private Sub Form_Load()
FileDialog = CD
End Sub
Private Sub mnuFExit_Click()
Unload Me
End Sub
Private Sub mnuFOpen_Click()
Dim OldFileName As String, OldMpqHeader As Long, fNum As Long, Text As String
FileDialog.Flags = &H1000 Or &H4 Or &H2
FileDialog.Filter = "All Files (*.*)|*.*"
OldFileName = FileDialog.FileName
OldMpqHeader = MpqHeader
FileDialog.hwndOwner = hWnd
If ShowOpen(FileDialog) = False Then GoTo Cancel
If FileLen(FileDialog.FileName) = 0 Then
    MsgBox "This is an empty file.", vbExclamation, "MPQ Embedder"
    GoTo Cancel
End If
fNum = FreeFile
Open FileDialog.FileName For Binary As #fNum
Text = String(2, Chr(0))
If LOF(fNum) >= 2 Then Get #fNum, 1, Text
Close #fNum
If Text = "MZ" Then IsEXE = True Else IsEXE = False
If IsEXE Then mnuRun.Enabled = True Else mnuRun.Enabled = False
MpqHeader = FindMpqHeader(FileDialog.FileName)
If MpqHeader <= -1 Then
    cmdAdd.Enabled = True
    cmdRemove.Enabled = False
    cmdSaveMPQ.Enabled = False
    cmdSaveEXE.Enabled = True
    MpqHeader = FileLen(FileDialog.FileName)
    Label1.Caption = "This file does not contain an MPQ archive."
ElseIf MpqHeader = 0 Then
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdSaveMPQ.Enabled = True
    cmdSaveEXE.Enabled = False
    Label1.Caption = "This file is an MPQ archive."
ElseIf MpqHeader > 0 Then
    cmdAdd.Enabled = False
    cmdRemove.Enabled = True
    cmdSaveMPQ.Enabled = True
    cmdSaveEXE.Enabled = True
    If MpqHeader / 512 = Int(MpqHeader / 512) Then
        Label1.Caption = "This file contains an MPQ archive."
    Else
        Label1.Caption = "This file contains an MPQ archive, but other programs may be unable to read it."
    End If
End If
Exit Sub
Cancel:
FileDialog.FileName = OldFileName
MpqHeader = OldMpqHeader
End Sub
Private Sub mnuHAbout_Click()
About.Show 1
End Sub
Private Sub mnuHReadme_Click()
Dim Path As String
Path = App.Path
If Right(Path, 1) <> "\" Then Path = Path + "\"
If Dir(Path + "WMpqEmbed.rtf") = "" Then MsgBox "Could not find WMpqEmbed.rtf!", vbCritical, "MPQ Embedder"
ShellExecute hWnd, vbNullString, Path + "WMpqEmbed.rtf", vbNullString, vbNullString, 1
End Sub
Private Sub mnuRun_Click()
On Error GoTo NotExecutable
Shell FileDialog.FileName, 1
Exit Sub
NotExecutable:
MsgBox "This file is not a .exe file.", vbInformation, "MPQ Embedder"
End Sub

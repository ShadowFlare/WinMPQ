VERSION 4.00
Begin VB.Form ChLCID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Changing Locale ID..."
   ClientHeight    =   1335
   ClientLeft      =   2670
   ClientTop       =   3180
   ClientWidth     =   3615
   Height          =   1740
   Icon            =   "ChLCID.frx":0000
   Left            =   2610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Top             =   2835
   Width           =   3735
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type in the new locale ID for the file(s) below."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
End
Attribute VB_Name = "ChLCID"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
MpqEx.ChangeLCID Text1
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Left = MpqEx.Left + 330 * 2
If Left < 0 Then Left = 0
If Left + Width > Screen.Width Then Left = Screen.Width - Width
Top = MpqEx.Top + 315 * 2
If Top < 0 Then Top = 0
If Top + Height > Screen.Height Then Top = Screen.Height - Height
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim NewValue As Long
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> Asc("-") Then KeyAscii = 0
On Error GoTo TooBig
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = Asc("-") Then NewValue = CLng(Text1 + Chr(KeyAscii))
On Error GoTo 0
Exit Sub
TooBig:
KeyAscii = 0
End Sub
Private Sub Text1_LostFocus()
If Text1 = "" Then Text1 = 0
End Sub

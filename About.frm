VERSION 4.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1575
   ClientLeft      =   1890
   ClientTop       =   2265
   ClientWidth     =   5820
   Height          =   1980
   Icon            =   "About.frx":0000
   Left            =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Top             =   1920
   Width           =   5940
   Begin VB.CommandButton Command1 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This program uses "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:  blakflare@hotmail.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ShadowFlare's Realm - http://sfsrealm.hopto.org/"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © ShadowFlare Software"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WinMPQ v"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "About"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Function GetAppVersionString() As String
GetAppVersionString = CStr(App.Major) + "."
If Len(CStr(App.Minor)) < 2 Then
    GetAppVersionString = GetAppVersionString + "0" + CStr(App.Minor)
Else
    GetAppVersionString = GetAppVersionString + CStr(App.Minor)
End If
If App.Revision <> 0 Then
    If Len(CStr(App.Revision)) < 4 Then
        GetAppVersionString = GetAppVersionString + "." + String(4 - Len(CStr(App.Revision)), "0") + CStr(App.Revision)
    Else
        GetAppVersionString = GetAppVersionString + "." + CStr(App.Revision)
    End If
End If
End Function
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = False
End Sub
Private Sub Form_Load()
Dim SFmpqString As String
Left = MpqEx.Left + 330
If Left < 0 Then Left = 0
If Left + Width > Screen.Width Then Left = Screen.Width - Width
Top = MpqEx.Top + 315
If Top < 0 Then Top = 0
If Top + Height > Screen.Height Then Top = Screen.Height - Height
Caption = "About " + App.Title
Label1 = App.Title + " v"
Label1 = Label1 + GetAppVersionString
Label2 = App.LegalCopyright
SFmpqString = String(SFMpqGetVersionString2(NullPtr, 0) - 1, Chr(0))
SFMpqGetVersionString2 SFmpqString, SFMpqGetVersionString2(NullPtr, 0)
Label5 = Label5 + SFmpqString
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = False
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = False
End Sub
Private Sub Label3_Click()
ShellExecute hWnd, vbNullString, "http://sfsrealm.hopto.org/", vbNullString, vbNullString, 1
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF00&
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF&
Label3.Font.underline = True
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = False
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
End Sub
Private Sub Label4_Click()
ShellExecute hWnd, vbNullString, "mailto:blakflare@hotmail.com", vbNullString, vbNullString, 1
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF00&
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF&
Label4.Font.underline = True
Label5.Font.underline = False
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
End Sub
Private Sub Label5_Click()
AboutSFMpq
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFF0000
Label3.Font.underline = False
Label4.ForeColor = &HFF0000
Label4.Font.underline = False
Label5.Font.underline = True
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Font.underline = False
End Sub

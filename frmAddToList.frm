VERSION 4.00
Begin VB.Form frmAddToList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add file to listing..."
   ClientHeight    =   1695
   ClientLeft      =   2190
   ClientTop       =   2610
   ClientWidth     =   4335
   Height          =   2100
   Icon            =   "frmAddToList.frx":0000
   Left            =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Top             =   2265
   Width           =   4455
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "If you know the name of a file, but it is not listed, type in the name here and it will be added to the list of files shown."
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAddToList"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MpqEx.List.Sorted = False
MpqEx.AddToListing Text1
MpqEx.List.Sorted = True
MpqEx.RemoveDuplicates
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

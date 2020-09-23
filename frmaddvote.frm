VERSION 5.00
Begin VB.Form frmaddvote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Votes"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddvote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Votes"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox T4 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox T3 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox T1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox T2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Current Vote:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Vote:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Candidate ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmaddvote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
T4.Text = Val(T4.Text) + Val(T3.Text)
Set rsAdd = New ADODB.Recordset
rsAdd.Open "Update tvote set votes ='" & T4.Text & "'" & _
"where candidateid=" & T1.Text & "", db, 3, 3
MsgBox T3.Text & " is add to " & T2.Text, vbInformation
T3.Text = ""
frmCandidateList.Timer1.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Set rsCurrentvote = New ADODB.Recordset
rsCurrentvote.Open "Select * from tvote where candidateid=" & frmCandidateList.T1.Text & "", db, 3, 3
T4.Text = rsCurrentvote!votes
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCurrentvote = Nothing
Set rsAdd = Nothing
End Sub

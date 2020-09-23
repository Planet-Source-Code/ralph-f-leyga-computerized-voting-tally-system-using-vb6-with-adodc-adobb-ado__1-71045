VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCandidateList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidates Votes"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCandidateList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   5280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   5760
      Picture         =   "frmCandidateList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add Votes"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Candidates"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Votes Information"
      Height          =   4575
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   6255
      Begin VB.TextBox T3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox T2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox T1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   0
         ScaleHeight     =   3075
         ScaleWidth      =   6195
         TabIndex        =   6
         Top             =   1440
         Width           =   6255
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3135
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Over-all Votes:"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Candidate ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Number of Candidates:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Position:"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select Year:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmCandidateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Call Combo2_Click
End Sub

Private Sub Combo2_Click()
Set rs = New ADODB.Recordset
rs.Open "Select * from qvote where canposition='" & Combo2.Text & "' and voteyear='" & Combo1.Text & "' order by votes desc", db, 3, 3
Set DataGrid1.DataSource = rs
Set rsVote = New ADODB.Recordset
totalresult
dbgrid
End Sub

Private Sub Command1_Click()
frmaddcan.Show vbModal
End Sub

Private Sub Command2_Click()
frmaddvote.T1.Text = T1.Text
frmaddvote.T2.Text = T2.Text
frmaddvote.Show vbModal
End Sub

Private Sub Command3_Click()
'On Error Resume Next
Dim strSearch1 As String
'Dim str1, str2, str3, str4 As String
strSearch1 = InputBox("Search for the ID.", "Search Option")
Set rsSearch = New ADODB.Recordset
'dbase
rsSearch.Open "Select * from qvote where Fullname='" & strSearch1 & "' and Voteyear='" & Combo1.Text & "'", db, 3, 3
Set DataGrid1.DataSource = rsSearch
'totalresult
'db1
dbgrid
End Sub

Private Sub Command4_Click()
Call Combo2_Click
DataGrid1.Enabled = True
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then

Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tvote where candidateID=" & T1.Text & "", db, 3, 3
rsRemove.Delete
Call Command4_Click
MsgBox "Data is remove.", vbInformation
End If

End Sub

Private Sub Command6_Click()
DataGrid1.Enabled = False
Set DataReport1.DataSource = rs
DataReport1.Sections("Section2").Controls("L1").Caption = Combo1.Text
DataReport1.Sections("Section2").Controls("L2").Caption = Combo2.Text
 'DataReport1.Sections("Section5").Controls("L5").Caption = Text1.Text
DataReport1.Show vbModal
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
T1.Text = rs!candidateid
T2.Text = rs!fullname
T3.Text = rs!votes
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Call DataGrid1_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
Call DataGrid1_Click
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Call DataGrid1_Click
End Sub

Private Sub Form_Load()
dbase
Set rsYear = New ADODB.Recordset
rsYear.Open "Select * from tYear order by yr asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsYear.EOF
        Combo1.AddItem rsYear!yr
        rsYear.MoveNext
    Loop
Set rsYear = Nothing
Set rsPositon = New ADODB.Recordset
rsposition.Open "Select * from tposition order by [position] asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsposition.EOF
        Combo2.AddItem rsposition!Position
rsposition.MoveNext
    Loop
Set rsYear = Nothing
End Sub
Public Sub totalresult()
Set rsVote = New ADODB.Recordset
rsVote.Open "Select count(candidateid) as totalCandidate from qvote where voteyear='" & Combo1.Text & "' and canposition='" & Combo2.Text & "'", db, 3, 3
Text1.Text = rsVote!totalCandidate
End Sub

Public Sub dbgrid()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1400
DataGrid1.Columns(3).Width = 1000
DataGrid1.Columns(4).Width = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsposition = Nothing
Set rsYear = Nothing
Set rs = Nothing
Set rsSearch = Nothing
End Sub

Private Sub Timer1_Timer()
Call Combo2_Click
Timer1.Enabled = False
End Sub

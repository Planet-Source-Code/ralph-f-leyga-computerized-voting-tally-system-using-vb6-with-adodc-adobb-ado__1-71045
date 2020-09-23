VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidates Records"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   5760
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   5640
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmRecord.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton Command7 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   28
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   27
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   4800
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00404040&
         Height          =   1335
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   6435
         TabIndex        =   20
         Top             =   3000
         Width           =   6495
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1335
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   2355
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
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   6495
         Begin VB.ComboBox Combo1 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Caption         =   "Personal Information"
            Height          =   1815
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   6255
            Begin VB.TextBox t5 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   4
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox T7 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox t6 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   5
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox t4 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox t3 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   2
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox t2 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               TabIndex        =   1
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox t1 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   0
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Date File:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   23
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Gender:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   17
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Age:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   16
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "First Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "ID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Width           =   255
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Select Year:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Total Information:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
'dbase
Set rs = New ADODB.Recordset
rs.Open "Select * from tinfo where voteyear='" & Combo1.Text & "' order by lastname asc", db, 3, 3
Set DataGrid1.DataSource = rs
'Set rs = Nothing

db1
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "&Modify" Then
Command1.Caption = "&Update"
Command7.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
T2.BackColor = &HFFFFFF
T3.BackColor = &HFFFFFF
T4.BackColor = &HFFFFFF
t5.BackColor = &HFFFFFF
t6.BackColor = &HFFFFFF
t7.BackColor = &HFFFFFF
T2.Locked = False
T3.Locked = False
T4.Locked = False
t5.Locked = False
t6.Locked = False
t7.Locked = False
Else
Command1.Caption = "&Modify"
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command7.Enabled = True
T2.BackColor = &HE0E0E0
T3.BackColor = &HE0E0E0
T4.BackColor = &HE0E0E0
t5.BackColor = &HE0E0E0
t6.BackColor = &HE0E0E0
t7.BackColor = &HE0E0E0
T2.Locked = True
T3.Locked = True
T4.Locked = True
t5.Locked = True
t6.Locked = True
t7.Locked = True
Set rsUpdate = New ADODB.Recordset
rsUpdate.Open "Update tInfo set lastname='" & T2.Text & "', firstname='" & T3.Text & "', middlename='" & T4.Text & "', gndr='" & t5.Text & "', currentage='" & t6.Text & "', datefile='" & t7.Text & "'" & _
"where candidateID=" & T1.Text & "", db, 3, 3
MsgBox "Successfully  Updated.", vbInformation
Call Command5_Click
'Command6.Enabled = False
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim repp As String
repp = MsgBox("Do you want to remove?", vbYesNo, "Confirm Delete")
If repp = vbYes Then

Set rsRemove = New ADODB.Recordset
rsRemove.Open "Select * from tinfo where candidateID=" & T1.Text & "", db, 3, 3
rsRemove.Delete
Call Command5_Click
MsgBox "Data is remove.", vbInformation
End If

End Sub

Private Sub Command3_Click()
DataGrid1.Enabled = False
Set DataReport2.DataSource = rs
DataReport2.Sections("Section3").Controls("L3").Caption = Text7.Text
DataReport2.Sections("Section2").Controls("L2").Caption = Combo1.Text
 'DataReport1.Sections("Section5").Controls("L5").Caption = Text1.Text
DataReport2.Show vbModal
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim strSearch As String
'Dim str1, str2, str3, str4 As String
strSearch = InputBox("Search for the lastname.", "Search Option")
Set rs = New ADODB.Recordset
dbase
rs.Open "Select * from tinfo where lastname='" & strSearch & "' and Voteyear='" & Combo1.Text & "'", db, 3, 3
Set DataGrid1.DataSource = rs
totalresult
db1
End Sub

Private Sub Command5_Click()
Call Combo1_Click
DataGrid1.Enabled = True
db1
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
frmaddinfo.Show vbModal
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
textselect
'Err:
'    MsgBox "No record.", vbExclamation
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
End Sub

Public Sub db1()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
DataGrid1.Columns(1).Width = 1300
DataGrid1.Columns(2).Width = 1300
DataGrid1.Columns(3).Width = 1300
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 900
totalresult
'Set rsYear = Nothing
End Sub

Public Sub textselect()
T1.Text = rs!candidateid
T2.Text = rs!lastname
T3.Text = rs!firstname
T4.Text = rs!middlename
t5.Text = rs!gndr
t6.Text = rs!currentage
t7.Text = rs!datefile
'Exit Sub
End Sub

Public Sub totalresult()
Set rsYear = New ADODB.Recordset
rsYear.Open "Select count(candidateid) as totalCandidate from tinfo where voteyear='" & Combo1.Text & "'", db, 3, 3
Text7.Text = rsYear!totalCandidate
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
Set rsYear = Nothing
Set rsUpdate = Nothing
End Sub

Private Sub Timer1_Timer()
Call Command5_Click
Timer1.Enabled = False
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   3840
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Year"
      TabPicture(0)   =   "frmSetting.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Position"
      TabPicture(1)   =   "frmSetting.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   2295
         Begin VB.CommandButton Command4 
            Caption         =   "Update"
            Height          =   375
            Left            =   1200
            TabIndex        =   14
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Top             =   1680
            Width           =   1815
         End
         Begin VB.PictureBox Picture2 
            Height          =   1335
            Left            =   240
            ScaleHeight     =   1275
            ScaleWidth      =   1755
            TabIndex        =   5
            Top             =   240
            Width           =   1815
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   1335
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   2355
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   17
               RowDividerStyle =   0
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
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
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
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
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
         Begin VB.CommandButton Command3 
            Caption         =   "Update"
            Height          =   375
            Left            =   1200
            TabIndex        =   11
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   1680
            Width           =   1815
         End
         Begin VB.PictureBox Picture1 
            Height          =   1335
            Left            =   240
            ScaleHeight     =   1275
            ScaleWidth      =   1755
            TabIndex        =   2
            Top             =   240
            Width           =   1815
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   1335
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   2355
               _Version        =   393216
               AllowUpdate     =   -1  'True
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   17
               RowDividerStyle =   0
               AllowAddNew     =   -1  'True
               AllowDelete     =   -1  'True
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
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
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
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Set rsAdd = New ADODB.Recordset
rsAdd.Open "Insert into tyear (yr) values ('" & Text1.Text & "');", db, 3, 3
MsgBox Text1.Text & " New Year is added", vbInformation
Call Form_Load
Text1.Text = ""
Set rsAdd = Nothing
End Sub

Private Sub Command2_Click()
Set rsAdd = New ADODB.Recordset
rsAdd.Open "Insert into tposition ([position]) values ('" & Text2.Text & "');", db, 3, 3
MsgBox Text2.Text & " Position is Added.", vbInformation
Call Form_Load
Text2.Text = ""
Set rsAdd = Nothing
End Sub

Private Sub Command3_Click()
Set rsupdateyear = New ADODB.Recordset
rsupdateyear.Open "Update tyear set yr='" & Text1.Text & "'" & _
"where id=" & Text3.Text & "", db, 3, 3
MsgBox "Successfully Changes", vbInformations
Call Form_Load
End Sub

Private Sub Command4_Click()
Set rsupDateposition = New ADODB.Recordset
rsupDateposition.Open "Update tposition set [position]='" & Text2.Text & "'" & _
"where id=" & Text4.Text & "", db, 3, 3
MsgBox "Successfully Changes", vbInformations
Call Form_Load
End Sub

Private Sub DataGrid1_Click()
Text3.Text = rs!id
End Sub

Private Sub DataGrid2_Click()
Text4.Text = rsposition!id
End Sub

Private Sub Form_Load()
dbase
yearsetting
positionsetting
End Sub

Public Sub yearsetting()
Set rs = New ADODB.Recordset
rs.Open "Select * from tyear order by yr asc", db, 3, 3
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Width = 1200
DataGrid1.Columns(1).Visible = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rs = Nothing
Set rsAdd = Nothing
Set rsposition = Nothing
Set rsupDateposition = Nothing
Set rsupdateyear = Nothing
End Sub

Public Sub positionsetting()
'Public Sub yearsetting()
Set rsposition = New ADODB.Recordset
rsposition.Open "Select * from tposition order by [position] asc", db, 3, 3
Set DataGrid2.DataSource = rsposition
DataGrid2.Columns(0).Width = 1200
DataGrid2.Columns(1).Visible = False
End Sub


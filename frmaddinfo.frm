VERSION 5.00
Begin VB.Form frmaddinfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Candidates"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
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
      TabIndex        =   17
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   2520
      TabIndex        =   16
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
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
      Left            =   1440
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERSONAL INFORMATION"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox t7 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox t6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox t5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox t4 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmaddinfo.frx":08CA
         Left            =   1560
         List            =   "frmaddinfo.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox t3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox t2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox t1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   2655
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
         Left            =   240
         TabIndex        =   7
         Top             =   2520
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
         Left            =   240
         TabIndex        =   6
         Top             =   2160
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
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
         Left            =   240
         TabIndex        =   4
         Top             =   1800
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
         TabIndex        =   3
         Top             =   1080
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
         TabIndex        =   2
         Top             =   720
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
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmaddinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
EnableBtn True
EnableFld Me, False
End Sub

Private Sub btnNew_Click()
EnableBtn False
EnableFld Me, True
T1.Text = ""
T2.Text = ""
T3.Text = ""
t5.Text = ""
't2.Text = ""
't2.Text = ""
End Sub
Private Sub EnableBtn(bVal As Boolean)
    btnNew.Enabled = bVal
    btnSave.Enabled = Not bVal
    btnCancel.Enabled = Not bVal
End Sub

Private Sub btnSave_Click()
If T1.Text <> "" And T2.Text <> "" And T3.Text <> "" And T4.Text <> "" And t5.Text <> "" And t6.Text <> "" And t7.Text <> "" Then
EnableBtn True
EnableFld Me, False
Set rsAdd = New ADODB.Recordset
rsAdd.Open "Insert into tinfo (lastname,firstname,middlename,Gndr,currentage,datefile,voteyear) values ('" & T1.Text & "','" & T2.Text & "','" & T3.Text & "','" & T4.Text & "','" & t5.Text & "','" & t6.Text & "','" & t7.Text & "');", db, 3, 3
MsgBox "Data is save.", vbInformation
Else
MsgBox "All fields are required.", vbExclamation
End If
End Sub

Private Sub Form_Load()
Set rsYear = New ADODB.Recordset
rsYear.Open "Select * from tYear order by yr asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsYear.EOF
       t7.AddItem rsYear!yr
        rsYear.MoveNext
    Loop
Set rsYear = Nothing
t6.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsAdd = Nothing
frmRecord.Timer1.Enabled = True
Unload Me
End Sub


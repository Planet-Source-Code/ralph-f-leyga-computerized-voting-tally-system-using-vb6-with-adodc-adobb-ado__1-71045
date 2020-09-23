VERSION 5.00
Begin VB.Form frmaddcan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Candidates"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddcan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Candidate"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Year:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Position:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Candidate ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmaddcan"
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
End Sub

Private Sub btnSave_Click()
On Error GoTo err
If Combo1.Text <> "" And Combo2.Text <> "" And Combo3.Text <> "" Then
EnableBtn True
EnableFld Me, False
Set rsAdd = New ADODB.Recordset
rsAdd.Open "Insert into tvote (candidateid,canposition,voteyear) values ('" & Combo1.Text & "','" & Combo2.Text & "','" & Combo3.Text & "');", db, 3, 3
frmCandidateList.Timer1.Enabled = True
Else
MsgBox "All fields are required!", vbInformation
End If
Exit Sub
err:
    MsgBox "Duplicate", vbInformation
    Unload Me
End Sub

Private Sub Form_Load()
dbase
Set rsCan = New ADODB.Recordset
rsCan.Open "Select * from tinfo order by candidateid asc", db, 3, 3
If rsCan.RecordCount > 0 Then
    Do Until rsCan.EOF
        Combo1.AddItem rsCan!candidateid
        rsCan.MoveNext
    Loop
    End If

Set rsSelectPosition = New ADODB.Recordset
rsSelectPosition.Open "Select * from tposition order by [position] asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsSelectPosition.EOF
        Combo2.AddItem rsSelectPosition!Position
rsSelectPosition.MoveNext
    Loop
Set rsSelectYear = New ADODB.Recordset
rsSelectYear.Open "Select * from tyear order by yr asc", db, 3, 3
'If rsSY.RecordCount > 0 Then
    Do Until rsSelectYear.EOF
        Combo3.AddItem rsSelectYear!yr
rsSelectYear.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsSelectPosition = Nothing
Set rsCan = Nothing
Set rsAdd = Nothing
End Sub
Private Sub EnableBtn(bVal As Boolean)
    btnNew.Enabled = bVal
    btnSave.Enabled = Not bVal
    btnCancel.Enabled = Not bVal
End Sub

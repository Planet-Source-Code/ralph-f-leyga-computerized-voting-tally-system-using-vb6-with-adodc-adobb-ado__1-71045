Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rsVote As New ADODB.Recordset
Public rsYear As New ADODB.Recordset
Public rsSearch As New ADODB.Recordset
Public rsUpdate As New ADODB.Recordset
Public rsAdd As New ADODB.Recordset
Public rsRemove As New ADODB.Recordset
Public rsaddYear As New ADODB.Recordset
Public rsposition As New ADODB.Recordset
Public rsCan As New ADODB.Recordset
Public rsSelectPosition As New ADODB.Recordset
Public rsSelectYear As New ADODB.Recordset
Public rstotalvote As New ADODB.Recordset
Public rsCurrentvote As New ADODB.Recordset
Public rsupdateyear As New ADODB.Recordset
Public rsupDateposition As New ADODB.Recordset
'Public rspos As New ADODB.Recordset
Public bol As Boolean
'Global rpt_header As report_header
'Public rpt_header As report_header
Public Sub dbase()
Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= dbase.mdb ;Persist Security Info=False;Jet OLEDB:Database Password=cheese"
End Sub
Public Sub EnableFld(FormName As Form, bVal As Boolean)
    Dim ObjCtrl As Control
    
    For Each ObjCtrl In FormName.Controls
        If TypeOf ObjCtrl Is TextBox Then
            ObjCtrl.Enabled = bVal
        ElseIf TypeOf ObjCtrl Is ComboBox Then
            ObjCtrl.Enabled = bVal
       ' ElseIf TypeOf ObjCtrl Is DTPicker Then
            ObjCtrl.Enabled = bVal
       ' ElseIf TypeOf ObjCtrl Is DataList Then
           ' ObjCtrl.Enabled = bVal
       '' ElseIf TypeOf ObjCtrl Is DataCombo Then
           ' ObjCtrl.Enabled = bVal
        End If
    Next ObjCtrl
    
    Set ObjCtrl = Nothing
End Sub


Attribute VB_Name = "Module1"
Public rs As ADODB.Recordset
Public cn As ADODB.Connection
Public sql As String
Public Category As String

Public Function SetCon()
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\LibrarySys.mdb;Persist Security Info=False"
    cn.Open
End Function
 Public Function SetRs()
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
End Function

Public Function CloseCon()
    cn.close
    Set cn = Nothing
End Function


Public Function CloseRS()
    rs.close
    Set rs = Nothing
End Function

Public Function CheckNum(text As TextBox)
    If text.text >= 48 And text.text <= 57 Then
        text.Locked = False
    ElseIf vbKeyBack Then
        text.Locked = False
    Else
        text.Locked = True
    End If
End Function

Public Function ClearText(sForm As Form)
Dim control As control
 For Each control In sForm
    If (TypeOf control Is TextBox) Then control = vbNullString
 Next
End Function

Public Function fillStudentID()
    SetCon
    sql = "SELECT * FROM tblStudents"
    SetRs
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        StudentFrm.lstStudID.clear
        Do While Not rs.EOF
            StudentFrm.lstStudID.AddItem rs.Fields("StudID")
            rs.MoveNext
        Loop
    End If
    CloseRS
    CloseCon
End Function

Attribute VB_Name = "UserSession"
Option Explicit
Public username As String
Public role As String
Public foreChange As String

Public Function getLoginUser() As String

  If (username <> vbNullString) Then
   getLoginUser = username
  Else
    getLoginUser = "System"
  End If
  
End Function
Public Function getAllUsers() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "select ID, USERNAME, ROLE, PASSWORD, FORCE_CHANGE from users order by ID "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getAllUsers = rs

End Function

Public Function getUserByUserName(username As String) As ADODB.Recordset
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "select ID, USERNAME, ROLE, PASSWORD, FORCE_CHANGE from users where USERNAME = '" & username & "'"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getUserByUserName = rs
End Function




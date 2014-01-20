Attribute VB_Name = "UserSession"
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

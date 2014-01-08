Attribute VB_Name = "DbInstance"
Option Explicit
Private Const DB_HOST As String = "localhost"
Private Const DB_NAME As String = "LIBRARY_SYSTEM"
Private Const DB_USERNAME As String = "root"
Private Const DB_PASSWORD As String = "mysqladmin"
Private con As ADODB.Connection
Public Function getDBConnetion() As ADODB.Connection

  If (Not con Is Nothing) Then
    If (con.State = adStateOpen) Then
      Set getDBConnetion = con
      Exit Function
    End If
  End If
  
  Set getDBConnetion = createConnection

End Function

Private Function createConnection() As ADODB.Connection
 Set con = New ADODB.Connection
 
 Dim strCon As String
 
 strCon = "Driver={MySQL ODBC 3.51 Driver}; " _
         & "SERVER=" & DB_HOST & "; " _
         & "Database=" & DB_NAME & "; " _
         & "User=" & DB_USERNAME & "; " _
         & "Password=" & DB_PASSWORD & "; " _
         & "Option=3"
 con.ConnectionString = strCon
 con.CursorLocation = adUseClient
 
 con.Open

 Set createConnection = con
End Function

Public Sub closeRecordSet(rs As ADODB.Recordset)
  
  If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
           rs.Close
        End If
  End If
  
End Sub


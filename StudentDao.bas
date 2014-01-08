Attribute VB_Name = "StudentDao"
Option Explicit
Public Function getAllRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.LRN, a.FIRST_NAME, a.MIDDLE_NAME, a.LAST_NAME, a.SECTION_ID   " & _
              "       , CONCAT(b.name, ' - ', b.level) as Section  " & _
              "       , a.CREATED_BY, a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID " & _
              "Order by LAST_MOD_DATE desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getAllRs = rs

End Function

Public Function qucikSearchRs(lrn As Integer, sectionID As Integer, lastName As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.LRN,CONCAT (a.LAST_NAME, ', ', a.FIRST_NAME, ' ', a.MIDDLE_NAME) as Full_Name   " & _
              "       , b.Adviser, CONCAT(b.name, ' - ', b.level) as Section, a.ID  " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID "
              
   If (Not IsNull(lrn) And lrn > 0) Then
     sqlQuery = sqlQuery & " and Cast(a.LRN as char) Like '" & lrn & "%' "
   End If
   
   If (Not IsNull(sectionID) And sectionID > 0) Then
     sqlQuery = sqlQuery & " and a.SECTION_ID Like " & sectionID & " "
   End If
   
                
   If (Not IsNull(lastName) And lastName <> vbNullString) Then
     sqlQuery = sqlQuery & " and Cast(a.LAST_NAME as char) Like '" & lastName & "%' "
   End If
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   Set qucikSearchRs = rs
              
End Function

Public Function searchRs(lrn As Integer, sectionID As Integer, lastName As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.LRN, a.FIRST_NAME, a.MIDDLE_NAME, a.LAST_NAME, a.SECTION_ID   " & _
              "       , CONCAT(b.name, ' - ', b.level) as Section  " & _
              "       , a.CREATED_BY, a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID "
              
   If (Not IsNull(lrn) And lrn > 0) Then
     sqlQuery = sqlQuery & " and Cast(a.LRN as char) Like '" & lrn & "%' "
   End If
   
   If (Not IsNull(sectionID) And sectionID > 0) Then
     sqlQuery = sqlQuery & " and a.SECTION_ID Like " & sectionID & " "
   End If
   
                
   If (Not IsNull(lastName) And lastName <> vbNullString) Then
     sqlQuery = sqlQuery & " and Cast(a.LAST_NAME as char) Like '" & lastName & "%' "
   End If
   
          
   sqlQuery = sqlQuery & " Order by LAST_MOD_DATE desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   Set searchRs = rs
              
End Function
Public Function getFakeRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from STUDENTS " & _
              "Where 1 = 2 "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getFakeRs = rs

End Function
Public Function getRsByID(StudentID As Integer) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from STUDENTS " & _
              "Where ID = " & StudentID
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getRsByID = rs

End Function



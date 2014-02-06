Attribute VB_Name = "StudentDao"
Option Explicit
Public Function getAllRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.LRN, a.FIRST_NAME, a.MIDDLE_NAME, a.LAST_NAME, a.SECTION_ID   " & _
              "       , CONCAT(b.name, ' - ', b.level) as SECTION, a.STATUS  " & _
              "       , a.CREATED_BY, a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID " & _
              "Order by LAST_MOD_DATE desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getAllRs = rs

End Function

Public Function qucikSearchRs(lrn As String, sectionID As Integer, lastName As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.LRN,CONCAT (a.LAST_NAME, ', ', a.FIRST_NAME, ' ', a.MIDDLE_NAME) as Full_Name   " & _
              "       , b.Adviser, CONCAT(b.name, ' - ', b.level) as Section, a.ID  " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID " & _
              "      and a.status = 'Enrolled' "
              
   If (CommonHelper.hasValidValue(lrn)) Then
     sqlQuery = sqlQuery & " and a.LRN Like '" & lrn & "%' "
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

Public Function searchRs(lrn As String, sectionID As Integer, lastName As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.LRN, a.FIRST_NAME, a.MIDDLE_NAME, a.LAST_NAME, a.SECTION_ID   " & _
              "       , CONCAT(b.name, ' - ', b.level) as SECTION, a.STATUS  " & _
              "       , a.CREATED_BY, a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE " & _
              "from STUDENTS a, sections b " & _
              "Where a.SECTION_ID = b.ID "
              
   If (CommonHelper.hasValidValue(lrn)) Then
     sqlQuery = sqlQuery & " and a.LRN Like '" & lrn & "%' "
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
Public Function getRsByID(studentID As Integer) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from STUDENTS " & _
              "Where ID = " & studentID
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getRsByID = rs

End Function
Public Function getRsByLrn(lrn As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from STUDENTS " & _
              "Where LRN = '" & lrn & "'"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getRsByLrn = rs

End Function
Public Function isStudentBeingUsed(id As Integer) As Boolean
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * from transactions where STUDENT_ID = " & id & _
              " limit 1 "

   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   If (rs.RecordCount > 0) Then
     isStudentBeingUsed = True
   Else
     isStudentBeingUsed = False
   End If
   Call closeRecordSet(rs)
   
End Function


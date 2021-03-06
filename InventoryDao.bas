Attribute VB_Name = "InventoryDao"
Option Explicit
Public Function getAllRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.DONATED_BY, a.PURCHASE_COST, a.STATUS, a.CREATED_BY, a.CREATED_DATE " & _
              "       , a.LAST_MOD_BY, a.LAST_MOD_DATE, a.ITEM_TYPE_ID, a.LOCATION_ID, a.CATEGORY_ID " & _
              "From ITEMS a, ITEM_TYPES b " & _
              "     , LOCATION_MAPPINGS c, CATEGORIES d " & _
              "Where a.ITEM_TYPE_ID = b.ID " & _
              "      And a.LOCATION_ID = c.ID " & _
              "      AND a.CATEGORY_ID = d.ID " & _
              "Order by a.LAST_MOD_DATE Desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getAllRs = rs

End Function
Public Function getDashboardEmptyRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.STATUS, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.DONATED_BY, a.PURCHASE_COST, a.CREATED_BY, a.CREATED_DATE " & _
              "       , a.LAST_MOD_BY, a.LAST_MOD_DATE, a.ITEM_TYPE_ID, a.LOCATION_ID, a.CATEGORY_ID, a.ID " & _
              "From ITEMS a, ITEM_TYPES b " & _
              "     , LOCATION_MAPPINGS c, CATEGORIES d " & _
              "Where a.ITEM_TYPE_ID = b.ID " & _
              "      And a.LOCATION_ID = c.ID " & _
              "      AND a.CATEGORY_ID = d.ID " & _
              "      AND 1 = 2 " & _
              "Order by a.LAST_MOD_DATE Desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getDashboardEmptyRs = rs

End Function
Public Function dashboardSearch(itemCode As String, itemTypeID As Integer, author As String, name As String, categoryID As Integer, status As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.STATUS, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.AQUISITION_TYPE, a.DONATED_BY, a.PURCHASE_COST, a.PUBLISHER, a.COPYRIGHT_YEAR " & _
              "       , a.VOLUME, a.ITEM_TYPE_ID, a.LOCATION_ID, a.CATEGORY_ID, a.ID " & _
              "From ITEMS a, ITEM_TYPES b " & _
              "     , LOCATION_MAPPINGS c, CATEGORIES d " & _
              "Where a.ITEM_TYPE_ID = b.ID " & _
              "      And a.LOCATION_ID = c.ID " & _
              "      AND a.CATEGORY_ID = d.ID "
              
   If (CommonHelper.hasValidValue(CStr(categoryID))) Then
      sqlQuery = sqlQuery & " And a.CATEGORY_ID = " & categoryID
   End If
          
   If (CommonHelper.hasValidValue(itemCode)) Then
     sqlQuery = sqlQuery & " And a.ITEM_CODE Like '" & itemCode & "%'"
   End If
   
   If (CommonHelper.hasValidValue(CStr(itemTypeID))) Then
      sqlQuery = sqlQuery & " And a.ITEM_TYPE_ID = " & itemTypeID
   End If
   
   If (CommonHelper.hasValidValue(author)) Then
     sqlQuery = sqlQuery & " And a.AUTHOR Like '" & author & "%'"
   End If
   
   If (CommonHelper.hasValidValue(name)) Then
     sqlQuery = sqlQuery & " And a.name Like '" & name & "%'"
   End If
   
    If (CommonHelper.hasValidValue(status)) Then
     If (status = "Archive") Then
       sqlQuery = sqlQuery & " And a.STATUS in ('Loss', 'Damaged') "
     Else
       sqlQuery = sqlQuery & " And a.STATUS = '" & status & "'"
     End If
   End If
   
   sqlQuery = sqlQuery & " Order by a.LAST_MOD_DATE Desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set dashboardSearch = rs

End Function

Public Function searchItem(Optional itemCode As String, Optional itemTypeID As Integer, Optional author As String, Optional name As String, Optional categoryID As Integer, Optional status As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.STATUS, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.AQUISITION_TYPE, a.DONATED_BY, a.PURCHASE_COST, a.PUBLISHER, a.COPYRIGHT_YEAR " & _
              "       , a.VOLUME, a.CREATED_BY, a.CREATED_DATE, a.LAST_MOD_BY, a.LAST_MOD_DATE, a.ITEM_TYPE_ID " & _
              "       , a.LOCATION_ID, a.CATEGORY_ID " & _
              "From ITEMS a, ITEM_TYPES b " & _
              "     , LOCATION_MAPPINGS c, CATEGORIES d " & _
              "Where a.ITEM_TYPE_ID = b.ID " & _
              "      And a.LOCATION_ID = c.ID " & _
              "      AND a.CATEGORY_ID = d.ID "
              
   If (CommonHelper.hasValidValue(CStr(categoryID))) Then
      sqlQuery = sqlQuery & " And a.CATEGORY_ID = " & categoryID
   End If
          
   If (CommonHelper.hasValidValue(itemCode)) Then
     sqlQuery = sqlQuery & " And a.ITEM_CODE Like '" & itemCode & "%'"
   End If
   
   If (CommonHelper.hasValidValue(CStr(itemTypeID))) Then
      sqlQuery = sqlQuery & " And a.ITEM_TYPE_ID = " & itemTypeID
   End If
   
   If (CommonHelper.hasValidValue(author)) Then
     sqlQuery = sqlQuery & " And a.AUTHOR Like '" & author & "%'"
   End If
   
   If (CommonHelper.hasValidValue(name)) Then
     sqlQuery = sqlQuery & " And a.name Like '" & name & "%'"
   End If
   
   If (CommonHelper.hasValidValue(status)) Then
     If (status = "Archive") Then
       sqlQuery = sqlQuery & " And a.STATUS in ('Loss', 'Damaged') "
     Else
       sqlQuery = sqlQuery & " And a.STATUS = '" & status & "'"
     End If
   End If
   
   sqlQuery = sqlQuery & " Order by a.LAST_MOD_DATE Desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set searchItem = rs

End Function
Public Function getFakeRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from ITEMS " & _
              "Where 1 = 2 "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getFakeRs = rs

End Function
Public Function getRsByItemCode(itemCode As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from ITEMS " & _
              "Where ITEM_CODE = '" & itemCode & "'"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getRsByItemCode = rs

End Function
Public Function getRsByID(itemID As Integer) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from ITEMS " & _
              "Where ID = " & itemID
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getRsByID = rs

End Function
Public Function getTransactionsRS() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from ITEMS " & _
              "Where 1 = 2 "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getTransactionsRS = rs

End Function
Public Function getFakeTransactionRS() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from transactions " & _
              "Where 1 = 2 "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getFakeTransactionRS = rs

End Function
Public Function getOverdueTransactions()
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select  " & _
              "        CASE   " & _
              "        When WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') <= 0 Then 'Over Due' " & _
              "        Else WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') " & _
              "        END as REMAINING_DAYS " & _
              "       , item.ITEM_CODE as ISBN, item.Name as TITLE, itype.NAME as TYPE, categ.name as CATEGORY, stud.LRN " & _
              "       , CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , REQUESTED_RETURN_DATE as DUE_DATE, tran.ID as TRANSACTION_ID " & _
              "From transactions tran, items item " & _
              "     , item_types as itype, STUDENTS stud " & _
              "     , categories categ " & _
              "Where tran.ITEM_ID = item.ID " & _
              "      And itype.ID = item.ITEM_TYPE_ID " & _
              "      And categ.ID = item.CATEGORY_ID " & _
              "      And tran.STUDENT_ID = stud.ID " & _
              "      And tran.RETURN_DATE is null " & _
              "      AND WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') <= 0"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
      
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   Set getOverdueTransactions = rs
End Function

Public Function getTransactionDashboardRs(Optional lrn As String, Optional itemCode As String, Optional title As String, Optional itemType As String, Optional category As String) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select  " & _
              "        CASE   " & _
              "        When WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') <= 0 Then 'Over Due' " & _
              "        Else WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') " & _
              "        END as REMAINING_DAYS " & _
              "       , item.ITEM_CODE as ISBN, item.Name as TITLE, itype.NAME as TYPE, categ.name as CATEGORY, stud.LRN " & _
              "       , CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , REQUESTED_RETURN_DATE as DUE_DATE, tran.ID as TRANSACTION_ID " & _
              "From transactions tran, items item " & _
              "     , item_types as itype, STUDENTS stud " & _
              "     , categories categ " & _
              "Where tran.ITEM_ID = item.ID " & _
              "      And itype.ID = item.ITEM_TYPE_ID " & _
              "      And categ.ID = item.CATEGORY_ID " & _
              "      And tran.STUDENT_ID = stud.ID " & _
              "      And tran.RETURN_DATE is null "
              
    If (CommonHelper.hasValidValue(lrn)) Then
       sqlQuery = sqlQuery & " And stud.LRN like '" & lrn & "%' "
    End If
    
    If (CommonHelper.hasValidValue(itemCode)) Then
       sqlQuery = sqlQuery & " And item.ITEM_CODE like '" & itemCode & "%' "
    End If
    
    If (CommonHelper.hasValidValue(title)) Then
       sqlQuery = sqlQuery & " And item.Name like '%" & title & "%' "
    End If
    
    If (CommonHelper.hasValidValue(itemType)) Then
       sqlQuery = sqlQuery & " And itype.NAME = '" & itemType & "' "
    End If
    
    If (CommonHelper.hasValidValue(category)) Then
       sqlQuery = sqlQuery & " And categ.name = '" & category & "' "
    End If
              
    sqlQuery = sqlQuery & "ORDER BY REMAINING_DAYS "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getTransactionDashboardRs = rs

End Function
Public Function getStudentBorrower(itemID As Integer) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select stud.LRN, CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , sec.Adviser, CONCAT(sec.name, ' - ', sec.level) as Section " & _
              "from transactions tran, STUDENTS stud " & _
              "     , sections sec " & _
              "where tran.ITEM_ID = " & itemID & _
              "      and tran.STUDENT_ID = stud.ID " & _
              "      and stud.SECTION_ID = sec.ID "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getStudentBorrower = rs

End Function
Public Function getTransaction(transactionID As Integer)
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * " & _
              "from transactions " & _
              "Where ID = " & transactionID
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getTransaction = rs
End Function
Public Function getTransactionInfo(transactionID As Integer)
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select item.ITEM_CODE, itype.NAME as ITEM_TYPE, cat.name as CATEGORY, item.Name as ITEM_NAME, item.AUTHOR " & _
              "       , stud.LRN, CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , sec.Adviser, CONCAT(sec.name, ' - ', sec.level) as Section, tran.LEND_DATE as BORROWED_DATE " & _
              "       ,REQUESTED_RETURN_DATE as DUE_DATE, WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') as REMAINING_DAYS " & _
              "from transactions tran, STUDENTS stud " & _
              "     , sections sec, items item " & _
              "     , item_types as itype " & _
              "     , categories cat " & _
              "where tran.ID = " & transactionID & _
              "      and tran.STUDENT_ID = stud.ID " & _
              "      and tran.ITEM_ID = item.ID " & _
              "      and stud.SECTION_ID = sec.ID " & _
              "      and itype.ID = item.ITEM_TYPE_ID " & _
              "      and cat.ID = item.CATEGORY_ID "
           
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getTransactionInfo = rs
End Function
Public Function getTransactionReport(startDate As Date, endDate As Date, remFilter As String)
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   Dim sqlQuery As String
   
   sqlQuery = "Select " & _
              "        CASE   " & _
              "        When tran.RETURN_DATE is not null then 'Returned' " & _
              "        When WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') <= 0 Then 'Over Due' " & _
              "        Else WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') " & _
              "        END as REMAINING_DAYS " & _
              "       , item.ITEM_CODE as ISBN, itype.NAME as ITEM_TYPE, cat.name as CATEGORY, item.id as Accession_No, item.Name as TITLE, item.AUTHOR " & _
              "       , stud.LRN, CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , sec.ADVISER, CONCAT(sec.name, ' - ', sec.level) as SECTION, tran.LEND_BY as RELEASED_BY " & _
              "       , tran.LEND_DATE as BORROWED_DATE, REQUESTED_RETURN_DATE as DUE_DATE, tran.RETURN_DATE, tran.RECEIVED_BY  " & _
              "from transactions tran, STUDENTS stud " & _
              "     , sections sec, items item " & _
              "     , item_types as itype " & _
              "     , categories cat " & _
              "where tran.STUDENT_ID = stud.ID " & _
              "      and tran.ITEM_ID = item.ID " & _
              "      and stud.SECTION_ID = sec.ID " & _
              "      and itype.ID = item.ITEM_TYPE_ID " & _
              "      and cat.ID = item.CATEGORY_ID " & _
              "      and tran.LEND_DATE between STR_TO_DATE('" & Format(startDate, "mm/dd/yyyy") & "','%m/%d/%Y') and STR_TO_DATE('" & Format(endDate, "mm/dd/yyyy") & "','%m/%d/%Y') "

   If (CommonHelper.hasValidValue(remFilter)) Then
              sqlQuery = sqlQuery & "      and  (CASE   " & _
              "        When tran.RETURN_DATE is not null then 'Returned' " & _
              "        When WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') <= 0 Then 'Over Due' " & _
              "        Else WORKDAYS_LEFT(REQUESTED_RETURN_DATE, '') " & _
              "        END)  = '" & remFilter & "' "
   End If
              
   sqlQuery = sqlQuery & " Order By  BORROWED_DATE "
           
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getTransactionReport = rs
End Function

Public Function getFakeTransactionReportRs()
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select item.ITEM_CODE as ISBN, itype.NAME as ITEM_TYPE, cat.name as CATEGORY, item.Name as TITLE, item.AUTHOR " & _
              "       , stud.LRN, CONCAT (stud.LAST_NAME, ', ', stud.FIRST_NAME, ' ', stud.MIDDLE_NAME) as STUDENT_NAME " & _
              "       , sec.ADVISER, CONCAT(sec.name, ' - ', sec.level) as SECTION, tran.LEND_BY as RELEASED_BY " & _
              "       , tran.LEND_DATE as BORROWED_DATE, REQUESTED_RETURN_DATE as DUE_DATE, tran.RETURN_DATE, tran.RECEIVED_BY  " & _
              "from transactions tran, STUDENTS stud " & _
              "     , sections sec, items item " & _
              "     , item_types as itype " & _
              "     , categories cat " & _
              "where tran.STUDENT_ID = stud.ID " & _
              "      and tran.ITEM_ID = item.ID " & _
              "      and stud.SECTION_ID = sec.ID " & _
              "      and itype.ID = item.ITEM_TYPE_ID " & _
              "      and cat.ID = item.CATEGORY_ID " & _
              "      and 1 = 2 "
       
           
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getFakeTransactionReportRs = rs
End Function

Public Function getBookStatRs() As ADODB.Recordset
    Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "select STATUS as BOOKS, COUNT(*) as Total " & _
              "from items " & _
              "GROUP BY STATUS " & _
              "ORDER BY STATUS "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getBookStatRs = rs
End Function
Public Function isItemBeingUsed(itemID As Integer) As Boolean
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select * from transactions where ITEM_ID = " & itemID & _
              " limit 1 "

   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   If (rs.RecordCount > 0) Then
     isItemBeingUsed = True
   Else
     isItemBeingUsed = False
   End If
   Call closeRecordSet(rs)
   
End Function

Public Function getItemNewID() As Long
   
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select ID + 1 as newID from items order by ID Desc " & _
              " limit 1 "

   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   If (rs.RecordCount > 0) Then
    getItemNewID = rs!newID
   Else
     getItemNewID = 1
   End If
   Call closeRecordSet(rs)
   
End Function
Public Function getQuantityReport(Optional isbn As String, Optional title As String, Optional itemType As String, Optional category As String) As ADODB.Recordset
  Dim con As ADODB.Connection
  Set con = DbInstance.getDBConnetion
  Dim sqlQuery As String
  
  sqlQuery = "Select a.ITEM_CODE as ISBN, b.name as ITEM_TYPE , a.NAME as Title, d.name AS CATEGORY, count(*) as QUANTITY " & _
             "       , (select count(*) from items  where item_code = a.item_code and name = a.name and status = 'Available') as AVAILABLE " & _
             "       , (select count(*) from items  where item_code = a.item_code and name = a.name and status = 'Borrowed') as BORROWED " & _
             "       , (select count(*) from items  where item_code = a.item_code and name = a.name and status = 'Damaged') as DAMAGED " & _
             "       , (select count(*) from items  where item_code = a.item_code and name = a.name and status = 'Loss') as LOSS " & _
             "From items a, ITEM_TYPES b, CATEGORIES d " & _
             "Where a.ITEM_TYPE_ID = b.ID  " & _
             "      AND a.CATEGORY_ID = d.ID "
             
             
  If (CommonHelper.hasValidValue(isbn)) Then
    sqlQuery = sqlQuery & " And a.ITEM_CODE like '" & isbn & "%' "
  End If
  
  If (CommonHelper.hasValidValue(title)) Then
    sqlQuery = sqlQuery & " And a.NAME like '%" & title & "%' "
  End If
  
  If (CommonHelper.hasValidValue(itemType)) Then
    sqlQuery = sqlQuery & " And b.NAME = '" & itemType & "' "
  End If
  
  If (CommonHelper.hasValidValue(category)) Then
    sqlQuery = sqlQuery & " And d.name = '" & category & "' "
  End If
  
  sqlQuery = sqlQuery & " Group by a.ITEM_CODE, a.name " & _
             "         , b.name, d.name "
             
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
  Set getQuantityReport = rs
End Function

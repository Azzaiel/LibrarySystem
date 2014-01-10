Attribute VB_Name = "InventoryDao"
Option Explicit
Public Function getAllRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.DONATED_BY, a.STATUS, a.CREATED_BY, a.CREATED_DATE " & _
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
Public Function getEmptyRs() As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.DONATED_BY, a.STATUS, a.CREATED_BY, a.CREATED_DATE " & _
              "       , a.LAST_MOD_BY, a.LAST_MOD_DATE, a.ITEM_TYPE_ID, a.LOCATION_ID, a.CATEGORY_ID " & _
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
   
   Set getEmptyRs = rs

End Function

Public Function search(itemCode As String, itemTypeID As Integer, author As String, name As String, categoryID As Integer) As ADODB.Recordset

   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select a.ID, a.ITEM_CODE, b.name as ITEM_TYPE, a.NAME, c.NAME as LOCATION,  d.name AS CATEGORY " & _
              "       , a.DESCRIPTION, a.AUTHOR, a.DONATED_BY, a.STATUS, a.CREATED_BY, a.CREATED_DATE " & _
              "       , a.LAST_MOD_BY, a.LAST_MOD_DATE, a.ITEM_TYPE_ID, a.LOCATION_ID, a.CATEGORY_ID " & _
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
   
   sqlQuery = sqlQuery & " Order by a.LAST_MOD_DATE Desc "
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set search = rs

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


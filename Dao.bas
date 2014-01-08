Attribute VB_Name = "LookupDao"
Option Explicit
Public Function getLocationMappingRS() As ADODB.Recordset
   
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "select ID, NAME, FILE_NAME, CREATED_BY, CREATED_DATE, LAST_MOD_BY, LAST_MOD_DATE " & _
              "from location_mappings " & _
              "Order by ID"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getLocationMappingRS = rs
   
End Function
Public Function getItemTypesRs() As ADODB.Recordset
   
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select ID, NAME, DESCRIPTION, CREATED_BY, CREATED_DATE " & _
              "       , LAST_MOD_BY, LAST_MOD_DATE " & _
              "From ITEM_TYPES " & _
              "Order by ID"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getItemTypesRs = rs
   
End Function
Public Function getLocationImgName(locationID As String) As String
  
  Dim con As ADODB.Connection
  Set con = DbInstance.getDBConnetion
  
  Dim sqlQuery As String
  sqlQuery = "select FILE_NAME from location_mappings Where ID = " & locationID
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
   
  rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic

  If (rs.RecordCount > 0) Then
    getLocationImgName = rs!FILE_NAME
  Else
    getLocationImgName = ""
  End If
  
  Call DbInstance.closeRecordSet(rs)
  
End Function
Public Function getCategoriesItemList() As Variant
  Dim itemList() As Variant
  Dim con As ADODB.Connection
  Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   sqlQuery = "select ID, NAME from categories; "
   
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   ReDim itemList(0 To rs.RecordCount, 0 To 1) As Variant
   Dim index As Integer
   index = 0
   While Not rs.EOF
     itemList(index, Constants.ITEM_VALUE_INDEX) = rs!ID
     itemList(index, Constants.ITEM_LABEL_INDEX) = rs!Name
     index = index + 1
     rs.MoveNext
   Wend
   
   getCategoriesItemList = itemList
   Call DbInstance.closeRecordSet(rs)
End Function
Public Function getLocationMappingItemList() As Variant
  Dim itemList() As Variant
  Dim con As ADODB.Connection
  Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   sqlQuery = "select ID, NAME from location_mappings order by NAME "
   
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   ReDim itemList(0 To rs.RecordCount, 0 To 1) As Variant
   Dim index As Integer
   index = 0
   While Not rs.EOF
     itemList(index, Constants.ITEM_VALUE_INDEX) = rs!ID
     itemList(index, Constants.ITEM_LABEL_INDEX) = rs!Name
     index = index + 1
     rs.MoveNext
   Wend
   
   getLocationMappingItemList = itemList
   Call DbInstance.closeRecordSet(rs)
End Function

Public Function getItemTypeItemList() As Variant

  Dim itemList() As Variant
  Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   sqlQuery = "select ID, NAME from item_types"
   
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   ReDim itemList(0 To rs.RecordCount, 0 To 1) As Variant
   Dim index As Integer
   index = 0
   While Not rs.EOF
     itemList(index, Constants.ITEM_VALUE_INDEX) = rs!ID
     itemList(index, Constants.ITEM_LABEL_INDEX) = rs!Name
     index = index + 1
     rs.MoveNext
   Wend
   
   getItemTypeItemList = itemList
   Call DbInstance.closeRecordSet(rs)
   
End Function


Public Function getSectionsItemList() As Variant

  Dim itemList() As Variant
  Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   sqlQuery = "select ID, CONCAT(name, ' - ', level) as Section  from sections"
   
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   ReDim itemList(0 To rs.RecordCount, 0 To 1) As Variant
   Dim index As Integer
   index = 0
   While Not rs.EOF
     itemList(index, Constants.ITEM_VALUE_INDEX) = rs!ID
     itemList(index, Constants.ITEM_LABEL_INDEX) = rs!Section
     index = index + 1
     rs.MoveNext
   Wend
   
   getSectionsItemList = itemList
   Call DbInstance.closeRecordSet(rs)
   
End Function
Public Function getCategoriesRs() As ADODB.Recordset
   
   Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select ID, NAME, DESCRIPTION, CREATED_BY, CREATED_DATE " & _
              "       , LAST_MOD_BY, LAST_MOD_DATE " & _
              "From CATEGORIES " & _
              "Order by ID"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getCategoriesRs = rs
   
End Function

Public Function getSections() As ADODB.Recordset
 Dim con As ADODB.Connection
   Set con = DbInstance.getDBConnetion
   
   Dim sqlQuery As String
   
   sqlQuery = "Select ID, NAME, LEVEL, ADVISER, CREATED_BY " & _
              "       , CREATED_DATE, LAST_MOD_BY " & _
              "       , LAST_MOD_DATE " & _
              "From Sections " & _
              "Order by ID"
              
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   rs.Open sqlQuery, con, adOpenDynamic, adLockPessimistic
   
   Set getSections = rs
End Function



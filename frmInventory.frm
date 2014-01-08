VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInventory 
   Caption         =   "Inventory"
   ClientHeight    =   11970
   ClientLeft      =   3765
   ClientTop       =   450
   ClientWidth     =   20655
   LinkTopic       =   "Form1"
   ScaleHeight     =   11970
   ScaleWidth      =   20655
   Begin VB.Frame Frame2 
      Caption         =   "Search Form"
      Height          =   1215
      Left            =   7560
      TabIndex        =   35
      Top             =   240
      Width           =   12855
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         Height          =   435
         Left            =   9840
         TabIndex        =   47
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   435
         Left            =   6840
         TabIndex        =   46
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtSearchAuthor 
         Height          =   285
         Left            =   7920
         TabIndex        =   44
         Top             =   240
         Width           =   4815
      End
      Begin VB.ComboBox cmSearchCategory 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchName 
         Height          =   285
         Left            =   4680
         TabIndex        =   40
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmSearchType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchItemCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FF80&
         Caption         =   "Author"
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "Category"
         Height          =   255
         Left            =   3480
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FF80&
         Caption         =   " Name"
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FF80&
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Item Code"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmbNewRec 
      Caption         =   "Add"
      Height          =   495
      Left            =   480
      TabIndex        =   34
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmbEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1800
      TabIndex        =   33
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmbDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   32
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmbClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmbClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   11160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item"
      Height          =   10695
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox cmStatus 
         Height          =   315
         ItemData        =   "frmInventory.frx":0000
         Left            =   1560
         List            =   "frmInventory.frx":0010
         TabIndex        =   48
         Text            =   "cmStatus"
         Top             =   8760
         Width           =   1935
      End
      Begin VB.TextBox txtDonatedBy 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   8400
         Width           =   4815
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   7920
         Width           =   4815
      End
      Begin VB.TextBox txtDescription 
         Height          =   1005
         Left            =   1560
         TabIndex        =   16
         Top             =   6720
         Width           =   4815
      End
      Begin VB.ComboBox cmCategory 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   6240
         Width           =   1935
      End
      Begin VB.PictureBox imgLoc 
         Height          =   3735
         Left            =   360
         Picture         =   "frmInventory.frx":0038
         ScaleHeight     =   3675
         ScaleWidth      =   6195
         TabIndex        =   12
         Top             =   2400
         Width           =   6255
      End
      Begin VB.ComboBox cmLocation 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox cmItemType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   10200
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   9840
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   9480
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   9120
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   10200
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod By"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   9840
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "Created Date"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   9480
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "Created By"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   9120
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "Status"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   8760
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Donated By"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   8400
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Author"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   7920
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Description"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "*Category"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H0080FF80&
         Caption         =   "*Location"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "* Name"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "*Type"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "Item Code"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
   End
   Begin MSDataGridLib.DataGrid dgItems 
      Height          =   9375
      Left            =   7560
      TabIndex        =   0
      Top             =   1560
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   16536
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private itemTypeItemList() As Variant
Private locationItemList() As Variant
Private categoriesItemList() As Variant
Private Function getLocationID() As Integer
  Dim index As Integer
  index = cmLocation.ListIndex
  If (index <> -1) Then
    getLocationID = locationItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getLocationID = 0
  End If
End Function
Private Function getItemTypeID() As Integer
  Dim index As Integer
  index = cmItemType.ListIndex
  If (index <> -1) Then
    getItemTypeID = itemTypeItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getItemTypeID = 0
  End If
End Function
Private Function getCategoryID() As Integer
  Dim index As Integer
  index = cmCategory.ListIndex
  If (index <> -1) Then
    getCategoryID = categoriesItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getCategoryID = 0
  End If
End Function

Private Sub cmbClear_Click()
  Call clearForm
End Sub

Private Sub cmbClose_Click()
 Unload Me
End Sub

Private Sub cmbDelete_Click()
 'Call resetFromSkin
 Dim response As String
 response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
  If (response = vbOK) Then
    Set tempRs = InventoryDao.getRsByID(rs!ID)
    tempRs.Delete
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Deleted", vbInformation
    Call clearForm
    Call populateDataGrid
  End If
End Sub
Private Sub clearForm()

     lblID.Caption = ""
    txtName.Text = ""
    txtItemCode.Text = ""
    txtDescription.Text = ""
    txtDonatedBy.Text = ""
    txtAuthor.Text = ""
    lblCreatedBy.Caption = ""
    lblCreatedDate.Caption = ""
    lblLatModBy.Caption = ""
    lblLastModDate.Caption = ""
    cmStatus.ListIndex = -1
    
    cmItemType.ListIndex = -1
    cmLocation.ListIndex = -1
    cmCategory.ListIndex = -1

End Sub

Private Sub cmbEdit_Click()
    Set tempRs = InventoryDao.getRsByID(rs!ID)
    tempRs!name = txtName.Text
    tempRs!ITEM_CODE = txtItemCode.Text
    tempRs!Description = txtDescription.Text
    tempRs!DONATED_BY = txtDonatedBy.Text
    tempRs!author = txtAuthor.Text
    tempRs!Status = cmStatus.Text
    tempRs!LOCATION_ID = getLocationID
    tempRs!ITEM_TYPE_ID = getItemTypeID
    tempRs!CATEGORY_ID = getCategoryID
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Updated!!", vbInformation
    Call populateDataGrid

End Sub

Private Sub cmbNewRec_Click()

    Set tempRs = InventoryDao.getFakeRs
    tempRs.AddNew
    tempRs!name = txtName.Text
    tempRs!ITEM_CODE = txtItemCode.Text
    tempRs!Description = txtDescription.Text
    tempRs!DONATED_BY = txtDonatedBy.Text
    tempRs!author = txtAuthor.Text
    tempRs!Status = cmStatus.Text
    tempRs!LOCATION_ID = getLocationID
    tempRs!ITEM_TYPE_ID = getItemTypeID
    tempRs!CATEGORY_ID = getCategoryID
    tempRs!CREATED_BY = UserSession.getLoginUser
    tempRs!CREATED_DATE = Now
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Created!!", vbInformation
    Call populateDataGrid

End Sub

Private Sub cmdClearSearch_Click()
  txtSearchItemCode.Text = ""
  cmSearchType.ListIndex = -1
  txtSearchName.Text = ""
  cmSearchCategory.ListIndex = -1
  txtSearchAuthor = ""
End Sub

Private Sub cmdSearch_Click()
  Set dgItems.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = InventoryDao.search(txtSearchItemCode, getSearchItemTypeID, txtSearchAuthor, txtSearchName, getSearchCategoryID)
  Set dgItems.DataSource = rs
  dgItems.Refresh
  Call clearForm
  Call formatDataGrid
End Sub
Private Function getSearchCategoryID() As Integer
  Dim index As Integer
  index = cmSearchCategory.ListIndex
  If (index <> -1) Then
    getSearchCategoryID = categoriesItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getSearchCategoryID = 0
  End If
End Function

Private Function getSearchItemTypeID() As Integer
  Dim index As Integer
  index = cmSearchType.ListIndex
  If (index <> -1) Then
    getSearchItemTypeID = itemTypeItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getSearchItemTypeID = 0
  End If
End Function

Private Sub cmLocation_Click()
  Dim fileName As String
  fileName = LookupDao.getLocationImgName(getLocationID)
  If (fileName <> vbNullString) Then
    imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & fileName)
  Else
     imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & Constants.MISSING_LOC_IMAGE_NAME)
  End If
End Sub

Private Sub cmSearchCategory_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

Private Sub cmSearchType_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

Private Sub dgItems_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedData
End Sub

Private Sub dgItems_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Private Sub Form_Load()
  Call populateDropDown
  Call populateDataGrid
End Sub
Private Sub populateDataGrid()
  Set rs = InventoryDao.getAllRs
  Set dgItems.DataSource = rs
  dgItems.Refresh
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  Else
    Call clearForm
  End If
  
  With dgItems
     'ID - 0
    .Columns(0).Width = 400
    .Columns(0).Alignment = dbgCenter

     'CREATED DATE - 11
    .Columns(11).Width = 1500
    .Columns(11).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(11).Alignment = dbgCenter
    
    'LAST MOD DATE - 13
    .Columns(13).Width = 1500
    .Columns(13).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(13).Alignment = dbgCenter
    
    .Columns(14).Visible = False
    .Columns(15).Visible = False
    .Columns(16).Visible = False
    
  End With
End Sub
Private Sub showSelectedData()
  
    lblID.Caption = rs!ID
    txtName.Text = CommonHelper.extractStringValue(rs!name)
    txtItemCode.Text = CommonHelper.extractStringValue(rs!ITEM_CODE)
    txtDescription.Text = CommonHelper.extractStringValue(rs!Description)
    txtDonatedBy.Text = CommonHelper.extractStringValue(rs!DONATED_BY)
    txtAuthor.Text = CommonHelper.extractStringValue(rs!author)
    lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
    lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
    lblLatModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
    lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
    cmStatus.Text = CommonHelper.extractStringValue(rs!Status)
    
    cmItemType.ListIndex = -1
    cmLocation.ListIndex = -1
    cmCategory.ListIndex = -1
    
    Dim index As Integer

   For index = 0 To UBound(itemTypeItemList)
     If (rs!ITEM_TYPE_ID = Val(itemTypeItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmItemType.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(locationItemList)
     If (rs!LOCATION_ID = Val(locationItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmLocation.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(categoriesItemList)
     If (rs!CATEGORY_ID = Val(categoriesItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmCategory.ListIndex = index
     End If
   Next index
   
End Sub
Private Sub populateDropDown()
  Dim index As Integer

  itemTypeItemList = LookupDao.getItemTypeItemList
  cmItemType.Clear
  cmSearchType.Clear
  For index = 0 To UBound(itemTypeItemList)
    cmItemType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
    cmSearchType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  locationItemList = LookupDao.getLocationMappingItemList
  cmLocation.Clear
  For index = 0 To UBound(locationItemList)
    cmLocation.AddItem (locationItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  categoriesItemList = LookupDao.getCategoriesItemList
  cmCategory.Clear
  cmSearchCategory.Clear
  For index = 0 To UBound(categoriesItemList)
    cmCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
    cmSearchCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
End Sub

Private Sub txtSearchAuthor_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

Private Sub txtSearchItemCode_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

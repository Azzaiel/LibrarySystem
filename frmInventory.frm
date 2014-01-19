VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInventory 
   Caption         =   "Inventory"
   ClientHeight    =   10245
   ClientLeft      =   330
   ClientTop       =   450
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   18765
   Begin VB.Frame Frame2 
      Caption         =   "Search Form"
      Height          =   1215
      Left            =   7200
      TabIndex        =   44
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmbExport 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearSearch 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtSearchAuthor 
         Height          =   285
         Left            =   7440
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox cmSearchCategory 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchName 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cmSearchType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchItemCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FF80&
         Caption         =   "Author"
         Height          =   255
         Left            =   6840
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "Category"
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FF80&
         Caption         =   " Name"
         Height          =   255
         Left            =   3360
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FF80&
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Item Code"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmbNewRec 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmbEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmbDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmbClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmbClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item"
      Height          =   9495
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   6975
      Begin VB.ComboBox cmStatus 
         Height          =   315
         ItemData        =   "frmInventory.frx":0000
         Left            =   1560
         List            =   "frmInventory.frx":0010
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Text            =   "cmStatus"
         Top             =   7680
         Width           =   1935
      End
      Begin VB.TextBox txtDonatedBy 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   7320
         Width           =   4815
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   6960
         Width           =   4815
      End
      Begin VB.TextBox txtDescription 
         Height          =   645
         Left            =   1560
         TabIndex        =   5
         Top             =   6240
         Width           =   4815
      End
      Begin VB.ComboBox cmCategory 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   5880
         Width           =   1935
      End
      Begin VB.PictureBox imgLoc 
         Height          =   3735
         Left            =   360
         Picture         =   "frmInventory.frx":0038
         ScaleHeight     =   3675
         ScaleWidth      =   6195
         TabIndex        =   30
         Top             =   2040
         Width           =   6255
      End
      Begin VB.ComboBox cmLocation 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   5055
      End
      Begin VB.ComboBox cmItemType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   9120
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   8040
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   9120
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod By"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   8760
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "Created Date"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   8400
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "Created By"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   8040
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "Status"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   7680
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Donated By"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   7320
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Author"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   6960
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Description"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "*Category"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H0080FF80&
         Caption         =   "*Location"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "* Name"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "*Type"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "*Item Code"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
   End
   Begin MSDataGridLib.DataGrid dgItems 
      Height          =   8775
      Left            =   7200
      TabIndex        =   22
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15478
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
  Call restoreFormDefaultSkin
  Call clearForm
  Call toogelInsertMode(False)
End Sub

Private Sub cmbClose_Click()
 Unload Me
End Sub

Private Sub cmbDelete_Click()
 Dim response As String
 response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
 
  If (InventoryDao.isItemBeingUsed(rs!id)) Then
    MsgBox "Cannot delete an Item that is already used for a transaction", vbCritical
    Exit Sub
  End If
 
  If (response = vbOK) Then
    Set tempRs = InventoryDao.getRsByID(rs!id)
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
    Call restoreFormDefaultSkin
    If (isFormDetailValid) Then
      If (isItemCodeAlreadyExist(rs!ITEM_CODE)) Then
        MsgBox "Item Code Already in use", vbCritical
        Exit Sub
      End If
      Set tempRs = InventoryDao.getRsByID(rs!id)
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
    End If
End Sub
Private Function isItemCodeAlreadyExist(Optional excludeItemCode As String = "") As Boolean
  Set tempRs = InventoryDao.getRsByItemCode(txtItemCode)
  Dim inUse As Boolean
  inUse = False
  If (tempRs.RecordCount > 0) Then
    If (tempRs!ITEM_CODE <> excludeItemCode) Then
      inUse = True
    End If
  End If
  isItemCodeAlreadyExist = inUse
  Call DbInstance.closeRecordSet(tempRs)
End Function

Private Sub cmbExport_Click()

  Dim excelApp As New Excel.Application
  Dim oBook As New Excel.Workbook
  Dim oSheet As New Excel.Worksheet
  
  Set excelApp = CreateObject("Excel.Application")
  Set oBook = excelApp.Workbooks.Open(CommonHelper.getTemplatesPath & "\" & Constants.INVENTORY_TEMPLATE)
  Set oSheet = excelApp.Worksheets(1)
  
  oSheet.name = "Transaction Report"
  
  oSheet.Range("A11").CopyFromRecordset dgItems.DataSource
  oSheet.Columns.AutoFit
  oSheet.Range("O1:Q1").EntireColumn.Hidden = True
  
  Dim availableCount As Long
  Dim borrowedCount As Long
  Dim damageCount As Long
  Dim lossCount As Long
    
  availableCount = 0
  borrowedCount = 0
  damageCount = 0
  lossCount = 0
  rs.MoveFirst
  Dim itemStatus As String
  While Not rs.EOF
    itemStatus = CommonHelper.extractStringValue(rs!Status)
    If (itemStatus = "Available") Then
      availableCount = availableCount + 1
    ElseIf (itemStatus = "Borrowed") Then
      borrowedCount = borrowedCount + 1
    ElseIf (itemStatus = "Damaged") Then
      damageCount = damageCount + 1
    ElseIf (itemStatus = "Loss") Then
      lossCount = lossCount + 1
    End If
    rs.MoveNext
  Wend
  rs.MoveFirst
  
  oSheet.Range("A4").value = "Available"
  oSheet.Range("C4").value = availableCount
  
  oSheet.Range("A5").value = "Borrowed"
  oSheet.Range("C5").value = borrowedCount
  
  oSheet.Range("A6").value = "Damaged"
  oSheet.Range("C6").value = damageCount
  
  oSheet.Range("A7").value = "Loss"
  oSheet.Range("C7").value = lossCount
  
  oSheet.Range("C8").value = availableCount + borrowedCount + damageCount + lossCount
  
  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\" & Constants.TEMP_WORK_BOOK
  
  If (UserSession.role = "Admin") Then
    excelApp.Visible = True
  Else
    Dim pdfFilePat As String
    pdfFilePat = CommonHelper.getTempPath & "\temp_" & Format(Now, "mmhhyysssh") & ".pdf"
    Call oBook.ExportAsFixedFormat(xlTypePDF, pdfFilePat, xlQualityStandard, False, True)
    oBook.Close
    Call CommonHelper.openFile(pdfFilePat, Me.hWnd)
  End If

End Sub

Private Sub cmbNewRec_Click()
      Call restoreFormDefaultSkin
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
    txtItemCode.SetFocus
  Else
    If (isFormDetailValid) Then
      If (isItemCodeAlreadyExist) Then
        MsgBox "Item Code Already in use", vbCritical
        Exit Sub
      End If
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
      Call toogelInsertMode(False)
    End If
  End If
End Sub
Private Sub restoreFormDefaultSkin()
  Call CommonHelper.toDefaultSkin(txtItemCode)
  Call CommonHelper.toComboBoxDefaultSkin(cmItemType)
  Call CommonHelper.toDefaultSkin(txtName)
  Call CommonHelper.toComboBoxDefaultSkin(cmLocation)
  Call CommonHelper.toComboBoxDefaultSkin(cmCategory)
End Sub

Private Function isFormDetailValid() As Boolean
  If (Not CommonHelper.hasValidValue(txtItemCode)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtItemCode, "Please enter the an Item Code")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmItemType.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmItemType, "Please select an Item Type Name")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(txtName)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtName, "Please enter the a Name")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmLocation.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmLocation, "Please select a Locatio ")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmCategory.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmCategory, "Please select a Category")
    Exit Function
  End If
  
  isFormDetailValid = True
End Function
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "Add"
    cmbEdit.Enabled = False
    cmbDelete.Enabled = False
    cmStatus.Text = "Available"
  Else
    cmbNewRec.Caption = "New"
    cmbEdit.Enabled = True
    cmbDelete.Enabled = True
  End If
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
  Dim FileName As String
  FileName = LookupDao.getLocationImgName(getLocationID)
  If (FileName <> vbNullString) Then
    imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & FileName)
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

    If (rs.EOF) Then
      Exit Sub
    End If
  
    lblID.Caption = rs!id
    txtName.Text = CommonHelper.extractStringValue(rs!name)
    txtItemCode.Text = CommonHelper.extractStringValue(rs!ITEM_CODE)
    txtDescription.Text = CommonHelper.extractStringValue(rs!Description)
    txtDonatedBy.Text = CommonHelper.extractStringValue(rs!DONATED_BY)
    txtAuthor.Text = CommonHelper.extractStringValue(rs!author)
    lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
    lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
    lblLatModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
    lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
    
    
    If (cmbNewRec.Caption = "New") Then
      cmStatus.Text = CommonHelper.extractStringValue(rs!Status)
    Else
      cmStatus.Text = "Available"
    End If
    
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

Private Sub Form_Unload(Cancel As Integer)
   Call frmMain.reloadBookStats
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

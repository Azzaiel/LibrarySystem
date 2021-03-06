VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInventory 
   Caption         =   "Inventory"
   ClientHeight    =   10605
   ClientLeft      =   -180
   ClientTop       =   450
   ClientWidth     =   20025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   20025
   Begin VB.Frame Frame2 
      Caption         =   "Search Form"
      Height          =   1695
      Left            =   8520
      TabIndex        =   43
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cmbSearchStatus 
         Height          =   315
         ItemData        =   "frmInventory.frx":0000
         Left            =   7440
         List            =   "frmInventory.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
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
         Left            =   7440
         TabIndex        =   21
         Top             =   1200
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
         Left            =   5040
         TabIndex        =   20
         Top             =   1200
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
         Left            =   2760
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtSearchAuthor 
         Height          =   285
         Left            =   7440
         TabIndex        =   17
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox cmSearchCategory 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchName 
         Height          =   285
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cmSearchType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSearchItemCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H0080FF80&
         Caption         =   "Status"
         Height          =   255
         Left            =   6840
         TabIndex        =   49
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FF80&
         Caption         =   "Author"
         Height          =   255
         Left            =   6840
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   "Category"
         Height          =   255
         Left            =   3360
         TabIndex        =   47
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080FF80&
         Caption         =   " Title"
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080FF80&
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "ISBN"
         Height          =   255
         Left            =   120
         TabIndex        =   44
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
      Left            =   1080
      TabIndex        =   8
      Top             =   9960
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
      Left            =   2280
      TabIndex        =   9
      Top             =   9960
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
      Left            =   3480
      TabIndex        =   10
      Top             =   9960
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
      Left            =   5880
      TabIndex        =   12
      Top             =   9960
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
      Left            =   4680
      TabIndex        =   11
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item"
      Height          =   9855
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtVolume 
         Height          =   285
         Left            =   5280
         MaxLength       =   9
         TabIndex        =   59
         Top             =   9360
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Acquisition"
         Height          =   1215
         Left            =   3840
         TabIndex        =   54
         Top             =   8040
         Width           =   4215
         Begin VB.ComboBox cmAquiType 
            Height          =   315
            ItemData        =   "frmInventory.frx":0056
            Left            =   1440
            List            =   "frmInventory.frx":0060
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtAqui 
            Height          =   285
            Left            =   1440
            MaxLength       =   9
            TabIndex        =   55
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label22 
            BackColor       =   &H0080FF80&
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblAqui 
            BackColor       =   &H0080FF80&
            Caption         =   "Purchase Cost"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCprYr 
         Height          =   285
         Left            =   1560
         TabIndex        =   52
         Top             =   7680
         Width           =   3015
      End
      Begin VB.TextBox txtPublisher 
         Height          =   285
         Left            =   1560
         TabIndex        =   50
         Top             =   7320
         Width           =   3015
      End
      Begin VB.ComboBox cmStatus 
         Height          =   315
         ItemData        =   "frmInventory.frx":0078
         Left            =   1560
         List            =   "frmInventory.frx":008B
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "cmStatus"
         Top             =   8040
         Width           =   1935
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   6960
         Width           =   3015
      End
      Begin VB.TextBox txtDescription 
         Height          =   645
         Left            =   1560
         TabIndex        =   5
         Top             =   6240
         Width           =   6255
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
         Left            =   1080
         Picture         =   "frmInventory.frx":00BD
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
      Begin VB.Label lblVolumne 
         BackColor       =   &H0080FF80&
         Caption         =   "Volume"
         Height          =   255
         Left            =   4080
         TabIndex        =   60
         Top             =   9360
         Width           =   735
      End
      Begin VB.Label Label21 
         BackColor       =   &H0080FF80&
         Caption         =   "Copyright Yr"
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080FF80&
         Caption         =   "Publisher"
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   7320
         Width           =   855
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   9480
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   9120
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   9480
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod By"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "Created Date"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   8760
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "Created By"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   8400
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "Status"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   8040
         Width           =   495
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
         Caption         =   "*Title"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Width           =   375
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
         Caption         =   "*ISBN"
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
         Caption         =   "Accession no."
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dgItems 
      Height          =   8655
      Left            =   8520
      TabIndex        =   22
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   15266
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

Private Sub cmAquiType_Click()
  txtAqui = ""
  If (cmAquiType.Text = "Purchased") Then
    lblAqui = "Purchase Cost"
  Else
    lblAqui = "Donated By"
  End If
End Sub

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
    txtAqui.Text = ""
    txtAuthor.Text = ""
    lblCreatedBy.Caption = ""
    lblCreatedDate.Caption = ""
    lblLatModBy.Caption = ""
    lblLastModDate.Caption = ""
    'txtPurchaseCost = ""
    cmStatus.ListIndex = -1
    
    cmItemType.ListIndex = -1
    cmLocation.ListIndex = -1
    cmCategory.ListIndex = -1
    cmAquiType.ListIndex = 0
    cmStatus.ListIndex = 0

End Sub

Private Sub cmbEdit_Click()
    Call restoreFormDefaultSkin
    If (isFormDetailValid) Then
      'If (isItemCodeAlreadyExist(rs!ITEM_CODE)) Then
       ' MsgBox "Item Code Already in use", vbCritical
        'Exit Sub
      'End If
      Set tempRs = InventoryDao.getRsByID(rs!id)
      tempRs!name = txtName.Text
      tempRs!ITEM_CODE = txtItemCode.Text
      tempRs!Description = txtDescription.Text
      tempRs!AQUISITION_TYPE = cmAquiType.Text
      
      If (cmAquiType.Text = "Purchased") Then
        tempRs!PURCHASE_COST = Val(txtAqui)
        tempRs!DONATED_BY = Null
      Else
        tempRs!DONATED_BY = txtAqui.Text
        tempRs!PURCHASE_COST = Null
      End If
      
      tempRs!PUBLISHER = txtPublisher
      tempRs!COPYRIGHT_YEAR = txtCprYr
      
      If (cmItemType.Text = "CD") Then
        tempRs!Volume = txtVolume
      Else
        tempRs!Volume = Null
      End If
      tempRs!author = txtAuthor.Text

      tempRs!status = cmStatus.Text
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
  
  oSheet.Range("A12").CopyFromRecordset dgItems.DataSource
  oSheet.Columns.AutoFit
  oSheet.Range("P1:AA1").EntireColumn.Hidden = True
  
  Dim availableCount As Long
  Dim borrowedCount As Long
  Dim damageCount As Long
  Dim lossCount As Long
  Dim obsoleteCount As Long
    
  availableCount = 0
  borrowedCount = 0
  damageCount = 0
  lossCount = 0
  obsoleteCount = 0
  
  rs.MoveFirst
  Dim itemStatus As String
  While Not rs.EOF
    itemStatus = CommonHelper.extractStringValue(rs!status)
    If (itemStatus = "Available") Then
      availableCount = availableCount + 1
    ElseIf (itemStatus = "Borrowed") Then
      borrowedCount = borrowedCount + 1
    ElseIf (itemStatus = "Damaged") Then
      damageCount = damageCount + 1
    ElseIf (itemStatus = "Loss") Then
      lossCount = lossCount + 1
    ElseIf (itemStatus = "Obsolete") Then
      obsoleteCount = obsoleteCount + 1
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
  
  oSheet.Range("A7").value = "Obsolete"
  oSheet.Range("C7").value = obsoleteCount
  
  oSheet.Range("C9").value = availableCount + borrowedCount + damageCount + lossCount + obsoleteCount
  
  oSheet.Range("L11:Z" & rs.RecordCount + 11).NumberFormat = Constants.DEFAULT_CURRENCY_FORMAT
  
  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\" & Constants.TEMP_WORK_BOOK
  
  'If (UserSession.role = "Admin") Then
  '  excelApp.Visible = True
  'Else
    Dim pdfFilePat As String
    pdfFilePat = CommonHelper.getTempPath & "\temp_" & Format(Now, "mmhhyysssh") & ".pdf"
    Call oBook.ExportAsFixedFormat(xlTypePDF, pdfFilePat, xlQualityStandard, False, True)
    oBook.Close
    Call CommonHelper.openFile(pdfFilePat, Me.hWnd)
  'End If

End Sub

Private Sub cmbNewRec_Click()
    Call restoreFormDefaultSkin
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
    lblID = InventoryDao.getItemNewID
    txtItemCode.SetFocus
  Else
    If (isFormDetailValid) Then
      'If (isItemCodeAlreadyExist) Then
        'MsgBox "Item Code Already in use", vbCritical
        'Exit Sub
      'End If
      Set tempRs = InventoryDao.getFakeRs
      tempRs.AddNew
      tempRs!name = txtName.Text
      tempRs!ITEM_CODE = txtItemCode.Text
      tempRs!Description = txtDescription.Text
      
      tempRs!AQUISITION_TYPE = cmAquiType.Text
      
      If (cmAquiType.Text = "Purchased") Then
        tempRs!PURCHASE_COST = Val(txtAqui)
        tempRs!DONATED_BY = Null
      Else
        tempRs!DONATED_BY = txtAqui.Text
        tempRs!PURCHASE_COST = Null
      End If
      
      tempRs!PUBLISHER = txtPublisher
      tempRs!COPYRIGHT_YEAR = txtCprYr
      
      If (cmItemType.Text = "CD") Then
        tempRs!Volume = txtVolume
      Else
        tempRs!Volume = Null
      End If
      
      tempRs!author = txtAuthor.Text
      tempRs!status = cmStatus.Text
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
    Call CommonHelper.sendWarning(txtItemCode, "Please enter the an ISBN")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmItemType.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmItemType, "Please select an Item Type Name")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(txtName)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtName, "Please enter the a Title")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmLocation.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmLocation, "Please select a Location ")
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

Private Sub cmbSearchStatus_Click()
  Call cmdSearch_Click
End Sub

Private Sub cmbSearchStatus_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmdSearch_Click
  End If
End Sub

Private Sub cmdClearSearch_Click()
  txtSearchItemCode.Text = ""
  cmSearchType.ListIndex = -1
  txtSearchName.Text = ""
  cmSearchCategory.ListIndex = -1
  txtSearchAuthor = ""
  cmbSearchStatus.ListIndex = -1
End Sub

Private Sub cmdSearch_Click()
  Set dgItems.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = InventoryDao.searchItem(txtSearchItemCode, getSearchItemTypeID, txtSearchAuthor, txtSearchName, getSearchCategoryID, cmbSearchStatus.Text)
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

Private Sub cmItemType_Click()
   lblVolumne.Visible = False
   txtVolume.Visible = False
   
   If (cmItemType.Text = "CD") Then
     lblVolumne.Visible = True
     txtVolume.Visible = True
   End If
End Sub

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
  cmAquiType.ListIndex = 0
  Call populateDropDown
  Call populateDataGrid
End Sub
Private Sub populateDataGrid()
  Set rs = InventoryDao.searchItem
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
    .Columns(0).Width = 1100
    .Columns(0).Alignment = dbgCenter
    .Columns(0).Caption = "Accession no."
    
    .Columns(2).Caption = "ISBN"
    
    .Columns(3).Caption = "TITLE"
    
     'CREATED DATE - 9
    .Columns(10).Width = 1500
    .Columns(10).NumberFormat = DEFAULT_CURRENCY_FORMAT
    
    
     'CREATED DATE - 12
    .Columns(16).Width = 1500
    .Columns(16).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(16).Alignment = dbgCenter
    
    'LAST MOD DATE - 14
    .Columns(18).Width = 1500
    .Columns(18).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(18).Alignment = dbgCenter
    
    .Columns(19).Visible = False
    .Columns(20).Visible = False
    .Columns(21).Visible = False
    
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
    'txtDonatedBy.Text = CommonHelper.extractStringValue(rs!DONATED_BY)
    txtAuthor.Text = CommonHelper.extractStringValue(rs!author)
    lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
    lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
    lblLatModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
    lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
    'txtPurchaseCost = CommonHelper.extractStringValue(rs!PURCHASE_COST)
    
    If (cmbNewRec.Caption = "New") Then
      cmStatus.Text = CommonHelper.extractStringValue(rs!status)
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
   
   lblVolumne.Visible = False
   txtVolume.Visible = False
   
   If (cmItemType.Text = "CD") Then
     lblVolumne.Visible = True
     txtVolume.Visible = True
   End If
   
   If (CommonHelper.extractStringValue(rs!AQUISITION_TYPE) <> vbNullString) Then
     cmAquiType.Text = CommonHelper.extractStringValue(rs!AQUISITION_TYPE)
   Else
     cmAquiType.Text = "Purchased"
   End If
      
   If (cmAquiType.Text = "Purchased") Then
     txtAqui = Val(CommonHelper.extractStringValue(rs!PURCHASE_COST))
   Else
     txtAqui = CommonHelper.extractStringValue(rs!DONATED_BY)
   End If
   
   
   
   txtPublisher = CommonHelper.extractStringValue(rs!PUBLISHER)
   
   txtCprYr = CommonHelper.extractStringValue(rs!COPYRIGHT_YEAR)
   
   txtVolume = CommonHelper.extractStringValue(rs!Volume)
      
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

Private Sub txtAqui_KeyPress(KeyAscii As Integer)
  If (cmAquiType.Text = "Purchased") Then
    If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii) Or Len(txtAqui) > 11)) Then
      KeyAscii = 0
      Beep
    End If
  End If
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
   If (Not CommonHelper.isFunctionAscii(KeyAscii) And (Not CommonHelper.isNumberAscii(KeyAscii)) And (Not CommonHelper.isLetterAscii(KeyAscii))) Then
    KeyAscii = 0
    Beep
  End If
End Sub

Private Sub txtPurchaseCost_KeyPress(KeyAscii As Integer)
   
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

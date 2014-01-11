VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Library System"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   20040
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   20040
   Begin VB.Frame frmControl 
      Height          =   11895
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   20175
      Begin VB.Frame Frame1 
         Caption         =   "Quick Search"
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.ComboBox cmSearchType 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtSearchItemCode 
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSearchName 
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox cmSearchCategory 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtSearchAuthor 
            Height          =   285
            Left            =   4200
            TabIndex        =   9
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmItemsQuickSearch 
            Caption         =   "Search"
            Height          =   435
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CommandButton cmdClearSearch 
            Caption         =   "Clear"
            Height          =   435
            Left            =   3360
            TabIndex        =   7
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label Label15 
            BackColor       =   &H0080FF80&
            Caption         =   " Name"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label14 
            BackColor       =   &H0080FF80&
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label8 
            BackColor       =   &H0080FF80&
            Caption         =   "Item Code"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label16 
            BackColor       =   &H0080FF80&
            Caption         =   "Category"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            BackColor       =   &H0080FF80&
            Caption         =   "Author"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Detail"
         Height          =   5535
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   6975
         Begin VB.Frame fmStudentInfo 
            Caption         =   "Student info"
            Height          =   1335
            Left            =   120
            TabIndex        =   38
            Top             =   3720
            Width           =   6735
            Begin VB.Label Label13 
               BackColor       =   &H0080FF80&
               Caption         =   "LRN"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label txtLRN 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   48
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label txtAdviser 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   45
               Top             =   600
               Width           =   2175
            End
            Begin VB.Label txtSection 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   44
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label txtStudentName 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   43
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblSelectStudent 
               Caption         =   "Select Student"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2520
               TabIndex        =   42
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label11 
               BackColor       =   &H0080FF80&
               Caption         =   "Adviser"
               Height          =   255
               Left            =   3720
               TabIndex        =   41
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label10 
               BackColor       =   &H0080FF80&
               Caption         =   "Section"
               Height          =   255
               Left            =   3720
               TabIndex        =   40
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label1 
               BackColor       =   &H0080FF80&
               Caption         =   "Student Name"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.ComboBox cmCategory 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   32
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtDescription 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2040
            Width           =   5055
         End
         Begin VB.TextBox txtAuthor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2640
            Width           =   5055
         End
         Begin VB.TextBox txtDonatedBy 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3000
            Width           =   5055
         End
         Begin VB.ComboBox cmStatus 
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
            ItemData        =   "Form1.frx":0000
            Left            =   1440
            List            =   "Form1.frx":0010
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   28
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtItemCode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cmItemType 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   960
            Width           =   1935
         End
         Begin VB.ComboBox cmLocation 
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   20
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label lblChekOut 
            Caption         =   "Check out Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            Caption         =   "Category"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FF80&
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label6 
            BackColor       =   &H0080FF80&
            Caption         =   "Author"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label7 
            BackColor       =   &H0080FF80&
            Caption         =   "Donated By"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label9 
            BackColor       =   &H0080FF80&
            Caption         =   "Status"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label lblName 
            BackColor       =   &H0080FF80&
            Caption         =   "Item Code"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FF80&
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080FF80&
            Caption         =   "Item Name"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H0080FF80&
            Caption         =   "Location"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Location Map"
         Height          =   4095
         Left            =   6960
         TabIndex        =   4
         Top             =   5640
         Width           =   6495
         Begin VB.PictureBox imgLoc 
            Height          =   3690
            Left            =   120
            Picture         =   "Form1.frx":0038
            ScaleHeight     =   3630
            ScaleWidth      =   6195
            TabIndex        =   6
            Top             =   240
            Width           =   6255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dashboard"
         ClipControls    =   0   'False
         Height          =   9615
         Left            =   13800
         TabIndex        =   3
         Top             =   120
         Width           =   6135
         Begin MSDataGridLib.DataGrid dgTransactionDash 
            Height          =   9255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   16325
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
      Begin VB.Frame Frame2 
         Caption         =   "Result"
         Height          =   7575
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   6495
         Begin MSDataGridLib.DataGrid dgItems 
            Height          =   7215
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   12726
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
   End
   Begin VB.Menu mnLookups 
      Caption         =   "Lookups"
      Begin VB.Menu Categuries 
         Caption         =   "Categories"
      End
      Begin VB.Menu mnItemType 
         Caption         =   "Itemn Type"
      End
      Begin VB.Menu sections 
         Caption         =   "Sections"
      End
      Begin VB.Menu mnLocationMapping 
         Caption         =   "Location Mapping"
      End
   End
   Begin VB.Menu mnStudents 
      Caption         =   "Students"
   End
   Begin VB.Menu mnInvetory 
      Caption         =   "Inventory"
   End
   Begin VB.Menu Account 
      Caption         =   "Account"
   End
   Begin VB.Menu mnLogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private itemsRs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private transactionRS As ADODB.Recordset

Private itemTypeItemList() As Variant
Private locationItemList() As Variant
Private categoriesItemList() As Variant

Public selectedStudentID As Integer
Public selectedReturnDate As Date
Private Sub Categuries_Click()
  frmCategories.Show vbModal
End Sub

Private Sub cmdClearSearch_Click()
  txtSearchItemCode.Text = ""
  cmSearchType.ListIndex = -1
  txtSearchName.Text = ""
  cmSearchCategory.ListIndex = -1
  txtSearchAuthor = ""
End Sub

Private Sub cmItemsQuickSearch_Click()
  Set dgItems.DataSource = Nothing
  Call DbInstance.closeRecordSet(itemsRs)
  Set itemsRs = InventoryDao.search(txtSearchItemCode, getSearchItemTypeID, txtSearchAuthor, txtSearchName, getSearchCategoryID)
  Set dgItems.DataSource = itemsRs
  If (itemsRs.RecordCount = 0) Then
    MsgBox "No record found", vbInformation
  End If
  dgItems.Refresh
  'Call clearForm
  Call formatIemsDataGrid
End Sub

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
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub cmSearchType_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub
Private Sub clearDetailForm()

   txtName.Text = ""
   txtItemCode.Text = ""
   txtDescription.Text = ""
   txtDonatedBy.Text = ""
   txtAuthor.Text = ""
   cmStatus.Text = ""
    
   cmItemType.ListIndex = -1
   cmLocation.ListIndex = -1
   cmCategory.ListIndex = -1
   
End Sub
Private Sub dgItems_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedItem
End Sub

Private Sub dgItems_SelChange(Cancel As Integer)
  Call showSelectedItem
End Sub
Private Sub showSelectedItem()

    Call clearStudentInfo
    Call clearDetailForm

    If (itemsRs.RecordCount = 0) Then
      Exit Sub
    End If
 
    txtName.Text = CommonHelper.extractStringValue(itemsRs!name)
    txtItemCode.Text = CommonHelper.extractStringValue(itemsRs!ITEM_CODE)
    txtDescription.Text = CommonHelper.extractStringValue(itemsRs!Description)
    txtDonatedBy.Text = CommonHelper.extractStringValue(itemsRs!DONATED_BY)
    txtAuthor.Text = CommonHelper.extractStringValue(itemsRs!author)
    cmStatus.Text = CommonHelper.extractStringValue(itemsRs!Status)
    
    cmItemType.ListIndex = -1
    cmLocation.ListIndex = -1
    cmCategory.ListIndex = -1
    
    Dim index As Integer

   For index = 0 To UBound(itemTypeItemList)
     If (itemsRs!ITEM_TYPE_ID = Val(itemTypeItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmItemType.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(locationItemList)
     If (itemsRs!LOCATION_ID = Val(locationItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmLocation.ListIndex = index
     End If
   Next index
   
   For index = 0 To UBound(categoriesItemList)
     If (itemsRs!CATEGORY_ID = Val(categoriesItemList(index, Constants.ITEM_VALUE_INDEX))) Then
       cmCategory.ListIndex = index
     End If
   Next index
   
   If (cmStatus = "Available") Then
     Call toogelItemCheckOutUI(True)
   Else
     Call toogelItemCheckOutUI(False)
   End If
   
   If (cmStatus = "Borrowed") Then
      Set tempRs = InventoryDao.getStudentBorrower(itemsRs!ID)
      txtLRN = tempRs!lrn
      txtStudentName = tempRs!Student_Name
      txtAdviser = tempRs!Adviser
      txtSection = tempRs!Section
      Call DbInstance.closeRecordSet(tempRs)
   End If
   
End Sub

Private Sub toogelItemCheckOutUI(isAvailable As Boolean)
  fmStudentInfo.Enabled = isAvailable
  lblChekOut.Enabled = isAvailable
  lblSelectStudent.Enabled = isAvailable
  If (isAvailable) Then
  
    txtStudentName.BackColor = vbWhite
    txtAdviser.BackColor = vbWhite
    txtSection.BackColor = vbWhite
    txtLRN.BackColor = vbWhite
    
    txtStudentName.ForeColor = vbBlack
    txtAdviser.ForeColor = vbBlack
    txtSection.ForeColor = vbBlack
    txtLRN.ForeColor = vbBlack
    
  Else
  
    txtStudentName.BackColor = vbGrayText
    txtAdviser.BackColor = vbGrayText
    txtSection.BackColor = vbGrayText
    txtLRN.BackColor = vbGrayText
    
    txtStudentName.ForeColor = vbWhite
    txtAdviser.ForeColor = vbWhite
    txtSection.ForeColor = vbWhite
    txtLRN.ForeColor = vbWhite
    
  End If
End Sub
Private Sub dgTransactionDash_DblClick()
   If (transactionRS.RecordCount > 0) Then
     frmItemReturn.transactionID = transactionRS!Transaction_ID
     frmItemReturn.Show vbModal
   End If
End Sub

Private Sub fmStudentInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lblSelectStudent.ForeColor = vbBlue
End Sub

Private Sub Form_Load()
  Call populateDropDown
  Call initiateItemsRs
  Call populateTransactionDatagrid
End Sub
Private Sub populateTransactionDatagrid()
  Set transactionRS = InventoryDao.getTransactionDashboardRs
  Set dgTransactionDash.DataSource = transactionRS
  dgTransactionDash.Refresh
  Call formatTransactionDashDatagrid
End Sub
Private Sub formatTransactionDashDatagrid()
    With dgTransactionDash
     'LRN - 0
    .Columns(0).Width = 1500
    .Columns(0).Alignment = dbgCenter

     'DUE DATE - 5
    .Columns(5).Width = 1500
    .Columns(5).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(5).Alignment = dbgCenter
    
    'TRANSACTION_ID
    .Columns(6).Visible = False
    
  End With
End Sub
Private Sub populateDropDown()
  Dim index As Integer

  itemTypeItemList = LookupDao.getItemTypeItemList
  cmSearchType.Clear
  cmItemType.Clear
  For index = 0 To UBound(itemTypeItemList)
    cmSearchType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
    cmItemType.AddItem (itemTypeItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  locationItemList = LookupDao.getLocationMappingItemList
  cmLocation.Clear
  For index = 0 To UBound(locationItemList)
     cmLocation.AddItem (locationItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
  categoriesItemList = LookupDao.getCategoriesItemList
  cmSearchCategory.Clear
  cmCategory.Clear
  For index = 0 To UBound(categoriesItemList)
    cmSearchCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
    cmCategory.AddItem (categoriesItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
  
End Sub
Private Sub initiateItemsRs()
  Set itemsRs = InventoryDao.getEmptyRs
  Set dgItems.DataSource = itemsRs
  dgItems.Refresh
  Call formatIemsDataGrid
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
Private Sub formatIemsDataGrid()
  If (itemsRs.RecordCount > 0) Then
    itemsRs.MoveFirst
    'Call showSelectedItem
  Else
    'Call clearForm
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
Private Function getLocationID() As Integer
  Dim index As Integer
  index = cmLocation.ListIndex
  If (index <> -1) Then
    getLocationID = locationItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    getLocationID = 0
  End If
End Function



Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
End Sub
Private Sub clearStudentInfo()
    txtStudentName = ""
    txtSection = ""
    txtAdviser = ""
    selectedStudentID = 0
    txtLRN = ""
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblChekOut.ForeColor = vbBlue
End Sub

Private Sub lblChekOut_Click()
  If (selectedStudentID > 0) Then
    selectedReturnDate = vbNull
    frmReturnDate.Show vbModal
    If (selectedReturnDate <> vbNull) Then
      Set tempRs = InventoryDao.getFakeTransactionRS
      tempRs.AddNew
      tempRs!ITEM_ID = itemsRs!ID
      tempRs!STUDENT_ID = selectedStudentID
      tempRs!LEND_DATE = Now
      tempRs!LEND_BY = UserSession.getLoginUser
      tempRs!REQUESTED_RETURN_DATE = selectedReturnDate
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      Set tempRs = InventoryDao.getRsByID(itemsRs!ID)
      tempRs!Status = "Borrowed"
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      MsgBox "Transaction Successful"
      Call cmItemsQuickSearch_Click
      Call clearDetailForm
    Else
      MsgBox "System cannot procced without retrun date", vbCritical
    End If
  Else
    MsgBox "Please select a Student", vbCritical
  End If
End Sub

Private Sub lblChekOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblChekOut.ForeColor = vbRed
End Sub

Private Sub lblSelectStudent_Click()
 Call clearStudentInfo
 frmStudentSelect.Show vbModal
 lblSelectStudent.ForeColor = vbBlue
End Sub

Private Sub lblSelectStudent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblSelectStudent.ForeColor = vbRed
End Sub

Private Sub mnInvetory_Click()
  frmInventory.Show vbModal
End Sub

Private Sub mnItemType_Click()
  frmItemTypes.Show vbModal
End Sub

Private Sub mnLocationMapping_Click()
  frmLocationMapping.Show vbModal
End Sub

Private Sub mnStudents_Click()
  frmStudents.Show vbModal
End Sub

Private Sub sections_Click()
  frmSections.Show vbModal
End Sub

Private Sub txtSearchAuthor_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub txtSearchItemCode_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    Call cmItemsQuickSearch_Click
  End If
End Sub

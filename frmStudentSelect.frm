VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStudentSelect 
   Caption         =   "Student Selection"
   ClientHeight    =   6180
   ClientLeft      =   6705
   ClientTop       =   1395
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8610
   Begin VB.Frame Frame2 
      Caption         =   "Search Panel"
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtSearchLastName 
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmSearchSection 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtSearchLrn 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmbSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmbClearSearch 
         Caption         =   "Clear"
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "LAST_NAME"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "SECTION"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "LRN"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dgStudents 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Doule Click to select the record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1680
      Width           =   3375
   End
End
Attribute VB_Name = "frmStudentSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private sectionItemList() As Variant

Private Sub cmbSearch_Click()
  Set dgStudents.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = StudentDao.qucikSearchRs(Val(txtSearchLrn.Text), getSearchSectionID, txtSearchLastName.Text)
  If (rs.RecordCount = 0) Then
    MsgBox "No Record Found", vbInformation
  End If
  Set dgStudents.DataSource = rs
  dgStudents.Refresh
  Call formatDataGrid
End Sub
Private Sub dgStudents_DblClick()
  If (rs.RecordCount > 0) Then
    frmMain.txtStudentName.Caption = rs!FULL_NAME
    frmMain.txtSection = rs!Section
    frmMain.txtAdviser = rs!Adviser
    frmMain.selectedStudentID = rs!ID
    frmMain.txtLrn = rs!lrn
  End If
  Unload Me
End Sub

Private Sub Form_Load()
  Call populateDropDown
  'fake rs search to initiate the data grid
  Call DbInstance.closeRecordSet(rs)
  Set rs = StudentDao.qucikSearchRs(-1, -1, "I do no Exist")
  Set dgStudents.DataSource = rs
  Call formatDataGrid
End Sub
Private Sub populateDropDown()
  sectionItemList = LookupDao.getSectionsItemList
  Dim index As Integer
  cmSearchSection.Clear
  For index = 0 To UBound(sectionItemList)
    cmSearchSection.AddItem (sectionItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
End Sub

Private Function getSearchSectionID() As Integer

  Dim index As Integer
  index = cmSearchSection.ListIndex
  getSearchSectionID = identifySectionID(index)

End Function
Private Function identifySectionID(index As Integer) As Integer
   If (index <> -1) Then
    identifySectionID = sectionItemList(index, Constants.ITEM_VALUE_INDEX)
  Else
    identifySectionID = 0
  End If
End Function
Private Sub formatDataGrid()
  With dgStudents
    .Columns(0).Width = 950
    .Columns(1).Width = 2500
    .Columns(2).Width = 2500
    .Columns(3).Width = 2000
    .Columns(4).Visible = False
  End With
End Sub


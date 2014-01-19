VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCategories 
   Caption         =   "Categories"
   ClientHeight    =   3975
   ClientLeft      =   735
   ClientTop       =   1485
   ClientWidth     =   18930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   18930
   Begin VB.Frame Frame1 
      Caption         =   "Created By"
      Height          =   3015
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "* Name"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Description"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Created By"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "Created Date"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod By"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   2400
         Width           =   2535
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
      Left            =   240
      TabIndex        =   2
      Top             =   3240
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
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
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgCategories 
      Height          =   3615
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   6376
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
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset


Private Sub showSelectedItemInForm()

  Call restoreFormDefaultSkin
  lblID.Caption = rs!id
  txtName.Text = rs!name
  txtDescription.Text = CommonHelper.extractStringValue(rs!Description)
  lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
  lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
  lblLatModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
  lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
  
End Sub

Private Sub cmbClear_Click()
 Call restoreFormDefaultSkin
 Call toogelInsertMode(False)
 Call clearForm
End Sub

Private Sub clearForm()
  lblID.Caption = ""
  txtName.Text = ""
  txtDescription.Text = ""
  lblCreatedBy.Caption = ""
  lblCreatedDate.Caption = ""
  lblLatModBy.Caption = ""
  lblLastModDate.Caption = ""
End Sub

Private Sub cmbDelete_Click()
  Dim response As String
  response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
  If (response = vbOK) Then
    If (LookupDao.isCategoryBeingUsed(rs!id)) Then
      MsgBox "Record cannot be deleted. It is being used by another record", vbCritical
      Exit Sub
    End If
    rs.Delete
    MsgBox "Record Deleted", vbInformation
    Call populateDataGrid
  End If
End Sub
Private Sub cmbEdit_Click()
  Call restoreFormDefaultSkin
  If (isFormDetailValid) Then
    If (LookupDao.isCategoryExist(txtName, rs!id)) Then
        MsgBox "Category Already Exist", vbCritical
        Exit Sub
    End If
    rs!name = txtName.Text
    rs!Description = txtDescription.Text
    rs!LAST_MOD_BY = UserSession.getLoginUser
    rs!LAST_MOD_DATE = Now
    rs.Update
    MsgBox "Record updated", vbInformation
    Call populateDataGrid
  End If
End Sub
Private Sub cmbNewRec_Click()
  Call restoreFormDefaultSkin
  If cmbNewRec.Caption = "New" Then
    Call toogelInsertMode(True)
    txtName.SetFocus
  Else
    If (LookupDao.isCategoryExist(txtName)) Then
        MsgBox "Category Already Exist", vbCritical
        Exit Sub
    End If
    If (isFormDetailValid) Then
       rs.AddNew
       rs!name = txtName.Text
       rs!Description = txtDescription.Text
       rs!CREATED_BY = UserSession.getLoginUser
       rs!CREATED_DATE = Now
       rs!LAST_MOD_BY = UserSession.getLoginUser
       rs!LAST_MOD_DATE = Now
       rs.Update
       MsgBox "Record Created!", vbInformation
       Call toogelInsertMode(False)
       Call populateDataGrid
    End If
  End If
End Sub
Private Sub restoreFormDefaultSkin()
  Call CommonHelper.toDefaultSkin(txtName)
End Sub
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "Add"
    cmbEdit.Enabled = False
    cmbDelete.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmbEdit.Enabled = True
    cmbDelete.Enabled = True
  End If
End Sub

Private Function isFormDetailValid() As Boolean
  If (Not CommonHelper.hasValidValue(txtName)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtName, "Please enter the Category Name")
  Else
    isFormDetailValid = True
  End If
End Function
Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub dgCategories_SelChange(Cancel As Integer)
   Call showSelectedItemInForm
End Sub

Private Sub Form_Load()
 Call populateDataGrid
 
End Sub
Public Sub populateDataGrid()
  Call restoreFormDefaultSkin
  Set rs = LookupDao.getCategoriesRs
  Set dgCategories.DataSource = rs
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedItemInForm
  Else
    Call clearForm
  End If
  Call formatDataGrid
  dgCategories.Refresh
End Sub
Private Sub formatDataGrid()
  With dgCategories
    .Columns(0).Width = 400
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 2500
    .Columns(2).Width = 3500
    .Columns(3).Width = 1300
    .Columns(4).Width = 1500
    .Columns(4).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(4).Alignment = dbgCenter
    .Columns(3).Alignment = dbgCenter
    .Columns(5).Width = 1300
    .Columns(5).Alignment = dbgCenter
    .Columns(6).Width = 1500
    .Columns(6).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(6).Alignment = dbgCenter
  End With
End Sub


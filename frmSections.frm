VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSections 
   Caption         =   "Sections"
   ClientHeight    =   5340
   ClientLeft      =   330
   ClientTop       =   2295
   ClientWidth     =   19605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   19605
   Begin VB.Frame Frame1 
      Caption         =   "Created By"
      Height          =   3615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtAdviser 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "* Name"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "*Level"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "*Adviser"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "Created by"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Created date"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod by"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblLastModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
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
      Left            =   120
      TabIndex        =   3
      Top             =   3960
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
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
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
      TabIndex        =   5
      Top             =   3960
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
      Left            =   1920
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   3960
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dbSections 
      Height          =   5055
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   8916
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
Attribute VB_Name = "frmSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset


Private Sub showSelectedItemInForm()

  lblID.Caption = rs!id
  txtName.Text = rs!name
  txtLevel.Text = CommonHelper.extractStringValue(rs!level)
  txtAdviser.Text = CommonHelper.extractStringValue(rs!Adviser)
  lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
  lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
  lblLastModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
  lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
  
End Sub

Private Sub cmbClear_Click()
 
 Call clearForm
  
End Sub

Private Sub clearForm()
  lblID.Caption = ""
  txtName.Text = ""
  txtLevel.Text = ""
txtAdviser.Text = ""
  lblCreatedBy.Caption = ""
  lblCreatedDate.Caption = ""
  lblLastModBy.Caption = ""
  lblLastModDate.Caption = ""
End Sub

Private Sub cmbDelete_Click()
 Dim response As String
  response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
  If (response = vbOK) Then
    If (LookupDao.isSectionBeingUsed(rs!id)) Then
      MsgBox "Record cannot be deleted. It is being used by another record", vbCritical
      Exit Sub
    End If
    rs.Delete
    MsgBox "Record Deleted", vbInformation
  End If
End Sub

Private Sub cmbEdit_Click()
   Call restoreFormDefaultSkin
  If (isFormDetailValid) Then
    If (hasValidForm) Then
      If (LookupDao.isSectionExist(txtName, txtLevel, rs!id)) Then
        MsgBox "Section Already Exist", vbCritical
        Exit Sub
      End If
      rs!name = txtName.Text
      rs!level = txtLevel.Text
      rs!Adviser = txtAdviser.Text
      rs!LAST_MOD_BY = UserSession.getLoginUser
      rs!LAST_MOD_DATE = Now
      rs.Update
      MsgBox "Record updated", vbInformation
      Call populateDataGrid
    End If
  End If
 
End Sub

Private Sub cmbNewRec_Click()
   Call restoreFormDefaultSkin
   If cmbNewRec.Caption = "New" Then
     Call toogelInsertMode(True)
     txtName.SetFocus
   Else
     If (hasValidForm) Then
       If (LookupDao.isSectionExist(txtName, txtLevel)) Then
          MsgBox "Section Already Exist", vbCritical
          Exit Sub
       End If
       rs.AddNew
       rs!level = txtLevel.Text
       rs!name = txtName.Text
       rs!Adviser = txtAdviser.Text
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
  Call CommonHelper.toDefaultSkin(txtLevel)
  Call CommonHelper.toDefaultSkin(txtAdviser)
End Sub
Private Function hasValidForm() As Boolean

  If (Not CommonHelper.hasValidValue(txtName)) Then
    Call CommonHelper.sendWarning(txtName, "Name is a required Field")
    hasValidForm = False
    Exit Function
  End If
  
    If (Not CommonHelper.hasValidValue(txtLevel)) Then
    Call CommonHelper.sendWarning(txtLevel, "Level is a required Field")
    hasValidForm = False
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(txtAdviser)) Then
    Call CommonHelper.sendWarning(txtAdviser, "Adviser is a required Field")
    hasValidForm = False
    Exit Function
  End If
  
  hasValidForm = True
  
End Function

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

Private Sub cmdClear_Click()
  Call restoreFormDefaultSkin
  Call toogelInsertMode(False)
  Call clearForm
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub
Private Sub dbSections_SelChange(Cancel As Integer)
  Call showSelectedItemInForm
End Sub

Private Sub Form_Load()
 Call populateDataGrid
End Sub
Public Sub populateDataGrid()
  Set rs = LookupDao.getSections
  Set dbSections.DataSource = rs
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedItemInForm
  End If
  dbSections.Refresh
  Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  With dbSections
    .Columns(0).Visible = False
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
Public Function isFormDetailValid() As Boolean
  isFormDetailValid = True
End Function


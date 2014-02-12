VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccount 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbReset 
      Caption         =   "Reset"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2040
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
      TabIndex        =   6
      Top             =   2640
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Form"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cmRole 
         Height          =   315
         ItemData        =   "frmAccount.frx":0000
         Left            =   1560
         List            =   "frmAccount.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   0
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "*Role"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "*Username"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dgUsers 
      Height          =   2895
      Left            =   5040
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
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
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private Sub toogelInsertMode(isInisilization As Boolean)
  If (isInisilization) Then
    Call clearForm
    cmbNewRec.Caption = "Add"
    cmbDelete.Enabled = False
    cmbReset.Enabled = False
  Else
    cmbNewRec.Caption = "New"
    cmbDelete.Enabled = True
    cmbReset.Enabled = True
  End If
End Sub
Private Sub showSelectedData()
  lblID = CommonHelper.extractStringValue(rs!id)
  txtUserName = CommonHelper.extractStringValue(rs!username)
  cmRole.Text = CommonHelper.extractStringValue(rs!role)
End Sub
Private Sub clearForm()
  lblID = ""
  txtUserName = ""
  cmRole.ListIndex = -1
End Sub

Private Sub cmbClear_Click()
  Call resetFormSkin
  Call toogelInsertMode(False)
  Call clearForm
End Sub

Private Sub cmbDelete_Click()
  If (CommonHelper.hasValidValue(lblID)) Then
    Dim response As String
    response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
    If (response = vbOK) Then
      rs.Delete
      MsgBox "Record Deleted", vbInformation
      Call populateDataGrid
    End If
  End If
End Sub
Private Sub cmbNewRec_Click()
  Call resetFormSkin
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
    txtUserName.SetFocus
  Else
    If (isFormDetailValid) Then
      If (isUsernameAlreadyExist) Then
        MsgBox "Username Already exist", vbCritical
        Exit Sub
      End If
      rs.AddNew
      rs!username = txtUserName
      rs!role = cmRole.Text
      Dim bytBlock() As Byte
      Dim Hash As New MD5Hash
      bytBlock = StrConv(txtUserName, vbFromUnicode)
      rs!Password = Hash.HashBytes(bytBlock)
      rs!FORCE_CHANGE = "T"
      rs.Update
      MsgBox "Record Added! Default password was set", vbInformation
      Call toogelInsertMode(False)
      Call populateDataGrid
    End If
  End If
End Sub
Private Function isUsernameAlreadyExist() As Boolean
  Set tempRs = UserSession.getUserByUserName(txtUserName)
  
  If (tempRs.RecordCount > 0) Then
    isUsernameAlreadyExist = True
  Else
    isUsernameAlreadyExist = False
  End If
  
  Call DbInstance.closeRecordSet(tempRs)
End Function
Private Function resetFormSkin()
  Call CommonHelper.toDefaultSkin(txtUserName)
  Call CommonHelper.toComboBoxDefaultSkin(cmRole)
End Function
Private Function isFormDetailValid() As Boolean
  If (Not CommonHelper.hasValidValue(txtUserName)) Then
    isFormDetailValid = False
    Call CommonHelper.sendWarning(txtUserName, "Please enter a Username")
    Exit Function
  End If
  
  If (Not CommonHelper.hasValidValue(cmRole.Text)) Then
    isFormDetailValid = False
    Call CommonHelper.sendComboBoxWarning(cmRole, "Please select a Role")
    Exit Function
  End If
  
  isFormDetailValid = True
End Function
Private Sub cmbReset_Click()
  If (CommonHelper.hasValidValue(lblID)) Then
    Dim response As String
    response = MsgBox("Are you sure you want to reset the password?", vbYesNo, "Question")
    If (response = vbYes) Then
      Dim bytBlock() As Byte
      Dim Hash As New MD5Hash
      bytBlock = StrConv(txtUserName, vbFromUnicode)
      rs!Password = Hash.HashBytes(bytBlock)
      rs!FORCE_CHANGE = "T"
      rs.Update
      MsgBox "Default password was set", vbInformation
    End If
  End If
End Sub
Private Sub Command4_Click()
  Unload Me
End Sub

Private Sub dgUsers_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedData
End Sub

Private Sub dgUsers_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub
Private Sub formatDataGrid()
  With dgUsers
    .Columns(0).Visible = False
    .Columns(3).Visible = False
    .Columns(4).Visible = False
  End With
End Sub
Private Sub Form_Load()
  Call populateDataGrid
End Sub
Public Sub populateDataGrid()
  Set rs = UserSession.getAllUsers
  Set dgUsers.DataSource = rs
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  End If
  dgUsers.Refresh
  Call formatDataGrid
End Sub

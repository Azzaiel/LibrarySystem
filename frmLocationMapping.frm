VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLocationMapping 
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   4575
   ClientTop       =   2280
   ClientWidth     =   20670
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   20670
   Begin VB.CommandButton cmbNewRec 
      Caption         =   "Add"
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmbEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmbDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmbClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loc Form"
      Height          =   7575
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdLoadImg 
         Caption         =   "Load Image"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog cdJpegBrowser 
         Left            =   5400
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox imgLoc 
         Height          =   3735
         Left            =   720
         Picture         =   "frmLocationMapping.frx":0000
         ScaleHeight     =   3675
         ScaleWidth      =   6195
         TabIndex        =   1
         Top             =   1920
         Width           =   6255
      End
      Begin VB.Label txtFileName 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "File Name"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Created By"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "Created Date"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod By"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Label lblLatModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "* Name"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dgLocationMapping 
      Height          =   5655
      Left            =   8040
      TabIndex        =   20
      Top             =   1320
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9975
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
Attribute VB_Name = "frmLocationMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private Sub cmbClear_Click()
  Call restoreFormDefaultSkin
  Call clearForm
End Sub

Private Sub cmbDelete_Click()
  Dim response As String
  response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
  If (response = vbOK) Then
    rs.Delete
    MsgBox "Record Deleted", vbInformation
    Call showSelectedData
  End If
End Sub

Private Function isFormValid() As Boolean
  
  Dim isValid As Boolean
  isValid = True
  
  If (Not CommonHelper.hasValidValue(txtName.Text)) Then
    
    isValid = False
    Call CommonHelper.sendWarning(txtName, "Name is a required Field")
    
  End If
  
  isFormValid = isValid
  
End Function

Private Sub cmbEdit_Click()
  Call restoreFormDefaultSkin
  If (isFormValid) Then
    rs!Name = txtName.Text
    rs!FILE_NAME = txtFileName.Caption
    rs!LAST_MOD_BY = UserSession.getLoginUser
    rs!LAST_MOD_DATE = Now
    rs.Update
    If (cdJpegBrowser.fileName <> vbNullString And txtFileName.Caption <> vbNullString) Then
      Call FileCopy(cdJpegBrowser.fileName, CommonHelper.getImgPath & "\" & txtFileName.Caption)
    End If
    MsgBox "Record Updated!", vbInformation
    Call showSelectedData
  End If
End Sub

Private Sub cmbNewRec_Click()
  Call restoreFormDefaultSkin
  If (isFormValid) Then
    rs.AddNew
    rs!Name = txtName.Text
    rs!FILE_NAME = txtFileName.Caption
    rs!CREATED_BY = UserSession.getLoginUser
    rs!CREATED_DATE = Now
    rs.Update
    If (cdJpegBrowser.fileName <> vbNullString) Then
      Call FileCopy(cdJpegBrowser.fileName, CommonHelper.getImgPath & "\" & txtFileName.Caption)
    End If
    MsgBox "Record Added!", vbInformation
    Call showSelectedData
  End If
  
End Sub
Private Sub clearForm()
  lblID.Caption = ""
  txtName.Text = ""
  txtFileName.Caption = ""
  lblCreatedBy.Caption = ""
  lblCreatedDate.Caption = ""
  lblLatModBy.Caption = ""
  lblLastModDate.Caption = ""
  imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & Constants.MISSING_LOC_IMAGE_NAME)
End Sub
Private Sub cmdLoadImg_Click()
  cdJpegBrowser.Filter = "JPEG FIle (*.jpg)|*.jpg"
  cdJpegBrowser.DialogTitle = "Select a JPEG image file"
  cdJpegBrowser.ShowOpen
  
  imgLoc.Picture = LoadPicture(cdJpegBrowser.fileName)
  txtFileName.Caption = CommonHelper.getFileName(cdJpegBrowser.fileName)
End Sub
Private Sub showSelectedData()
  lblID.Caption = rs!ID
  txtName.Text = rs!Name
  txtFileName.Caption = CommonHelper.extractStringValue(rs!FILE_NAME)
  lblCreatedBy.Caption = CommonHelper.extractStringValue(rs!CREATED_BY)
  lblCreatedDate.Caption = CommonHelper.extractDateValue(rs!CREATED_DATE)
  lblLatModBy.Caption = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
  lblLastModDate.Caption = CommonHelper.extractDateValue(rs!LAST_MOD_DATE)
  If (txtFileName.Caption <> vbNullString) Then
    imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & txtFileName.Caption)
  Else
    imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & Constants.MISSING_LOC_IMAGE_NAME)
  End If
End Sub

Private Sub dgLocationMapping_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedData
End Sub
Private Sub restoreFormDefaultSkin()
  Call CommonHelper.toDefaultSkin(txtName)
End Sub
Private Sub dgLocationMapping_SelChange(Cancel As Integer)
  Call showSelectedData
End Sub

Private Sub Form_Load()
   imgLoc.Picture = LoadPicture(CommonHelper.getImgPath & "\" & Constants.MISSING_LOC_IMAGE_NAME)
   Call populateDataGrid
End Sub
Public Sub populateDataGrid()
  Set rs = LookupDao.getLocationMappingRS
  Set dgLocationMapping.DataSource = rs
  dgLocationMapping.Refresh
End Sub

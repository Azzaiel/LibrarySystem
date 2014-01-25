VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStudents 
   Caption         =   "Student Information"
   ClientHeight    =   5895
   ClientLeft      =   525
   ClientTop       =   660
   ClientWidth     =   18855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   18855
   Begin VB.Frame Frame2 
      Caption         =   "Search Panel"
      Height          =   855
      Left            =   5280
      TabIndex        =   33
      Top             =   120
      Width           =   13335
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
         Height          =   495
         Left            =   12120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmbClearSearch 
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
         Left            =   11160
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmbSearch 
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
         Height          =   495
         Left            =   10200
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtSearchLrn 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmSearchSection 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSearchLastName 
         Height          =   285
         Left            =   7800
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FF80&
         Caption         =   "LRN"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "SECTION"
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "LAST_NAME"
         Height          =   255
         Left            =   6720
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
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
      Left            =   3960
      TabIndex        =   8
      Top             =   4560
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
      Left            =   2040
      TabIndex        =   9
      Top             =   5160
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
      Left            =   2760
      TabIndex        =   7
      Top             =   4560
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
      Left            =   1560
      TabIndex        =   6
      Top             =   4560
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
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detail Form"
      Height          =   4335
      Left            =   240
      TabIndex        =   17
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cmSections 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TxtLAST_NAME 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtLrn 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtMIDDLE_NAME 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "*SECTION"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "*LAST_NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblLastModBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblCreatedDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblCreatedBy 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod by"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FF80&
         Caption         =   "Created date"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label LA 
         BackColor       =   &H0080FF80&
         Caption         =   "Created by"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "MIDDLE_NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LBFIRST_NAME 
         BackColor       =   &H0080FF80&
         Caption         =   "*FIRST_NAME"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblName 
         BackColor       =   &H0080FF80&
         Caption         =   "* LRN"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblID 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLastModDate 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Last Mod Date"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dgStudents 
      Height          =   4575
      Left            =   5280
      TabIndex        =   16
      Top             =   1080
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
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
End
Attribute VB_Name = "frmStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As ADODB.Recordset
Private tempRs As ADODB.Recordset
Private sectionItemList() As Variant



Private Sub DataGrid1_Click()

End Sub

Private Sub Command4_Click()
  
End Sub

Private Sub cmbClear_Click()
  Call toogelInsertMode(False)
  Call resetFromSkin
  Call clearForm
End Sub
Public Sub clearForm()
   lblID.Caption = ""
   txtLRN.Text = ""
   txtFirstName.Text = ""
   txtMIDDLE_NAME.Text = ""
   TxtLAST_NAME.Text = ""
   lblCreatedBy = ""
   lblCreatedDate = ""
   lblLastModBy = ""
   lblLastModDate = ""
   If (dgStudents.SelBookmarks.Count > 0) Then
     dgStudents.SelBookmarks.Remove (0)
   End If
   cmSections.ListIndex = -1
End Sub

Private Sub cmbClearSearch_Click()
  txtSearchLrn.Text = ""
  cmSearchSection.ListIndex = -1
  txtSearchLastName.Text = ""
End Sub

Private Sub cmbClose_Click()
  Unload Me
End Sub

Private Sub cmbDelete_Click()
 Call resetFromSkin
 Dim response As String
 response = MsgBox("Are you sure you want to delete the record?", vbOKCancel, "Question")
  If (response = vbOK) Then
  
    If (StudentDao.isStudentBeingUsed(rs!id)) Then
    
      MsgBox "Cannot delete record, It has a reference to the transaction data", vbCritical
      Exit Sub
    
    End If
  
    Set tempRs = StudentDao.getRsByID(rs!id)
    tempRs.Delete
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Deleted", vbInformation
    Call clearForm
    Call populateDataGrid
  End If
End Sub

Private Sub cmbEdit_Click()
  Call resetFromSkin
  If (isFormValid) Then
  
    If (isLrnAlreadyInUse(rs!id)) Then
      MsgBox "Lrn is Already in use", vbCritical
      Exit Sub
    End If
      
    Set tempRs = StudentDao.getRsByID(rs!id)
    tempRs!lrn = Val(txtLRN.Text)
    tempRs!FIRST_NAME = txtFirstName.Text
    tempRs!MIDDLE_NAME = txtMIDDLE_NAME.Text
    tempRs!LAST_NAME = TxtLAST_NAME.Text
    tempRs!SECTION_ID = getSectionID
    tempRs!LAST_MOD_BY = UserSession.getLoginUser
    tempRs!LAST_MOD_DATE = Now
    tempRs.Update
    Call DbInstance.closeRecordSet(tempRs)
    MsgBox "Record Updated!!", vbInformation
    Call clearForm
    Call populateDataGrid
  End If
End Sub

Private Sub cmbExport_Click()
  
  Dim excelApp As New Excel.Application
  Dim oBook As New Excel.Workbook
  Dim oSheet As New Excel.Worksheet
  
  Set excelApp = CreateObject("Excel.Application")
  Set oBook = excelApp.Workbooks.Open(CommonHelper.getTemplatesPath & "\" & Constants.STUDENT_TEMPLATE)
  Set oSheet = excelApp.Worksheets(1)
  
  oSheet.Range("A2").CopyFromRecordset dgStudents.DataSource
  oSheet.Columns.AutoFit
  oSheet.Range("F1:F1").EntireColumn.Hidden = True

  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\" & Constants.TEMP_WORK_BOOK
  
  'If (UserSession.role = "Admin") Then
  ' excelApp.Visible = True
  'Else
    Dim pdfFilePat As String
    pdfFilePat = CommonHelper.getTempPath & "\temp_" & Format(Now, "mmhhyysssh") & ".pdf"
    Call oBook.ExportAsFixedFormat(xlTypePDF, pdfFilePat, xlQualityStandard, False, True)
    oBook.Close
    Call CommonHelper.openFile(pdfFilePat, Me.hWnd)
  'End If

End Sub

Private Sub cmbNewRec_Click()
  Call resetFromSkin
  If (cmbNewRec.Caption = "New") Then
    Call toogelInsertMode(True)
    txtLRN.SetFocus
  Else
    If (isFormValid) Then
     
      If (isLrnAlreadyInUse) Then
        MsgBox "Lrn is Already in use", vbCritical
        Exit Sub
      End If
    
      Set tempRs = StudentDao.getFakeRs
      tempRs.AddNew
      tempRs!lrn = Val(txtLRN.Text)
      tempRs!FIRST_NAME = txtFirstName.Text
      tempRs!MIDDLE_NAME = txtMIDDLE_NAME.Text
      tempRs!LAST_NAME = TxtLAST_NAME.Text
      tempRs!SECTION_ID = getSectionID
      tempRs!CREATED_BY = UserSession.getLoginUser
      tempRs!CREATED_DATE = Now
      tempRs!LAST_MOD_BY = UserSession.getLoginUser
      tempRs!LAST_MOD_DATE = Now
      tempRs.Update
      Call DbInstance.closeRecordSet(tempRs)
      MsgBox "Record Created!!", vbInformation
      Call toogelInsertMode(False)
      Call clearForm
      Call populateDataGrid
    End If
  End If
End Sub
Private Function isLrnAlreadyInUse(Optional studentID As Integer = -1) As Boolean
   Set tempRs = StudentDao.getRsByLrn(txtLRN)
   If (tempRs.RecordCount > 0) Then
     If (tempRs!id = studentID) Then
       isLrnAlreadyInUse = False
     Else
       isLrnAlreadyInUse = True
     End If
   Else
     isLrnAlreadyInUse = False
   End If
   Call DbInstance.closeRecordSet(tempRs)
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
Private Function getSectionID() As Integer
  Dim index As Integer
  index = cmSections.ListIndex
  getSectionID = identifySectionID(index)
End Function

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

Private Sub cmbSearch_Click()
  Set dgStudents.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
  Set rs = StudentDao.searchRs(Val(txtSearchLrn.Text), getSearchSectionID, txtSearchLastName.Text)
  Set dgStudents.DataSource = rs
  dgStudents.Refresh
  Call clearForm
  Call formatDataGrid
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dgStudents_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call showSelectedData
End Sub

Private Sub dgStudents_SelChange(Cancel As Integer)
 Call showSelectedData
End Sub
Private Sub showSelectedData()
   lblID.Caption = rs!id
   txtLRN.Text = rs!lrn
   txtFirstName.Text = rs!FIRST_NAME
   txtMIDDLE_NAME.Text = rs!MIDDLE_NAME
   TxtLAST_NAME.Text = rs!LAST_NAME
   lblCreatedBy = rs!CREATED_BY
   lblCreatedDate = CommonHelper.extractDateValue(rs!CREATED_DATE)
   lblLastModBy = CommonHelper.extractStringValue(rs!LAST_MOD_BY)
   lblLastModDate = CommonHelper.extractStringValue(rs!LAST_MOD_DATE)

   Dim index As Integer
   cmSections.ListIndex = -1
   For index = 0 To UBound(sectionItemList)
     If (rs!SECTION_ID = sectionItemList(index, Constants.ITEM_VALUE_INDEX)) Then
       cmSections.ListIndex = index
     End If
   Next index
End Sub

Private Sub Form_Load()
  Call populateDropDown
  Call populateDataGrid
End Sub
Private Sub populateDropDown()
  sectionItemList = LookupDao.getSectionsItemList
  Dim index As Integer
  cmSections.Clear
  cmSearchSection.Clear
  For index = 0 To UBound(sectionItemList)
    cmSections.AddItem (sectionItemList(index, Constants.ITEM_LABEL_INDEX))
    cmSearchSection.AddItem (sectionItemList(index, Constants.ITEM_LABEL_INDEX))
  Next index
End Sub
Private Sub populateDataGrid()
  Set rs = StudentDao.getAllRs
  Set dgStudents.DataSource = rs
  dgStudents.Refresh
  Call formatDataGrid
End Sub

Private Sub formatDataGrid()
  If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Call showSelectedData
  Else
    Call resetFromSkin
    Call clearForm
  End If
  
  With dgStudents
    .Columns(0).Width = 400
    .Columns(0).Alignment = dbgCenter
    .Columns(5).Visible = False
    .Columns(8).Width = 1500
    .Columns(8).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(8).Alignment = dbgCenter
    .Columns(10).Width = 1500
    .Columns(10).NumberFormat = Constants.DEFAULT_FORMAT
    .Columns(10).Alignment = dbgCenter
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dgStudents.DataSource = Nothing
  Call DbInstance.closeRecordSet(rs)
End Sub
Private Function isFormValid() As Boolean
  Dim isValid As Boolean
  isValid = True
  If (Not CommonHelper.hasValidValue(txtLRN.Text)) Then
     Call CommonHelper.sendWarning(txtLRN, "LRN is required field")
     isFormValid = False
     Exit Function
  End If
  If (Not CommonHelper.hasValidValue(txtFirstName.Text)) Then
     Call CommonHelper.sendWarning(txtFirstName, "First Name is required field")
     isFormValid = False
     Exit Function
  End If
  If (Not CommonHelper.hasValidValue(TxtLAST_NAME.Text)) Then
     Call CommonHelper.sendWarning(TxtLAST_NAME, "Last Name is required field")
     isFormValid = False
     Exit Function
  End If
  If (Not CommonHelper.hasValidValue(CStr(getSectionID))) Then
     Call CommonHelper.sendComboBoxWarning(cmSections, "Please select a section")
     isFormValid = False
     Exit Function
  End If
  isFormValid = isValid
End Function
Private Sub resetFromSkin()

 Call CommonHelper.toDefaultSkin(txtLRN)
 Call CommonHelper.toDefaultSkin(txtFirstName)
 Call CommonHelper.toDefaultSkin(TxtLAST_NAME)
 Call CommonHelper.toComboBoxDefaultSkin(cmSections)

End Sub

Private Sub txtFirstName_Change()
   If (Len(txtFirstName) = 1) Then
    txtFirstName = StrConv(txtFirstName, vbProperCase)
    txtFirstName.SelStart = 1
  End If
End Sub

Private Sub TxtLAST_NAME_Change()
  If (Len(TxtLAST_NAME) = 1) Then
    TxtLAST_NAME = StrConv(TxtLAST_NAME, vbProperCase)
    TxtLAST_NAME.SelStart = 1
  End If
End Sub

Private Sub txtMIDDLE_NAME_Change()
  If (Len(txtMIDDLE_NAME) = 1) Then
    txtMIDDLE_NAME = StrConv(txtMIDDLE_NAME, vbProperCase)
    txtMIDDLE_NAME.SelStart = 1
  End If
End Sub

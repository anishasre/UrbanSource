VERSION 5.00
Begin VB.Form frmEmployeeSubledger 
   Caption         =   "Create Officials"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmEmployeeSubledger.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2033
      TabIndex        =   5
      Top             =   2130
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   3413
      TabIndex        =   6
      Top             =   2130
      Width           =   1335
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1710
      Width           =   4245
   End
   Begin VB.ComboBox cmbDesignation 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1365
      Width           =   4245
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1035
      Width           =   4245
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   315
      Width           =   4245
   End
   Begin VB.TextBox txtDDOCode 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   675
      Width           =   4245
   End
   Begin VB.Label lblSubledgerCode 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1350
      TabIndex        =   13
      Top             =   2610
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subledger Code"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2610
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Left            =   1170
      TabIndex        =   11
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   1275
      TabIndex        =   10
      Top             =   360
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "DDO Code"
      Height          =   195
      Left            =   930
      TabIndex        =   9
      Top             =   735
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      Height          =   195
      Left            =   885
      TabIndex        =   8
      Top             =   1095
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
      Height          =   195
      Left            =   870
      TabIndex        =   7
      Top             =   1455
      Width           =   840
   End
End
Attribute VB_Name = "frmEmployeeSubledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private mEmployeeID As Variant
    Dim mSubLedgerTypeID    As Integer
    
    Private Sub GetEmployeeDetails(EmpID As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim obJDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        
        '*********************************************************************************************'
        '       Procedure to Fill all the details of a particular Employee from Sthapana DB           '
        '*********************************************************************************************'
        On Error GoTo Err
        mSQL = " SELECT  TB_EmployeeDetails_TRN.intEmpId, TB_EmployeeDetails_TRN.chvEmpId, TB_EmployeeDetails_TRN.chvEmpName,TB_EmployeeDetails_TRN.chvPFNo, TB_EmployeeDetails_TRN.intDesigId, TB_EmployeeDetails_TRN.intDeptId, TB_EmployeeDetails_TRN.intSecId, TB_Designation_Lcl_Mst.chvDesigCode, TB_Designation_Lcl_Mst.chvDesigName, TB_Department_Lcl_Mst.chvDeptCode, TB_Department_Lcl_Mst.chvDeptName, TB_Section_Lcl_Mst.chvSectionCode, TB_Section_Lcl_Mst.chvSectionName "
        mSQL = mSQL + " FROM TB_EmployeeDetails_TRN INNER JOIN  TB_Designation_Lcl_Mst ON TB_EmployeeDetails_TRN.intDesigId = TB_Designation_Lcl_Mst.intDesigId INNER JOIN"
        mSQL = mSQL + " TB_Department_Lcl_Mst ON TB_EmployeeDetails_TRN.intDeptId = TB_Department_Lcl_Mst.intDeptId INNER JOIN"
        mSQL = mSQL + " TB_Section_Lcl_Mst ON TB_EmployeeDetails_TRN.intSecId = TB_Section_Lcl_Mst.intSectionId"
        mSQL = mSQL + " Where TB_EmployeeDetails_TRN.intEmpId = " & EmpID
       If obJDB.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
            Rec.Open mSQL, mCnn, adOpenStatic, adLockReadOnly
            If Not (Rec.EOF And Rec.BOF) Then
                txtName.Text = IIf(IsNull(Rec!chvEmpName), "", Rec!chvEmpName)
                txtName.Tag = IIf(IsNull(Rec!intEmpID), "", Rec!intEmpID)
                txtDDOCode.Text = IIf(IsNull(Rec!chvEmpID), "", Rec!chvEmpID)
                cmbDepartment.Text = IIf(IsNull(Rec!chvDeptName), "", Rec!chvDeptName)
                cmbDesignation.Text = IIf(IsNull(Rec!chvDesigName), "", Rec!chvDesigName)
                cmbSection.Text = IIf(IsNull(Rec!chvSectionName), "", Rec!chvSectionName)
            End If
            Rec.Close
        End If
        mCnn.Close
        If txtName.Tag <> "" Then
            If obJDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSQL = "Select vchSubLedgerCode From faSubSidiaryAccountHeads"
                mSQL = mSQL + " Where numEmpID = " & txtName.Tag
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    txtDDOCode.Tag = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
                    lblSubledgerCode.Caption = txtDDOCode.Tag
                End If
                Rec.Close
            End If
        End If
        mCnn.Close
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
        
    Private Sub FormInitialize()
        txtName.Text = ""
        txtName.Tag = ""
        txtDDOCode.Text = ""
        txtDDOCode.Tag = ""
        cmbDesignation.ListIndex = 0
        cmbDepartment.ListIndex = 0
        cmbSection.ListIndex = 0
    End Sub
    
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdSave_Click()
        On Error GoTo Err:
            Dim obJDB As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim aryOut As Variant
            
            '*********************************************************************************************'
            '                  Procedure to Fill all the details of a particular demand                   '
            '*********************************************************************************************'
            On Error GoTo Err
            If txtName.Text = "" Then
                MsgBox "Please enter the Name", vbInformation
                txtName.SetFocus
                Exit Sub
            End If
            If cmbDesignation.ListIndex <= 0 Then
                MsgBox "Please select the Designation", vbInformation
                cmbDesignation.SetFocus
                Exit Sub
            End If
            If cmbDepartment.ListIndex <= 0 Then
                MsgBox "Please select the Department", vbInformation
                cmbDepartment.SetFocus
                Exit Sub
            End If
            If cmbSection.ListIndex <= 0 Then
                MsgBox "Please select the Section", vbInformation
                cmbSection.SetFocus
                Exit Sub
            End If
            
            If obJDB.SetConnection(mCnn) Then
                aryIn = Array(mSubLedgerTypeID, _
                    val(txtDDOCode.Tag), _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    txtName.Tag, _
                    txtDDOCode.Text, _
                    Null, _
                    Trim(txtName.Text), _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, gbFinancialYearID, cmbDesignation.Text, cmbDepartment.Text, cmbSection.Text, cmbDesignation.ItemData(cmbDesignation.ListIndex), cmbDepartment.ItemData(cmbDepartment.ListIndex), cmbSection.ItemData(cmbSection.ListIndex))
                obJDB.ExecuteSP "spSaveSubSidiaryAccountHeads", aryIn, aryOut, , mCnn
                lblSubledgerCode.Caption = aryOut(0, 0)
                txtDDOCode.Tag = aryOut(0, 0)
            Else
                MsgBox "Connection to Finance doesnot Exist, Please contact your System Administrator", vbInformation
            End If
            cmdSave.Enabled = False
            MsgBox "Successfully Saved", vbInformation
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim obJDB   As New clsDB
        Dim mSQL    As String
        
        On Error GoTo Err
        mSubLedgerTypeID = 10 'Officials
        If obJDB.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
        
            mSQL = "Select chvDesigName,intDesigID From TB_Designation_Lcl_Mst Order By chvDesigName"
            PopulateList cmbDesignation, mSQL, , True, True, True, enuSourceString.Sthapana
            
            mSQL = "Select chvDeptName,intDeptID From TB_Department_Lcl_Mst Order By chvDeptName"
            PopulateList cmbDepartment, mSQL, , True, True, True, enuSourceString.Sthapana
           
            mSQL = "Select chvSectionName,intSectionID From TB_Section_Lcl_Mst Order By chvSectionName"
            PopulateList cmbSection, mSQL, , True, True, True, enuSourceString.Sthapana
            
           If EmployeeID <> "" Then
                Call GetEmployeeDetails(EmployeeID)
            Else
                Call FormInitialize
            End If
        Else
            MsgBox "Didn't able to Establish connection to Sthapana - Establishment Module", vbInformation
        End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Public Property Let EmployeeID(mData As Variant)
        mEmployeeID = mData
    End Property
    
    Public Property Get EmployeeID() As Variant
        EmployeeID = mEmployeeID
    End Property

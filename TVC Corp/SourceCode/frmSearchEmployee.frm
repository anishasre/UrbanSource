VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchEmployee 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seach Employees"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   6360
      Width           =   6135
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   -3630
      Top             =   1680
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.ListBox lstEmployees 
      BackColor       =   &H80000013&
      Height          =   4860
      Left            =   30
      TabIndex        =   4
      Top             =   1440
      Width           =   6945
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   -90
      Width           =   6975
      Begin VB.ComboBox cmbDesignations 
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   5445
      End
      Begin VB.ComboBox cmbDepartment 
         BackColor       =   &H80000004&
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   5445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designations :"
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department :"
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   6420
      Width           =   615
   End
End
Attribute VB_Name = "frmSearchEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private mUserFlag As Boolean
    Public Property Let CommonUser(mFlag As Boolean)
        mUserFlag = mFlag
    End Property
    Private Sub FillSeachCombo()
        Dim mDepartment As String
        Dim mDesig As String
        Dim mSQL As String
        Dim mCommonSql As String
        
        If cmbDepartment.ListIndex < 1 Then
            mDepartment = "%"
        Else
            mDepartment = CStr(cmbDepartment.itemData(cmbDepartment.ListIndex))
        End If
        
        If cmbDesignations.ListIndex < 1 Then
            mDesig = "%"
        Else
            mDesig = CStr(cmbDesignations.itemData(cmbDesignations.ListIndex))
        End If
        mSQL = "Select chvEmpName,intEmpId from TB_EmployeeDetails_Trn where Convert(varchar(10),intDeptID) Like '" & mDepartment & "' And Convert(varchar(10),intDesigID) Like '" & mDesig & "'And chvEmpName Like '" & "%" & Trim(txtSearch.Text) & "%" & "' Order By Ltrim(chvEmpName)"
        mCommonSql = "SELECT vchEmpName,numUserID FROM DB_Masters.dbo.GM_User "
        mCommonSql = mCommonSql + " Where tnyDeletedStatus = 0 And Convert(varchar(10),isnull(intDeptID,0)) Like '" & mDepartment & "' And Convert(varchar(10),intDesignationID) Like '" & mDesig & "'And vchEmpName Like '" & "%" & Trim(txtSearch.Text) & "%" & "' Order By Ltrim(vchEmpName)"
        If mUserFlag = False Then
            PopulateList lstEmployees, mSQL, , False, True, True, enuSourceString.Sthapana
        Else
            PopulateList lstEmployees, mCommonSql, , False, , True, enuSourceString.DBMaster
        End If
        If lstEmployees.ListCount > 0 Then
            lstEmployees.Selected(0) = True
        End If
    End Sub
    Private Sub cmbDepartment_Click()
        Dim mCommonSql As String
        If cmbDepartment.ListIndex > 0 Then
            If mUserFlag = False Then
                PopulateList cmbDesignations, "Select Distinct chvDesigName,TB_Designation_Lcl_Mst.intDesigId [intDesignationID] from TB_EmployeeDetails_TRN Inner join TB_Designation_Lcl_Mst On TB_Designation_Lcl_Mst.intDesigId=TB_EmployeeDetails_TRN.intDesigId where intDeptId=" & cmbDepartment.itemData(cmbDepartment.ListIndex) & " Order By chvDesigName", , True, True, True, enuSourceString.Sthapana
            Else
                mCommonSql = "SELECT Distinct case when DB_Masters..GM_User.intDesignationID =0 then 'Temperary Staff'else chvDesigName end as chvDesigName,DB_Masters..GM_User.intDesignationID [intDesignationID] FROM DB_Masters.dbo.GM_User "
                mCommonSql = mCommonSql + "LEFT JOIN DB_Sthapana..TB_Designation_Lcl_Mst ON DB_Sthapana..TB_Designation_Lcl_Mst.intDesigId=DB_Masters.dbo.GM_User.intDesignationID "
                mCommonSql = mCommonSql + "INNER JOIN DB_Sthapana..TB_Department_Lcl_Mst ON DB_Sthapana..TB_Department_Lcl_Mst.intDeptId=DB_Masters.dbo.GM_User.intDeptID"
                mCommonSql = mCommonSql + " Where tnyDeletedStatus = 0 And DB_Sthapana..TB_Department_Lcl_Mst.intDeptID = " & cmbDepartment.itemData(cmbDepartment.ListIndex) & " Order By chvDesigName"
        
                PopulateList cmbDesignations, mCommonSql, , True, True, True, enuSourceString.DBMaster
                cmbDesignations.AddItem "Temperary Staff", 1
            End If
        End If
        Call FillSeachCombo
        txtSearch.SetFocus
    End Sub
    Private Sub cmbDesignations_Click()
        Call FillSeachCombo
        txtSearch.SetFocus
    End Sub
    Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
        End If
    End Sub
    Private Sub Form_Load()
        winXPC.InitIDESubClassing
        PopulateList cmbDepartment, "SELECT chvDeptName,intDeptID FROM TB_Department_Lcl_Mst Order By chvDeptName", , True, True, True, enuSourceString.Sthapana
        cmbDepartment.AddItem "Unknown Department"
        cmbDepartment.itemData(cmbDepartment.NewIndex) = 1
        gbSearchID = 0
        gbSearchStr = ""
        Call FillSeachCombo
    End Sub
    Private Sub lstEmployees_DblClick()
        Dim mSQL As String
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        
        If lstEmployees.ListIndex > -1 Then
            If mUserFlag = False Then
                objDB.CreateNewConnection mCnn, enuSourceString.Sthapana
                mSQL = "SELECT TB_Designation_Lcl_Mst.chvDesigName,chvDeptName FROM  TB_Designation_Lcl_Mst INNER JOIN TB_EmployeeDetails_TRN ON TB_EmployeeDetails_TRN.intDesigId = TB_Designation_Lcl_Mst.intDesigId INNER JOIN TB_Department_Lcl_Mst ON TB_Department_Lcl_Mst.intDeptId=TB_EmployeeDetails_TRN.intDeptId WHERE TB_EmployeeDetails_TRN.intEmpId = " & lstEmployees.itemData(lstEmployees.ListIndex)
            Else
                objDB.CreateNewConnection mCnn, enuSourceString.DBMaster
                mSQL = "SELECT case when DB_Masters..GM_User.intDesignationID=0 then 'Temperary Staff' else chvDesigName end as chvDesigName,case when isnull(DB_Masters..GM_User.intDeptID,0)=0 Then 'Unknown Department' else chvDeptName End As chvDeptName FROM DB_Masters.dbo.GM_User "
                mSQL = mSQL + "LEFT JOIN DB_Sthapana..TB_Designation_Lcl_Mst ON DB_Sthapana..TB_Designation_Lcl_Mst.intDesigId=DB_Masters.dbo.GM_User.intDesignationID "
                mSQL = mSQL + "LEFT JOIN DB_Sthapana..TB_Department_Lcl_Mst ON DB_Sthapana..TB_Department_Lcl_Mst.intDeptId=DB_Masters.dbo.GM_User.intDeptID"
                mSQL = mSQL + " Where tnyDeletedStatus = 0 And DB_Masters..GM_User.numUserID = " & lstEmployees.itemData(lstEmployees.ListIndex)
            End If
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                gbSearchID = lstEmployees.itemData(lstEmployees.ListIndex)
                gbSearchStr = lstEmployees.List(lstEmployees.ListIndex)
                If mUserFlag = False Then
                    cmbDepartment.Text = Rec!chvDeptName
                    cmbDesignations.Text = Rec!chvDesigName
                End If
            End If
            Rec.Close
            mCnn.Close
            gbSearchStr = gbSearchStr + "^" + cmbDepartment.Text + "^" + cmbDesignations.Text
            Unload Me
        End If
    End Sub
    Private Sub txtSearch_Change()
        Call FillSeachCombo
    End Sub
    Private Sub txtSearch_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 And lstEmployees.ListIndex > -1 Then
            lstEmployees.ListIndex = 0
            lstEmployees_DblClick
        End If
    End Sub

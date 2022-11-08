VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchEmplyees 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   10875
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7035
      Width           =   2205
   End
   Begin VB.ComboBox cmbDepartment 
      Height          =   315
      Left            =   8565
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7035
      Width           =   2310
   End
   Begin VB.ComboBox cmbDesignation 
      Height          =   315
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7035
      Width           =   2400
   End
   Begin VB.TextBox txtEmpCode 
      Height          =   315
      Left            =   4650
      TabIndex        =   6
      Top             =   7035
      Width           =   1500
   End
   Begin VB.TextBox txtEmpName 
      Height          =   315
      Left            =   570
      TabIndex        =   5
      Top             =   7035
      Width           =   4050
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   13410
      TabIndex        =   3
      Top             =   0
      Width           =   13410
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   13350
      TabIndex        =   1
      Top             =   7515
      Width           =   13410
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   450
         Left            =   45
         TabIndex        =   10
         Top             =   30
         Width           =   1530
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   450
         Left            =   5910
         TabIndex        =   2
         Top             =   45
         Width           =   1530
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid Grid 
      Height          =   5640
      Left            =   60
      TabIndex        =   0
      Top             =   1125
      Width           =   13290
      _cx             =   23442
      _cy             =   9948
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14335672
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchEmplyees.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Name                          Emp.Code       Designation            Department            Sections"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   6810
      Width           =   12465
   End
End
Attribute VB_Name = "frmSearchEmplyees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '*********************************************************************************************'
    '                  Form to search Employees from DB_Sthapana                                  '
    '*********************************************************************************************'
    Private Sub FillEmployees()
        Dim mSQL        As String
        Dim mSQL1       As String
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim obJDB       As New clsDB
        Dim mCount      As Integer
        
        mSQL = " SELECT  TB_EmployeeDetails_TRN.intEmpId, TB_EmployeeDetails_TRN.chvEmpId, TB_EmployeeDetails_TRN.chvEmpName,TB_EmployeeDetails_TRN.chvPFNo, TB_EmployeeDetails_TRN.intDesigId, TB_EmployeeDetails_TRN.intDeptId, TB_EmployeeDetails_TRN.intSecId, TB_Designation_Lcl_Mst.chvDesigCode, TB_Designation_Lcl_Mst.chvDesigName, TB_Department_Lcl_Mst.chvDeptCode, TB_Department_Lcl_Mst.chvDeptName, TB_Section_Lcl_Mst.chvSectionCode, TB_Section_Lcl_Mst.chvSectionName "
        mSQL = mSQL + " FROM TB_EmployeeDetails_TRN INNER JOIN  TB_Designation_Lcl_Mst ON TB_EmployeeDetails_TRN.intDesigId = TB_Designation_Lcl_Mst.intDesigId INNER JOIN"
        mSQL = mSQL + " TB_Department_Lcl_Mst ON TB_EmployeeDetails_TRN.intDeptId = TB_Department_Lcl_Mst.intDeptId INNER JOIN"
        mSQL = mSQL + " TB_Section_Lcl_Mst ON TB_EmployeeDetails_TRN.intSecId = TB_Section_Lcl_Mst.intSectionId"
        If Trim(txtEmpName.Text) <> "" Then
            mSQL1 = " chvEmpName Like '%" & txtEmpName.Text & "%'"
        End If
        If Trim(txtEmpCode.Text) <> "" Then
            mSQL1 = mSQL1 + " And chvEmpID ='" & Trim(txtEmpCode.Text) & "'"
        End If
        If cmbDesignation.ListIndex > 0 Then
            mSQL1 = mSQL1 + " And chvDesigName ='" & cmbDesignation.Text & "'"
        End If
        If cmbDepartment.ListIndex > 0 Then
            mSQL1 = mSQL1 + " And chvDeptName ='" & cmbDepartment.Text & "'"
        End If
        If cmbSection.ListIndex > 0 Then
            mSQL1 = mSQL1 + " And chvSectionName ='" & cmbSection.Text & "'"
        End If
        If Trim(mSQL1) <> "" Then
            If (Trim(Left(mSQL1, 4)) = "And") Then
                mSQL1 = mID(mSQL1, 5)
            End If
            mSQL = mSQL + " Where " + mSQL1
        End If
        If obJDB.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
            Rec.Open mSQL, mCnn, adOpenStatic, adLockReadOnly
            Grid.Rows = 1
            If Not (Rec.BOF And Rec.EOF) Then
                
                While Not Rec.EOF
                    mCount = mCount + 1
                    If Grid.Rows <= mCount Then
                        Grid.Rows = Grid.Rows + 50
                    End If
                    Grid.TextMatrix(mCount, 1) = IIf(IsNull(Rec!chvEmpName), "", Rec!chvEmpName)
                    Grid.TextMatrix(mCount, 2) = IIf(IsNull(Rec!chvEmpID), "", Rec!chvEmpID)
                    Grid.TextMatrix(mCount, 3) = IIf(IsNull(Rec!chvDesigName), "", Rec!chvDesigName)
                    Grid.TextMatrix(mCount, 4) = IIf(IsNull(Rec!chvDeptName), "", Rec!chvDeptName)
                    Grid.TextMatrix(mCount, 5) = IIf(IsNull(Rec!chvSectionName), "", Rec!chvSectionName)
                    Grid.TextMatrix(mCount, 6) = IIf(IsNull(Rec!intEmpID), "", Rec!intEmpID)
                    Grid.TextMatrix(mCount, 7) = IIf(IsNull(Rec!intDesigID), "", Rec!intDesigID)
                    Grid.TextMatrix(mCount, 8) = IIf(IsNull(Rec!intDeptID), "", Rec!intDeptID)
                    Grid.TextMatrix(mCount, 9) = IIf(IsNull(Rec!intSecID), "", Rec!intSecID)
                    Rec.MoveNext
                Wend
            End If
            Rec.Close
        Else
            MsgBox "Didn't able to Establish connection to Sthapana - Establishment Module", vbInformation
        End If
    End Sub
    
    Private Sub cmdNew_Click()
        frmEmployeeSubledger.EmployeeID = ""
        frmEmployeeSubledger.Show vbModal
    End Sub

    Private Sub cmdSearch_Click()
        Call FillEmployees
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim obJDB   As New clsDB
        Dim mSQL    As String
        
        If obJDB.CreateNewConnection(mCnn, enuSourceString.Sthapana) Then
        
            mSQL = "Select chvDesigName,intDesigID From TB_Designation_Lcl_Mst Order By chvDesigName"
            PopulateList cmbDesignation, mSQL, , True, True, True, enuSourceString.Sthapana
            
            mSQL = "Select chvDeptName,intDeptID From TB_Department_Lcl_Mst Order By chvDeptName"
            PopulateList cmbDepartment, mSQL, , True, True, True, enuSourceString.Sthapana
           
            mSQL = "Select chvSectionName,intSectionID From TB_Section_Lcl_Mst Order By chvSectionName"
            PopulateList cmbSection, mSQL, , True, True, True, enuSourceString.Sthapana
            
            Call FillEmployees
        Else
            MsgBox "Didn't able to Establish connection to Sthapana - Establishment Module", vbInformation
        End If
    End Sub

    Private Sub Grid_DblClick()
        If Grid.TextMatrix(Grid.Row, 6) <> "" Then
            frmEmployeeSubledger.EmployeeID = Grid.TextMatrix(Grid.Row, 6)
            frmEmployeeSubledger.Show vbModal
        End If
    End Sub

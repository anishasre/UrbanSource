VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStockRegisterOfReceiptBooks 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Register of Receipt Books"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Receipt Issued From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   6000
      Begin VB.ComboBox cmbDesignation 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1260
         Width           =   4080
      End
      Begin VB.ComboBox cmbDepartment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   915
         Width           =   4080
      End
      Begin VB.ComboBox cmbStaffName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmStockRegisterOfReceiptBooks.frx":0000
         Left            =   1800
         List            =   "frmStockRegisterOfReceiptBooks.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1590
         Width           =   4080
      End
      Begin VB.ComboBox cmbStaff 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmStockRegisterOfReceiptBooks.frx":0004
         Left            =   1800
         List            =   "frmStockRegisterOfReceiptBooks.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   570
         Width           =   4080
      End
      Begin VB.ComboBox cmbBookIssueTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   4080
      End
      Begin VB.ComboBox cmbCounters 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmStockRegisterOfReceiptBooks.frx":0008
         Left            =   1800
         List            =   "frmStockRegisterOfReceiptBooks.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   25
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   23
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lblTempStaff 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Staff Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   630
         TabIndex        =   22
         Top             =   1590
         Width           =   1110
      End
      Begin VB.Label lblcaptions 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Staff Type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   720
         TabIndex        =   21
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Book Issue to:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblcaptions 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Counter:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   900
         TabIndex        =   19
         Top             =   1935
         Width           =   840
      End
   End
   Begin VB.Frame fraStaff 
      BackColor       =   &H80000018&
      Caption         =   "Receipt Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   60
      TabIndex        =   0
      Top             =   2355
      Width           =   5970
      Begin MSComCtl2.DTPicker dtissuedDate 
         Height          =   285
         Left            =   1380
         TabIndex        =   6
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58064897
         CurrentDate     =   39475
      End
      Begin VB.TextBox txtStartingNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   8
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox txtEndingNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   9
         Top             =   555
         Width           =   1455
      End
      Begin VB.TextBox txtBookNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1425
      End
      Begin VB.CheckBox chkClosed 
         BackColor       =   &H80000018&
         Caption         =   "Closing the Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3600
         TabIndex        =   10
         Top             =   1080
         Width           =   1905
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2340
         TabIndex        =   12
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3390
         TabIndex        =   13
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   11
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Ending No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3375
         TabIndex        =   18
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Starting No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3270
         TabIndex        =   17
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Issued Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Book No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   15
         Top             =   630
         Width           =   780
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgODCRBook 
      Height          =   1695
      Left            =   60
      TabIndex        =   14
      Top             =   4380
      Width           =   5970
      _cx             =   10530
      _cy             =   2990
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorSel    =   -2147483635
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStockRegisterOfReceiptBooks.frx":000C
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
End
Attribute VB_Name = "frmStockRegisterOfReceiptBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Option Explicit
            Dim mCon                        As New ADODB.Connection
            Dim mCom                        As New ADODB.Command
            Dim Rec                         As New ADODB.Recordset
            Dim Rec2                        As New ADODB.Recordset
            Dim Rec3                        As New ADODB.Recordset
            Dim Rec4                        As New ADODB.Recordset
            Dim Rec5                        As New ADODB.Recordset
            Private marr1()                 As String
            Private marr2()                 As String
            Private marr3()                 As String
            Dim mEditFlag                   As Boolean

    Private Sub FormInitialize()
'            Dim mCrl As Control
'            For Each mCrl In Me.Controls
'                If TypeOf mCrl Is TextBox Then
'                    mCrl.Text = ""
'                    mCrl.Tag = ""
'                ElseIf TypeOf mCrl Is OptionButton Then
'                    mCrl.Value = False
'                ElseIf TypeOf mCrl Is ComboBox Then
'                    If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
'                ElseIf TypeOf mCrl Is ComboBox Then
'                    mCrl.ListIndex = -1
'                End If
'            Next
            mEditFlag = False
            
    End Sub
    
    Private Sub cmbBookIssueTo_click()
'        If cmbBookIssueTo.ListIndex > 0 Then
'            If cmbBookIssueTo.ListIndex = 7 Then
'                cmbCounters.Enabled = True
'            Else
'                cmbCounters.Enabled = False
'            End If
            cmbStaff.ListIndex = -1
            txtBookNo.Text = ""
            txtStartingNo.Text = ""
            txtEndingNo.Text = ""
            dtissuedDate.Value = Date
            Call PopulateReceiptBookInfo
        'End If
            
    End Sub

    Private Sub cmbCounters_Click()
'        Dim mCount As Integer
'        Dim arrIn As Variant
'        Dim Rec As New ADODB.Recordset
'        Dim mCon As New ADODB.Connection
'        Dim objDB As New clsDB
'
'        fgODCRBook.Rows = 1
'            arrIn = Array(cmbCounters.ListIndex, cmbStaffName.ListIndex)
'        Set Rec = objDB.ExecuteSP("spGetCounterReceiptBookInfo", arrIn, , , mCon, adCmdStoredProc)
'        mCount = 1
'        If Not Rec.EOF Then
'            While Not Rec.EOF
'                fgODCRBook.Rows = fgODCRBook.Rows + 1
'                'fgODCRBook.Cell(flexcpText, mCount, 0) = mCount
'                fgODCRBook.Cell(flexcpText, mCount, 1) = Rec!dtissuedDate
'                fgODCRBook.Cell(flexcpText, mCount, 2) = Rec!vchBookNo
'                fgODCRBook.Cell(flexcpText, mCount, 3) = Rec!intStartingNo
'                fgODCRBook.Cell(flexcpText, mCount, 4) = Rec!intEndingNo
'                fgODCRBook.Cell(flexcpText, mCount, 5) = Rec!numReceiptBookID
'                mCount = mCount + 1
'                Rec.MoveNext
'            Wend
'        Else
'            Exit Sub
'        End If
    End Sub

        Private Sub cmbStaff_click()
                Dim mSQL    As String
                Dim Rec     As New ADODB.Recordset
                Dim mIndex  As Long
                Dim mCount As Integer
                Dim arrIn As Variant
                Dim RecPermanent As New ADODB.Recordset
                Dim RecTemporary As New ADODB.Recordset
                Dim mCon As New ADODB.Connection
                Dim objDB As New clsDB
                cmbStaffName.Clear
                
                If cmbStaff.ListIndex = 1 Then
                    mSQL = "SELECT vchEmpName, numEmployeeID From faStaffs Where faStaffs.tnyEmpType=0"
                    Call PopulateList(cmbStaffName, mSQL, , True, , True)
                    cmbDepartment.Enabled = True
                    cmbDesignation.Enabled = True
                ElseIf cmbStaff.ListIndex = 0 Then
                    cmbDepartment.Enabled = False
                    cmbDesignation.Enabled = False
                    mSQL = "SELECT vchEmpName, numEmployeeID From faStaffs Where faStaffs.tnyEmpType=1"
                    Call PopulateList(cmbStaffName, mSQL, , True, , True)
                End If
                Exit Sub
                
                
                
                
                                        '                    fgODCRBook.Rows = 1
                                        '                    mSQL = "Select * "
                                        '                    mSQL = mSQL & " FROM faStockRegisterReceipts INNER JOIN "
                                        '                    mSQL = mSQL & " faStaffs ON faStockRegisterReceipts.numEmployeeID = faStaffs.numEmployeeID AND faStaffs.tnyEmpType = 0 "
                                        '                    mSQL = mSQL & " Where faStockRegisterReceipts.intVoucherTypeID='" & cmbBookIssueTo.ListIndex & "'"
                                        '                    mCount = 1
                                        '                    Set Rec = GetRecordSet(mSQL)
                                        '                    If Rec.RecordCount > 0 Then
                                        '                        If Not Rec.EOF Then
                                        '                            While Not Rec.EOF
                                        '                                fgODCRBook.Rows = fgODCRBook.Rows + 1
                                        '                                'fgODCRBook.Cell(flexcpText, mCount, 0) = mCount
                                        '                                fgODCRBook.Cell(flexcpText, mCount, 1) = Rec!dtissuedDate
                                        '                                fgODCRBook.Cell(flexcpText, mCount, 2) = Rec!vchBookNo
                                        '                                fgODCRBook.Cell(flexcpText, mCount, 3) = Rec!intStartingNo
                                        '                                fgODCRBook.Cell(flexcpText, mCount, 4) = Rec!intEndingNo
                                        '                                fgODCRBook.Cell(flexcpText, mCount, 5) = Rec!numReceiptBookID
                                        '                                mCount = mCount + 1
                                        '                                Rec.MoveNext
                                        '                            Wend
                                        '                        Else
                                        '                            Exit Sub
                                        '                        End If
                                        '                    End If
                                        '            End If
                                        '           cmbStaffName.ListIndex = mIndex
                                        '        'End If
                                        ' Exit Sub
        End Sub
    
        Private Sub cmbStaffName_click()
        
'            Dim mSQL As String
'            Dim mCount As Integer
'            Dim arrIn As Variant
'            Dim Rec As New ADODB.Recordset
'            Dim mCon As New ADODB.Connection
'            Dim objDB As New clsDB
'
'            fgODCRBook.Rows = 1
'            mSQL = "Select * "
'            mSQL = mSQL & " FROM faStockRegisterReceipts INNER JOIN "
'            mSQL = mSQL & " faStaffs ON faStockRegisterReceipts.numEmployeeID = faStaffs.numEmployeeID "
'            mSQL = mSQL & " Where faStockRegisterReceipts.intVoucherTypeID='" & cmbBookIssueTo.ListIndex & "'"
'            mSQL = mSQL & " AND faStockRegisterReceipts.numEmployeeID='" & cmbStaffName.ItemData(cmbStaffName.ListIndex) & "'"
'            mCount = 1
'            Set Rec = GetRecordSet(mSQL)
'            If Rec.RecordCount > 0 Then
'                If Not Rec.EOF Then
'                    While Not Rec.EOF
'                        fgODCRBook.Rows = fgODCRBook.Rows + 1
'                        'fgODCRBook.Cell(flexcpText, mCount, 0) = mCount
'                        fgODCRBook.Cell(flexcpText, mCount, 1) = Rec!dtissuedDate
'                        fgODCRBook.Cell(flexcpText, mCount, 2) = Rec!vchBookNo
'                        fgODCRBook.Cell(flexcpText, mCount, 3) = Rec!intStartingNo
'                        fgODCRBook.Cell(flexcpText, mCount, 4) = Rec!intEndingNo
'                        fgODCRBook.Cell(flexcpText, mCount, 5) = Rec!numReceiptBookID
'                        mCount = mCount + 1
'                        txtBookNo.Text = Rec!vchBookNo
'                        txtStartingNo.Text = Rec!intStartingNo
'                        txtEndingNo.Text = Rec!intEndingNo
'                        dtissuedDate.Value = Rec!dtissuedDate
'                        Rec.MoveNext
'                    Wend
'                Else
'                    Exit Sub
'                End If
'            End If
    
    End Sub

    Private Sub cmdClose_Click()
        Call cmdNew_Click
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmbStaff.ListIndex = -1
        cmbBookIssueTo.ListIndex = -1
        cmbStaffName.ListIndex = -1
        cmbCounters.ListIndex = -1
        txtBookNo.Text = ""
        txtBookNo.Tag = ""
        txtStartingNo.Text = ""
        txtEndingNo.Text = ""
        dtissuedDate.Value = Date
        chkClosed.Value = 0
    End Sub

    Private Sub cmdSave_Click()
       
            Dim mintReceiptBookID       As Double
            Dim mintVoucherTypeID       As Long
            Dim mintEmployeeID          As Double
            Dim mintCounterID           As Long
            Dim mintFinancialYearID     As Long
            Dim arrInput                As Variant
                   
            Dim objDB As New clsDB
            Set Rec = New ADODB.Recordset
            '------------------------------------------
            '   Validations
            '------------------------------------------
            If cmbBookIssueTo.ListIndex = -1 Then
                cmbBookIssueTo.SetFocus
                Exit Sub
            End If
            If cmbStaff.ListIndex = -1 Then
                cmbStaff.SetFocus
                Exit Sub
            End If
            If cmbStaffName.ListIndex = -1 Then
                cmbStaffName.SetFocus
                Exit Sub
            End If
            If cmbCounters.ListIndex = -1 Then
                cmbCounters.SetFocus
                Exit Sub
            End If
            If Trim(txtBookNo.Text) = "" Then
                txtBookNo.SetFocus
                Exit Sub
            End If
            If Trim(txtStartingNo.Text) = "" Then
                txtStartingNo.SetFocus
                Exit Sub
            End If
            If Trim(txtEndingNo.Text) = "" Then
                txtEndingNo.SetFocus
                Exit Sub
            End If
            If mEditFlag And Val(txtBookNo.Tag) < 1 Then
                MsgBox "Error: Try again!", vbInformation
                Exit Sub
            ElseIf mEditFlag And Val(txtBookNo.Tag) > 0 Then
                MsgBox "Editing an Existing Receipt Book!"
            ElseIf mEditFlag = False And Val(txtBookNo.Tag) = 0 Then
                MsgBox "Creating a new Receipt Book!"
            End If
            '------------------------------------------
            '   Saving a New ReceiptBook
            '------------------------------------------
            objDB.SetConnection mCon
            
            mintVoucherTypeID = IIf(cmbBookIssueTo.ItemData(cmbBookIssueTo.ListIndex) > -1, cmbBookIssueTo.ItemData(cmbBookIssueTo.ListIndex), 0)
            mintEmployeeID = IIf(cmbStaffName.ItemData(cmbStaffName.ListIndex) > -1, cmbStaffName.ItemData(cmbStaffName.ListIndex), 0)
            mintCounterID = IIf(cmbCounters.ItemData(cmbCounters.ListIndex) > -1, cmbCounters.ItemData(cmbCounters.ListIndex), 0)
            'mintFinancialYearID = IIf(Val(txtFinancialYear.Tag) > -1, Val(txtFinancialYear.Tag), 0)
                '------------------------------------------------------'
                                'faStockRegisterReceipts'
                '------------------------------------------------------'
                        '@numReceiptBookID      numeric            ,
                        '@intVoucherTypeID      int                ,
                        '@numEmployeeID         numeric    =Null   ,
                        '@intCounterID          int                ,
                        '@intFinancialYearID    int                ,
                        '@dtIssuedDate          smalldatetime      ,
                        '@vchBookNo             varchar            ,
                        '@intStartingNo         int                ,
                        '@intEndingNo           int                ,
                        '@tnyStatus             tinyint            ,
                        '@intLocalBodyID        int
                '------------------------------------------------------'
            arrInput = Array((IIf(mEditFlag = True, Val(txtBookNo.Tag), -1)), _
                                mintVoucherTypeID, _
                                mintEmployeeID, _
                                mintCounterID, _
                                gbFinancialYearID, _
                                DdMmmYy(dtissuedDate.Value), _
                                txtBookNo.Text, _
                                Val(txtStartingNo.Text), _
                                Val(txtEndingNo.Text), _
                                IIf(chkClosed, 1, 0), _
                                gbLocalBodyID _
                                )
                                
         
            objDB.ExecuteSP "spSaveStockRegisterOfReceipts", arrInput
            
            Call FormInitialize
    End Sub
    
    Private Sub Form_Activate()
        Me.Top = 0
        frmStockRegisterOfReceiptBooks.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    
    Private Sub Form_Load()
        Call FillcmbBookIssueTo
        Call FillcmbStaffType
        Call FillcmbCounter
        fgODCRBook.AutoSizeMode = flexAutoSizeColWidth
        fgODCRBook.AutoSize 0, fgODCRBook.Cols - 1, , True
    End Sub
    Private Sub FillcmbBookIssueTo()
            Dim mSQL    As String
            Dim Rec     As New ADODB.Recordset
            Dim mIndex  As Long
            
            mSQL = "SELECT * From faVoucherType Order By intVoucherTypeID"
            Set Rec = GetRecordSet(mSQL)
            cmbBookIssueTo.AddItem ""
            mIndex = 0
            If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                    cmbBookIssueTo.AddItem Rec!vchVoucherType
                    cmbBookIssueTo.ItemData(cmbBookIssueTo.NewIndex) = Rec!intVoucherTypeID
                    Rec.MoveNext
                Wend
            End If
            cmbBookIssueTo.ListIndex = mIndex
    End Sub
    Private Sub FillcmbStaffType()
        cmbStaff.AddItem "Temporary"
        cmbStaff.AddItem "Permanent"
    End Sub
    
    Private Sub FillcmbCounter()
            Dim mSQL    As String
            Dim Rec     As New ADODB.Recordset
            Dim mIndex  As Long
            mSQL = "SELECT * From faCounters Order By intCounterID"
            Set Rec = GetRecordSet(mSQL)
            cmbCounters.AddItem ""
            mIndex = 0
            If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                    cmbCounters.AddItem Rec!vchDescription & "    " & Rec!intCounterNo
                    cmbCounters.ItemData(cmbCounters.NewIndex) = Rec!intcounterID
                    Rec.MoveNext
                Wend
            End If
            cmbCounters.ListIndex = mIndex
    End Sub
    
    Private Sub PopulateReceiptBookInfo()
        Dim mCount As Integer
        Dim arrIn As Variant
        Dim Rec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        Dim objDB As New clsDB
        
        If cmbBookIssueTo.ListIndex > 0 Then
            If cmbBookIssueTo.ListIndex = 7 Then
                cmbCounters.Enabled = True
            Else
                cmbCounters.Enabled = False
            End If
        fgODCRBook.Rows = 1
            arrIn = Array(cmbBookIssueTo.ListIndex)
        Set Rec = objDB.ExecuteSP("spGetReceiptBookInfo", arrIn, , , mCon, adCmdStoredProc)
        mCount = 1
        'If Not Rec.BOF And Rec.EOF Then
        If Not Rec.EOF Then
            While Not Rec.EOF
                fgODCRBook.Rows = fgODCRBook.Rows + 1
                'fgODCRBook.Cell(flexcpText, mCount, 0) = mCount
                fgODCRBook.Cell(flexcpText, mCount, 1) = Rec!dtissuedDate
                fgODCRBook.Cell(flexcpText, mCount, 2) = Rec!vchBookNo
                fgODCRBook.Cell(flexcpText, mCount, 3) = Rec!intStartingNo
                fgODCRBook.Cell(flexcpText, mCount, 4) = Rec!intEndingNo
                fgODCRBook.Cell(flexcpText, mCount, 7) = Rec!numReceiptBookID
                mCount = mCount + 1
                Rec.MoveNext
            Wend
        Else
            Exit Sub
        End If
       End If
    End Sub
    Private Sub fgODCRBook_Click()
        If fgODCRBook.Rows > 1 Then
            txtBookNo.Text = fgODCRBook.TextMatrix(fgODCRBook.Row, 2)
            txtBookNo.Tag = fgODCRBook.TextMatrix(fgODCRBook.Row, 7)
        End If
        mEditFlag = True
        Call Display
    End Sub

    Private Sub Display()
        Dim objDB As New clsDB
        Dim mCon As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mCounterID As Integer
        Dim RecCounter As New ADODB.Recordset
        objDB.SetConnection mCon
        
        mSQL = "Select * "
        mSQL = mSQL & " FROM faStockRegisterReceipts LEFT JOIN "
        mSQL = mSQL & " faStaffs ON faStockRegisterReceipts.numEmployeeID = faStaffs.numEmployeeID  LEFT JOIN "
        mSQL = mSQL & " faVoucherType ON faStockRegisterReceipts.intVoucherTypeID = faVoucherType.intVoucherTypeID  LEFT JOIN "
        mSQL = mSQL & " faCounters ON faStockRegisterReceipts.intCounterID = faCounters.intCounterID "
        mSQL = mSQL & " Where faStockRegisterReceipts.numReceiptBookID='" & Val(txtBookNo.Tag) & "'"
        Set Rec = GetRecordSet(mSQL)
        
        If Rec.RecordCount > 0 Then
            If Not Rec.EOF Then
                txtBookNo.Tag = Rec!numReceiptBookID
                txtBookNo.Text = Rec!vchBookNo
                dtissuedDate = Rec!dtissuedDate
                txtStartingNo.Text = Rec!intStartingNo
                txtEndingNo.Text = Rec!intEndingNo
                    If (Rec!tnyEmpType = 1) Then
                    cmbStaff.ListIndex = 1
                    ElseIf (Rec!tnyEmpType = 0) Then
                    cmbStaff.ListIndex = 0
                    End If
                Call gSubSetComboItem(cmbStaffName, Rec!numEmployeeID)
                If Val(txtBookNo.Tag) > 0 Then
                    mEditFlag = True
                    txtBookNo.Tag = Rec!numReceiptBookID
                    txtBookNo.Text = Rec!vchBookNo
                    dtissuedDate = Rec!dtissuedDate
                    txtStartingNo.Text = Rec!intStartingNo
                    txtEndingNo.Text = Rec!intEndingNo
                        If (Rec!tnyStatus = 1) Then
                            chkClosed.Value = 1
                        Else
                            chkClosed.Value = 0
                        End If
                    
                        'mCounterID = Rec!intcounterID
                        If Not IsNull(Rec!intcounterID) Then
                         cmbCounters.ListIndex = Rec!intcounterID
                        Else
                          cmbCounters.ListIndex = -1
                        
'                        mSQL = "Select * "
'                        mSQL = mSQL & " FROM faCounters "
'                        mSQL = mSQL & " Where faCounters.intCounterID='" & mCounterID & "'"
'                        Set RecCounter = GetRecordSet(mSQL)
'                        If RecCounter.RecordCount > 0 Then
'                            If Not RecCounter.EOF Then
'                                cmbCounters.ListIndex = Rec!intcounterID
'                            End If
'                        End If
                    End If
                    Else
                        mEditFlag = False
                        txtBookNo.Tag = ""
                        txtBookNo.Text = ""
                        dtissuedDate = ""
                        txtStartingNo.Text = ""
                        txtEndingNo.Text = ""
                        chkClosed.Value = 0
                End If
            End If
    End If
    End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListOfAllotmentLetters 
   BackColor       =   &H00EBF3F3&
   Caption         =   "List of Allotment Letters"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   Icon            =   "frmListOfAllotmentLetters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpAllotmentDate 
      Height          =   315
      Left            =   6210
      TabIndex        =   6
      Top             =   645
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      _Version        =   393216
      Format          =   64815105
      CurrentDate     =   40106
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   8010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtAllotmentDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4365
      TabIndex        =   5
      Top             =   645
      Width           =   1830
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8010
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2430
      TabIndex        =   7
      Top             =   1260
      Width           =   1635
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Top             =   615
      Width           =   1830
   End
   Begin VB.TextBox txtProjectNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4365
      TabIndex        =   2
      Top             =   315
      Width           =   1830
   End
   Begin VB.TextBox txtAllotmentNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1230
      TabIndex        =   0
      Top             =   315
      Width           =   1830
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridAllotments 
      Height          =   3930
      Left            =   15
      TabIndex        =   8
      Top             =   1905
      Width           =   6960
      _cx             =   12277
      _cy             =   6932
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483634
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483634
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfAllotmentLetters.frx":1CCA
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   2
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
   Begin VB.Label lblAllotmentDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3915
      TabIndex        =   14
      Top             =   660
      Width           =   405
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7110
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblProjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6840
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   11
      Top             =   615
      Width           =   675
   End
   Begin VB.Label lblProjectNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3420
      TabIndex        =   10
      Top             =   315
      Width           =   915
   End
   Begin VB.Label lblAllotmentNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allotment No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   9
      Top             =   285
      Width           =   1125
   End
End
Attribute VB_Name = "frmListOfAllotmentLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mPreviousYearMode   As Integer
Dim mRemitBackMode      As Integer
Dim mUnAuthorizedDrawal As Integer
    Private Sub FillAllotments()
        'On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mRowCount As Integer
            
            If objDB.SetConnection(mCnn) Then
                '                mSql = "Select *,faAllotments.intID[AllotmentID],suExpenditures.intID[ExpenditureID] from faAllotments "
                '                mSql = mSql + " Inner Join SuSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
                '                mSql = mSql + " Inner Join suProjectDetails On faAllotments.numProjectID = suProjectDetails.decProjectID"
                '                mSql = mSql + " Left Join suExpenditures On faAllotments.intID = suExpenditures.intAllotmentID "
                '                mSql = mSql + " Where vchAllotmentNo Like '" & txtAllotmentNo.Text & "%'"
                '                mSql = mSql + " And chvProjectSlNo Like '" & txtProjectNo.Text & "%'"
                '                mSql = mSql + " And chvProjectNameEnglish Like '" & txtProjectName.Text & "%'"
                '                If txtAmount.Text <> "" Then
                '                    mSql = mSql + " And fltAmount = " & txtAmount.Text
                '                End If
                '                If cmbCategory.ListIndex <> -1 Then
                '                    If cmbCategory.ItemData(cmbCategory.ListIndex) <> 0 Then
                '                        mSql = mSql + " And faAllotments.intSourceID = " & cmbCategory.ItemData(cmbCategory.ListIndex)
                '                    End If
                '                End If
                '                If txtAllotmentDate.Text <> "" Then
                '                    mSql = mSql + " And dtAllotmentDate = '" & txtAllotmentDate.Text & "'"
                '                End If
                '                mSql = mSql + " Order by vchAllotmentNo"
                
                
                'SELECT * FROM faAllotments
                'LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID
                'Where IsNull(faPayOrder.tnyCancelled, 0) <> 1 And faPayOrder.intPayOrderID Is Null
                                
''                mSQL = " SELECT faAllotments.intID[AllotmentID], * FROM faAllotments"
''                mSQL = mSQL + " INNER JOIN suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
''                mSQL = mSQL + " LEFT JOIN faTransactionCategory ON faTransactionCategory.intCategoryID = faAllotments.intFundCategoryID"
''                mSQL = mSQL + " LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID "
''                mSQL = mSQL + " WHERE IsNull(faPayOrder.tnyCancelled, 0) <> 1 " 'And faPayOrder.intPayOrderID Is Null "
''                mSQL = mSQL + " AND faAllotments.intFinancialYearID = " & gbFinancialYearID & " AND vchAllotmentNo Like '" & txtAllotmentNo.Text & "%'"
''                If txtAmount.Text <> "" Then
''                    mSQL = mSQL + " And fltAmount = " & txtAmount.Text
''                End If
''                mSQL = mSQL + " Order by vchAllotmentNo"
''
''                mSQL = " SELECT vchAllotmentNo, dtAllotmentDate,vchTransactionCategory,fltAuthorizedAmt,vchProjectNo,intAllotmentID, intPayOrderID,intFinancialYearID,numProjectID FROM ("
''                mSQL = mSQL + " Select  vchAllotmentNo, dtAllotmentDate,fltAuthorizedAmt,vchProjectNo,intID intAllotmentID,intFundCategoryID,intPayOrderID,faAllotments.intFinancialYearID,numProjectID  From faAllotments"
''                mSQL = mSQL + " LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID "
''                mSQL = mSQL + " Where faPayOrder.intPayOrderID Is Null"
''                mSQL = mSQL + " Union Select   vchAllotmentNo, dtAllotmentDate,fltAuthorizedAmt,vchProjectNo,intID intAllotmentID,intFundCategoryID ,intPayOrderID,faAllotments.intFinancialYearID,numProjectID From faAllotments"
''                mSQL = mSQL + " LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID"
''                mSQL = mSQL + " WHERE tnyCancelled <> 1 AND intAllotmentID IN ("
''                mSQL = mSQL + "     Select intAllotmentID  From faAllotments"
''                mSQL = mSQL + "     LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID AND tnyCancelled = 1"
''                mSQL = mSQL + "     Where Not faPayOrder.intPayOrderID Is Null"
''                mSQL = mSQL + "     ) AND faPayOrder.intPayOrderID is NULL"
''                mSQL = mSQL + " ) A  LEFT JOIN faTransactionCategory ON faTransactionCategory.intCategoryID = A.intFundCategoryID"
''                mSQL = mSQL + "     Where A.intPayOrderID Is Null"
''                mSQL = mSQL + " AND A.intFinancialYearID = " & gbFinancialYearID & "AND vchAllotmentNo Like '%'"
''                mSQL = mSQL + " Order by vchAllotmentNo"


                If mRemitBackMode = 1 Then
                    mSQL = "Select intID AllotmentID,* From faAllotments"
                    mSQL = mSQL + " INNER JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID"
                    mSQL = mSQL + " LEFT JOIN faTransactionCategory ON faTransactionCategory.intCategoryID = faAllotments.intFundCategoryID"
                    mSQL = mSQL + " Where isnull(tnyCancelled,0) <> 1   "
                    mSQL = mSQL + " And isnull(faPayOrder.intVoucherID,0)<>0"
'                    If mUnAuthorizedDrawal = 1 Then
'                        mSQL = mSQL + " And isnull(tnyTypeID,0) =3"
'                    Else
                        mSQL = mSQL + " And isnull(tnyTypeID,0) not in (1,2)"
'                    End If
                Else
                    mSQL = "Select intID AllotmentID,* From faAllotments "
                    mSQL = mSQL + "  LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID = faAllotments.intID" & vbNewLine
                    mSQL = mSQL + "  LEFT JOIN faTransactionCategory ON faTransactionCategory.intCategoryID = faAllotments.intFundCategoryID" & vbNewLine
                    mSQL = mSQL + "  Where (faPayOrder.intPayOrderID Is Null Or tnyCancelled = 1) "
                    
'                    If mUnAuthorizedDrawal = 1 Then
'                        mSQL = mSQL + " And isnull(tnyTypeID,0) =3"
'                    Else
                        mSQL = mSQL + " And isnull(tnyTypeID,0)<>1 "
'                    End If
                End If
  
                If mPreviousYearMode = 1 Then
                    mSQL = mSQL + "  And faAllotments.intFinancialYearID = " & gbFinancialYearID - 1
                    If mUnAuthorizedDrawal = 1 Then
                        mSQL = mSQL + " And isnull(tnyTypeID,0) =3"
                    End If
                Else
                    mSQL = mSQL + "  And faAllotments.intFinancialYearID = " & gbFinancialYearID
                End If
                
                mSQL = mSQL + "  AND vchAllotmentNo Like '" & txtAllotmentNo.Text & "%'"
                mSQL = mSQL + "  And faAllotments.tnyStatus=1 And faAllotments.tnyStage=2"
                If txtAmount.Text <> "" Then
                    mSQL = mSQL + " And fltAuthorizedAmt = " & txtAmount.Text
                End If
                If txtProjectNo.Text <> "" Then
                    mSQL = mSQL + " And vchProjectNo like  '" & txtProjectNo.Text & " %'"
                End If
                If txtAllotmentDate.Text <> "" Then
                    mSQL = mSQL + " And dtRequisitionDate ='  " & txtAllotmentDate.Text & " ' "
                End If
                mSQL = mSQL + " Order by vchAllotmentNo"

                Rec.Open mSQL, mCnn
                
                mRowCount = 1
                vsGridAllotments.Rows = 2
                vsGridAllotments.Clear 1, 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGridAllotments.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    vsGridAllotments.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate)
                    vsGridAllotments.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    vsGridAllotments.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
                    vsGridAllotments.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
                    vsGridAllotments.TextMatrix(mRowCount, 5) = "" 'IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
                    vsGridAllotments.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!AllotmentID), "", Rec!AllotmentID)
                    vsGridAllotments.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                    vsGridAllotments.TextMatrix(mRowCount, 8) = "" 'IIf(IsNull(Rec!ExpenditureID), "", Rec!ExpenditureID)
                    'vsGridAllotments.Cell(flexcpFontName, mRowCount, 5) = "ML-TTRevathi"
                    mRowCount = mRowCount + 1
                    vsGridAllotments.Rows = vsGridAllotments.Rows + 1
                    Rec.MoveNext
                Wend
                Rec.Close
                
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your System Administrator", vbInformation
            End If
                    
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearch_Click()
        Call FillAllotments
    End Sub
    Private Sub dtpAllotmentDate_CloseUp()
        txtAllotmentDate.Text = CheckDateInMMM(dtpAllotmentDate.value)
    End Sub

    Private Sub Form_Load()
        dtpAllotmentDate = Now
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        mPreviousYearMode = 0
        mRemitBackMode = 0
        mUnAuthorizedDrawal = 0
    End Sub

    Private Sub vsGridAllotments_DblClick()
        On Error GoTo Err:
            If vsGridAllotments.TextMatrix(vsGridAllotments.Row, 0) <> "" Then
                If val(vsGridAllotments.TextMatrix(vsGridAllotments.Row, 8)) > 0 Then
                    MsgBox "There is already a Payment Generated against this Allotment", vbInformation
                    Exit Sub
                End If
                gbSearchCode = vsGridAllotments.TextMatrix(vsGridAllotments.Row, 0) ' Allotment NO
                gbSearchID = vsGridAllotments.TextMatrix(vsGridAllotments.Row, 6) ' Allotment ID
                gbSearchStr = vsGridAllotments.TextMatrix(vsGridAllotments.Row, 7) ' ProjectID
                Unload Me
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property
    Public Property Let RemitBackMode(mRemitData As Integer)
        mRemitBackMode = mRemitData
    End Property

    Public Property Let UnAuthorizedDrawal(mUnAuthorized As Integer)
        mUnAuthorizedDrawal = mUnAuthorized
    End Property

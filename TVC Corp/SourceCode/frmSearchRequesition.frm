VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchRequesition 
   Caption         =   "Search Requesition"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   Icon            =   "frmSearchRequesition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7005
      TabIndex        =   17
      Top             =   1275
      Width           =   1335
   End
   Begin VB.TextBox txtRequistionNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7965
      TabIndex        =   11
      Top             =   270
      Width           =   1485
   End
   Begin VB.TextBox txtDatefrom 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2385
      TabIndex        =   10
      Top             =   270
      Width           =   1500
   End
   Begin VB.TextBox txtDateTo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4710
      TabIndex        =   9
      Top             =   270
      Width           =   1500
   End
   Begin VB.TextBox txtIMPOName 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2385
      TabIndex        =   8
      Top             =   615
      Width           =   3000
   End
   Begin VB.TextBox txtSourceofFund 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2385
      TabIndex        =   7
      Top             =   1365
      Width           =   3015
   End
   Begin VB.TextBox txtCatogory 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2385
      TabIndex        =   6
      Top             =   990
      Width           =   3030
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6990
      TabIndex        =   5
      Top             =   780
      Width           =   1335
   End
   Begin VB.CommandButton cmdImplementingOfc 
      Caption         =   "..."
      Height          =   300
      Left            =   5460
      TabIndex        =   4
      Top             =   630
      Width           =   360
   End
   Begin VB.CommandButton cmdSourceOfFund 
      Caption         =   "..."
      Height          =   300
      Left            =   5460
      TabIndex        =   3
      Top             =   1395
      Width           =   360
   End
   Begin VB.CommandButton cmdCatogory 
      Caption         =   "..."
      Height          =   300
      Left            =   5460
      TabIndex        =   2
      Top             =   1005
      Width           =   360
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3150
      Left            =   15
      TabIndex        =   0
      Top             =   2040
      Width           =   9900
      _cx             =   17462
      _cy             =   5556
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchRequesition.frx":1CCA
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
   Begin VB.Label Label6 
      Caption         =   "DateTo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      TabIndex        =   16
      Top             =   300
      Width           =   780
   End
   Begin VB.Shape Shape1 
      Height          =   1740
      Left            =   60
      Top             =   180
      Width           =   9810
   End
   Begin VB.Label Label1 
      Caption         =   "Requistion NO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6420
      TabIndex        =   15
      Top             =   285
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1275
      TabIndex        =   14
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Implementing Officer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   300
      TabIndex        =   13
      Top             =   645
      Width           =   2070
   End
   Begin VB.Label Label4 
      Caption         =   "Source of Fund"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   12
      Top             =   1410
      Width           =   1560
   End
   Begin VB.Label Label5 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1425
      TabIndex        =   1
      Top             =   1035
      Width           =   945
   End
End
Attribute VB_Name = "frmSearchRequesition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearTaskID As Integer

    Private Sub cmdCatogory_Click()
                frmSearchMasters.SQLQry = "Select intCategoryID,vchTransactionCategory from  faTransactionCategory"
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.QrySP = Qyery
                frmSearchMasters.Show vbModal
                txtCatogory.Text = gbSearchStr
                txtCatogory.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchCode = -1
    End Sub
    
    Private Sub cmdClear_Click()
        txtRequistionNo.Text = ""
        txtIMPOName.Text = ""
        txtSourceOfFund.Text = ""
        txtCatogory.Text = ""
        If mPreviousYearMode Then
            txtDatefrom.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
        Else
            txtDatefrom.Text = DdMmmYy(gbEndingDate)
            txtDateTo.Text = DdMmmYy(gbEndingDate)
        End If
        Call FillGrid
    End Sub
    
    Private Sub cmdImplementingOfc_Click()
            gbSearchID = -1                                         ''  Setting the Search ID to -1
            frmSearchSubsidiaryAccountHeads.SubLedgerType = 1       ''  1. Implementing Officer
            frmSearchSubsidiaryAccountHeads.Show vbModal
            txtIMPOName.Text = gbSearchStr
            txtIMPOName.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchCode = -1
    End Sub
    
    Private Sub cmdSearch_Click()
         If Trim(txtDatefrom.Text) = "" Then
            MsgBox "From date is Mandatory"
            Exit Sub
        End If
        If Trim(txtDateTo.Text) = "" Then
            MsgBox "To date is Mandatory"
            Exit Sub
        End If
        Call FillGrid
    End Sub
    
    Private Sub cmdSourceOfFund_Click()
                frmSearchMasters.SQLQry = "Select intSourceFundID,vchSourceFundName From suSourceOfFund Where intSourceFundID in(1,3,4,16,17, 25, 26, 27, 28)"
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.QrySP = Qyery
                frmSearchMasters.Show vbModal
                txtSourceOfFund.Text = gbSearchStr
                txtSourceOfFund.Tag = gbSearchID
                gbSearchCode = -1
                gbSearchStr = ""
    End Sub
    
    Private Sub FillGrid()
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim mRow As Integer
            Dim mSql As String
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
          
            mSql = "        SELECT intID,vchRequisitionNo,dtRequisitionDate,faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID,"
            mSql = mSql + " faSubSidiaryAccountHeads.vchName,suSourceOfFund.intSourceFundID,suSourceOfFund.vchSourceFundName,"
            mSql = mSql + " faTransactionCategory.intCategoryID,faTransactionCategory.vchTransactionCategory "
            mSql = mSql + " FROM faAllotments "
            mSql = mSql + " INNER JOIN faSubSidiaryAccountHeads on faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAllotments.intImplementingOfficersID"
            mSql = mSql + " INNER JOIN suSourceOfFund on suSourceOfFund.intSourceFundID=faAllotments.intSourceID"
            mSql = mSql + " LEFT JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faAllotments.intFundCategoryID"
            mSql = mSql + " WHERE dtRequisitionDate BETWEEN '" & txtDatefrom & "' And '" & txtDateTo & "'"
            
            If mPreviousYearMode = 0 Then
                mSql = mSql + " And tnyStatus <> 2 "
            Else
                If mPreviousYearTaskID = 3 Then
                    mSql = mSql + " And tnyStatus = 0 "
                Else
                    mSql = mSql + " And tnyStatus <> 2 "
                End If
            End If

            
            If txtRequistionNo.Text <> "" Then
                  mSql = mSql + " And vchRequisitionNo=" & txtRequistionNo.Text
            End If
            If txtIMPOName.Text <> "" Then
                mSql = mSql + " And IsNull(faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID,0)=" & txtIMPOName.Tag
            End If
            If txtSourceOfFund.Text <> "" Then
                mSql = mSql + " And IsNull(suSourceOfFund.intSourceFundID,0)=" & txtSourceOfFund.Tag
            End If
            If txtCatogory.Text <> "" Then
                mSql = mSql + " And IsNull(faTransactionCategory.intCategoryID,0)=" & txtCatogory.Tag
            End If
            mSql = mSql + " Order by dtRequisitionDate "
           Rec.Open mSql, mCnn
           vsGrid.Clear
           mRow = 1
          ' vsGrid.Clear 1, 1
           vsGrid.Rows = 1
           vsGrid.TextMatrix(0, 1) = "Requisition No"
           vsGrid.TextMatrix(0, 2) = "Date"
           vsGrid.TextMatrix(0, 3) = "Implimenting Officer"
           vsGrid.TextMatrix(0, 4) = "Source of Fund"
           vsGrid.TextMatrix(0, 5) = "Category"
           
           If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                      vsGrid.Rows = vsGrid.Rows + 1
                      vsGrid.TextMatrix(mRow, 0) = IIf(IsNull(Rec!intID), "", Rec!intID)
                      vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
                      vsGrid.TextMatrix(mRow, 2) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
                      vsGrid.TextMatrix(mRow, 3) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                      vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                      vsGrid.TextMatrix(mRow, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                      Rec.MoveNext
                      mRow = mRow + 1
                Wend
               
               
           End If
    End Sub
    Private Sub Form_Load()
        
        If mPreviousYearMode Then
            txtDatefrom.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
        Else
            txtDatefrom.Text = DdMmmYy(gbTransactionDate)
            txtDateTo.Text = DdMmmYy(gbTransactionDate)
        End If
    
        'txtDatefrom.Text = DdMmmYy(Date - 31) 'DdMmmYy(gbTransactionDate)
        'txtDateTo.Text = DdMmmYy(gbTransactionDate)
        txtCatogory.Locked = True
        txtSourceOfFund.Locked = True
        txtIMPOName.Locked = True
        Call FillGrid
    End Sub

    Private Sub txtDatefrom_KeyPress(KeyAscii As Integer)
'        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
'            KeyAscii = 0
'        End If
    End Sub

    Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
'        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
'            KeyAscii = 0
'        End If
    End Sub

    Private Sub txtDateTo_LostFocus()
        
        Dim mDate As Date
        If IsDate(txtDateTo) Then
            mDate = txtDateTo
            txtDateTo.Text = DdMmmYy(mDate)
        End If
        If mPreviousYearMode Then
              If Not (mDate >= DateAdd("yyyy", -1, gbStartingDate) And mDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                txtDateTo.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            End If
        Else
            If Not (mDate >= gbStartingDate And mDate <= gbEndingDate) Then
                txtDateTo.Text = DdMmmYy(gbTransactionDate)
            End If
        End If
        
        '        If txtDateTo.Text <> "" Then
        '            If CheckDateInMMM(txtDateTo.Text) >= gbTransactionDate Then
        '                txtDateTo.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
        '            Else
        '                txtDateTo.Text = Format(CheckDateInMMM(txtDateTo.Text), "dd/mmm/yyyy")
        '            End If
        '       End If
    End Sub
    Private Sub txtDateFrom_LostFocus()
        Dim mDate As Date
        If IsDate(txtDatefrom) Then
            txtDatefrom.Text = DdMmmYy(txtDatefrom.Text)
        Else
            txtDatefrom.Text = CheckDateInMMM(txtDatefrom)
        End If
        
        
        If IsDate(txtDatefrom) Then
            mDate = txtDatefrom
        End If
        If mPreviousYearMode Then
              If Not (mDate >= DateAdd("yyyy", -1, gbStartingDate) And mDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                txtDatefrom.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
            End If
        Else
            If Not (mDate >= gbStartingDate And mDate <= gbEndingDate) Then
                txtDatefrom.Text = DdMmmYy(gbStartingDate)
            End If
        End If
        
        '        If txtDateFrom.Text <> "" Then
        '           If CheckDateInMMM(txtDateFrom.Text) >= gbTransactionDate Then
        '               txtDateFrom.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
        '           Else
        '               txtDateFrom.Text = Format(CheckDateInMMM(txtDateFrom.Text), "dd/mmm/yyyy")
        '           End If
        '        End If
        
    End Sub
    
    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then 'vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 1)
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 0)
            Unload Me
        End If
    End Sub
    
    Public Property Let PreviousYearMode(mData As Variant)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearTaskID(mData As Variant)
        mPreviousYearTaskID = mData
    End Property

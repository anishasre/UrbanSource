VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearchLetterOfAuthority 
   Caption         =   "Search Letter Of Authority"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9600
      Begin VB.TextBox txtAllotmentDate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2070
         TabIndex        =   14
         Top             =   360
         Width           =   1830
      End
      Begin VB.CommandButton cmdCatogory 
         Caption         =   "..."
         Height          =   300
         Left            =   5115
         TabIndex        =   7
         Top             =   1320
         Width           =   360
      End
      Begin VB.CommandButton cmdSourceOfFund 
         Caption         =   "..."
         Height          =   300
         Left            =   5130
         TabIndex        =   6
         Top             =   855
         Width           =   360
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
         Left            =   6690
         TabIndex        =   5
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox txtCatogory 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2070
         TabIndex        =   4
         Top             =   1305
         Width           =   3030
      End
      Begin VB.TextBox txtSourceofFund 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2070
         TabIndex        =   3
         Top             =   825
         Width           =   3015
      End
      Begin VB.TextBox txtAuthorityNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7665
         TabIndex        =   2
         Top             =   360
         Width           =   1485
      End
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
         Left            =   6705
         TabIndex        =   1
         Top             =   1305
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpAllotmentDate 
         Height          =   360
         Left            =   3915
         TabIndex        =   13
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   635
         _Version        =   393216
         Format          =   17694721
         CurrentDate     =   40106
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
         Left            =   1080
         TabIndex        =   11
         Top             =   1350
         Width           =   945
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
         Left            =   510
         TabIndex        =   10
         Top             =   870
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   900
         TabIndex        =   9
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Authority NO:"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   375
         Width           =   1530
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2745
      Left            =   45
      TabIndex        =   12
      Top             =   2025
      Width           =   9630
      _cx             =   16986
      _cy             =   4842
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchLetterOfAuthority.frx":0000
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
Attribute VB_Name = "frmSearchLetterOfAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim mvarYearID As Integer
    
    Private Sub FillGrid()
        Dim mcnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objDb As New clsDB
        Dim mRowCount As Integer
        
        If objDb.SetConnection(mcnn) Then
            mSql = " SELECT * FROM faAllotmentLetters"
            mSql = mSql + " INNER JOIN suSourceOfFund ON suSourceOfFund.intSourceFundID=faAllotmentLetters.intSourceOfFundID"
            mSql = mSql + " LEFT JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faAllotmentLetters.intCategoryID"
            mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID=faAllotmentLetters.intTransactionTypeID"
            mSql = mSql + " WHERE tnyStatus<>8 "
            
            If mvarYearID > 2011 Then
                mSql = mSql + " AND faAllotmentLetters.intFinancialYearID = " & mvarYearID
            End If
    
            If txtAuthorityNo.Text <> "" Then
                mSql = mSql + "  AND vchAllotmentNo Like '" & txtAuthorityNo.Text & "%'"
            End If
    
            If txtAllotmentDate.Text <> "" Then
                mSql = mSql + " AND dtAllotmentDate = '" & txtAllotmentDate.Text & "' "
            End If
            
            If txtSourceofFund.Text <> "" Then
                mSql = mSql + " And IsNull(suSourceOfFund.intSourceFundID,0)=" & txtSourceofFund.Tag
            End If
            If txtCatogory.Text <> "" Then
                mSql = mSql + " And IsNull(faTransactionCategory.intCategoryID,0)=" & txtCatogory.Tag
            End If
            mSql = mSql + " ORDER BY intAllotmentID"
    
            Rec.Open mSql, mcnn
            
            mRowCount = 1
            vsGrid.Rows = 2
            vsGrid.Clear 1, 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                vsGrid.TextMatrix(mRowCount, 1) = DdMmmYy(IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate))
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                mRowCount = mRowCount + 1
                vsGrid.Rows = vsGrid.Rows + 1
                Rec.MoveNext
            Wend
            Rec.Close
            
        Else
            MsgBox "Connection to Finance does not Exist, Please contact your System Administrator", vbInformation
        End If
                        
    End Sub
    
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
        txtAuthorityNo.Text = ""
        txtSourceofFund.Text = ""
        txtCatogory.Text = ""
        txtAllotmentDate.Text = ""
        Call FillGrid
    End Sub
    
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub
    
    Private Sub cmdSourceOfFund_Click()
        frmSearchMasters.SQLQry = "Select intSourceFundID,vchSourceFundName From suSourceOfFund Where intSourceFundID in(1,3,4,16,17, 25, 26, 27, 28)"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        txtSourceofFund.Text = gbSearchStr
        txtSourceofFund.Tag = gbSearchID
        gbSearchCode = -1
        gbSearchStr = ""
    End Sub
    
    Private Sub dtpAllotmentDate_CloseUp()
        txtAllotmentDate.Text = CheckDateInMMM(dtpAllotmentDate.value)
    End Sub
    
    Private Sub Form_Load()
          Call FillGrid
    End Sub
    
    Private Sub vsGrid_DblClick()
         If vsGrid.Row > 0 Then
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 0)
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 5)
            Unload Me
        End If
End Sub

    Public Property Let YearID(mData As Integer)
        mvarYearID = mData
    End Property


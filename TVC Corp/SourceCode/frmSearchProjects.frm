VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchProjects 
   BackColor       =   &H00F6FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projects"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProjectDetails 
      Height          =   6060
      Left            =   120
      ScaleHeight     =   6000
      ScaleWidth      =   10320
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   10380
      Begin VB.TextBox txtSelProjName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F8F8&
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   675
         TabIndex        =   31
         Top             =   915
         Width           =   8805
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridFund 
         Height          =   3465
         Left            =   630
         TabIndex        =   27
         Top             =   1845
         Width           =   9180
         _cx             =   16192
         _cy             =   6112
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
         BackColor       =   16054520
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16054520
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSearchProjects.frx":0000
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
      Begin VB.TextBox txtSelProjCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00EDF1F1&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   675
         TabIndex        =   26
         Top             =   540
         Width           =   4515
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8055
         TabIndex        =   30
         Top             =   5340
         Width           =   1410
      End
      Begin VB.Label lblUtilized 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6720
         TabIndex        =   29
         Top             =   5340
         Width           =   1335
      End
      Begin VB.Label lblEstimate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5355
         TabIndex        =   28
         Top             =   5340
         Width           =   1365
      End
      Begin VB.Line Line1 
         X1              =   5355
         X2              =   9465
         Y1              =   5670
         Y2              =   5670
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "PROJECT  DETAILS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   675
         TabIndex        =   25
         Top             =   285
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E3EDED&
      Height          =   3060
      Left            =   120
      ScaleHeight     =   3000
      ScaleWidth      =   10320
      TabIndex        =   1
      Top             =   4200
      Width           =   10380
      Begin VB.CommandButton cmdIMPO 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   34
         Top             =   2160
         Width           =   315
      End
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2040
         TabIndex        =   32
         Top             =   2160
         Width           =   2730
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         TabIndex        =   22
         Top             =   1425
         Width           =   3435
      End
      Begin VB.CommandButton cmdSearchFund 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   21
         Top             =   1800
         Width           =   315
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   1800
         Width           =   2730
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   2640
         Width           =   1710
      End
      Begin VB.CommandButton cmdSearchSubsector 
         Caption         =   "..."
         Height          =   300
         Left            =   8850
         TabIndex        =   17
         Top             =   1815
         Width           =   315
      End
      Begin VB.TextBox txtSubsector 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6090
         TabIndex        =   16
         Top             =   1800
         Width           =   2730
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E3EDED&
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   6585
         TabIndex        =   11
         Top             =   0
         Width           =   2250
         Begin VB.OptionButton optNew 
            Caption         =   "New"
            Height          =   240
            Left            =   315
            TabIndex        =   13
            Top             =   375
            Width           =   1515
         End
         Begin VB.OptionButton optSpillOver 
            Caption         =   "Spill Over"
            Height          =   255
            Left            =   315
            TabIndex        =   12
            Top             =   630
            Width           =   1515
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E3EDED&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   1245
         TabIndex        =   7
         Top             =   0
         Width           =   2250
         Begin VB.OptionButton optGen 
            Caption         =   "General"
            Height          =   240
            Left            =   420
            TabIndex        =   10
            Top             =   360
            Width           =   1425
         End
         Begin VB.OptionButton optSCP 
            Caption         =   "SCP"
            Height          =   240
            Left            =   420
            TabIndex        =   9
            Top             =   630
            Width           =   1425
         End
         Begin VB.OptionButton optTSP 
            Caption         =   "TSP"
            Height          =   240
            Left            =   420
            TabIndex        =   8
            Top             =   900
            Width           =   1425
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E3EDED&
         Caption         =   "Sector Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   3945
         TabIndex        =   3
         Top             =   0
         Width           =   2250
         Begin VB.OptionButton optInfrastructure 
            Caption         =   "Infrastructure"
            Height          =   240
            Left            =   420
            TabIndex        =   6
            Top             =   900
            Width           =   1440
         End
         Begin VB.OptionButton optService 
            Caption         =   "Service"
            Height          =   240
            Left            =   420
            TabIndex        =   5
            Top             =   630
            Width           =   1440
         End
         Begin VB.OptionButton optProductive 
            Caption         =   "Productive"
            Height          =   255
            Left            =   420
            TabIndex        =   4
            Top             =   345
            Width           =   1440
         End
      End
      Begin VB.TextBox txtProjNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2070
         TabIndex        =   2
         Top             =   1425
         Width           =   2730
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPO:"
         Height          =   255
         Left            =   1545
         TabIndex        =   33
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   240
         Left            =   4965
         TabIndex        =   23
         Top             =   1470
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         Height          =   255
         Left            =   1545
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subsector"
         Height          =   240
         Left            =   5355
         TabIndex        =   15
         Top             =   1845
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Sl. No:"
         Height          =   240
         Left            =   960
         TabIndex        =   14
         Top             =   1440
         Width           =   1080
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3435
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   10380
      _cx             =   18309
      _cy             =   6059
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
      BackColorFixed  =   13095120
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchProjects.frx":00D4
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
Attribute VB_Name = "frmSearchProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mProjectID As Variant
Dim mCategoryID As Variant
Dim mSourceOfFundID As Variant
Dim mSubsectorID As Variant
Dim mPreviousYearMode As Variant
Dim mSearchBy As Variant


'Dim mchvFullName As Variant
'Dim mchvDesignation As Variant


Private Sub FormInitialize()
    Dim mCrl As Control
    vsGrid.Rows = 1
    vsGrid.Rows = 15
    'vsGrid.Clear 1, 0
    For Each mCrl In Me
       If TypeOf mCrl Is TextBox Then
           mCrl.Text = ""
           mCrl.Tag = ""
       ElseIf TypeOf mCrl Is OptionButton Then
           mCrl.value = 0
       End If
    Next
    
    mProjectID = Null
    mCategoryID = Null
    mSourceOfFundID = Null
    mSubsectorID = Null
    
'   mchvFullName = Null
'   mchvDesignation = Null
    
End Sub
Private Sub FillProjectDetails(numProjID As Double)
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mWHERE As String
    Dim mLoop As Integer
    Dim mYearID As Integer
    
    Dim mEstAmt As Double
    Dim mUtiAmt As Double
    Dim mBalAmt As Double
    
    Dim arrInput As Variant
    
    mEstAmt = 0
    mUtiAmt = 0
    mBalAmt = 0
    
    If mPreviousYearMode Then
        mYearID = gbFinancialYearID - 1
    Else
        mYearID = gbFinancialYearID
    End If
        
    mSql = " SELECT * FROM ProjectDetails "
    mSql = mSql + " INNER JOIN FundDetails ON FundDetails.decProjectID = ProjectDetails.decProjectID "
    mSql = mSql + " INNER JOIN M_FundSource ON M_FundSource.intFundSrcID = FundDetails.intFundSrcID "
    'mSQL = mSQL + " INNER JOIN  SubjectCheckList ON  SubjectCheckList.decProjectID= ProjectDetails.decProjectID "
    mSql = mSql + " Where FundDetails.intYearID = " & mYearID & " And ProjectDetails.decProjectID = " & numProjID
    mSql = mSql + " Order by intSlNo"
    
    vsGridFund.Rows = 1
    If objDb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
        'Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
        'mLoop = 0
        'vsGridFund.Rows = Rec.RecordCount + 1
        
        arrInput = Array(numProjID, mYearID)
        Set Rec = objDb.ExecuteSP("FundDetails_S", arrInput, , , mCnn, adCmdStoredProc)
        mLoop = 0
        
        If Not (Rec.BOF And Rec.EOF) Then
            
            mProjectID = Rec!decProjectID
            mCategoryID = Rec!intProjCatID
            mSubsectorID = Rec!intSubSecID
            mSourceOfFundID = Null
        
'           mchvFullName = Rec!chvFullName
'           mchvDesignation = Rec!chvDesignation
        
            txtSelProjCode.Text = Rec!chvProjectSlNo
            txtSelProjName.Text = Rec!chvProjectNameEng
            vsGridFund.Rows = 1
            While Not Rec.EOF
                vsGridFund.Rows = vsGridFund.Rows + 1
                mLoop = mLoop + 1
                vsGridFund.TextMatrix(mLoop, 0) = Rec!chvCode
                vsGridFund.TextMatrix(mLoop, 1) = Rec!chvFundSourceEnglish
                vsGridFund.TextMatrix(mLoop, 2) = Rec!fltAmt
                vsGridFund.TextMatrix(mLoop, 5) = Rec!intFundSrcID
                mEstAmt = mEstAmt + Rec!fltAmt
                Rec.MoveNext
            Wend
            Rec.Close
         
            lblEstimate.Caption = Format(mEstAmt, "0.00")
            lblUtilized.Caption = ""
            lblBalance.Caption = ""
        End If
    End If
    picProjectDetails.Visible = True
End Sub

Private Sub FillGrid()
    Dim objDb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mWHERE As String
    
    Dim mLoop As Integer
    Dim mCatID As Integer
    Dim mSecTypeID As Integer
    Dim mProjTypeID As Integer
    
    Dim arrInput As Variant
    Dim mYearID As Integer
    Dim mProjectName As Variant
    Dim mProjectNo As Variant
    Dim mFundSrcID As Variant
    Dim mSubsectorID As Variant
    Dim mMircoSectorID As Variant
    Dim mIMPOID As Variant
    'Dim mCatID As Variant       ' GEN/SCP/TSP
    'Dim mMajorSecID As Variant  ' Productive/Infra/Service
    'Dim mProjTypeID As Variant  ' NEW /Spillover
    
    
    mProjectNo = 0
    mFundSrcID = 0
    mSubsectorID = 0
    mMircoSectorID = 0
    mIMPOID = 0
    vsGrid.Rows = 1
    
    ' FILTER : YEAR ID
    If mPreviousYearMode = 1 Then
         mWHERE = " WHERE ProjectDetails.intYearID = " & gbFinancialYearID - 1
         mYearID = gbFinancialYearID - 1
    Else
        mWHERE = " WHERE ProjectDetails.intYearID = " & gbFinancialYearID
        mYearID = gbFinancialYearID
    End If
   
    ' CATEGORY ID
    mCatID = 0
    If optGen.value = True Then mCatID = 1
    If optSCP.value = True Then mCatID = 2
    If optTSP.value = True Then mCatID = 3
    If mCatID > 0 Then
        mWHERE = mWHERE + " AND intProjCatID = " & mCatID
        
    End If
    
    ' SECTOR TYPE ID
    mSecTypeID = 0
    If optProductive.value = True Then mSecTypeID = 1
    If optService.value = True Then mSecTypeID = 2
    If optInfrastructure.value = True Then mSecTypeID = 3
    If mSecTypeID > 0 Then
        mWHERE = mWHERE + " AND intMajorSecID = " & mSecTypeID
    End If
    
    ' PROJECT TYPE (NEW or SPILL OVER)
    mProjTypeID = 0
    If optNew.value = True Then mProjTypeID = 1
    If optSpillOver.value = True Then mProjTypeID = 2
    If mProjTypeID > 0 Then
        mWHERE = mWHERE + " AND intProjTypeID = " & mProjTypeID
    End If
    If txtProjNo.Text <> "" Then
        mWHERE = mWHERE + " AND chvProjectSlNo like '%" & Trim(txtProjNo.Text) & "%' "
        mProjectNo = val(Trim(txtProjNo.Text))
    End If
    If txtFund.Text <> "" Then
        mWHERE = mWHERE + " AND intFundSrcID=" & val(txtFund.Tag)
        mFundSrcID = val(txtFund.Tag)
    End If
    If txtName.Text <> "" Then
        mWHERE = mWHERE + " AND chvProjectNameEng like '%" & Trim(txtName.Text) & "%' "
        mProjectName = "%" & Trim(txtName.Text) & "%"
    End If
    If txtSubsector.Text <> "" Then
        mWHERE = mWHERE + " AND intSubSecID=" & val(txtSubsector.Tag)
        mSubsectorID = val(txtSubsector.Tag)
    End If
    
    If txtIMPO.Text <> "" Then
        'mWHERE = mWHERE + " AND intSubSecID=" & val(txtSubsector.Tag)
        mIMPOID = val(txtIMPO.Tag)
    End If
    
    
    mMircoSectorID = 0
    
    mSql = " SELECT  ProjectDetails.decProjectID, ProjectDetails.intLBID, ProjectDetails.intYearID, intProjectSlNo, chvProjectSlNo, chvProjectName, "
    mSql = mSql + " chvProjectNameEng, intProjTypeID, intProjCatID, intSecID, intSubSecID, intCSSID,"
    mSql = mSql + " intSSSID , tnyConstruction, tnyMaintanance, tnyPurchase, tnyGOProject"
    mSql = mSql + " FROM ProjectDetails "
    mSql = mSql + " inner JOIN FundDetails On ProjectDetails.decProjectID=FundDetails.decProjectID"
    mSql = mSql + mWHERE
    mSql = mSql + " gROUP BY ProjectDetails.decProjectID, ProjectDetails.intLBID, ProjectDetails.intYearID, intProjectSlNo, chvProjectSlNo,"
    mSql = mSql + " chvProjectName,  chvProjectNameEng, intProjTypeID, intProjCatID, intSecID, intSubSecID, intCSSID, intSSSID ,"
    mSql = mSql + " tnyConstruction , tnyMaintanance, tnyPurchase, tnyGOProject"
    mSql = mSql + " Order By ProjectDetails.decProjectID "
    
    arrInput = Array(mYearID, mProjectName, mProjectNo, mFundSrcID, mSubsectorID, mMircoSectorID, mCatID, mSecTypeID, mProjTypeID, mIMPOID)
    
    
    If objDb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
        'Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
        
        Rec.CursorType = adOpenDynamic
        Rec.CursorLocation = adUseClient
        Set Rec = objDb.ExecuteSP("ProjectDetails_Search", arrInput, , , mCnn, adCmdStoredProc)
        mLoop = 0
        'Rec.MoveLast
        'Rec.MoveFirst
        'vsGrid.Rows = Rec.RecordCount + 1
        vsGrid.Rows = 1
        While Not Rec.EOF
            vsGrid.Rows = vsGrid.Rows + 1
            mLoop = mLoop + 1
            vsGrid.TextMatrix(mLoop, 0) = mLoop
            vsGrid.TextMatrix(mLoop, 1) = Rec!chvProjectSlNo
            vsGrid.TextMatrix(mLoop, 2) = Rec!chvProjectNameEng
            vsGrid.TextMatrix(mLoop, 3) = IIf(IsNull(Rec!EstimatedCost), "", Rec!EstimatedCost)
            vsGrid.TextMatrix(mLoop, 4) = IIf(IsNull(Rec!UtilisedCost), "", Rec!UtilisedCost)
            vsGrid.TextMatrix(mLoop, 5) = Rec!decProjectID
            Rec.MoveNext
        Wend
        Rec.Close
    End If
End Sub

Private Sub cmdIMPO_Click()
    frmSearchSubsidiaryAccountHeads.checkIMPO = 1 ' 1=Implementing Officer
    frmSearchSubsidiaryAccountHeads.SubLedgerType = 1
    frmSearchSubsidiaryAccountHeads.Show vbModal
    txtIMPO.Text = Trim(gbSearchStr)
    txtIMPO.Tag = gbSearchID
    Call FillGrid
    gbSearchStr = ""
    gbSearchID = -1
End Sub

Private Sub cmdSearch_Click()
    FillGrid
End Sub

Private Sub cmdSearchFund_Click()
    frmSearchMasters.Connection = enuSourceString.Saankhya
    If mPreviousYearMode = 1 And gbFinancialYearID = 2017 Then
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund Where intsourceFundId<>41"
    Else
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund "
    End If
    frmSearchMasters.QrySP = Qyery
    frmSearchMasters.Show vbModal
    If gbSearchID <> -1 Then
        txtFund.Text = gbSearchStr
        txtFund.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub

Private Sub cmdSearchSubsector_Click()
    frmSearchMasters.Connection = enuSourceString.Saankhya
    frmSearchMasters.SQLQry = "SELECT intSubSecID, vchSubSectorEng FROM faSubSector"
    frmSearchMasters.QrySP = Qyery
    frmSearchMasters.Show vbModal
    If gbSearchID <> -1 Then
        txtSubsector.Text = gbSearchStr
        txtSubsector.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gbSearchStr = ""
        gbSearchID = -1
        If KeyCode = vbKeyEscape Then
            If picProjectDetails.Visible Then
                picProjectDetails.Visible = False
                mSourceOfFundID = Null
            Else
                Unload Me
            End If
        ElseIf KeyCode = 13 Then
            If picProjectDetails.Visible = True Then
                Call vsGridFund_DblClick
            End If
            If IsNumeric(mSourceOfFundID) And IsNumeric(mProjectID) Then
                gbSearchStr = mProjectID
                gbSearchID = mSourceOfFundID
                Unload Me
            End If
        End If
End Sub

Private Sub Form_Load()
    FormInitialize
    FillGrid
    'FillProjectDetails (118600160081#)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mSearchBy = Nothing
End Sub

Private Sub optGen_DblClick()
    optGen.value = False
End Sub
Private Sub optInfrastructure_DblClick()
    optInfrastructure.value = False
End Sub
Private Sub optNew_DblClick()
    optNew.value = False
End Sub
Private Sub optProductive_DblClick()
    optProductive.value = False
End Sub
Private Sub optSCP_DblClick()
    optSCP.value = False
End Sub
Private Sub optService_DblClick()
    optService.value = False
End Sub
Private Sub optSpillOver_DblClick()
    optSpillOver.value = False
End Sub
Private Sub optTSP_DblClick()
    optTSP.value = False
End Sub
Private Sub vsGrid_DblClick()
    Dim mRow As Integer
    
    mRow = vsGrid.Row
On Error GoTo SkipSulekha:
    If mRow > 0 Then
        If val(vsGrid.TextMatrix(mRow, 5)) > 0 Then
            FillProjectDetails (val(vsGrid.TextMatrix(mRow, 5)))
            
            
            ' NOTE : MODIFY Based on Changes Made in SULEKHA -2013-14
            
            '            Dim objProj As New clsProject
            '            objProj.SetProject (val(vsGrid.TextMatrix(mRow, 5)))
            '
            '            Dim mCol As Collection
            '            Dim objProFund As clsProjectFund
            '            Dim mYearID As Integer
            '
            '            If mPreviousYearMode Then
            '                mYearID = gbFinancialYearID - 1
            '            Else
            '                mYearID = gbFinancialYearID
            '            End If
            '
            '            Set mCol = objProj.GetFundDetails(mYearID, objProj.ProjectID) '(2012, objProj.ProjectID)
            '            For mRow = 1 To mCol.count
            '                    Set objProFund = mCol.Item(mRow)
            '                    Debug.Print objProFund.SourceName
            '                    Debug.Print objProFund.SourceWiseAmount
            '                    Debug.Print
            '                    Set objProFund = Nothing
            '            Next mRow
            
            
        End If
        
    End If
    Exit Sub
SkipSulekha:
    frmRequisition.FundErSulekha = 1
    
    'MsgBox "Some Mistakes in Ported data from Sulekha"
End Sub

Private Sub vsGridFund_DblClick()
    ' CHECKING THE PROJECT IS SELECTED : ANY EXCEPTIONS HERE
        If Not IsNumeric(mProjectID) Then Exit Sub
        If Not IsNumeric(mCategoryID) Then Exit Sub
    ' END OF CHECKING
    
    Dim mRow As Integer
    mRow = vsGridFund.Row
    If mRow > 0 Then
        If val(vsGridFund.TextMatrix(mRow, 5)) > 0 Then
            mSourceOfFundID = val(vsGridFund.TextMatrix(mRow, 5))
            If IsNumeric(mSourceOfFundID) And IsNumeric(mProjectID) Then
                gbSearchStr = mProjectID
                gbSearchID = mSourceOfFundID
                Unload Me
            End If
            picProjectDetails.Visible = False
        End If
    End If
    
End Sub

Public Property Let ProjectID(mData As Variant)
    mProjectID = mData
End Property
Public Property Get ProjectID() As Variant
    ProjectID = mProjectID
End Property

Public Property Let CategoryID(mData As Variant)
    mCategoryID = mData
End Property
Public Property Get CategoryID() As Variant
    CategoryID = mCategoryID
End Property

Public Property Let SourceOfFundID(mData As Variant)
    mSourceOfFundID = mData
End Property
Public Property Get SourceOfFundID() As Variant
    SourceOfFundID = mSourceOfFundID
End Property

Public Property Let SubSectorID(mData As Variant)
    mSubsectorID = mData
End Property
Public Property Get SubSectorID() As Variant
    SubSectorID = mSubsectorID
End Property

Public Property Let PreviousYearMode(mData As Variant)
    mPreviousYearMode = mData
End Property


Public Property Let SearchBy(mData As Variant)
    mSearchBy = mData
End Property

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmViewBalanceSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmViewBalanceSheet"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   Icon            =   "frmViewBalanceSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   13980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   8295
      Left            =   11640
      TabIndex        =   12
      Top             =   360
      Width           =   2295
      Begin VB.Frame Frame7 
         Height          =   540
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   2085
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "####-####"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   60
            TabIndex        =   30
            Top             =   150
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "BACK"
         Height          =   345
         Left            =   150
         TabIndex        =   27
         Top             =   2130
         Width           =   1995
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "NEXT"
         Height          =   345
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   "test2"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   8010
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CommandButton Command1 
         Caption         =   "test"
         Enabled         =   0   'False
         Height          =   495
         Left            =   210
         TabIndex        =   24
         Top             =   8520
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "VERIFY"
         Height          =   345
         Left            =   150
         TabIndex        =   23
         Top             =   6810
         Width           =   1995
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "SUBMIT"
         Height          =   345
         Left            =   180
         TabIndex        =   22
         Top             =   7320
         Width           =   1995
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   90
         TabIndex        =   14
         Top             =   120
         Width           =   2085
         Begin VB.CommandButton cmdYearUp 
            Caption         =   ">>"
            Height          =   345
            Left            =   1500
            TabIndex        =   16
            Top             =   143
            Width           =   525
         End
         Begin VB.CommandButton cmdYearDown 
            Caption         =   "<<"
            Height          =   345
            Left            =   30
            TabIndex        =   15
            Top             =   143
            Width           =   525
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "####-####"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   660
            TabIndex        =   17
            Top             =   210
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "EXTRACT"
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2205
         Left            =   360
         TabIndex        =   28
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Liability"
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11625
      Begin VB.Frame Frame5 
         Height          =   405
         Left            =   210
         TabIndex        =   9
         Top             =   8670
         Width           =   7695
         Begin VB.Label Label4 
            Caption         =   "Total Liability"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   120
            Width           =   2325
         End
         Begin VB.Label Label3 
            Height          =   315
            Left            =   5340
            TabIndex        =   10
            Top             =   60
            Width           =   2325
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridLiability 
         Height          =   3315
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   11415
         _cx             =   20135
         _cy             =   5847
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
         BackColorSel    =   14271125
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15790320
         BackColorAlternate=   -2147483643
         GridColor       =   15987699
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmViewBalanceSheet.frx":1CCA
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
      Begin VB.Label lblTotal1 
         Caption         =   "Liability  Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   3630
         Width           =   4215
      End
      Begin VB.Label lblLiaTot 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """INR""#,##0.00;(""INR""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4950
         TabIndex        =   20
         Top             =   3630
         Width           =   6675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asset"
      Height          =   4185
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   11625
      Begin VB.Frame Frame4 
         Height          =   405
         Left            =   90
         TabIndex        =   6
         Top             =   8700
         Width           =   7695
         Begin VB.Label Label2 
            Height          =   315
            Left            =   5340
            TabIndex        =   8
            Top             =   60
            Width           =   2325
         End
         Begin VB.Label Label1 
            Caption         =   "Total Assets"
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   120
            Width           =   2325
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridAsset 
         Height          =   3375
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   11415
         _cx             =   20135
         _cy             =   5953
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
         BackColorSel    =   14271125
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15790320
         BackColorAlternate=   -2147483643
         GridColor       =   15987699
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
         GridLines       =   0
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
         FormatString    =   $"frmViewBalanceSheet.frx":1D9C
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
      Begin VB.Label lblTotal2 
         Caption         =   "Asset Total"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   3660
         Width           =   4095
      End
      Begin VB.Label lblAssetTot 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """INR""#,##0.00;(""INR""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   3600
         Width           =   6195
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   13950
      TabIndex        =   0
      Top             =   0
      Width           =   13980
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE SHEET"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   6780
         TabIndex        =   1
         Top             =   -30
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmViewBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mStatus     As Integer
    Dim mLBCode     As String
    Dim mMode       As Integer '1=Balance sheet,2=Income Expenditure,3=RP
    Dim mBLID       As Integer
    Dim mIEID       As Integer
    Dim mRPID       As Integer
    Dim mVerifyCount As Integer
    Dim mAVerifyCount As Integer
    Dim mLBMerge As Integer ''''' To set 1 for Merged Localbody (old lb )(Panchayath to Munc or Munc to Corp)
    
    
    Public Function GetStatus0fLFA() As Boolean
    Dim mStatus As Boolean
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim mRec As New ADODB.Recordset
    Dim objdb As New clsDB
    
    objdb.SetConnection mCnn
    
    mSql = "SELECT tnyStatus FROM faBLSubmission WHERE intYearID=" & gbFinancialYearID - 1
    mRec.Open mSql, mCnn
    GetStatus0fLFA = False
    If Not mRec.EOF Or mRec.BOF Then
        While Not mRec.EOF
            If mRec!tnyStatus = 4 Or mRec!tnyStatus = 3 Then
                GetStatus0fLFA = True
                GoTo break
            Else
                GetStatus0fLFA = False
            End If
            mRec.MoveNext
        Wend
    End If
break: mRec.Close
    mCnn.Close
    End Function
    
    Private Sub cmdBack_Click()
        If mMode = 1 Then
            mMode = 3
        ElseIf mMode = 2 Then
            mMode = 1
        ElseIf mMode = 3 Then
            mMode = 2
        ElseIf mMode = 0 Then
            mMode = 3
        End If
        Call Initialize
        Call FillBL
    End Sub
    
    Private Sub cmdGo_Click()
        If mMode = 1 Then
            Call ExtractBL
        ElseIf mMode = 2 Then
            Call ExtractIE
        ElseIf mMode = 3 Then
            Call ExtractRP
        End If
    End Sub
    Private Sub ExtractBL()
       
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim Rec1        As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mAmt        As Double
        Dim mAsOnDate   As String
        Dim mDate       As Date
        Dim mArr        As Variant
        Dim mTotAsset   As Variant
        Dim mtotLiab    As Variant
        Dim mAccHeadCode    As Variant
        Dim mAccHeadType    As Variant
        Dim mAccHeadGroup   As Variant
        Dim mSchedule       As Variant
        Dim mScheduleGroupID As Variant
        Dim mArrIN          As Variant
        Dim mArrOut          As Variant
        Dim mintID          As Integer
        Dim mDescip         As String
        Dim mSql            As String
        mTotAsset = 0
        mtotLiab = 0
            vsGridAsset.Rows = 1
            vsGridLiability.Rows = 1
            Call ClearGrid
            If (val(lblYear.Tag)) < 1 Then
                MsgBox "Please Select Finanacial Year ", vbInformation
                Exit Sub
            ElseIf val(lblYear.Tag) = gbFinancialYearID Then
                     MsgBox "Current Year AFS can't Extract " & val(lblYear.Tag) & "-" & val(lblYear.Tag) + 1, vbInformation
                     Exit Sub
                
            End If
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mAsOnDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mAsOnDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            If IsDate(mAsOnDate) Then
                mDate = CDate(mAsOnDate)
            Else
                mDate = gbEndingDate
            End If
            
           If objdb.SetConnection(mCnn) Then
                mArr = Array(mAsOnDate, gbFundID)
                Set Rec = objdb.ExecuteSP("spBLExtract", mArr, , False, mCnn, adCmdStoredProc)
                If Not (Rec.EOF And Rec.BOF) Then
                   mArrIN = Array(mBLID, gbLBID, val(lblYear.Tag), Null, Null, Null, 0, 1)
                   objdb.ExecuteSP "spSaveBLSubmission", mArrIN, mArrOut, , mCnn
                   'mIntID = mArrOut(0, 0)
                   mBLID = mArrOut(0, 0)
                   mSql = ""
                   mSql = "SELECT * FROM faBLSubmissionChild WHERE intID=" & mBLID & " AND tnyCategoryFlag=" & mMode & " Order By vchMajorAccountHeadCode Asc"
                   Rec1.Open mSql, mCnn
                   If Not Rec1.EOF Then
                        mSql = ""
                        mSql = "DELETE FROM faBLSubmissionChild WHERE intID=" & mBLID & " AND tnyCategoryFlag=" & mMode
                        mCnn.Execute mSql
                   End If
                    While Not Rec.EOF
                       Set mArrIN = Nothing
                       mAccHeadCode = IIf(IsNull(Rec!vchMajorAccountHeadCode), "", Rec!vchMajorAccountHeadCode)
                       mAmt = IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
                       mAccHeadGroup = IIf(IsNull(Rec!vchScheduleGroup), "", Rec!vchScheduleGroup)
                       mAccHeadType = IIf(IsNull(Rec!accountHeadType), "", Rec!accountHeadType)
                       mSchedule = IIf(IsNull(Rec!vchScheduleTitle), "", Rec!vchScheduleTitle)
                       mDescip = IIf(IsNull(Rec!Accounts), "", Rec!Accounts)
                       mScheduleGroupID = IIf(IsNull(Rec!intScheduleGroupID), "", Rec!intScheduleGroupID)
                       mArrIN = Array(mBLID, mAccHeadCode, mSchedule, mDescip, mAmt, mAccHeadGroup, mAccHeadType, Null, 1, mScheduleGroupID)
                       objdb.ExecuteSP "spSaveBLSubmissionChild", mArrIN, , False, mCnn
                       Rec.MoveNext
                    Wend
                End If
            End If
            lblAssetTot.Caption = mTotAsset
            lblLiaTot.Caption = mtotLiab
            Call FillBL
            'cmdGo.Enabled = False
    End Sub
    Private Sub ExtractIE()
        Dim mToDate As Date
        Dim mFromDate   As Date
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim Rec1        As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mAmt        As String
        Dim mAsOnDate   As String
        Dim mDate       As Date
        Dim mArr        As Variant
        Dim mTotAsset   As Variant
        Dim mtotLiab    As Variant
        Dim mAccHeadCode    As String
        Dim mAccHeadType    As Integer
        Dim mAccHeadGroup   As String
        Dim mSchedule       As String
        Dim mArrIN          As Variant
        Dim mArrOut          As Variant
        Dim mintID          As Integer
        Dim mDescip         As String
        Dim mSql            As String
        mTotAsset = 0
        mtotLiab = 0
        
            vsGridAsset.Rows = 1
            vsGridLiability.Rows = 1
            Call ClearGrid
            If (val(lblYear.Tag)) < 1 Then
                MsgBox "Please Select Finanacial Year ", vbInformation
                Exit Sub
            ElseIf val(lblYear.Tag) = gbFinancialYearID Then
                     MsgBox "Current Year AFS can't Extract " & val(lblYear.Tag) & "-" & val(lblYear.Tag) + 1, vbInformation
                     Exit Sub
            End If
            
            mFromDate = "01/Apr/" + CStr(val(lblYear.Tag))
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            'mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            mFromDate = CDate(mFromDate)
            mToDate = CDate(mToDate)
            
            If objdb.SetConnection(mCnn) Then
                mArr = Array(mFromDate, mToDate)
                Set Rec = objdb.ExecuteSP("spRptIncomeExpenditureExtract", mArr, , False, mCnn, adCmdStoredProc)
                If Not (Rec.EOF And Rec.BOF) Then
                   mArrIN = Array(mIEID, gbLBID, val(lblYear.Tag), Null, Null, Null, 0, 2)
                   objdb.ExecuteSP "spSaveBLSubmission", mArrIN, mArrOut, , mCnn
                   'mIntID = mArrOut(0, 0)
                   mIEID = mArrOut(0, 0)
                   mSql = ""
                   mSql = "SELECT * FROM faBLSubmissionChild WHERE intID=" & mIEID & "AND tnyCategoryFlag=" & mMode
                   Rec1.Open mSql, mCnn
                   If Not Rec1.EOF Then
                        mSql = ""
                        mSql = "DELETE FROM faBLSubmissionChild WHERE intID=" & mIEID & "AND tnyCategoryFlag=" & mMode
                        mCnn.Execute mSql
                   End If
                    While Not Rec.EOF
                       Set mArrIN = Nothing
                       mAccHeadCode = IIf(IsNull(Rec!MAjorCode), "", Rec!MAjorCode)
                       mAmt = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                       mAccHeadGroup = 0 'IIf(IsNull(Rec!vchSchedulegroup), "", Rec!vchSchedulegroup)
                       mAccHeadType = IIf(IIf(IsNull(Rec!Transactiongroup), "", Rec!Transactiongroup) = "Income", 3, 4)
                       mSchedule = IIf(IsNull(Rec!SCHEDULE), "", Rec!SCHEDULE)
                       mDescip = IIf(IsNull(Rec!Head), "", Rec!Head)
                       mArrIN = Array(mIEID, mAccHeadCode, mSchedule, mDescip, mAmt, mAccHeadGroup, mAccHeadType, Null, 2, Null)
                       objdb.ExecuteSP "spSaveBLSubmissionChild", mArrIN, , False, mCnn
                       Rec.MoveNext
                    Wend
                End If
            End If
            lblAssetTot.Caption = mTotAsset
            lblLiaTot.Caption = mtotLiab
            Call FillBL
    End Sub
    
    Private Sub ExtractRP()
        Dim mToDate As Date
        Dim mFromDate   As Date
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim Rec1        As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mAmt        As String
        Dim mAsOnDate   As String
        Dim mDate       As Date
        Dim mArr        As Variant
        Dim mTotAsset   As Variant
        Dim mtotLiab    As Variant
        Dim mAccHeadCode    As String
        Dim mAccHeadType    As Integer
        Dim mAccHeadGroup   As String
        Dim mSchedule       As String
        Dim mOperating      As Integer
        Dim mArrIN          As Variant
        Dim mArrOut          As Variant
        Dim mintID          As Integer
        Dim mDescip         As String
        Dim mSql            As String
        mTotAsset = 0
        mtotLiab = 0
        
            vsGridAsset.Rows = 1
            vsGridLiability.Rows = 1
            Call ClearGrid
            If (val(lblYear.Tag)) < 1 Then
                MsgBox "Please Select Finanacial Year ", vbInformation
                Exit Sub
            ElseIf val(lblYear.Tag) = gbFinancialYearID Then
                MsgBox "Current Year AFS can't Extract " & val(lblYear.Tag) & "-" & val(lblYear.Tag) + 1, vbInformation
                Exit Sub
            End If
            
            mFromDate = "01/Apr/" + CStr(val(lblYear.Tag))
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            'mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            mFromDate = CDate(mFromDate)
            mToDate = CDate(mToDate)
            
            If objdb.SetConnection(mCnn) Then
                mArr = Array(mFromDate, mToDate)
                Set Rec = objdb.ExecuteSP("spRptRPExtract", mArr, , False, mCnn, adCmdStoredProc)
                If Not (Rec.EOF And Rec.BOF) Then
                   mArrIN = Array(mRPID, gbLBID, val(lblYear.Tag), Null, Null, Null, 0, 3)
                   objdb.ExecuteSP "spSaveBLSubmission", mArrIN, mArrOut, , mCnn
                   'mIntID = mArrOut(0, 0)
                   mRPID = mArrOut(0, 0)
                   mSql = ""
                   mSql = "SELECT * FROM faBLSubmissionChild WHERE intID=" & mRPID & " AND tnyCategoryFlag=" & mMode & " Order By vchMajorAccountHeadCode Asc"
                   Rec1.Open mSql, mCnn
                   If Not Rec1.EOF Then
                        mSql = ""
                        mSql = "DELETE FROM faBLSubmissionChild WHERE intID=" & mRPID & " AND tnyCategoryFlag=" & mMode
                        mCnn.Execute mSql
                   End If
                    While Not Rec.EOF
                       Set mArrIN = Nothing
                       mAccHeadCode = IIf(IsNull(Rec!vchMajorAccountHeadCode), "", Rec!vchMajorAccountHeadCode)
                       mAmt = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                       mAccHeadGroup = "" 'IIf(IsNull(Rec!vchSchedulegroup), "", Rec!vchSchedulegroup)
                       If Rec!intMajSec = 1 Then
                            mAccHeadType = 5
                       Else
                            mAccHeadType = 6
                       End If
                       'mAccHeadType = IIf(IIf(IsNull(Rec!Transactiongroup), "", Rec!Transactiongroup) = "Receipt", 5, 6)
                       mSchedule = IIf(IsNull(Rec!vchScheduleCode), "", Rec!vchScheduleCode)
                       mDescip = IIf(IsNull(Rec!vchMajorAccountHead), "", Rec!vchMajorAccountHead)
                       mOperating = IIf(IsNull(Rec!intOperating), "", Rec!intOperating)
                       mArrIN = Array(mRPID, mAccHeadCode, mSchedule, mDescip, mAmt, mAccHeadGroup, mAccHeadType, mOperating, 3, Null)
                       objdb.ExecuteSP "spSaveBLSubmissionChild", mArrIN, , False, mCnn
                       Rec.MoveNext
                    Wend
                End If
            End If
            lblAssetTot.Caption = mTotAsset
            lblLiaTot.Caption = mtotLiab
            Call FillBL
    End Sub


    Private Sub cmdNext_Click()
        If mMode = 1 Then
            mMode = 2
        ElseIf mMode = 2 Then
            mMode = 3
        ElseIf mMode = 3 Then
            mMode = 1
        End If
        Call Initialize
        Call FillBL
    End Sub
    Private Sub Initialize()
        Call ClearGrid
        If mMode = 1 Then
            lblTitle.Caption = "BALANCE SHEET"
            cmdGo.Caption = "EXTRACT BL"
            Frame1.Caption = "Asset"
            Frame3.Caption = "Liability"
            lblTotal1.Caption = "Liability Total"
            lblTotal2.Caption = "Asset Total"
        ElseIf mMode = 2 Then
            lblTitle.Caption = "INCOME AND EXPENDITURE"
            cmdGo.Caption = "EXTRACT I&E"
            Frame1.Caption = "Expenditure"
            Frame3.Caption = "Income"
            lblTotal1.Caption = "Income Total"
            lblTotal2.Caption = "Expenditure Total"
        ElseIf mMode = 3 Then
            lblTitle.Caption = "RECEIPT AND PAYMENT"
            cmdGo.Caption = "EXTRACT RP"
            Frame1.Caption = "Payment"
            Frame3.Caption = "Receipt"
            lblTotal1.Caption = "Receipt Total"
            lblTotal2.Caption = "Payment Total"
        End If
        
    End Sub
    Public Function SetOldLB()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mRec    As New ADODB.Recordset
        If objdb.SetConnection(mCnn) Then
            mSql = " Select * From faLBSettings Where intLBID in (354,1132,113,1137,1138,1141,222,342,415,"
            mSql = mSql + " 478,512,535,605,692,693,728,733,816,826,839,916,917,956,964,966,976,1019,1051,1061,"
            mSql = mSql + "1066,1069,1071,1082,1128,1151,1152,1156,1169)"
            mRec.Open mSql, mCnn
            If Not (mRec.EOF And mRec.BOF) Then
                mLBMerge = 1
            Else
                mLBMerge = 0
            End If
        End If
    End Function
'''''    Function CheckProxy() As Boolean
'''''        Dim HostName As String
'''''        Dim mResult As Boolean
'''''        Dim proxyserver
'''''        Dim ProxyUser
'''''        Dim ProxyPassword
'''''        Dim SoapObject As New SoapClient30
'''''
'''''        mResult = False
'''''        HostName = "http://54.201.75.215:8080/esubmission/getAccountStatus.action"
'''''        SoapObject.MSSoapInit HostName 'Host name like you said ie:"http://yourserverdomain/yourservice.asmx?"
'''''        SoapObject.ConnectorProperty("EndPointURL") = HostName 'This was the bit I was missing!
'''''        SoapObject.ConnectorProperty("ProxyServer") = proxyserver
'''''        If ProxyUser < "" Then
'''''            SoapObject.ConnectorProperty("ProxyUser") = ProxyUser
'''''            If ProxyPassword < "" Then
'''''                SoapObject.ConnectorProperty("ProxyPassword") = ProxyPassword
'''''            End If
'''''            mResult = True
'''''        End If
'''''
'''''        ''''The variable ProxyServer would be:
'''''        ''''"<CURRENT_USER>" 'If you're using IE default settings
'''''        ''''or
'''''        ''''"http://myserver01:8080" 'If you're specifying the
'''''        ''''proxyserver: proxyport
'''''    End Function
    Private Sub cmdSubmit_Click()
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim param As String
        Dim Record1 As String
        Dim RecordSub1 As String
        Dim RecordSubT1 As String
        Dim RecordSub2 As String
        Dim RecordSubT2 As String
        Dim RecordSub3 As String
        Dim RecordSubT3 As String
        Dim RecordSub4 As String
        Dim RecordSubT4 As String
        Dim RecordSub4Ass As String
        Dim RecordSubT4Ass As String
        Dim RecordSub5 As String
        Dim RecordSubT5 As String
        Dim RecordSub6 As String
        Dim RecordSubT6 As String
        Dim RecordSub7 As String
        Dim RecordSubT7 As String
        Dim RecordSub8 As String
        Dim RecordSubT8 As String
        Dim RecordSub9 As String
        Dim RecordSubT9 As String
        Dim RecordSub10 As String
        Dim RecordSubT10 As String
        Dim RecordSub11 As String
        Dim RecordSubT11 As String
        Dim RecordT1 As String
        Dim Record2 As String
        Dim RecordT2 As String
        Dim Record3 As String
        Dim RecordT3 As String
        Dim Record4 As String
        Dim RecordT4 As String
        Dim Record5 As String
        Dim RecordT5 As String
        Dim Record6 As String
        Dim RecordT6 As String
        Dim mCnn As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
        Dim mRec1 As New ADODB.Recordset
        
        Dim objdb As New clsDB
        Dim aryIn As Variant
        Dim mCnt    As Integer
        Dim mTotAmt As Double
        Dim params
        Dim mRowCnt As Integer
        Dim Index As Integer
        Dim mResult As String
        Dim mSql As String
        Dim mtotAss As Double
        Dim mtotLiab As Double
        Dim mtotIncom As Double
        Dim mtotExpen As Double
        Dim mtotRece As Double
        Dim mtotPay As Double
        Dim mPrior As Double
        Dim mPriorRecord As String
        Dim mTransfer As Double
        Dim mTransferRecord As String
        Dim mGrossSurDef As Double
        Dim mGrossSurDefRecord As String
        Dim mGrossSurDef1 As Double
        Dim mGrossSurDef1Record As String
        Dim mNetBalance As Double
        Dim mNetBalanceRecord As String
        Dim mMessage As String
        mRec1.CursorLocation = adUseClient
        
        
            mRowCnt = 0
            If objdb.SetConnection(mCnn) Then
                aryIn = Array(val(lblYear.Tag))
                
                mSql = "SELECT * FROM faBLSubmission where tnyStatus=1 AND intYearID=" & val(lblYear.Tag)   'intLBID, intYearID, tnyStatus
                mRec1.Open mSql, mCnn
                If mRec1.RecordCount = 3 Then
                    'mRec.CursorLocation = adUseClient
                    Set mRec = objdb.ExecuteSP("spSubmitToLFA", aryIn, , False, mCnn, adCmdStoredProc)
                    'mRowCnt = mRec.RecordCount()
                    While Not mRec.EOF
                        If mRec!tnyAccHeadType = 1 Then
                            If mRec!intScheduleGroupID = 4 Then
                                RecordSub4Ass = RecordSub4Ass + mRec!String
                            ElseIf mRec!intScheduleGroupID = 5 Then
                                RecordSub5 = RecordSub5 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 6 Then
                                RecordSub6 = RecordSub6 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 7 Then
                                RecordSub7 = RecordSub7 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 8 Then
                                RecordSub8 = RecordSub8 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 9 Then
                                RecordSub9 = RecordSub9 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 10 Then
                                RecordSub10 = RecordSub10 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 11 Then
                                RecordSub11 = RecordSub11 + mRec!String
                            End If
                                'Record1 = Record1 + mRec!String
                                mtotAss = mtotAss + (mRec!fltAmount)
                        ElseIf mRec!tnyAccHeadType = 2 Then
                            If mRec!intScheduleGroupID = 1 Then
                                RecordSub1 = RecordSub1 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 2 Then
                                RecordSub2 = RecordSub2 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 3 Then
                                RecordSub3 = RecordSub3 + mRec!String
                            ElseIf mRec!intScheduleGroupID = 4 Then
                                RecordSub4 = RecordSub4 + mRec!String
                            End If
                                'Record2 = Record2 + mRec!String
                                mtotLiab = mtotLiab + (mRec!fltAmount)
                        ElseIf mRec!tnyAccHeadType = 3 Then
                            If mRec!vchMajorAccountHeadCode = "280000000" Then
                                mPrior = mRec!fltAmount
                                mPriorRecord = mRec!String
                            ElseIf mRec!vchMajorAccountHeadCode = "290000000" Then
                                mTransfer = mRec!fltAmount
                                mTransferRecord = mRec!String
                            Else
                                Record3 = Record3 + mRec!String
                                mtotIncom = mtotIncom + (mRec!fltAmount)
                            End If
                        ElseIf mRec!tnyAccHeadType = 4 Then
                            If mRec!vchMajorAccountHeadCode = "280000000" Then
                                mPrior = mRec!fltAmount
                                mPriorRecord = mRec!String
                            ElseIf mRec!vchMajorAccountHeadCode = "290000000" Then
                                mTransfer = mRec!fltAmount
                                mTransferRecord = mRec!String
                            Else
                                Record4 = Record4 + mRec!String
                                mtotExpen = mtotExpen + (mRec!fltAmount)
                            End If
                        ElseIf mRec!tnyAccHeadType = 5 Then
                                Record5 = Record5 + mRec!String
                                mtotRece = mtotRece + (mRec!fltAmount)
                        ElseIf mRec!tnyAccHeadType = 6 Then
                                Record6 = Record6 + mRec!String
                                mtotPay = mtotPay + (mRec!fltAmount)
                        End If
                        
''                        intScheduleGroupID vchScheduleGroup
''1   Reserve& Surplus
''2   Grants,Contributions for specific purposes
''3   Loans
''4   Current Liabilities And Provisions
''5   Gross Block Fixed Assets
''6   Net Block
''7   Investments
''8   Current Assets, Loans And Advances
''9   Fixed Assets
''10  Other Assets
''11  Miscellaneous Expenditure(To the Extent not written off)
                    
                        'mTotAmt = mTotAmt + mRec!fltAmount + mRec!fltAmount
                        
                        mRowCnt = mRowCnt + 1
                        mRec.MoveNext
                    Wend
                    
                    'mRowCnt = mRec.RecordCount
                    
                    mRec1.Close
                    
                    mSql = "SELECT  intScheduleGroupID,Sum (fltAmount) as fltAmount,tnyAccHeadType"
                    mSql = mSql + " FROM faBLSubmission INNER JOIN faBLSubmissionChild "
                    mSql = mSql + " ON faBLSubmission.intID=faBLSubmissionChild.intID "
                    mSql = mSql + " Where intYearID = " & val(lblYear.Tag) & " And faBLSubmission.tnyCategoryFlag = 1"
                    mSql = mSql + " Group By intScheduleGroupID,tnyAccHeadType"
                    mSql = mSql + " Order By intScheduleGroupID"

''''                    msql = "SELECT  intScheduleGroupID,Sum (fltAmount) as fltAmount"
''''                    msql = msql + " FROM faBLSubmission INNER JOIN faBLSubmissionChild"
''''                    msql = msql + " ON faBLSubmission.intID=faBLSubmissionChild.intID  Where intYearID = 2014"
''''                    msql = msql + " And faBLSubmission.tnyCategoryFlag = 1 Group By intScheduleGroupID Order By intScheduleGroupID"
''''
                    mRec1.Open mSql, mCnn

                    While Not mRec1.EOF
                        If mRec1!intScheduleGroupID = 1 Then
                            RecordSubT1 = RecordSubT1 + "&record='$L$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 2 Then
                            RecordSubT2 = RecordSubT2 + "&record='$L$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 3 Then
                            RecordSubT3 = RecordSubT3 + "&record='$L$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 4 Then
                            If mRec1!tnyAccHeadType = 2 Then
                                RecordSubT4 = RecordSubT4 + "&record='$L$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            ElseIf mRec1!tnyAccHeadType = 1 Then
                                RecordSubT4Ass = RecordSubT4Ass + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            End If
                            'RecordSubT4 = RecordSubT4 + "&record='$L$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 5 Then
                            RecordSubT5 = RecordSubT5 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 6 Then
                            RecordSubT6 = RecordSubT6 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 7 Then
                            RecordSubT7 = RecordSubT7 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 8 Then
                            RecordSubT8 = RecordSubT8 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 9 Then
                            RecordSubT9 = RecordSubT9 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 10 Then
                            RecordSubT10 = RecordSubT10 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        ElseIf mRec1!intScheduleGroupID = 11 Then
                            RecordSubT11 = RecordSubT11 + "&record='$A$T$Total$Tot$" + CStr(mRec1!fltAmount) + "$$'"
                            mRowCnt = mRowCnt + 1
                        End If
                        mRec1.MoveNext
                    Wend
                        param = ""
                        'mLBCode = "G011104"
                        params = "lbCode=" & CStr(mLBCode) + "&year=" + CStr(val(lblYear.Tag)) + "-" + CStr(val(lblYear.Tag) + 1)
                        'params = params + "&rowCount=" + CStr(mRowCnt + 6) + "&totalAmount=" + CStr(mTotAmt)
                        params = params + "&rowCount=" + CStr(mRowCnt + 8) '+ "&totalAmount=" + CStr(mTotAmt)
                        Debug.Print params
                        
                        RecordT1 = RecordT1 + "&record='$A"
                        RecordT1 = RecordT1 + "$TOT"
                        RecordT1 = RecordT1 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT1 = RecordT1 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT1 = RecordT1 + "$" + CStr(mtotAss)
                        RecordT1 = RecordT1 + "$$'"
                        
                        
                        RecordT2 = RecordT2 + "&record='$L"
                        RecordT2 = RecordT2 + "$TOT"
                        RecordT2 = RecordT2 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT2 = RecordT2 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT2 = RecordT2 + "$" + CStr(mtotLiab)
                        RecordT2 = RecordT2 + "$$'"
                        
                        RecordT3 = RecordT3 + "&record='$I"
                        RecordT3 = RecordT3 + "$TOT"
                        RecordT3 = RecordT3 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT3 = RecordT3 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT3 = RecordT3 + "$" + CStr(mtotIncom)
                        RecordT3 = RecordT3 + "$$'"
                        
                        RecordT4 = RecordT4 + "&record='$E"
                        RecordT4 = RecordT4 + "$TOT"
                        RecordT4 = RecordT4 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT4 = RecordT4 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT4 = RecordT4 + "$" + CStr(mtotExpen)
                        RecordT4 = RecordT4 + "$$'"
                        
                        RecordT5 = RecordT5 + "&record='$R"
                        RecordT5 = RecordT5 + "$TOT"
                        RecordT5 = RecordT5 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT5 = RecordT5 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT5 = RecordT5 + "$" + CStr(mtotRece)
                        RecordT5 = RecordT5 + "$$'"

                        RecordT6 = RecordT6 + "&record='$P"
                        RecordT6 = RecordT6 + "$TOT"
                        RecordT6 = RecordT6 + "$Total" '+ vsGridLiability.TextMatrix(mCnt, 2)
                        RecordT6 = RecordT6 + "$Tot" '+ vsGridLiability.TextMatrix(mCnt, 3)
                        RecordT6 = RecordT6 + "$" + CStr(mtotPay)
                        RecordT6 = RecordT6 + "$$'"
                        
                        mGrossSurDef = mtotIncom - mtotExpen
                        mGrossSurDefRecord = "&record='$E$T$Total$Tot$" + CStr(mGrossSurDef) + "$$'"
                        
                        mGrossSurDef1 = mGrossSurDef - mPrior
                        mGrossSurDef1Record = "&record='$E$T$Total$Tot$" + CStr(mGrossSurDef1) + "$$'"
                        'mNetBalance=
                        
                        params = params + RecordSub4Ass + RecordSubT4Ass
                        params = params + RecordSub5 + RecordSubT5 + RecordSub6 + RecordSubT6 + RecordSub7 + RecordSubT7 + RecordSub8
                        params = params + RecordSubT8 + RecordSub9 + RecordSubT9 + RecordSub10 + RecordSubT10 + RecordSub11 + RecordSubT11
                        params = params + RecordT1
                        params = params + RecordSub1 + RecordSubT1 + RecordSub2 + RecordSubT2 + RecordSub3 + RecordSubT3 + RecordSub4 + RecordSubT4
                        params = params + RecordT2
                        params = params + Record3
                        params = params + RecordT3
                        params = params + Record4
                        params = params + RecordT4
                        params = params + mGrossSurDefRecord
                        params = params + mPriorRecord
                        params = params + mGrossSurDef1Record
                        params = params + mTransferRecord
                        
                        params = params + Record5
                        params = params + RecordT5
                        params = params + Record6
                        params = params + RecordT6


                    
                       ' url = "http://54.201.75.215:8080/esubmission/submitAccountsFromSankhya.action"
                        If MsgBox("Are you Sure, No More Previous Year Transactions can be made?!", vbYesNo + vbDefaultButton2) = vbYes Then
                            On Error GoTo Message:
                            xmlHttp.Open "POST", "http://aims.ksad.kerala.gov.in/esubmission/submitAccountsFromSankhya.action", False
                            
                            'Call CheckProxy
                            xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                            xmlHttp.send params
                            'On Error GoTo Message:
                            'MsgBox xmlHttp.responseText
                            mResult = xmlHttp.responseText
                            If mResult = "$SUCCESS$" Then
                                If objdb.SetConnection(mCnn) Then
                                    mCnn.Execute ("Update faBLSubmission Set tnyStatus=2,dtSubmitedDate='" & Format(gbTransactionDate, "dd/MMM/YYYY") & "' Where intYearID=" & val(lblYear.Tag))
                                    MsgBox "Successfully Submitted", vbApplicationModal
                                    cmdSubmit.Enabled = False
                                Else
                                    MsgBox "Connection can not established"
                                    Exit Sub
                                End If
                            ElseIf Left(mResult, 16) = "$FAILED$rowCount" Then
                                MsgBox "Submission is Failed Due to Invalid Data, Please Check Your AFS (BL,RP and IE)Statements ", vbApplicationModal
                                Exit Sub
                            Else
                                'MsgBox mResult
                                MsgBox "CONNECTION TO WEB SERVICE CANNOT BE ESTABLISHED...KINDLY CONTACT LFA(LOCAL FUND AUDIT OFFICE)"
                                 'MsgBox "CONNECTION TO WEB SERVICE CANNOT BE ESTABLISHED..."
                            End If
                        End If
                    Else
                        MsgBox "Please Verify All Pages(BL,RP and IE)", vbApplicationModal
                        Exit Sub
                    End If
                    Call CheckSubmissionStatus
                Else
                    MsgBox "Connection Failed", vbApplicationModal
                    Exit Sub
                End If
            
            Exit Sub
Message:
        MsgBox ("Connection to the LFA Webservice can not be Established ")
    End Sub
    
'''''    Public Function CheckURL(ByVal HostAddress As String) As Boolean
'''''    CheckURL = False
'''''    Dim url As New System.Url(HostAddress)
'''''
'''''    Dim wRequest As System.Net.WebRequest
'''''
'''''    wRequest = System.Net.WebRequest.Create(url)
'''''
'''''    Dim wResponse As System.Net.WebResponse
'''''    Try
'''''        wResponse = wRequest.GetResponse()
'''''        Is the responding address the same as HostAddress to avoid false positive from an automatic redirect.
'''''        If wResponse.ResponseUri.AbsoluteUri().ToString = HostAddress Then 'include query strings
'''''            CheckURL = True
'''''        End If
'''''        wResponse.Close()
'''''        wRequest = Nothing
'''''    Catch ex As Exception
'''''        wRequest = Nothing
'''''        MsgBox (Ex.ToString)
'''''    End Try
'''''    Return CheckURL
'''''End Function
    Private Sub GlValues()
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mRec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim mSql As String
                mSql = " Select * From faBLSubmission"
                mSql = mSql + " Inner Join falbSettings On falbSettings.intLbID=faBLSubmission.intLBID Where intYearID=" & val(lblYear.Tag) & " and faBLSubmission.intLBID=" & gbLBID
                'msql = msql + " And faBLSubmission.tnyCategoryFlag=" & mMode
                If objdb.SetConnection(mCnn) Then
                    Rec.Open mSql, mCnn
                    While Not (Rec.EOF Or Rec.BOF)
                        If Rec!tnyCategoryFlag = 1 Then
                            mBLID = Rec!intID
                        ElseIf Rec!tnyCategoryFlag = 2 Then
                            mIEID = Rec!intID
                        ElseIf Rec!tnyCategoryFlag = 3 Then
                            mRPID = Rec!intID
                        End If
                        mLBCode = Rec!chvLBCode
                        Rec.MoveNext
                     Wend
                    End If
    End Sub
    Private Sub FillBL()
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mRec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim mCnt        As Integer
            Dim mAmt        As String
            Dim mAsOnDate   As String
            Dim mDate       As Date
            Dim mArr        As Variant
            Dim mTotAsset   As Variant
            Dim mtotLiab    As Variant
            Dim mSql        As String
            Dim mAccHeadCode    As String
            Dim mAccHeadType    As Integer
            Dim mAccHeadGroup   As String
            Dim mSchedule       As String
            Dim mArrIN          As Variant
            Dim mArrOut          As Variant
            Dim mintID          As Integer
            Dim mDescip         As String
            
            mTotAsset = 0
            mtotLiab = 0
            vsGridAsset.Rows = 1
            vsGridLiability.Rows = 1
            Call ClearGrid
            If (val(lblYear.Tag)) < 1 Then
'                MsgBox "Please Select Finanacial Year ", vbInformation
'                Exit Sub
                lblYear.Tag = gbFinancialYearID - 1
            End If
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mAsOnDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mAsOnDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            If IsDate(mAsOnDate) Then
                mDate = CDate(mAsOnDate)
            Else
                mDate = gbEndingDate
            End If
            
           If objdb.SetConnection(mCnn) Then
                mSql = " Select CASE When vchSchedule='RP-40(a)' then 1 "
                mSql = mSql + " When  vchSchedule='RP-40(a)' then 2 "
                mSql = mSql + " When  vchSchedule='RP-40(b)' then 7 "
                mSql = mSql + " When  vchSchedule='RP-40(b)' then 8 "
                mSql = mSql + " Else 3 end as intSecID,* From faBLSubmission "
                mSql = mSql + " Inner Join faBLSubmissionChild On faBLSubmission.intID=faBLSubmissionChild.intID "
                mSql = mSql + " Inner Join falbSettings On falbSettings.intLbID=faBLSubmission.intLBID Where intYearID=" & val(lblYear.Tag) & " and faBLSubmission.intLBID=" & gbLBID
                mSql = mSql + " And faBLSubmission.tnyCategoryFlag=" & mMode & "Order By intSecID,tnyOperating,vchMajorAccountHeadCode"
                mRec.Open mSql, mCnn
                If Not (mRec.EOF And mRec.BOF) Then
                    lblYear.Tag = mRec!intYearID
                    'vsGridLiability.Tag = mRec!intID
                    If mMode = 1 Then
                        mBLID = mRec!intID
                    ElseIf mMode = 2 Then
                        mIEID = mRec!intID
                    ElseIf mMode = 3 Then
                        mRPID = mRec!intID
                    End If

                    
                    mStatus = mRec!tnyStatus
                    
'''''                    '--------------------------------------------
'''''                 If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
'''''                    If mStatus = 0 Or mStatus = 4 Then         'Extract data Completed
'''''                            cmdVerify.Enabled = True
'''''                            cmdSubmit.Enabled = True
'''''                            cmdGO.Enabled = False
'''''                        If mMode = 1 Then
'''''                            cmdGO.Caption = "REEXTRACT BL"
'''''                        ElseIf mMode = 2 Then
'''''                            cmdGO.Caption = "REEXTRACT IE"
'''''                        ElseIf mMode = 3 Then
'''''                            cmdGO.Caption = "REEXTRACT RP"
'''''                        End If
'''''                    ElseIf mStatus = 1 Then    'Verify  Completed
'''''                        cmdGO.Enabled = False
'''''                        cmdVerify.Enabled = False
'''''                        cmdSubmit.Enabled = True
'''''                    ElseIf mStatus = 2 Then    'Submission  Completed
'''''                        cmdSubmit.Enabled = False
'''''                        cmdVerify.Enabled = False
'''''                    ElseIf mStatus = 3 Then    'Accepted by LFA
'''''                        cmdSubmit.Enabled = False
'''''                        cmdGO.Enabled = False
'''''                        cmdVerify.Enabled = False
''''''''''                    ElseIf mStatus = 4 Then     'Rejected by LFA
''''''''''                        cmdSubmit.Enabled = True
''''''''''                        cmdGo.Enabled = True
''''''''''                        cmdVerify.Enabled = True
'''''                    End If
'''''                  ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
'''''                    '--------------------------------------------


                    'Ststus = 0-Extract, 1-Verification, 2-Submission From Saankhya, 3-Submission From E-Submission, 4-Acception, 9-Rejection
                    
                    If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Or mStatus = 3 Or mStatus = 4 Or mStatus = 9 Then
                            If mMode = 1 Then
                                cmdGo.Caption = "REEXTRACT BL"
                            ElseIf mMode = 2 Then
                                cmdGo.Caption = "REEXTRACT IE"
                            ElseIf mMode = 3 Then
                                cmdGo.Caption = "REEXTRACT RP"
                            End If
                    End If
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    
                        If mStatus = 0 Or mStatus = 2 Or mStatus = 9 Then
                            cmdGo.Enabled = False
                            cmdVerify.Enabled = True
                            'cmdSubmit.Enabled = False
                        ElseIf mStatus = 1 Then
                            cmdGo.Enabled = False
                            cmdVerify.Enabled = False
                            'cmdSubmit.Enabled = True
                        ElseIf mStatus = 3 Or mStatus = 4 Then
                            cmdGo.Enabled = False
                            cmdVerify.Enabled = False
                            'cmdSubmit.Enabled = False
                        End If
                    ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                        If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Or mStatus = 9 Then
                            cmdGo.Enabled = True
                            cmdVerify.Enabled = False
                            cmdSubmit.Enabled = False

                        ElseIf mStatus = 3 Or mStatus = 4 Then
                            cmdGo.Enabled = False
                            cmdVerify.Enabled = False
                            cmdSubmit.Enabled = False
                        End If
                    ElseIf gbSeatGroupID = gbSeatGroupAuditorsGroup Then
                        If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Or mStatus = 9 Or mStatus = 3 Or mStatus = 4 Then
                            cmdGo.Enabled = False
                            cmdVerify.Enabled = False
                            cmdSubmit.Enabled = False
                        End If
                    End If
                        
                    While Not mRec.EOF
                           mLBCode = mRec!chvLBCode
                           If mRec!tnyAccHeadType = 1 Or mRec!tnyAccHeadType = 4 Or mRec!tnyAccHeadType = 6 Then
                               vsGridAsset.Rows = vsGridAsset.Rows + 1
                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = IIf(IsNull(mRec!vchHeadGroup), "", mRec!vchHeadGroup)
                               If mMode = 2 Then
                                    vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = "Expenditure"
                               ElseIf mMode = 3 Then
                                    If mRec!tnyOperating = 0 Then
                                        vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = "Operating"
                                    ElseIf mRec!tnyOperating = 1 Then
                                        vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = "NonOperating"
                                    End If
                               End If
'                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = "Receipt"
                               vsGridAsset.Row = True
                               vsGridAsset.MergeCells = flexMergeFree
                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 1) = IIf(IsNull(mRec!vchMajorAccountHeadCode), "", mRec!vchMajorAccountHeadCode)
                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 2) = IIf(IsNull(mRec!vchDescription), "", mRec!vchDescription)
                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 3) = IIf(IsNull(mRec!vchSchedule), "", mRec!vchSchedule)
                               
                               If mRec!fltAmount < 0 Then
                                    vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 4) = IIf(IsNull(mRec!fltAmount), 0, "(" & Abs(mRec!fltAmount) & ")")
                               Else
                                    vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 4) = IIf(IsNull(mRec!fltAmount), 0, mRec!fltAmount)
                               End If
                               vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 5) = IIf(IsNull(mRec!intScheduleGroupID), 0, mRec!intScheduleGroupID)
                               mTotAsset = mTotAsset + IIf(IsNull(mRec!fltAmount), 0, mRec!fltAmount)
                               
                           ElseIf mRec!tnyAccHeadType = 2 Or mRec!tnyAccHeadType = 3 Or mRec!tnyAccHeadType = 5 Then
                               vsGridLiability.MergeCells = flexMergeFree
                               vsGridLiability.Rows = vsGridLiability.Rows + 1
                               vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = IIf(IsNull(mRec!vchHeadGroup), "", mRec!vchHeadGroup)
                               If mMode = 2 Then
                                    vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = "Income"
                               ElseIf mMode = 3 Then
                                    If mRec!tnyOperating = 0 Then
                                        vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = "Operating"
                                    ElseIf mRec!tnyOperating = 1 Then
                                        vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = "NonOperating"
                                    End If
                               End If
                               vsGridLiability.Col = True
                               vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 1) = IIf(IsNull(mRec!vchMajorAccountHeadCode), "", mRec!vchMajorAccountHeadCode)
                               vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 2) = IIf(IsNull(mRec!vchDescription), "", mRec!vchDescription)
                               vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 3) = IIf(IsNull(mRec!vchSchedule), "", mRec!vchSchedule)
                               If mRec!fltAmount < 0 Then
                                    vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 4) = IIf(IsNull(mRec!fltAmount), 0, "(" & Abs(mRec!fltAmount) & ")")
                               Else
                                    vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 4) = IIf(IsNull(mRec!fltAmount), 0, mRec!fltAmount)
                               End If
                               vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 5) = IIf(IsNull(mRec!intScheduleGroupID), 0, mRec!intScheduleGroupID)
                               mtotLiab = mtotLiab + IIf(IsNull(mRec!fltAmount), 0, mRec!fltAmount)
                           End If
                           mRec.MoveNext
                      Wend
                Else
                    cmdGo.Enabled = True
                End If
                mSql = ""
                
                Dim aryIn As Variant
                Dim Rec1 As New ADODB.Recordset
                Rec1.CursorLocation = adUseClient
                aryIn = Array(val(lblYear.Tag))
                
                mSql = "SELECT * FROM faBLSubmission where tnyStatus=1 AND intYearID=" & val(lblYear.Tag)   'intLBID, intYearID, tnyStatus
                Rec1.Open mSql, mCnn
                If Rec1.RecordCount = 3 Then
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        cmdSubmit.Enabled = True
                    Else
                        cmdSubmit.Enabled = False
                    End If
                Else
                    cmdSubmit.Enabled = False
                End If
            End If
            lblAssetTot.Caption = mTotAsset
            lblLiaTot.Caption = mtotLiab
            
            
        
            
    End Sub

    Private Sub cmdVerify_Click()    'tnyStatus 1 = Verify  2=Submit 3= E-submission 4=Accept 9=Reject
        Dim mSql    As String
        Dim objdb As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
  
 
        mVerifyCount = 0
        If objdb.SetConnection(mCnn) Then
                If mMode = 1 Then
                    If val(mBLID) > 0 Then
                        If val(lblAssetTot.Caption) = val(lblLiaTot.Caption) Then
                            mSql = ""
                            mSql = "SELECT tnyStatus from faBLSubmission WHERE intID= " & mBLID
                            mRec.Open mSql, mCnn
                            If mRec!tnyStatus = 0 Or mRec!tnyStatus = 2 Or mRec!tnyStatus = 9 Then
                                mCnn.Execute ("Update faBLSubmission Set tnyStatus=1 Where intID=" & mBLID)
                                MsgBox "Balancesheet Verified Successfully", vbApplicationModal
                                Call FillBL
                                
                                Exit Sub
                            End If
                            
                            mRec.Close
                        Else
                            MsgBox "Total asset and Liability not equal", vbApplicationModal
                            Exit Sub
                        End If
                    Else
                        MsgBox "Extract Balance sheet Values", vbApplicationModal
                    End If
                ElseIf mMode = 2 Then
                    If val(mIEID) > 0 Then
                        mSql = ""
                        mSql = "SELECT tnyStatus from faBLSubmission WHERE intID= " & mIEID
                        mRec.Open mSql, mCnn
                        If mRec!tnyStatus = 0 Or mRec!tnyStatus = 2 Or mRec!tnyStatus = 9 Then
                            mCnn.Execute ("Update faBLSubmission Set tnyStatus=1 Where intID=" & mIEID)
                            MsgBox "Income Expenditure Verified Successfully", vbApplicationModal
                            cmdGo.Enabled = False
                            Call FillBL
                            cmdVerify.Enabled = False
                            Exit Sub
                        End If
                        mRec.Close
                    Else
                        MsgBox "Extract Income And Expenditure Values", vbApplicationModal
                    End If
                ElseIf mMode = 3 Then
                    If val(mRPID) > 0 Then
                        If val(lblAssetTot.Caption) = val(lblLiaTot.Caption) Then
                            mSql = ""
                            mSql = "SELECT tnyStatus from faBLSubmission WHERE intID= " & mRPID
                            mRec.Open mSql, mCnn
                            If mRec!tnyStatus = 0 Or mRec!tnyStatus = 2 Or mRec!tnyStatus = 9 Then
                                cmdVerify.Enabled = True
                                cmdGo.Enabled = True
                                mCnn.Execute ("Update faBLSubmission Set tnyStatus=1 Where intID=" & mRPID)
                                MsgBox "Receipt and Payment Verified Successfully", vbApplicationModal
                                cmdGo.Enabled = False
                                Call FillBL
                                cmdVerify.Enabled = False
                                Exit Sub
                            End If
                        Else
                            MsgBox "Total Receipt and Payment not equal", vbApplicationModal
                            Exit Sub
                        End If
                    Else
                        MsgBox "Extract Receipt and Payment Values", vbApplicationModal
                    End If
                End If

            Else
                MsgBox "Connection not established", vbApplicationModal
            End If
    End Sub

    Private Sub cmdYearDown_Click()
        Dim mYear   As String
        Dim mFYear   As Integer
        Dim mToDate As String
        mFYear = val(lblYear.Tag) - 1
        '''Enable Previos AFS To send  Only for Kannur Corp
        If gbLBID = 1259 Then
            lblYear.Tag = val(lblYear.Tag) - 1
            mYear = CStr(val(lblYear.Tag)) + "-" + CStr(val(lblYear.Tag) + 1)
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            lblDate.Caption = "01/Apr/" + CStr(val(lblYear.Tag)) + "-" + CStr(mToDate)
            lblYear.Caption = mYear
            Call Initialize
            Call ClearGrid
            Call FillBL
            cmdYearDown.Enabled = True
            cmdYearUp.Enabled = True
        End If
        '''''---------------------------------------------------
        If mFYear >= gbFinancialYearID - 1 Then
            lblYear.Tag = val(lblYear.Tag) - 1
            mYear = CStr(val(lblYear.Tag)) + "-" + CStr(val(lblYear.Tag) + 1)
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            lblDate.Caption = "01/Apr/" + CStr(val(lblYear.Tag)) + "-" + CStr(mToDate)
            lblYear.Caption = mYear
            Call Initialize
            Call ClearGrid
            Call FillBL
            cmdYearDown.Enabled = True
            cmdYearUp.Enabled = True
            'cmdGO.Enabled = True
        Else
            cmdYearDown.Enabled = False
        End If
    End Sub

    Private Sub cmdYearUp_Click()
        Dim mYear   As String
        Dim mFYear   As Integer
        Dim mToDate As String
        mFYear = val(lblYear.Tag) + 1
        If mFYear <= gbFinancialYearID Then
            lblYear.Tag = val(lblYear.Tag) + 1
            mYear = CStr(val(lblYear.Tag)) + "-" + CStr(val(lblYear.Tag) + 1)
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            lblDate.Caption = "01/Apr/" + CStr(val(lblYear.Tag)) + "-" + CStr(mToDate)
            lblYear.Caption = mYear
            Call Initialize
            Call ClearGrid
            Call FillBL
            cmdYearDown.Enabled = True
            cmdYearUp.Enabled = True
            'cmdGO.Enabled = True
        Else
            cmdYearUp.Enabled = False
        End If
            
    End Sub

    Private Sub Command1_Click()
        Dim xmlHttp As Object
         Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim Para As String
        Dim X As Integer
        
        Para = ""
        Para = "lbCode=TVM&year=2015"
        For X = 1 To 24
            Para = Para + "&ielarpItem=L&codeNo=310000000&description=Panchayat Fund&schedule=B-1&amount=" & (5000 + X)
        Next
        xmlHttp.Open "GET", "http://202.88.240.97:8084/idms/home.action?b=" & Para, False
        xmlHttp.send
        MsgBox xmlHttp.responseText
    End Sub

    Private Sub Command2_Click()
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim params
        Dim Index
        Index = 0
        params = "lbCode=D010000&year=2015-2016&rowCount=3&totalAmount=2000"
        Do While Index < 1
        params = params + "&record='$L$310000000$Panchayath Fund$B-1$250$$'"
        Index = Index + 1
        Loop
        
        params = params + "&record='$L$312000000$Reserves$B-3$750$$'"
        
        params = params + "&record='$L$T$$$1000$$'"
        
        MsgBox params
        params = "lbCode=D010000&year=2014-2015&rowCount=2&totalAmount=1000&record='$L$310000000$Panchayath Fund$B-1$500$$'&record='$L$310000000$Panchayath Fund$B-1$500$$'"
             MsgBox params
        xmlHttp.Open "POST", "http://54.201.75.215:8080/esubmission/submitAccountsFromSankhya.action", False
        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlHttp.send params
        MsgBox xmlHttp.responseText
    End Sub



    Private Sub Form_Load()
        Dim objdb As New clsDB
        Dim aryIn As Variant
        Dim mCnn As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
        Dim mSql As String
        
        mLBMerge = 0
        mMode = 1
        Call SetOldLB
        Call Initialize
        FillYear
        GlValues
        FillBL
        If objdb.SetConnection(mCnn) Then
            mSql = "SELECT * FROM faBLSubmission WHERE intYearID=" & val(lblYear.Tag) '& " AND tnyStatus > 1 "
            mRec.Open mSql, mCnn
            If Not mRec.BOF Or Not mRec.BOF Then
                   Call CheckSubmissionStatus
            End If
        End If
        
    End Sub
    
     

    Private Sub CheckSubmissionStatus()
        Dim mYear As String
       ' Dim mLBCode As Integer
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim param As String
        Dim mCnn As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
        Dim mRec1 As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim aryIn As Variant
        Dim mCnt    As Integer
        Dim mTotAmt As Double
        Dim params
        Dim mRowCnt As Integer
        Dim Index As Integer
        Dim mResult As String
        Dim mSql As String
        Dim mState As Integer
            
        params = "lbCode=" & CStr(mLBCode) + "&year=" + CStr(val(lblYear.Tag)) + "-" + CStr(val(lblYear.Tag) + 1)
        'params = CStr(mLBCode) + CStr(val(lblyear.Tag))
        'xmlHttp.Open "POST", "http://aims.lfa.kerala.gov.in/esubmission/get AccountStatus.action", False
        
''''        If CheckProxy Then
''''            xmlHttp.Open "POST", "http://54.201.75.215:8080/esubmission/getAccountStatus.action", False
''''        Else
''''            xmlHttp.Open "POST", "http://54.201.75.215:8080/esubmission/getAccountStatus.action", False
''''        End If
        On Error GoTo Message:
            xmlHttp.Open "POST", "http://aims.ksad.kerala.gov.in/esubmission/getAccountStatus.action", False
            xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            xmlHttp.send params
            
            'MsgBox xmlHttp.responseText
            mResult = xmlHttp.responseText
            If mResult = "$NotSubmitted$$" Then
                mState = 0
            ElseIf mResult = "$SubmittedFromSankhya$$" Then
                mState = 1
            ElseIf mResult = "$Submitted$$" Then
                mState = 2
            ElseIf mResult = "$Accepted$$" Then
                mState = 3
            ElseIf Left(mResult, 10) = "$Rejected$" Then
                mState = 4
            End If
''        Else
''            MsgBox ("Connection to the LFA Webservice can not be Established ")
        
        If objdb.SetConnection(mCnn) Then
            If mState = 0 Then
                lblMsg.Caption = "AFS is not submitted from Saankhya"
            ElseIf mState = 1 Then
                lblMsg.Caption = "AFS is submitted from Saankhya"
            ElseIf mState = 2 Then
                lblMsg.Caption = "AFS is submitted from E-submission"
                mCnn.Execute ("UPDATE faBLSubmission SET tnyStatus=3 WHERE intYearID=" & val(lblYear.Tag))
            ElseIf mState = 3 Then
                mCnn.Execute ("UPDATE faBLSubmission SET tnyStatus=4 WHERE intYearID=" & val(lblYear.Tag))
                lblMsg.Caption = "AFS is Accepted by LFA"
            ElseIf mState = 4 Then
                mCnn.Execute ("UPDATE faBLSubmission SET tnyStatus=9 WHERE intYearID=" & val(lblYear.Tag))
                lblMsg.Caption = "AFS is Rejected by LFA"
                cmdGo.Enabled = True
            Else
                'MsgBox mResult
                'MsgBox "CONNECTION TO WEB SERVICE CANNOT BE ESTABLISHED...KINDLY CONTACT LFA(LOCAL FUND AUDIT OFFICE)"
                lblMsg.Caption = " "
            End If
        End If
       Exit Sub
Message:
        MsgBox ("Connection To The LFA Webservice Can Not be Established Now, Try Again !!!")
        
       
'        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
'            cmdSubmit.Enabled = True
'        ElseIf gbSeatGroupID = gbSeatGroupSecretary Then
'            cmdVerify.Enabled = True
'        End If
    End Sub
    
    Private Sub FillYear()
        Dim mYear   As String
        Dim mFYear  As Integer
        Dim mDate   As String
        Dim mToDate As String
        mFYear = gbFinancialYearID - 1
        lblYear.Tag = mFYear 'gbFinancialYearID
        mYear = CStr(mFYear) + "-" + CStr(mFYear + 1)
        lblYear.Caption = mYear

        If mLBMerge = 1 And lblYear.Tag = 2015 Then
            mToDate = "31/Oct/" + CStr(val(lblYear.Tag))
        Else
            mToDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
        End If
        lblDate.Caption = "01/Apr/" + CStr(val(lblYear.Tag)) + "-" + CStr(mToDate)
    End Sub
    Private Sub ClearGrid()
        vsGridAsset.Clear 1, 0
        vsGridLiability.Clear 1, 0
        lblAssetTot.Caption = ""
        lblLiaTot.Caption = ""
    End Sub
    
     Private Sub FillGrid()
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mRec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim mCnt        As Integer
            Dim mAmt        As String
            Dim mAsOnDate   As String
            Dim mDate       As Date
            Dim mArr        As Variant
            Dim mTotAsset   As Variant
            Dim mtotLiab    As Variant
            Dim mSql        As String
            Dim mAccHeadCode    As String
            Dim mAccHeadType    As Integer
            Dim mAccHeadGroup   As String
            Dim mSchedule       As String
            Dim mArrIN          As Variant
            Dim mArrOut          As Variant
            Dim mintID          As Integer
            Dim mDescip         As String
            mTotAsset = 0
            mtotLiab = 0
            vsGridAsset.Rows = 1
            vsGridLiability.Rows = 1
            Call ClearGrid
            If (val(lblYear.Tag)) < 1 Then
                MsgBox "Please Select Finanacial Year ", vbInformation
                Exit Sub
            End If
            If mLBMerge = 1 And lblYear.Tag = 2015 Then
                mAsOnDate = "31/Oct/" + CStr(val(lblYear.Tag))
            Else
                mAsOnDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
            End If
            'mAsOnDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
           
            If IsDate(mAsOnDate) Then
                mDate = CDate(mAsOnDate)
            Else
                mDate = gbEndingDate
            End If
            
           If objdb.SetConnection(mCnn) Then
                mSql = " Select * From faBLSubmission"
                mSql = mSql + " Inner Join faBLSubmissionChild On faBLSubmission.intID=faBLSubmissionChild.intID Where intYearID=" & val(lblYear.Tag) & "and intLBID=" & gbLBID
                mRec.Open mSql, mCnn
                If Not (mRec.EOF And mRec.BOF) Then
                    Call FillBL
                Else
                     mArr = Array(mAsOnDate, gbFundID)
                     Set Rec = objdb.ExecuteSP("spBLExtract", mArr, , False, mCnn, adCmdStoredProc)
                     If Not (Rec.EOF And Rec.BOF) Then
                        
                        mArrIN = Array(Null, gbLBID, val(lblYear.Tag), Null, Null, Null, 0)
                        objdb.ExecuteSP "spSaveBLSubmission", mArrIN, mArrOut, , mCnn
                        mintID = mArrOut(0, 0)

                         While Not Rec.EOF
                            Set mArrIN = Nothing
                            mAccHeadCode = IIf(IsNull(Rec!vchMajorAccountHeadCode), "", Rec!vchMajorAccountHeadCode)
                            mAmt = IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
                            mAccHeadGroup = IIf(IsNull(Rec!vchScheduleGroup), "", Rec!vchScheduleGroup)
                            mAccHeadType = IIf(IsNull(Rec!accountHeadType), "", Rec!accountHeadType)
                            mSchedule = IIf(IsNull(Rec!vchScheduleTitle), "", Rec!vchScheduleTitle)
                            mDescip = IIf(IsNull(Rec!Accounts), "", Rec!Accounts)
                            mArrIN = Array(mintID, mAccHeadCode, mSchedule, mDescip, mAmt, mAccHeadGroup, mAccHeadType)
                            objdb.ExecuteSP "spSaveBLSubmissionChild", mArrIN, , False, mCnn
                            Rec.MoveNext
                         Wend
                     End If
                 End If
             lblAssetTot.Caption = mTotAsset
             lblLiaTot.Caption = mtotLiab
                End If
             
    End Sub
'''
'''    Private Sub FillGrid()
'''            Dim objDB       As New clsDB
'''            Dim Rec         As New ADODB.Recordset
'''            Dim mRec         As New ADODB.Recordset
'''            Dim mCnn        As New ADODB.Connection
'''            Dim mCnt        As Integer
'''            Dim mAmt        As String
'''            Dim mAsOnDate   As String
'''            Dim mDate       As Date
'''            Dim mArr        As Variant
'''            Dim mTotAsset   As Variant
'''            Dim mTotLiab    As Variant
'''            Dim mSql        As String
'''
'''            mTotAsset = 0
'''            mTotLiab = 0
'''            vsGridAsset.Rows = 1
'''            vsGridLiability.Rows = 1
'''           Call ClearGrid
'''            If (val(lblYear.Tag)) < 1 Then
'''                MsgBox "Please Select Finanacial Year ", vbInformation
'''                Exit Sub
'''            End If
'''
'''           mAsOnDate = "31/Mar/" + CStr(val(lblYear.Tag) + 1)
'''
'''
'''           If IsDate(mAsOnDate) Then
'''            mDate = CDate(mAsOnDate)
'''           Else
'''            mDate = gbEndingDate
'''           End If
'''           If objDB.SetConnection(mCnn) Then
'''                mSql = "Select * From faBLSubmission Where intYearID=" & val(lblYear.Tag) & "and intLBID=" & gbLBID
'''                mRec.Open mSql, mCnn
'''                If Not (mRec.EOF And mRec.BOF) Then
'''                 While Not mRec.EOF
'''                             If mRec!AccountHeadCode = "ASSETS" Then
'''                                 vsGridAsset.Rows = vsGridAsset.Rows + 1
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = IIf(IsNull(mRec!vchSchedulegroup), "", mRec!vchSchedulegroup)
'''                                 vsGridAsset.Row = True
'''                                 vsGridAsset.MergeCells = flexMergeFree
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 1) = IIf(IsNull(mRec!vchmajoraccountheadcode), "", mRec!vchmajoraccountheadcode)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 2) = IIf(IsNull(mRec!accounts), "", mRec!accounts)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 3) = IIf(IsNull(mRec!vchscheduletitle), "", mRec!vchscheduletitle)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 4) = IIf(IsNull(mRec!transactionamount), 0, mRec!transactionamount)
'''                                 mTotAsset = mTotAsset + IIf(IsNull(mRec!transactionamount), 0, mRec!transactionamount)
'''                             Else
'''                                 vsGridLiability.MergeCells = flexMergeFree
'''                                 vsGridLiability.Rows = vsGridLiability.Rows + 1
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = IIf(IsNull(mRec!vchSchedulegroup), "", mRec!vchSchedulegroup)
'''                                 vsGridLiability.Col = True
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 1) = IIf(IsNull(mRec!vchmajoraccountheadcode), "", mRec!vchmajoraccountheadcode)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 2) = IIf(IsNull(mRec!accounts), "", mRec!accounts)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 3) = IIf(IsNull(mRec!vchscheduletitle), "", mRec!vchscheduletitle)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 4) = IIf(IsNull(mRec!transactionamount), 0, mRec!transactionamount)
'''                                 mTotLiab = mTotLiab + IIf(IsNull(mRec!transactionamount), 0, mRec!transactionamount)
'''                             End If
'''                             Rec.MoveNext
'''                         Wend
'''                Else
'''                     mArr = Array(mAsOnDate, gbFundID)
'''                     Set Rec = objDB.ExecuteSP("spBLExtract", mArr, , False, mCnn, adCmdStoredProc)
'''                     If Not (Rec.EOF And Rec.BOF) Then
'''                        mSql = "INSERT INTO faBLSubmission Values() "
'''                         While Not Rec.EOF
'''                              If Rec!AccountHeadCode = "ASSETS" Then
'''                                 vsGridAsset.Rows = vsGridAsset.Rows + 1
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 0) = IIf(IsNull(Rec!vchSchedulegroup), "", Rec!vchSchedulegroup)
'''                                 vsGridAsset.Row = True
'''                                 vsGridAsset.MergeCells = flexMergeFree
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 1) = IIf(IsNull(Rec!vchmajoraccountheadcode), "", Rec!vchmajoraccountheadcode)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 2) = IIf(IsNull(Rec!accounts), "", Rec!accounts)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 3) = IIf(IsNull(Rec!vchscheduletitle), "", Rec!vchscheduletitle)
'''                                 vsGridAsset.TextMatrix(vsGridAsset.Rows - 1, 4) = IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
'''                                 mTotAsset = mTotAsset + IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
'''                             Else
'''                                 vsGridLiability.MergeCells = flexMergeFree
'''                                 vsGridLiability.Rows = vsGridLiability.Rows + 1
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 0) = IIf(IsNull(Rec!vchSchedulegroup), "", Rec!vchSchedulegroup)
'''                                 vsGridLiability.Col = True
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 1) = IIf(IsNull(Rec!vchmajoraccountheadcode), "", Rec!vchmajoraccountheadcode)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 2) = IIf(IsNull(Rec!accounts), "", Rec!accounts)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 3) = IIf(IsNull(Rec!vchscheduletitle), "", Rec!vchscheduletitle)
'''                                 vsGridLiability.TextMatrix(vsGridLiability.Rows - 1, 4) = IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
'''                                 mTotLiab = mTotLiab + IIf(IsNull(Rec!transactionamount), 0, Rec!transactionamount)
'''                             End If
'''                             Rec.MoveNext
'''                         Wend
'''                     End If
'''                 End If
'''                End If
'''             lblAssetTot.Caption = mTotAsset
'''             lblLiaTot.Caption = mTotLiab
'''    End Sub


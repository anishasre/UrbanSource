VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmExtractedBalanceSheet 
   BorderStyle     =   0  'None
   Caption         =   "Extracted Balance Sheet"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13350
   Icon            =   "frmExtractedBalanceSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   0
      TabIndex        =   12
      Top             =   45
      Width           =   12915
      Begin VB.TextBox txtFinancialYear 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1530
         TabIndex        =   14
         Top             =   360
         Width           =   2580
      End
      Begin VB.TextBox txtClosingBalance 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9675
         TabIndex        =   13
         Top             =   360
         Width           =   2580
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Balance As On"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7425
         TabIndex        =   15
         Top             =   360
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   60
      TabIndex        =   5
      Top             =   6165
      Width           =   12900
      Begin VB.CommandButton cmdVerify 
         Caption         =   "VERIFY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7110
         TabIndex        =   19
         Top             =   585
         Width           =   1590
      End
      Begin VB.CommandButton cmdPublish 
         Caption         =   "PUBLISH"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8880
         TabIndex        =   18
         Top             =   585
         Width           =   1590
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10650
         TabIndex        =   17
         Top             =   585
         Width           =   1590
      End
      Begin VB.Frame Frame2 
         Caption         =   "COUNCIL RESOLUTION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   4875
         Begin VB.TextBox txtCouncilNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   855
            TabIndex        =   8
            Top             =   315
            Width           =   1590
         End
         Begin VB.TextBox txtCouncilDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3105
            TabIndex        =   7
            Top             =   315
            Width           =   1590
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2250
            TabIndex        =   10
            Top             =   315
            Width           =   690
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   315
            Width           =   465
         End
      End
   End
   Begin VB.TextBox txtLiabiliityTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   11010
      TabIndex        =   4
      Top             =   5835
      Width           =   1530
   End
   Begin VB.TextBox txtAssestTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4680
      TabIndex        =   2
      Top             =   5835
      Width           =   1545
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridLiability 
      Height          =   4785
      Left            =   45
      TabIndex        =   0
      Top             =   990
      Width           =   6450
      _cx             =   11377
      _cy             =   8440
      Appearance      =   1
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
      Rows            =   20
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExtractedBalanceSheet.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
   Begin VSFlex8LCtl.VSFlexGrid vsGridAssets 
      Height          =   4785
      Left            =   6555
      TabIndex        =   11
      Top             =   990
      Width           =   6360
      _cx             =   11218
      _cy             =   8440
      Appearance      =   1
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
      Rows            =   20
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExtractedBalanceSheet.frx":1D7E
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
   Begin VB.Label Label2 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4140
      TabIndex        =   3
      Top             =   5895
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10470
      TabIndex        =   1
      Top             =   5895
      Width           =   510
   End
End
Attribute VB_Name = "frmExtractedBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mYearID As Integer
Dim mMonthID As Integer
Dim dtClosingDate As Variant
Dim mLoadMode As Integer ' 1-Yearly 2-Monthly

Private Sub cmdClose_Click()
    If mLoadMode = 1 Then
        frmExtractYearWiseList.FillGrid
        frmExtractYearWiseList.frmProgressBar.Visible = False
    Else
        frmExtactedMonthWiseList.FillGrid
    End If
    Unload Me
End Sub

Private Sub cmdPublish_Click()
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objDB   As New clsDB
    Dim mArrIn  As Variant
    Dim mSQL    As String
    Dim MthID   As String
    
    MthID = IIf(mMonthID = 0, 3, mMonthID)
    
    If objDB.SetConnection(mCnn) Then
    
        If Trim(txtCouncilNo.Text) = "" Then
            MsgBox "Enter Council No", vbInformation, "Saankhya"
            Exit Sub
        End If
        If Trim(txtCouncilDate.Text) = "" Then
            MsgBox "Enter Council Date", vbInformation, "Saankhya"
            Exit Sub
        End If
   
        mSQL = "UPDATE faPostingIndex SET vchRefNo='" & Trim(txtCouncilNo.Text) & "',dtRefDate='" & Trim(txtCouncilDate.Text) & "',tnyStage=2 WHERE intFinYearID=" & mYearID & " AND intMonthID=" & MthID & " AND tnyVerifyCash=1  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        MsgBox "Verified Successfully!", vbInformation, "Saankhya"
        cmdVerify.Enabled = False
        cmdPublish.Enabled = False
        
        If mMonthID = 3 Then
            mArrIn = Array(mYearID + 1)
            objDB.ExecuteSP "spDailyExtract_Opening", mArrIn, , , mCnn, adCmdStoredProc
        End If
        
        mSQL = "SELECT MAX(dtPostingDate) dtPostingDate FROM faPostingIndex WHERE tnyStage=2"
        Set Rec = GetRecordSet(mSQL)
        If Not (Rec.BOF And Rec.EOF) Then
            gbLastPostingDate = Format(Rec!dtPostingDate, "dd-mmm-yyyy")
        End If
        Rec.Close
        
    End If
    
End Sub

Private Sub cmdVerify_Click()
    Dim mCnn    As New ADODB.Connection
    Dim objDB   As New clsDB
    Dim mArrIn  As Variant
    Dim mSQL    As String
    Dim MthID As String
    
    MthID = IIf(mMonthID = 0, 3, mMonthID)
    If val(txtAssestTotal.Text) <> val(txtLiabiliityTotal.Text) Then
        MsgBox "Total Assets And Liability Not Equal!!!!!!!!!", vbCritical
        Exit Sub
    End If
    If objDB.SetConnection(mCnn) Then
        mSQL = "UPDATE faPostingIndex SET tnyStage=1,tnyVerifyBS=1 WHERE intFinYearID=" & mYearID & " AND intMonthID=" & MthID & " AND tnyVerifyCash=1  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        MsgBox "Verified Successfully!", vbInformation, "Saankhya"
        cmdVerify.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Me.Left = 1850
    Me.Top = 2200
End Sub

Private Sub Form_Load()
    vsGridAssets.MergeCol(0) = True
    vsGridAssets.MergeRow(0) = True
    vsGridAssets.MergeRow(1) = True
    vsGridAssets.MergeRow(2) = True
    
    vsGridLiability.MergeCol(0) = True
    vsGridLiability.MergeRow(0) = True
    vsGridLiability.MergeRow(1) = True
    vsGridLiability.MergeRow(2) = True
    
    dtClosingDate = "31/Mar/" + CStr(mYearID + 1)
    txtClosingBalance.Text = dtClosingDate
    txtClosingBalance.Enabled = False
    txtFinancialYear.Text = mYearID
    txtFinancialYear.Enabled = False
    Call FillGridAssets
    Call FillGridLiability
    Call CalculateTotal
    
    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
        cmdPublish.Enabled = True
    Else
        cmdPublish.Enabled = False
    End If
End Sub
Public Property Let YearID(mData As Integer)
    mYearID = mData
End Property

Public Property Get YearID() As Integer
    YearID = mYearID
End Property
Public Property Let MonthID(mData As Integer)
    mMonthID = mData
End Property

Public Property Get MonthID() As Integer
    MonthID = mMonthID
End Property
Public Property Let LoadMode(mData As Integer)
    mLoadMode = mData
End Property

Public Property Get LoadMode() As Integer
    LoadMode = mLoadMode
End Property

Private Sub FillGridAssets()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    If mLoadMode = 1 Then
        dtClosingDate = "31/Mar/" + CStr(mYearID + 1)
    Else
       Call MonthLastDay(mMonthID)
    End If
 
    mSQL = " SELECT A.vchAccountHeadCode,A.vchAccountHead,A.intAccountHeadID,SUM(A.fltAmount) fltAmount FROM"
    mSQL = mSQL + "     ("
    mSQL = mSQL + "     SELECT   faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,faDailyExtracts.intAccountHeadID intAccountHeadID,fltAmount  FROM faDailyExtracts"
    mSQL = mSQL + "     INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faDailyExtracts.intAccountHeadID"
    mSQL = mSQL + "     WHERE   faDailyExtracts.intFinancialYearID=" & mYearID & ""
    
    If mLoadMode = 2 Then
        mSQL = mSQL + "         AND faDailyExtracts.intmonthID < = " & mMonthID & ""
    End If

    mSQL = mSQL + "         AND faAccountHeads.tinType = 4"
    mSQL = mSQL + "     ) A"
    mSQL = mSQL + "     Group By"
    mSQL = mSQL + "        A.vchAccountHeadCode,"
    mSQL = mSQL + "        A.vchAccountHead,"
    mSQL = mSQL + "        A.intAccountHeadID"
    mSQL = mSQL + "     Order By"
    mSQL = mSQL + "        A.vchAccountHeadCode"
        
        
   
    Rec.CursorLocation = adUseClient
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 2
    vsGridAssets.Clear 1, 1
    vsGridAssets.Rows = 2
    While Not (Rec.EOF Or Rec.BOF)
        vsGridAssets.Rows = vsGridAssets.Rows + 1
        vsGridAssets.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGridAssets.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        vsGridAssets.TextMatrix(mRowCnt, 2) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
        vsGridAssets.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close


End Sub
Private Sub FillGridLiability()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    
    
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    If mLoadMode = 1 Then
        dtClosingDate = "31/Mar/" + CStr(mYearID + 1)
    Else
       Call MonthLastDay(mMonthID)
    End If
 
    mSQL = " SELECT A.vchAccountHeadCode,A.vchAccountHead,A.intAccountHeadID,SUM(A.fltAmount)*-1 fltAmount FROM"
    mSQL = mSQL + "     ("
    mSQL = mSQL + "     SELECT   faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,faDailyExtracts.intAccountHeadID intAccountHeadID,fltAmount  FROM faDailyExtracts"
    mSQL = mSQL + "     INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faDailyExtracts.intAccountHeadID"
    mSQL = mSQL + "     WHERE   faDailyExtracts.intFinancialYearID=" & mYearID & ""
    
    If mLoadMode = 2 Then
        mSQL = mSQL + "         AND faDailyExtracts.intmonthID < = " & mMonthID & ""
    End If
    
    mSQL = mSQL + "         AND faAccountHeads.tinType=3"
    mSQL = mSQL + "     ) A"
    mSQL = mSQL + "     Group By"
    mSQL = mSQL + "        A.vchAccountHeadCode,"
    mSQL = mSQL + "        A.vchAccountHead,"
    mSQL = mSQL + "        A.intAccountHeadID"
    mSQL = mSQL + "     Order By"
    mSQL = mSQL + "        A.vchAccountHeadCode"
        
    Rec.CursorLocation = adUseClient
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 2
    vsGridLiability.Clear 1, 1
    vsGridLiability.Rows = 2
    While Not (Rec.EOF Or Rec.BOF)
        vsGridLiability.Rows = vsGridLiability.Rows + 1
        vsGridLiability.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGridLiability.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        vsGridLiability.TextMatrix(mRowCnt, 2) = Format(IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount), "0.00")
        vsGridLiability.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    Rec.Close

End Sub
Private Sub CalculateTotal()
    Dim mTotalAssets As Variant
    Dim mTotalLiability As Variant
    Dim mLoop As Integer
    
    
    mTotalAssets = 0
    mTotalLiability = 0
    For mLoop = 2 To vsGridAssets.Rows - 1
        mTotalAssets = mTotalAssets + val(vsGridAssets.TextMatrix(mLoop, 2))
    Next
    txtAssestTotal.Text = mTotalAssets

    For mLoop = 2 To vsGridLiability.Rows - 1
        mTotalLiability = mTotalLiability + val(vsGridLiability.TextMatrix(mLoop, 2))
    Next
    txtLiabiliityTotal.Text = mTotalLiability
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mLoadMode = 1 Then
        frmExtractYearWiseList.FillGrid
        frmExtractYearWiseList.frmProgressBar.Visible = False
    Else
        frmExtactedMonthWiseList.FillGrid
    End If
End Sub

Private Sub txtCouncilDate_LostFocus()
    If txtCouncilDate.Text <> "" Then
        txtCouncilDate.Text = CheckDateInMMM(txtCouncilDate.Text)
    End If
End Sub
Private Function mCheckStatus() As Integer
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mStatus As Integer
    
    mStatus = -1
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " SELECT * FROM faPostingIndex WHERE   intMonthID=" & mMonthID & " AND tnyVerifyBS=1 AND intFinYearID=" & mYearID
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mStatus = Rec!tnyStage
    End If
    If mStatus = 1 Then
        cmdVerify.Enabled = False
        cmdPublish.Enabled = True
    ElseIf mStatus = 2 Then
        cmdVerify.Enabled = False
        cmdPublish.Enabled = True
    End If
    Rec.Close
    mCnn.Close
End Function
Private Function MonthLastDay(intMonthID As Integer) As Variant

  Dim dtMonthLastDate  As Variant
  
  Select Case intMonthID
        
            Case 4
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 5
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 6
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 7
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 8
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 9
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 10
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 11
                    dtMonthLastDate = "30/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 12
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID)
            Case 1
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            Case 2
                    dtMonthLastDate = "28/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            Case 3
                    dtMonthLastDate = "31/" + Left(MonthName(intMonthID), 3) + "/" + CStr(mYearID + 1)
            End Select
            
            dtClosingDate = dtMonthLastDate
            txtClosingBalance.Text = dtClosingDate
            
End Function


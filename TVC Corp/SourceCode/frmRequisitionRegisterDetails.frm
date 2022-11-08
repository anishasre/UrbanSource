VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmRequisitionRegisterDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RequisitionRegister Details"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10230
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   9630
      Top             =   5040
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   45
      TabIndex        =   16
      Top             =   1755
      Width           =   10140
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2130
         Left            =   1215
         TabIndex        =   17
         Top             =   315
         Width           =   7395
         _cx             =   13044
         _cy             =   3757
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Rows            =   6
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRequisitionRegisterDetails.frx":0000
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   4905
      TabIndex        =   3
      Top             =   4815
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   3555
      TabIndex        =   2
      Top             =   4815
      Width           =   1320
   End
   Begin VB.Frame FrReqDetails 
      Height          =   1095
      Left            =   -45
      TabIndex        =   1
      Top             =   630
      Width           =   10230
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3915
         TabIndex        =   14
         Top             =   210
         Width           =   2535
      End
      Begin VB.TextBox txtSource 
         Appearance      =   0  'Flat
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
         Left            =   7830
         TabIndex        =   12
         Top             =   225
         Width           =   2175
      End
      Begin VB.TextBox txtReqNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   945
         TabIndex        =   10
         Top             =   225
         Width           =   1905
      End
      Begin VB.TextBox txtIMPODesig 
         Appearance      =   0  'Flat
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
         Left            =   3915
         TabIndex        =   8
         Top             =   615
         Width           =   2535
      End
      Begin VB.TextBox txtReqDate 
         Appearance      =   0  'Flat
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
         Left            =   945
         TabIndex        =   6
         Top             =   600
         Width           =   1905
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7830
         TabIndex        =   4
         Top             =   630
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "IMPO"
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
         Left            =   3375
         TabIndex        =   15
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Source Of Fund"
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
         Left            =   6570
         TabIndex        =   13
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Req.No"
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
         Left            =   270
         TabIndex        =   11
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Designation"
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
         Left            =   2970
         TabIndex        =   9
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Req.Date"
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
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   735
      End
      Begin VB.Label lblCategory 
         Caption         =   "Category"
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
         Left            =   6975
         TabIndex        =   5
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   0
      Width           =   10185
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2745
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   4650
   End
End
Attribute VB_Name = "frmRequisitionRegisterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim numTotalExpenditureExludingThisBill As Variant
    Dim ReqID   As Variant
        
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSave_Click()
            Dim mCnn    As New ADODB.Connection
            Dim objdb   As New clsDB
            Dim mArrIn  As Variant
            Dim mArrInChild  As Variant
            Dim mArrOut  As Variant
            Dim mSql As String
            Dim mReqNo As Variant
            Dim mAuthorizeNo As String
            Dim mAllotNo As Variant
            Dim mLen As Variant
            
            mReqNo = txtReqNo.Text
            mLen = Len(txtReqNo.Text)
            mAuthorizeNo = "7" + mID(txtReqNo.Text, 2, mLen)
            mAllotNo = "8" + mID(txtReqNo.Text, 2, mLen)
                    
            If objdb.SetConnection(mCnn) Then
                mArrIn = Array(txtReqNo.Tag, _
                                val(vsGrid.TextMatrix(2, 1)), _
                                val(vsGrid.TextMatrix(3, 1)), _
                                val(vsGrid.TextMatrix(1, 1)), _
                                val(vsGrid.TextMatrix(5, 1)), _
                                val(vsGrid.TextMatrix(4, 1)), _
                                numTotalExpenditureExludingThisBill, _
                                0)
                objdb.ExecuteSP "spSaveIssueLetterOfAllotment", mArrIn, , , mCnn, adCmdStoredProc
                mSql = "Update faAllotments set vchAuthorizationNo=" & mAuthorizeNo & ",vchAllotmentNo=" & mAllotNo & ",intCountOfVouchers=1 where intID=" & txtReqNo.Tag & "  "
                'UPDATE FIELD intCountOfVouchers TO IDENTIFY VERIFIED REQISITION DETAILS SINCE THE FIELD IS UNUSED
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                MsgBox "Updated Successfully ", vbInformation
                cmdSave.Enabled = False
                frmRequisitionRegister.FillGrid
                Unload Me
           End If
    End Sub
    
    Private Sub Form_Load()
         XPC.InitSubClassing
         lblMsg.Visible = False
         'Call FillGrid
    End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub

Public Sub FillGrid()
    Dim mCnn    As New ADODB.Connection
    Dim objdb   As New clsDB
    Dim Rec     As New ADODB.Recordset
    Dim mSql    As String
    Dim mYearID   As Integer
    
    mYearID = gbFinancialYearID - 1
    
    numTotalExpenditureExludingThisBill = Null
    objdb.CreateNewConnection mCnn, enuSourceString.Saankhya

    '******************TOTAL ALLOTMENT RECEIVED************************************
    Select Case val(txtSource.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 10, 11, 12, 13, 14 ' Development Fund (Gen/SPC/TSP) + Special Grant + Road Renovation
            mSql = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID in (1,27,28, 10, 11, 12, 13, 14) "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & " AND dtAllotmentDate <= '" & txtReqDate.Tag & "' "
            mSql = mSql + " AND ISNULL(tnyGroupID,0)<>90"
            mSql = mSql + " Union All"
            mSql = mSql + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID in (1,27,28, 10, 11, 12, 13, 14) "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
         Case 16, 17 'Road / Non Road
            mSql = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID in (16,17) "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & " AND dtAllotmentDate <= '" & txtReqDate.Tag & "' "
            mSql = mSql + " AND ISNULL(tnyGroupID,0)<>90"
            mSql = mSql + " Union All"
            mSql = mSql + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID in (16,17) "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
        Case 3 ' B-Fund
            mSql = "Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0 AND intSchemeID = " & val(txtCategory.Tag)
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & " AND dtAllotmentDate <= '" & txtReqDate.Tag & "' "
            mSql = mSql + " AND ISNULL(tnyGroupID,0)<>90"
        Case Else
            mSql = "Select Sum(A.AmountReceived) AmountReceived From (Select Sum(fltAmount) As AmountReceived From faAllotmentLetters Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID=" & val(txtSource.Tag) & "  "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 1 AND intFinancialYearID=" & mYearID & " AND dtAllotmentDate <= '" & txtReqDate.Tag & "' "
            mSql = mSql + " AND ISNULL(tnyGroupID,0)<>90"
            mSql = mSql + " Union All"
            mSql = mSql + " Select Sum(fltAmount) As AmountReceived from faExtractAllotments Where ISNULL(tnyCancelledFlag,0) = 0"
            mSql = mSql + " AND intSourceofFundID=" & val(txtSource.Tag) & " "
            mSql = mSql + " AND ISNULL(tnyStatus,0) = 2 AND intFinancialYearID=" & mYearID & " )A"
    End Select
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!AmountReceived), "0", Rec!AmountReceived)
    End If
    Rec.Close
    
    '********************TOTAL ALLOTMENT ISSUED*************************************
    
    Select Case val(txtSource.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 10, 11, 12, 13, 14
            mSql = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSql = mSql + " AND intSourceID IN (1, 27, 28, 10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
         Case 16, 17
            mSql = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 "
            mSql = mSql + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "'  And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
        Case 3
            mSql = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  "
            mSql = mSql + " AND intSourceID =" & val(txtSource.Tag) & "  AND intSchemeID = " & val(txtCategory.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "'  And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
        Case Else
            mSql = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 "
            mSql = mSql + " AND intSourceID =" & val(txtSource.Tag) & " AND intFinancialYearID=" & mYearID & " "
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
    End Select
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(3, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
    End If
    Rec.Close
    
    '*********************TOTAL ALLOTMENT ISSUED TO THE IMPO FOR THE CURRENT YEAR*************
    
       Select Case val(txtSource.Tag) 'SOURCE OF FUND
        Case 1, 27, 28, 10, 11, 12, 13, 14
            mSql = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & txtIMPO.Tag & " "
            mSql = mSql + " AND intSourceID IN (1, 27, 28, 10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " "
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
        Case 16, 17
            mSql = " Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & txtIMPO.Tag & " "
            mSql = mSql + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & " "
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
        Case 3
            mSql = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  = 1  AND intImplementingOfficersID=" & txtIMPO.Tag & "  "
            mSql = mSql + " AND intSourceID =" & val(txtSource.Tag) & "  AND intSchemeID = " & val(txtCategory.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
        Case Else
            mSql = "Select Sum(fltAuthorizedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) = 1 AND intImplementingOfficersID=" & txtIMPO.Tag & "  "
            mSql = mSql + " AND intSourceID =" & val(txtSource.Tag) & " AND intFinancialYearID=" & mYearID & ""
            mSql = mSql + " AND dtAuthorizationDate <= '" & txtReqDate.Tag & "' And intCountOfVouchers = 1 AND tnyStage = 2 "
            mSql = mSql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1,2)"
    End Select
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(4, 1) = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
    End If
    Rec.Close
    
    '*********************BALANCE AVAILABLE****************************************************
    vsGrid.TextMatrix(5, 1) = vsGrid.TextMatrix(2, 1) - vsGrid.TextMatrix(3, 1) - vsGrid.TextMatrix(0, 1)
    
    '*********************BALANCE AVAILABLE TO THE IMPO****************************************
    numTotalExpenditureExludingThisBill = val(vsGrid.TextMatrix(4, 1)) - val(vsGrid.TextMatrix(3, 1))
End Sub

Public Property Let RequisitionID(mData As Variant)
    ReqID = mData
End Property

Public Property Get RequisitionID() As Variant
    RequisitionID = ReqID
End Property




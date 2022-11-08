VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProjectJournalDetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15420
   Icon            =   "frmProjectJournalDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUndo 
      Caption         =   "CANCEL EXPENDITURE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   13635
      TabIndex        =   32
      Top             =   2655
      Width           =   1590
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   14760
      Top             =   8235
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   15420
      TabIndex        =   0
      Top             =   0
      Width           =   15420
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   12570
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   345
         Width           =   2160
      End
      Begin VB.ComboBox cmbTransactionType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frmProjectJournalDetails.frx":1CCA
         Left            =   300
         List            =   "frmProjectJournalDetails.frx":1CCC
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   390
         Width           =   10455
      End
      Begin VB.Label lblmsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PROJECT EXPENDITURE FOR SOURCE DEDUCTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   2925
         TabIndex        =   4
         Top             =   810
         Visible         =   0   'False
         Width           =   6525
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5070
      Left            =   0
      TabIndex        =   3
      Top             =   2655
      Width           =   11550
      _cx             =   20373
      _cy             =   8943
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProjectJournalDetails.frx":1CCE
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
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      Top             =   555
      Width           =   15420
      Begin VB.Frame fraVoucher 
         Height          =   1620
         Left            =   4770
         TabIndex        =   16
         Top             =   240
         Width           =   3945
         Begin VB.TextBox txtVrDate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1245
            TabIndex        =   24
            Top             =   825
            Width           =   2265
         End
         Begin VB.CommandButton cmdSearchVoucher 
            Caption         =   "..."
            Height          =   330
            Left            =   3540
            TabIndex        =   23
            Top             =   465
            Width           =   285
         End
         Begin VB.TextBox txtVoucherNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1245
            TabIndex        =   22
            Top             =   465
            Width           =   2265
         End
         Begin VB.TextBox txtAmount 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1245
            TabIndex        =   19
            Top             =   1185
            Width           =   2265
         End
         Begin VB.Label Label8 
            Caption         =   "DATE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   765
            TabIndex        =   25
            Top             =   870
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "VOUCHER NO"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "AMOUNT"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   480
            TabIndex        =   20
            Top             =   1230
            Width           =   750
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   14625
         TabIndex        =   2
         Top             =   330
         Width           =   675
      End
      Begin VB.Frame fraProject 
         Height          =   1620
         Left            =   8745
         TabIndex        =   5
         Top             =   240
         Width           =   5865
         Begin VB.TextBox txtPAmount 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3825
            TabIndex        =   30
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox txtCategory 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1215
            TabIndex        =   15
            Top             =   1155
            Width           =   1755
         End
         Begin VB.TextBox txtSourceOfFund 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1215
            TabIndex        =   14
            Top             =   810
            Width           =   4095
         End
         Begin VB.CommandButton cmdSearchProject 
            Caption         =   "..."
            Height          =   330
            Left            =   5340
            TabIndex        =   13
            Top             =   465
            Width           =   285
         End
         Begin VB.TextBox txtProjectNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1215
            TabIndex        =   12
            Top             =   465
            Width           =   1245
         End
         Begin VB.TextBox txtProjectNameEng 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2475
            TabIndex        =   11
            Top             =   465
            Width           =   2835
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "AMOUNT"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   3075
            TabIndex        =   29
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "PROJECT NO"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   135
            TabIndex        =   8
            Top             =   510
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "CATEGORY"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   270
            TabIndex        =   7
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "SOURCE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   510
            TabIndex        =   6
            Top             =   855
            Width           =   645
         End
      End
      Begin VB.Frame fraLetterOfAuthority 
         Height          =   1605
         Left            =   45
         TabIndex        =   9
         Top             =   255
         Width           =   4695
         Begin VB.TextBox txtLetterAuthorityAmt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2460
            TabIndex        =   27
            Top             =   1155
            Width           =   1965
         End
         Begin VB.TextBox txtLetterAuthorityTrType 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   300
            TabIndex        =   26
            Top             =   810
            Width           =   4125
         End
         Begin VB.TextBox txtLetterOfAuthority 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   300
            TabIndex        =   18
            Top             =   450
            Width           =   3795
         End
         Begin VB.CommandButton cmdSearchAuthority 
            Caption         =   "..."
            Height          =   315
            Left            =   4125
            TabIndex        =   17
            Top             =   450
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "AMOUNT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   1725
            TabIndex        =   28
            Top             =   1230
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "LETTER OF AUTHORITY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   315
            TabIndex        =   10
            Top             =   165
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmProjectJournalDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim mCrAccHeadId    As Integer
    Dim mFunctionId     As Integer
    Dim mFunctionaryID  As Integer
    Dim mGrossAccHeadID As Integer
    Dim mTrType         As Integer
    Dim mSourceOfFundID As Variant
    Dim mLoASource      As Variant
    Dim mLoACategory    As Variant
    Dim mSubsectorID    As Integer
    Dim mAllotmentDate  As Date
    Dim mPFunctionID    As Integer
    Dim mPFunctionaryID As Integer
    Dim mPMicroSectorID As Integer
    Dim mDPCApprovalNo  As Variant
    Dim mDPCApprovalDate As Variant
    Dim mPAccHeadID     As Integer
    Dim mPAccHeadCode   As Variant
    Dim mPProjectCost   As Variant
    Dim mLoadMode       As Integer      '50-For BENEFICIARY CONTRIBUTIONS
    Dim mYearID         As Integer
    Dim mTransactionTypeID As Integer
    
    Private Sub cmbTransactionType_Click()
        Call FormInitialize
        If cmbTransactionType.ListIndex > 0 Then
            mTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
        End If
        Call FillGrid
    End Sub
    
    Private Sub cmbYear_Change()
        'Call FormInitialize
    End Sub

    Private Sub cmbYear_Click()
    If cmbYear.ListIndex > -1 Then
        mYearID = cmbYear.ItemData(cmbYear.ListIndex)
    End If
    Call FormInitialize
    Call FillGrid
    End Sub

    Private Sub cmdAdd_Click()
        Dim mcnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim arrInput    As Variant
        Dim arrOutPut   As Variant
        Dim Reqn        As uRequisition
        Dim mSql        As String
        Dim mArrIn      As Variant
        Dim objVrSub    As uVoucherSub
        Dim mCount      As Integer
        Dim mLetterOfAuthority As Variant
        Dim mTypeID As Integer
        
    
        If SaveValidation = False Then Exit Sub
        
        If mLoadMode = 50 Then          'BENEFICIARY CONTRIBUTIONS
            If cmbTransactionType.ListIndex > -1 Then
                Select Case cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                    Case Is = 3008: mTypeID = 2 ' Beneficiary Contribution
                    Case Is = 3009: mTypeID = 3 ' Loan
                    Case Is = 3010: mTypeID = 4 ' State Sponsors Scheme
                    Case Is = 3011: mTypeID = 5 ' Centrally Sponsored Scheme
                    Case Is = 3012: mTypeID = 6 ' MNREGS
                    Case Else
                        MsgBox "Please select the type of Transaction", vbInformation
                        Exit Sub
                End Select
            End If
            
            With objVrSub
                .intVoucherID = val(txtVoucherNo.Tag)
                .decProjectID = val(txtProjectNo.Tag)
                .intSourceOfFundID = mSourceOfFundID
                .intCategoryID = val(txtCategory.Tag)
                .intSectorID = mSubsectorID
                .intAllotmentID = Null
                .intAgreementID = Null
                .intCashBookID = Null
                .intImplementingOfficerID = Null
                .intCreditorTypeID = Null
                .intCreditorsID = Null
                .intTypeID = mTypeID ' 2                  'To Identify Journals  for Beneficiary Contributions
                .intLocalBodyID = gbLocalBodyID
                
                arrInput = Array(.intVoucherID, _
                                .intLocalBodyID, _
                                .decProjectID, _
                                .intSourceOfFundID, _
                                .intCategoryID, _
                                .intSectorID, _
                                .intAllotmentID, _
                                .intAgreementID, _
                                .intCashBookID, _
                                .intImplementingOfficerID, _
                                .intCreditorTypeID, _
                                .intCreditorsID, _
                                .intTypeID)
                objDb.ExecuteSP "spSaveVoucherSub", arrInput, , , mcnn
            End With
        Else
            Call GetLetterOfAuthorityDetails
            '------------Save faVoucherSub---------------------------------
            If objDb.SetConnection(mcnn) Then
                mSql = "SELECT intVoucherID FROM faVoucherSub WHERE intVoucherID = " & val(txtVoucherNo.Tag) & "  "
                Rec.Open mSql, mcnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mCount = 1
                Else
                    mCount = 0
                End If
                Rec.Close
                If mCount = 1 Then
                    mSql = " UPDATE faVoucherSub SET "
                    mSql = mSql + " decProjectID = " & val(txtProjectNo.Tag) & ", intSourceOfFundID =" & mSourceOfFundID & ", intCategoryID =" & val(txtCategory.Tag) & ", intSectorID =" & mSubsectorID & ", intTypeID =1"
                    mSql = mSql + " WHERE intVoucherID=" & txtVoucherNo.Tag
                    objDb.ExecuteSP mSql, , , , mcnn, adCmdText
                Else
                    With objVrSub
                        .intVoucherID = val(txtVoucherNo.Tag)
                        .decProjectID = val(txtProjectNo.Tag)
                        .intSourceOfFundID = mSourceOfFundID
                        .intCategoryID = val(txtCategory.Tag)
                        .intSectorID = mSubsectorID
                        .intAllotmentID = Null
                        .intAgreementID = Null
                        .intCashBookID = Null
                        .intImplementingOfficerID = Null
                        .intCreditorTypeID = Null
                        .intCreditorsID = Null
                        .intTypeID = 1                  'To Identify Journals  inserted from this module
                        .intLocalBodyID = gbLocalBodyID
                        
                        arrInput = Array(.intVoucherID, _
                                        .intLocalBodyID, _
                                        .decProjectID, _
                                        .intSourceOfFundID, _
                                        .intCategoryID, _
                                        .intSectorID, _
                                        .intAllotmentID, _
                                        .intAgreementID, _
                                        .intCashBookID, _
                                        .intImplementingOfficerID, _
                                        .intCreditorTypeID, _
                                        .intCreditorsID, _
                                        .intTypeID)
                        objDb.ExecuteSP "spSaveVoucherSub", arrInput, , , mcnn
                    End With
                End If
            End If
            
            '--------------Save faAllotmentLetters-------------------------
            mLetterOfAuthority = Token(txtLetterOfAuthority.Text, "[")
            mArrIn = Array(-1, _
                            mLetterOfAuthority, _
                            mAllotmentDate, _
                            Null, _
                            mLoASource, mLoACategory, _
                            Null, Null, _
                            mCrAccHeadId, _
                            Null, _
                            Null, _
                            mFunctionaryID, _
                            mFunctionId, _
                            mGrossAccHeadID, _
                            val(txtAmount.Text), _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            gbUserID, _
                            gbTransactionDate, _
                            Null, _
                            gbLocalBodyID, _
                            mYearID, _
                            1, _
                            0, Null, Null, 90, mTrType, 0, val(txtVoucherNo.Tag) _
                        )
            objDb.ExecuteSP "spSaveAllotmentLetter", mArrIn, , , mcnn, adCmdStoredProc  'MODIFIED BY MINU ON 18/MAY/2013
        
            '--------------Save faAllotments-------------------------------
          
            With Reqn
                .tnyStage = 2
                .vchRequisition = Null
                .dtRequisitionDate = DdMmmYy(txtVrDate.Text)
                .intFinancialYearID = mYearID
                .intImplementingOfficersID = Null
                .vchDesignation = Null
                .vchNameofIMPO = Null
                .vchPlace = Null
                .vchDepartment = Null
                .vchDDOCode = Null
                .fltRequestedAmt = val(txtAmount.Text)
                .tnyPlanOrNonPlan = 1
                .numProjectID = val(txtProjectNo.Tag)
                .numProjectNo = txtProjectNo.Text
                .fltProjectCost = mPProjectCost
                .vchDPCApprovalNo = mDPCApprovalNo
                .dtDPCDate = mDPCApprovalDate
                .intSourceID = val(txtSourceOfFund.Tag)
                .intCategoryID = val(txtCategory.Tag)
                .intTreasuryID = Null
                .vchTreasuryCode = Null
                .vchTreasuryName = Null
                .vchGHeadofAccount = Null
                .vchGBudgetHead = Null
                .vchGDemandNo = Null
                .intFunctionaryID = mPFunctionaryID
                .intFunctionID = mPFunctionID
                .intAccountHeadID = mPAccHeadID
                .vchAccountHeadCode = mPAccHeadCode
                .intLBID = gbLocalBodyID
                .tnyStatus = 1
                .tnyInstallmentNo = Null
                .intSchemeID = Null
                .intSubSecID = mSubsectorID
                .intMircoSectorID = mPMicroSectorID
        
            arrInput = Array("", .tnyStage, .vchRequisition, _
                        .dtRequisitionDate, _
                        .intImplementingOfficersID, _
                        .vchDesignation, _
                        .vchNameofIMPO, _
                        .vchPlace, _
                        .vchDepartment, _
                        .vchDDOCode, _
                        .fltRequestedAmt, _
                        .tnyPlanOrNonPlan, _
                        .numProjectID, _
                        .numProjectNo, _
                        .fltProjectCost, _
                        .vchDPCApprovalNo, _
                        .dtDPCDate, _
                        .intSourceID, _
                        .intCategoryID, _
                        .intTreasuryID, _
                        .vchTreasuryCode, _
                        .vchTreasuryName, _
                        .vchGHeadofAccount, _
                        .vchGBudgetHead, _
                        .vchGDemandNo, _
                        .intFunctionaryID, .intFunctionID, .intAccountHeadID, .vchAccountHeadCode, .intLBID, .intFinancialYearID, .tnyStatus, Null, Null, .fltRequestedAmt, Null, Null, Null, Null, Null, .tnyInstallmentNo, _
                         Null, Null, Null, Null, Null, Null, Null, Null, Null, .intSchemeID, .intSubSecID, .intMircoSectorID, 1, val(txtVoucherNo.Tag))
        
                objDb.ExecuteSP "spSaveAllotmentRequisition", arrInput, arrOutPut, True, mcnn, adCmdStoredProc  'MODIFIED BY MINU ON 18/MAY/2013
            End With
        End If
        Call UpdateDetailsToSulekha
        Call FillGrid
        Call FormInitialize
    End Sub
    
    Private Sub cmdSearchAuthority_Click()
        If cmbYear.ItemData(cmbYear.ListIndex) > 2011 Then
            frmSearchLetterOfAuthority.YearID = cmbYear.ItemData(cmbYear.ListIndex)
            frmSearchLetterOfAuthority.Show vbModal
            If gbSearchID <> -1 Then
                'txtLetterOfAuthority.Text = gbSearchStr
                txtLetterOfAuthority.Tag = gbSearchID
                gbSearchCode = ""
                gbSearchID = -1
            End If
        End If
        
        Call GetLetterOfAuthorityDetails
        
    End Sub
    Private Sub GetLetterOfAuthorityDetails()
        Dim mcnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
       
     
        If objDb.SetConnection(mcnn) Then
            If val(txtLetterOfAuthority.Tag) > 0 Then
                mSql = " SELECT * FROM faAllotmentLetters "
                mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID=faAllotmentLetters.intTransactionTypeID"
                mSql = mSql + " WHERE intAllotmentID= " & val(txtLetterOfAuthority.Tag) & " "
                Rec.Open mSql, mcnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mLoASource = IIf(IsNull(Rec!intSourceOfFundID), 0, Rec!intSourceOfFundID)
                    mLoACategory = IIf(IsNull(Rec!intCategoryID), 0, Rec!intCategoryID)
                    mCrAccHeadId = IIf(IsNull(Rec!intCrAccountHeadID), 0, Rec!intCrAccountHeadID)
                    mFunctionId = IIf(IsNull(Rec!intFunctionID), 0, Rec!intFunctionID)
                    mFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), 0, Rec!intFunctionaryID)
                    mGrossAccHeadID = IIf(IsNull(Rec!intGrossAccountHeadID), 0, Rec!intGrossAccountHeadID)
                    mTrType = IIf(IsNull(Rec!intTransactionTypeID), 0, Rec!intTransactionTypeID)
                    mAllotmentDate = DdMmmYy(Rec!dtAllotmentDate)
                    txtLetterOfAuthority.Text = Rec!vchAllotmentNo & " [" & DdMmmYy(Rec!dtAllotmentDate) & "]"
                    txtLetterAuthorityTrType = IIf(IsNull(Rec!vchTransactionType), 0, Rec!vchTransactionType)
                    txtLetterAuthorityAmt = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                End If
                Rec.Close
            End If
       End If
    End Sub
    
    Private Sub cmdSearchProject_Click()
        Dim mcnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        
        
         'cmbYear.ItemData (cmbYear.ListIndex)
        If mYearID < 2012 Then
            Exit Sub
        End If
        
        If val(txtVoucherNo.Tag) > 0 Then
        
            If mYearID = gbFinancialYearID - 1 Then
                frmSearchProjects.PreviousYearMode = 1
            End If
            frmSearchProjects.Show vbModal
           
''''            If objDB.SetConnection(mCnn) Then   COMMENTED BY MINU ON 18/MAY/2013
''''                mSQL = " SELECT * FROM faVoucherSub "
''''                mSQL = mSQL + " WHERE decProjectID= " & val(gbSearchStr) & " "
''''                mSQL = mSQL + " And intSourceOfFundID=" & val(gbSearchID) & ""
''''                mSQL = mSQL + " And ISNULL(intTypeID,0)<>1 "
''''                Rec.Open mSQL, mCnn
''''                If Not (Rec.EOF And Rec.BOF) Then
''''                    MsgBox "Project Already Selected", vbInformation, "Saankhya"
''''                    Exit Sub
''''                Else

                    If mLoadMode = 50 Then          'BENEFICIARY CONTRIBUTIONS
                        Select Case cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                            Case 3008
                                If val(gbSearchID) <> 9 Then
                                    MsgBox "The Project Source is not Beneficiary Contribution", vbInformation, "Saankhya"
                                    Exit Sub
                                End If
                            Case 3009
                                If val(gbSearchID) <> 5 And val(gbSearchID) <> 6 Then
                                    MsgBox "The Project Source is not Loan", vbInformation, "Saankhya"
                                    Exit Sub
                                End If
                            Case 3010
                                If val(gbSearchID) <> 3 Then
                                    MsgBox "The Project Source is not State Sponsored Scheme Fund !", vbInformation, "Saankhya"
                                    Exit Sub
                                End If
                            Case 3012 ' MNREGS
                                If val(gbSearchID) <> 2 Then
                                    MsgBox "The Project Source is not Centrally Sponsored Scheme Fund !", vbInformation, "Saankhya"
                                    Exit Sub
                                End If
                            Case Else
                        End Select
                    End If
                    txtProjectNo.SetFocus
''''                End If
''''                Rec.Close
''''            End If
        Else
            MsgBox "Please Select a Voucher Number", vbInformation, "Saankhya"
            Exit Sub
        End If
    End Sub
    
    Private Sub cmdSearchVoucher_Click()
        Dim mcnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        Dim mSqlMsg As String
        
        
        If cmbYear.ItemData(cmbYear.ListIndex) > 2011 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
        Else
            Exit Sub
        End If
        
        If mLoadMode = 50 Then          'BENEFICIARY CONTRIBUTIONS
            frmSearchVouchers.CheckMode = 40
            'frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
            'frmSearchVouchers.txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
            
            frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateSerial(mYearID, 4, 1))
            frmSearchVouchers.txtToDate.Text = DdMmmYy(DateSerial(mYearID + 1, 3, 31))
            
            frmSearchVouchers.chkContra.Visible = False
            frmSearchVouchers.chkReceipt.Visible = False
            frmSearchVouchers.chkJournal.Visible = True
            frmSearchVouchers.chkJournal.value = 1
            frmSearchVouchers.chkPayment.Visible = False
            frmSearchVouchers.txtFromDate.Enabled = False
            frmSearchVouchers.txtToDate.Enabled = False
            frmSearchVouchers.Show vbModal
            If gbSearchID <> -1 Then
                txtVoucherNo.Text = gbSearchCode
                txtVoucherNo.Tag = gbSearchID
                gbSearchCode = ""
                gbSearchID = -1
            End If
            If val(txtVoucherNo.Tag) > 1 Then
                If ValidateBeneficiaryAccHead(val(txtVoucherNo.Tag)) = False Then
                
                    mSqlMsg = "JOURNAL NOT CREDITED TO  "
                    mSqlMsg = mSqlMsg + cmbTransactionType.Text
                    
                    MsgBox mSqlMsg, vbInformation, "Saankhya"
                    'MsgBox "Journal Not credited to Beneficiary Contribution", vbInformation, "Saankhya"
                    txtVoucherNo.Text = ""
                    txtVoucherNo.Tag = ""
                    txtVrDate.Text = ""
                    txtAmount.Text = ""
                    Exit Sub
                End If
            End If
        Else
             If val(txtLetterOfAuthority.Tag) > 0 Then
                 frmSearchVouchers.CheckMode = 40
                 
                 'frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbStartingDate))
                 'frmSearchVouchers.txtToDate.Text = DdMmmYy(DateAdd("yyyy", -1, gbEndingDate))
                 
                 frmSearchVouchers.txtFromDate.Text = DdMmmYy(DateSerial(mYearID, 4, 1))
                 frmSearchVouchers.txtToDate.Text = DdMmmYy(DateSerial(mYearID + 1, 3, 31))
                
                 frmSearchVouchers.chkContra.Visible = False
                 frmSearchVouchers.chkReceipt.Visible = False
                 frmSearchVouchers.chkJournal.Visible = True
                 frmSearchVouchers.chkJournal.value = 1
                 frmSearchVouchers.chkPayment.Visible = False
                 frmSearchVouchers.txtFromDate.Enabled = False
                 frmSearchVouchers.txtToDate.Enabled = False
                 frmSearchVouchers.Show vbModal
                 If gbSearchID <> -1 Then
                     txtVoucherNo.Text = gbSearchCode
                     txtVoucherNo.Tag = gbSearchID
                     gbSearchCode = ""
                     gbSearchID = -1
                 End If
             Else
                 MsgBox "Please Select a Letter Of Authority", vbInformation, "Saankhya"
                 Exit Sub
             End If
        End If
             If val(txtVoucherNo.Tag) > 0 Then
                 If objDb.SetConnection(mcnn) Then
                     mSql = " SELECT * FROM faVouchers "
                     mSql = mSql + " WHERE intVoucherID = " & txtVoucherNo.Tag & " "
                     Rec.Open mSql, mcnn
                     If Not (Rec.EOF And Rec.BOF) Then
                          txtAmount.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                          txtVrDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), 0, Rec!dtDate))
                     End If
                     Rec.Close
                  End If
                  
                  If val(txtLetterOfAuthority.Tag) <> 0 Then
                      If ValidateCreditAccHead = False Then
                         MsgBox "Credit Account Head Of Journal is not matching with Letter Of Authority", vbInformation, "Saankhya"
                         txtVoucherNo.Text = ""
                         txtVoucherNo.Tag = ""
                         txtVrDate.Text = ""
                         txtAmount.Text = ""
                         Exit Sub
                     End If
                  End If
                 
            End If
        
    End Sub
    
    Private Sub Command1_Click()
        Call FormInitialize
    End Sub
    
    Private Sub UpdateDetailsToSulekha()
         Dim mCnnSulekha   As New ADODB.Connection
         Dim objDb   As New clsDB
         Dim mSql As String
         Dim arrInput As Variant
         
         Call GetDebitAccHeadDetails
         If val(txtVoucherNo.Tag) > 0 Then
             If (objDb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                 arrInput = Array(gbLBID, _
                                mYearID, _
                                val(txtProjectNo.Tag), _
                                -1, val(txtSourceOfFund.Tag), _
                                val(txtAmount), _
                                val(txtVoucherNo.Tag), CDate(txtVrDate.Text))
        
                objDb.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
             Else
                MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
                Exit Sub
             End If
             mCnnSulekha.Close
         End If
    End Sub
    
    Private Sub FillCombo()
        Dim mSql As String
        mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType WHERE intTransactionTypeID IN ( 3008,3009,3010,3011,3012)"
        PopulateList cmbTransactionType, mSql, , True, , True
    End Sub
    
    Private Sub cmdUndo_Click()  'NEW CODE BY MINU ON 18/MAY/2013
        Dim mSql    As String
        Dim objDb   As New clsDB
        Dim mcnn    As New ADODB.Connection
        Dim mCnnSulekha    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        
        If mLoadMode = 50 Then
            mSql = "Update faVoucherSub Set intTypeID=NUll,intAllotmentID=NULL Where intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 7)) & ""
            objDb.ExecuteSP mSql, , , , mcnn, adCmdText
        Else
            mSql = "Update faVoucherSub Set intTypeID=NUll,intAllotmentID=NULL Where intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 7)) & ""
            objDb.ExecuteSP mSql, , , , mcnn, adCmdText
            
            mSql = ""
            mSql = " Update faAllotmentLetters Set tnyStatus=8 From faAllotmentLetters" 'intVoucherID
            mSql = mSql + " Where faAllotmentLetters.tnyGroupID=90 And intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 7)) & ""
            objDb.ExecuteSP mSql, , , , mcnn, adCmdText
            
            mSql = ""
            mSql = "Update faAllotments Set tnyStatus=2 Where intVoucherID=" & val(vsGrid.TextMatrix(vsGrid.Row, 7)) & ""
            objDb.ExecuteSP mSql, , , , mcnn, adCmdText
        End If
        If (objDb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            mSql = "Update ExpenseDetails set tnyTransfer=0,tnyCancelation=1 where intVoucherID = " & val(vsGrid.TextMatrix(vsGrid.Row, 7)) & ""
            objDb.ExecuteSP mSql, , , , mCnnSulekha, adCmdText
            mCnnSulekha.Close
        Else
            MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
            Exit Sub
        End If
        cmdUndo.Enabled = False
        Call FillGrid
    End Sub

    Private Sub Form_Load()
         XPC.InitSubClassing
         Call FillYear
         If mLoadMode = 50 Then
            'fraLetterOfAuthority.Enabled = False
            fraLetterOfAuthority.Visible = False
            fraVoucher.Enabled = True
            fraProject.Enabled = True
            lblmsg.Caption = "BENEFICIARY CONTRIBUTION DIRECT EXPENDITURE"
            lblmsg.Visible = True
            lblmsg.Left = 100   'MODIFIED BY MINU ON 18/MAY/2013
            lblmsg.Top = 100
            Call FillCombo
            Call FillGrid
         Else
            fraLetterOfAuthority.Enabled = True
            fraVoucher.Enabled = True
            fraProject.Enabled = True
            lblmsg.Caption = "PROJECT EXPENDITURE FOR SOURCE DEDUCTION"
            'Call FillCombo
            lblmsg.Visible = True  'MODIFIED BY MINU ON 18/MAY/2013
            lblmsg.Left = 100
            lblmsg.Top = 200
            cmbTransactionType.Visible = False
            Call FillGrid
         End If
         
    End Sub
        
    Private Sub FillYear()
        PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4)) + '-' + Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear WHERE intFinancialYearID > 2011", , , , True
        cmbYear.ListIndex = cmbYear.ListCount - 1
        vsGrid.SelectionMode = flexSelectionByRow
    End Sub
    
        
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    
    Private Sub txtLetterAuthorityAmt_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
     Private Sub KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        Else
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtLetterAuthorityTrType_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    
    Private Sub txtLetterOfAuthority_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    
    Private Sub txtPAmount_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    
    Private Sub txtProjectNameEng_KeyPress(KeyAscii As Integer)
        Call KeyPress(KeyAscii)
    End Sub
    
    Private Sub txtProjectNo_GotFocus()
        Dim objProj As New clsProject
        Dim objProFund As New clsProjectFund
        Dim mProjectID As Variant
        Dim mcnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mCnPlan As New ADODB.Connection
        Dim mSql  As String
        Dim mCol  As Collection
        Dim mRow As Integer
        
        mProjectID = gbSearchStr
        mSourceOfFundID = gbSearchID
    
        If val(gbSearchStr) > 0 Then
            objProj.SetProject mProjectID, mYearID
            If objProj.ProjectID > 0 Then
                txtProjectNameEng.Text = objProj.ProjectNameEnglish
                txtProjectNo.Text = objProj.ProjectSerialNo
                txtProjectNo.Tag = objProj.ProjectID
                txtCategory.Tag = objProj.ProjCatID
                txtCategory.Text = objProj.Category
                txtSourceOfFund.Tag = mSourceOfFundID
                txtSourceOfFund.Text = objProj.FindSourceOfFund(mSourceOfFundID)
                txtSourceOfFund.Enabled = False
                txtCategory.Enabled = False
                mSubsectorID = objProj.SubSectorID
                
                Set mCol = objProj.GetFundDetails(CInt(mYearID), objProj.ProjectID)
                For mRow = 1 To mCol.count
                    Set objProFund = mCol.Item(mRow)
                    If objProFund.SourceOfFundID = mSourceOfFundID Then
                        mPProjectCost = objProFund.SourceWiseAmount
                        txtPAmount.Text = mPProjectCost
                        Exit For
                    End If
                Next mRow
            End If
            
''''            If mLoadMode <> 50 Then ' 50 = Benifishery Contribution   'COMMENTED BY MINU ON 18/MAY/2013
''''            If val(txtAmount.Text) <> val(txtPAmount.Text) Then
''''                MsgBox "Amount Not Matching with the Journal", vbInformation, "Saankhya"
''''                txtProjectNo.Text = ""
''''                txtCategory.Text = ""
''''                txtSourceOfFund.Text = ""
''''                txtPAmount.Text = ""
''''                txtProjectNameEng.Text = ""
''''                Exit Sub
''''            End If
''''            End If
            
            If mSubsectorID > 0 Then
                If objDb.SetConnection(mcnn) Then
                    mSql = "SELECT * FROM faSubSectorHeads "
                    mSql = mSql + "  INNER JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faSubSectorHeads.intFunctionID"
                    mSql = mSql + " WHERE intSubSectorID = " & mSubsectorID & " And intCategoryID = " & val(txtCategory.Tag)
                    Rec.Open mSql, mcnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mPFunctionID = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                        mPFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                        mPAccHeadID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                        mPAccHeadCode = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    End If
                    Rec.Close
                End If
            End If
            If objDb.CreateNewConnection(mCnPlan, enuSourceString.Sulekha) Then
                mSql = " SELECT MicroSector.intMicroSecID  FROM MicroSector WHERE decProjectID = " & val(txtProjectNo.Tag)
                Rec.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
                If Not (Rec.BOF And Rec.EOF) Then
                    mPMicroSectorID = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                End If
                Rec.Close
                mSql = "SELECT     SubjectCheckList.nchApprovalNo, SubjectCheckList.dtApprovaldate"
                mSql = mSql + "  FROM  ProjectDetails INNER JOIN"
                mSql = mSql + "  SubjectCheckList ON ProjectDetails.decProjectID = SubjectCheckList.decProjectID"
                mSql = mSql + "  WHERE ProjectDetails.decProjectID = " & val(txtProjectNo.Tag)
                Rec.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
                If Not (Rec.BOF And Rec.EOF) Then
                    mDPCApprovalNo = IIf(IsNull(Rec!nchApprovalNo), "", Rec!nchApprovalNo)
                    mDPCApprovalDate = DdMmmYy(IIf(IsNull(Rec!dtApprovalDate), 0, Rec!dtApprovalDate))
                End If
                Rec.Close
            End If
        End If
    End Sub

Private Function ValidateCreditAccHead() As Boolean
    Dim mcnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mCrAccHeadId As Integer
    Dim mJrCrAccHeadId As Integer
    
    If objDb.SetConnection(mcnn) Then
        If val(txtLetterOfAuthority.Tag) > 0 Then
            mSql = " SELECT * FROM faAllotmentLetters "
            mSql = mSql + " WHERE intAllotmentID= " & val(txtLetterOfAuthority.Tag) & " "
            Rec.Open mSql, mcnn
            If Not (Rec.EOF And Rec.BOF) Then
                 'mCrAccHeadId = IIf(IsNull(Rec!intCrAccountHeadID), 0, Rec!intCrAccountHeadID)
                  mCrAccHeadId = IIf(IsNull(Rec!intGrossAccountHeadID), 0, Rec!intGrossAccountHeadID)
            End If
            Rec.Close
        End If
        If val(txtVoucherNo.Tag) > 0 Then
            mSql = " SELECT * FROM faTransactionChild "
            mSql = mSql + " WHERE intTransactionID = (Select intTransactionID from faTransactions Where "
            mSql = mSql + " intVoucherID= " & val(txtVoucherNo.Tag) & " ) And tinDebitOrCreditFlag=0 "
            Rec.Open mSql, mcnn
            If Not (Rec.EOF And Rec.BOF) Then
                 mJrCrAccHeadId = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
            End If
            Rec.Close
        End If
        If mCrAccHeadId = mJrCrAccHeadId Then
            ValidateCreditAccHead = True
        Else
            ValidateCreditAccHead = False
        End If
    End If
End Function

Private Sub FillGrid()
    Dim mcnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mRowCnt As Integer
    Dim objProj As New clsProject
    Dim mProjectID As Variant
    
    On Error GoTo err
    objDb.CreateNewConnection mcnn, enuSourceString.Saankhya
    
    mSql = " SELECT * FROM faVoucherSub"
    mSql = mSql + " INNER JOIN faVouchers ON faVouchers.intVoucherID=faVoucherSub.intVoucherID"
    mSql = mSql + " INNER JOIN suSourceOfFund ON  suSourceOfFund.intSourceFundID=faVoucherSub.intSourceOfFundID"
    mSql = mSql + " INNER JOIN faTransactionCategory ON faTransactionCategory.intCategoryID=faVoucherSub.intCategoryID"
    mSql = mSql + " WHERE tnyVoucherTypeID = 40 "
    mSql = mSql + " AND intFinancialYearID = " & mYearID
    If mLoadMode = 50 Then
        Select Case mTransactionTypeID
        Case Is = 3008
            mSql = mSql + " And faVoucherSub.intTypeID = 2"
        Case Is = 3009
            mSql = mSql + " And faVoucherSub.intTypeID = 3"
        Case Is = 3010
            mSql = mSql + " And faVoucherSub.intTypeID = 4"
        Case Is = 3011
            mSql = mSql + " And faVoucherSub.intTypeID = 5"
        Case Is = 3012
            mSql = mSql + " And faVoucherSub.intTypeID = 6"
        Case Else
            mSql = mSql + " And faVoucherSub.intTypeID = 2"
        End Select
    Else
        mSql = mSql + " And faVoucherSub.intTypeID = 1"
    End If
    
    Rec.CursorLocation = adUseClient
    Rec.Open mSql, mcnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        mProjectID = IIf(IsNull(Rec!decProjectID), "", Rec!decProjectID)
        objProj.SetProject mProjectID, mYearID
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtDate), "", Rec!dtDate))
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        vsGrid.TextMatrix(mRowCnt, 3) = objProj.ProjectSerialNo
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
        vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
        If CheckTransferToSulekha(IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)) = True Then
            vsGrid.TextMatrix(mRowCnt, 6) = vbChecked
        Else
            vsGrid.TextMatrix(mRowCnt, 6) = vbUnchecked
        End If
        vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
        Rec.MoveNext
        mRowCnt = mRowCnt + 1
    Wend
    
    Rec.Close
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function SaveValidation() As Boolean
    If mLoadMode <> 50 Then          'BENEFICIARY CONTRIBUTIONS
        If Trim(txtLetterOfAuthority.Text) = "" Then
            MsgBox "Select Letter Of Authority", vbInformation, "Saankhya"
            SaveValidation = False
            Exit Function
        End If
    End If
    If Trim(txtVoucherNo.Text) = "" Then
        MsgBox "Select Voucher", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
    End If
    If Trim(txtProjectNo.Text) = "" Then
        MsgBox "Select Project", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
    End If
    If mYearID < 2012 Then
        MsgBox "Select a Valid Year", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
    End If
    SaveValidation = True
End Function
Private Sub FormInitialize()
    Dim ctrl    As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
            ctrl.Tag = ""
        ElseIf TypeOf ctrl Is OptionButton Then
            ctrl.value = False
        ElseIf TypeOf ctrl Is ComboBox Then
            'If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
            'ctrl.Tag = ""
        End If
    Next
End Sub
Private Sub GetDebitAccHeadDetails()
    Dim mcnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mJrDrAccHeadId As Integer
    
    If objDb.SetConnection(mcnn) Then
        mSql = " SELECT * FROM faVoucherChild "
        mSql = mSql + " WHERE tnyDebitOrCredit=1 And intVoucherID= " & val(txtVoucherNo.Tag) & " "
        Rec.Open mSql, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
             mJrDrAccHeadId = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
        End If
        Rec.Close
    End If
End Sub
Private Function CheckTransferToSulekha(mVoucherID As Variant) As Boolean
    Dim mCnnSulekha   As New ADODB.Connection
    Dim objDb   As New clsDB
    Dim mSql As String
    Dim arrInput As Variant
    Dim Rec As New ADODB.Recordset
    
    If (objDb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
        mSql = "Select * from ExpenseDetails where intVoucherID = " & mVoucherID & "  "
        Rec.Open mSql, mCnnSulekha
        If Not (Rec.EOF And Rec.BOF) Then
            CheckTransferToSulekha = True
        Else
            CheckTransferToSulekha = False
        End If
        Rec.Close
    Else
        MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
        Exit Function
    End If
    mCnnSulekha.Close
End Function

Private Sub txtProjectNo_KeyPress(KeyAscii As Integer)
    Call KeyPress(KeyAscii)
End Sub

Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
    Call KeyPress(KeyAscii)
End Sub

Private Sub txtVrDate_KeyPress(KeyAscii As Integer)
    Call KeyPress(KeyAscii)
End Sub
 Public Property Let LoadMode(mData As Integer)
    mLoadMode = mData
End Property

Public Property Get LoadMode() As Integer
    LoadMode = mLoadMode
End Property
Private Function ValidateBeneficiaryAccHead(mVoucherID As Double) As Boolean
    Dim mcnn  As New ADODB.Connection
    Dim objDb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mBeneficiaryID As Integer
    Dim mAccHeadId  As Integer
    
    If objDb.SetConnection(mcnn) Then
        mSql = "SELECT * FROM faTransactionChild "
        mSql = mSql + " WHERE intTransactionID = (Select intTransactionID from faTransactions Where "
        mSql = mSql + " intVoucherID= " & val(txtVoucherNo.Tag) & " ) And tinDebitOrCreditFlag=0 "
        Rec.Open mSql, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
            mAccHeadId = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
        End If
        Rec.Close
    End If
    
    If gbLBPanchayat = 1 Then
        If cmbTransactionType.ListIndex > 0 Then
        Select Case cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                
        Case Is = 3009   ' LOAN AVAILED
            If mAccHeadId = 2185 Or mAccHeadId = 267 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        Case Is = 3010  ' State Sponsored Scheme Fund
            If mAccHeadId = 208 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        Case Is = 3011  ' Centrally Sponsored Scheme Fund
            If mAccHeadId = 235 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        
        Case Is = 3012  ' MNREGS - Centrally Sponsored Scheme Fund
            If mAccHeadId = 212 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        
        
        Case Else
            If mAccHeadId = 2185 Or mAccHeadId = 1047 Or mAccHeadId = 1668 Or mAccHeadId = 2185 Then   ''' added 1668,2185 on 3 aug 17
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        End Select
        Else
            ValidateBeneficiaryAccHead = False
        End If
    Else ':: MUNICIPALITY
        Select Case cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                
        Case Is = 3009   ' LOAN AVAILED
            If mAccHeadId = 2358 Or mAccHeadId = 252 Then 'If mAccHeadId = 2185 Or mAccHeadId = 267 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        Case Is = 3010  ' State Sponsored Scheme Fund
            If mAccHeadId = 231 Then 'If mAccHeadId = 208 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        Case Is = 3011  ' Centrally Sponsored Scheme Fund
            If mAccHeadId = 248 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        Case Else
            If mAccHeadId = 2358 Or mAccHeadId = 910 Then 'If mAccHeadId = 2355 Or mAccHeadId = 910 Then
                ValidateBeneficiaryAccHead = True
            Else
                ValidateBeneficiaryAccHead = False
            End If
        End Select
        
    End If
End Function

Private Sub vsGrid_Click()
    'If mLoadMode <> 50 Then
        cmdUndo.Enabled = True
    'End If
End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAuctionBidderList 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Auction Bidder List / Auction Day"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10740
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bidder List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   30
      TabIndex        =   6
      Top             =   135
      Width           =   10395
      Begin VB.TextBox txtTotBidAmt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   7605
         TabIndex        =   13
         Top             =   2865
         Width           =   1995
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   1065
         Left            =   1545
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   3030
      End
      Begin VSFlex8LCtl.VSFlexGrid fgBidderList 
         Height          =   2220
         Left            =   690
         TabIndex        =   7
         Top             =   330
         Width           =   8925
         _cx             =   15743
         _cy             =   3916
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColorFixed  =   -2147483629
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483629
         BackColorAlternate=   16761087
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAuctionBidderList.frx":0000
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
         Editable        =   2
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BID AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   6495
         TabIndex        =   12
         Top             =   2925
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   9
         Top             =   2760
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bidding Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   30
      TabIndex        =   2
      Top             =   4740
      Width           =   10410
      Begin VB.TextBox txtDemandNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1515
         TabIndex        =   11
         Top             =   900
         Width           =   2220
      End
      Begin VB.TextBox txtBidAmount 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1515
         TabIndex        =   5
         Top             =   540
         Width           =   2220
      End
      Begin VSFlex8LCtl.VSFlexGrid fgSuccessfulBidder 
         Height          =   900
         Left            =   4260
         TabIndex        =   3
         Top             =   420
         Width           =   5940
         _cx             =   10477
         _cy             =   1587
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColorAlternate=   16761087
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
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAuctionBidderList.frx":0121
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
         Editable        =   2
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demand No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   465
         TabIndex        =   10
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BID AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   420
         TabIndex        =   4
         Top             =   615
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdTopBidders 
      Caption         =   "Generate Demand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4020
      TabIndex        =   1
      Top             =   4200
      Width           =   2370
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   10485
      Top             =   6720
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSaveSuccessfulBidder 
      Caption         =   "Save Successful Bidder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4020
      TabIndex        =   0
      Top             =   6420
      Width           =   2370
   End
End
Attribute VB_Name = "frmAuctionBidderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mAuctionTypeID As Integer
Dim mAddressID As Integer
Dim mDepositAmount As Double
Dim mAuctionID As Integer
Dim mDepositID As Double

Private Sub cmdSaveSuccessfulBidder_Click()
    If txtBidAmount.Text <> "" Then
        If MsgBox("Are you sure that you have ticked the Successful Bidder", vbYesNo) = vbYes Then
            Dim mSQL As String
            Dim objDB As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.RecordSet
            Dim mRowCount As Integer
                objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
            For mRowCount = 1 To 2
                If fgSuccessfulBidder.Cell(flexcpChecked, mRowCount, 1) = vbChecked Then
                    mDepositID = fgSuccessfulBidder.TextMatrix(mRowCount, 2)
                End If
            Next
            mSQL = "Update smAuctionDeposit Set tnyBidderStatus = 1 where numDepositID = " & mDepositID
            'Rec.Open mSQL, mCnn
            mCnn.Execute mSQL
            Call DemandFirstFace
            MsgBox "Congratulations to the successful Bidder", vbInformation
        Else
            MsgBox "Please Select the Successful Bidder from the List", vbInformation
        End If
    Else
        MsgBox "Please Export the Top 2 Bidders before saving the successful Bidder", vbCritical
        cmdTopBidders.SetFocus
    End If
End Sub

Private Sub DemandFirstFace()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim aryIn As Variant
    Dim AryOut As Variant
    Dim mFirstFace As Double
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    aryIn = Array(171, _
                    107, _
                    200, _
                    1, _
                    47, _
                    gbFinancialYearID, _
                    9, _
                    Date, _
                    mAuctionTypeID, _
                    Null, _
                    Null, _
                    txtRemarks.Text, _
                    0, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null _
                )
    objDB.ExecuteSP "spSaveIDemandTBL", aryIn, AryOut, True, mCnn, adCmdStoredProc
    
    aryIn = Array(AryOut(0, 0), _
                    171, _
                    1, _
                    1039, _
                    340200200, _
                    (Val(txtBidAmount.Text) * 25) / 100, _
                    txtRemarks.Text, _
                    0, _
                    Date _
                   )
    objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
    mCnn.Close
    objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    aryIn = Array(mAuctionID, _
                    mDepositID, _
                    AryOut(0, 0), _
                    340100100, _
                    47, _
                    (Val(txtBidAmount.Text) * 25) / 100, _
                    Date, _
                    Null, _
                    Null, _
                    txtRemarks.Text, _
                    0 _
                )
    objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
    txtDemandNo.Text = AryOut(0, 0)
End Sub
Private Sub SaveDemand()
''''    Dim objDB As New clsDB
''''    Dim mCnn As New ADODB.Connection
''''    Dim Rec As New ADODB.Recordset
''''    Dim aryIn As Variant
''''    Dim aryOut As Variant
''''    Dim mSql As String
''''    '   ---------           Procedure to Save the Demand to Finance DataBase        ----------  '
''''    objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
''''    mSql = "Select * from smAuctionDeposit where tnyBidderStatus = 2"
''''    Rec.Open mSql, mCnn
''''
''''    aryIn = Array(171, _
''''                    107, _
''''                    200, _
''''                    1, _
''''                    47, _
''''                    gbFinancialYearID, _
''''                    9, _
''''                    Date, _
''''                    mAuctionTypeID, _
''''                    Null, _
''''                    mAddressID, _
''''                    txtRemarks.Text, _
''''                    0, _
''''                    Null, _
''''                    Null, _
''''                    Null, _
''''                    Null _
''''                )
''''        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
''''    objDB.ExecuteSP "spSaveIDemandTBL", aryIn, aryOut, True, mCnn, adCmdStoredProc
''''    aryIn = Array(aryOut(0, 0), _
''''                    171, _
''''                    1, _
''''                    1039, _
''''                    340200200, _
''''                    mDepositAmount, _
''''                    txtRemarks.Text, _
''''                    0, _
''''                    Date _
''''                   )
''''    objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
''''    '   ---             Demand of the Second successful Bidder      ---    '
''''    aryIn = Array(171, _
''''                    107, _
''''                    200, _
''''                    1, _
''''                    47, _
''''                    gbFinancialYearID, _
''''                    9, _
''''                    Date, _
''''                    mAuctionTypeID, _
''''                    Null, _
''''                    mAddressID, _
''''                    txtRemarks.Text, _
''''                    0, _
''''                    Null, _
''''                    Null, _
''''                    Null, _
''''                    Null _
''''                )
''''    objDB.ExecuteSP "spSaveIDemandTBL", aryIn, aryOut, True, mCnn, adCmdStoredProc
''''    aryIn = Array(aryOut(0, 0), _
''''                    171, _
''''                    1, _
''''                    1039, _
''''                    340200200, _
''''                    mDepositAmount, _
''''                    txtRemarks.Text, _
''''                    0, _
''''                    Date _
''''                   )
''''    objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
''''
''''    '''''''''''''''''''''''''''''''''''''''''''
End Sub

    Private Sub cmdTopBidders_Click()
        If txtTotBidAmt.Text <> "" Then
            If MsgBox("Are you sure that Top 2 Bidders are Marked", vbYesNo) = vbYes Then
                Dim objDB As New clsDB
                Dim mCnn As New ADODB.Connection
                Dim Rec As New ADODB.RecordSet
                Dim mSQL As String
                Dim mRowCount As Integer
                Dim mRowSecond As Integer
                Dim aryIn As Variant
                Dim AryOut As Variant
                Dim mSqlAddress As String
                
                '       Declarations for faIDemandTBL      '
                
                Dim intLBID As Integer
                Dim tnyExtAppID As Integer
                Dim tnyExtModuleID As Integer
                Dim tnyDemandType As Integer
                Dim intTransactionTypeID  As Integer
                Dim intYearID As Integer
                Dim tnyPeriodID As Integer
                Dim dtDemandDate As Variant
                Dim numSubLedgerID As Variant
                Dim intKeyID As Integer
                Dim intKeyID2 As Integer
                Dim vchRemarks As Variant
                Dim tnyStatus As Integer
                Dim intVoucherID As Integer
                Dim dtVoucherDate As Variant
                Dim tnyArrearFlag  As Integer       '     TinyInt = 0,
                Dim dtExpiryDate As Variant         '       SmallDateTime   = Null,
                Dim numDemandID  As Variant         '       Numeric     = Null Output,
                
                Dim intFinancialYearID As Integer   '     Int     = Null,
                Dim numSeatID As Variant            '     Numeric     = Null,
                Dim intSectionID As Integer         '       Int         = Null,
                Dim numUserID As Variant            '      Numeric     = Null,
                Dim numCounterID As Variant         '       Numeric     = Null,
                Dim vchAdminNote As Variant         '   varChar(100)    = Null,
                Dim vchDemandNo As Variant          '        varChar(20)     = Null,
                Dim numZoneID As Variant            '      Numeric     = Null,
                Dim intWardNo As Integer            '     Int         = Null,
                Dim intDoorNo As Integer            '      Int         = Null,
                Dim vchDoorNo2 As Variant           '     varChar(10)     = Null,
                Dim numForwardedSeatID As Variant   ' Numeric     = Null
                
                Dim fltAmount As Double
                
                Dim vchName As Variant
                Dim vchInit1 As Variant
                Dim vchInit2 As Variant
                Dim vchInit3 As Variant
                Dim vchInit4 As Variant
                Dim vchHouseName As Variant
                Dim vchStreet As Variant
                Dim vchLocalPlace As Variant
                Dim vchMainPlace As Variant
                Dim vchPost As Variant
                Dim vchPin As Variant
                Dim vchPhone As Variant
                Dim tnyFlag As Variant
                
                Dim tnySlNo As Variant
                
                Dim intAccountHeadID As Integer
                Dim vchAccountHeadCode As Variant
                
                Dim dtOnDate As Variant
                
                
                '---------------------------------------------------------------------'
                '                           Definitions                               '
                '---------------------------------------------------------------------'
                
                intLBID = gbLocalBodyID
                tnyExtAppID = AppID.Sanjaya
                tnyExtModuleID = 200
                tnyDemandType = 1
                intTransactionTypeID = 47
                intYearID = Year(Date)
                'tnyPeriodID = Null
                dtDemandDate = Date
                numSubLedgerID = mAuctionTypeID
                'intKeyID = Null
                'intKeyID2 = Null
                vchRemarks = txtRemarks.Text
                tnyStatus = 0
                'intVoucherID = Null
                dtVoucherDate = Null
                'tnyArrearFlag = Null
                dtExpiryDate = Null
                numDemandID = Null
                
                intFinancialYearID = gbFinancialYearID
                numSeatID = gbSeatID
                intSectionID = gbSectionID
                numUserID = Null
                numCounterID = Null
                vchAdminNote = Null
                vchDemandNo = Null
                numZoneID = Null
                'intWardNo = Null
                'intDoorNo = Null
                vchDoorNo2 = Null
                numForwardedSeatID = Null
                
                intAccountHeadID = 1010
                vchAccountHeadCode = 340100100
                
                dtOnDate = Date
                
                '---------------------------------------------------------------------'
                    objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
                
                Call SaveAuction        '       Auction Bidder Finalised        '
                txtBidAmount.Text = txtTotBidAmt.Text
                mRowSecond = 1
                For mRowCount = 1 To fgBidderList.Rows - 1
                    Dim mSlNo As Integer
                    mSlNo = 1
                    If fgBidderList.Cell(flexcpChecked, mRowCount, 5) = 1 Then
                        fgSuccessfulBidder.TextMatrix(mRowSecond, 0) = fgBidderList.TextMatrix(mRowCount, 2)
                        fgSuccessfulBidder.Cell(flexcpChecked, mRowSecond, 1) = 2
                        fgSuccessfulBidder.TextMatrix(mRowSecond, 2) = fgBidderList.TextMatrix(mRowCount, 6)
                        mRowSecond = mRowSecond + 1
                        mSQL = "Update smAuctionDeposit Set tnyBidderStatus = 2 where numDepositID = " & fgBidderList.TextMatrix(mRowCount, 6)
                        mCnn.Execute mSQL
                        mSqlAddress = "Select * from smAddressBook where intAddressID = " & fgBidderList.TextMatrix(mRowCount, 7)
                        Rec.Open mSqlAddress, mCnn
                        
                        vchName = Rec!vchName
                        vchInit1 = Null
                        vchInit2 = Null
                        vchInit3 = Null
                        vchInit4 = Null
                        vchHouseName = Rec!vchAddress1
                        vchStreet = Rec!vchAddress2
                        vchLocalPlace = Rec!vchPlace
                        vchMainPlace = Rec!vchAddress3
                        vchPost = Null
                        vchPin = Rec!vchPin
                        vchPhone = Rec!vchPhone
                        tnyFlag = Null
                        
                        '---------------------------------------------------'
                        '   ---     Demand to Finance Databse       ---     '
                        '---------------------------------------------------'
                        mCnn.Close
                        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                        aryIn = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                        tnyDemandType, intTransactionTypeID, intYearID, _
                                        tnyPeriodID, _
                                        dtDemandDate, _
                                        numSubLedgerID, _
                                        intKeyID, intKeyID2, vchRemarks, _
                                        tnyStatus, _
                                        intVoucherID, _
                                        dtVoucherDate, _
                                        tnyArrearFlag, _
                                        dtExpiryDate, _
                                        numDemandID, _
                                        intFinancialYearID, _
                                        numSeatID, _
                                        intSectionID, _
                                        numUserID, _
                                        numCounterID, _
                                        vchAdminNote, _
                                        vchDemandNo, _
                                        numZoneID, _
                                        intWardNo, _
                                        intDoorNo, _
                                        vchDoorNo2, _
                                        numForwardedSeatID _
                                )
                        objDB.ExecuteSP "spSaveIDemandTBL", aryIn, AryOut, True, mCnn, adCmdStoredProc
                        
                        numDemandID = AryOut(0, 0)
                        fltAmount = mDepositAmount
                        tnySlNo = mSlNo
                        
                        '-------------------------------------------------------------------------'
                        '------------- Save to Auction Demand Child Table in Saankhya-------------'
                        '-------------------------------------------------------------------------'
                        aryIn = Array(numDemandID, _
                                        intLBID, _
                                        tnySlNo, _
                                        intAccountHeadID, _
                                        vchAccountHeadCode, _
                                        fltAmount, _
                                        vchRemarks, _
                                        tnyStatus, _
                                        dtDemandDate, _
                                        intFinancialYearID, _
                                        tnyPeriodID, _
                                        tnyArrearFlag _
                                       )
                        objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
                        '-----------------------------------------------------------'
                        '           Saving the Address Table of Finance Demand      '
                        '-----------------------------------------------------------'
                        
                        
                        aryIn = Array(numDemandID, _
                                        numZoneID, _
                                        intWardNo, _
                                        intDoorNo, _
                                        vchDoorNo2, _
                                        vchName, _
                                        vchInit1, _
                                        vchInit2, _
                                        vchInit3, _
                                        vchInit4, _
                                        vchHouseName, _
                                        vchStreet, _
                                        vchLocalPlace, _
                                        vchMainPlace, _
                                        vchPost, _
                                        vchPin, _
                                        vchPhone, _
                                        tnyFlag _
                                        )
                        objDB.ExecuteSP "spSaveIDemandAddress", aryIn, , , mCnn, adCmdStoredProc
                        mCnn.Close
                        '------------------------------------------------------------------'
                        '------- Save to Auction Demand Table in iSaankhyaMasters ---------'
                        '------------------------------------------------------------------'
                        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
                        aryIn = Array(mAuctionID, _
                                        fgBidderList.TextMatrix(mRowCount, 6), _
                                        AryOut(0, 0), _
                                        340100100, _
                                        47, _
                                        mDepositAmount, _
                                        Date, _
                                        Null, _
                                        Null, _
                                        txtRemarks.Text, _
                                        1 _
                                    )
                        objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
                        txtDemandNo.Text = AryOut(0, 0)
                    End If
                    mSlNo = mSlNo + 1
                Next
            Else
                MsgBox "Plese Mention the Top 2 Bidders", vbInformation
            End If
        Else
            MsgBox "Please Give the Total Bid Amount", vbCritical
            txtTotBidAmt.SetFocus
        End If
    End Sub

Private Sub Form_Load()
    WindowsXPC1.InitIDESubClassing
    frmAuctionBidderList.Width = 10860
    frmAuctionBidderList.Height = 7425
    FillGrid
End Sub
Private Sub SaveAuction()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim aryIn As Variant
    Dim AryOut As Variant
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    aryIn = Array(mAuctionID, _
                            frmDemandForAuctionDeposit.txtAuctionTitle.Text, _
                            Date, _
                            frmDemandForAuctionDeposit.txtAuctionNo.Text, _
                            frmDemandForAuctionDeposit.txtFormNo.Text, _
                            mAuctionTypeID, _
                            Null, _
                            mDepositAmount, _
                            txtTotBidAmt.Text _
                            )
    objDB.ExecuteSP "spSaveAuction", aryIn, AryOut, True, mCnn, adCmdStoredProc
    mAuctionID = AryOut(0, 0)
End Sub
Private Sub FillGrid()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.RecordSet
    Dim mSQL As String
    Dim mRowCount As Integer
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "Select intAuctionTypeID,smAuctionDeposit.vchAuctionNo,vchFormNo,fltDepositAmount,vchName,numDepositID,tnyBidderStatus,smAuctionDeposit.intAuctionID,smAuctionDeposit.intAddressID from smAuctionDeposit"
    mSQL = mSQL + " Inner Join smAddressBook on smAuctionDeposit.intAddressID = smAddressBook.intAddressID"
    mSQL = mSQL + " Inner Join smAuctionTitles on smAuctionTitles.vchAuctionNo = smAuctionDeposit.vchAuctionNo"
    mSQL = mSQL + " Where smAuctionDeposit.vchAuctionNo = '" & frmDemandForAuctionDeposit.txtAuctionNo.Text & "'"
    Rec.Open mSQL, mCnn
    mRowCount = 1
    fgBidderList.Rows = 2
    While Not Rec.EOF And Not Rec.BOF
        fgBidderList.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAuctionNo), "", Rec!vchAuctionNo)
        fgBidderList.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchFormNo), "", Rec!vchFormNo)
        fgBidderList.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
        fgBidderList.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltDepositAmount), "", Rec!fltDepositAmount)
'''        If Rec!tnyBidderStatus = 1 Then
'''            fgBidderList.Cell(flexcpChecked, mRowCount, 4) = 1  '   First Successful Bidder     '
'''        Else
'''            fgBidderList.Cell(flexcpChecked, mRowCount, 4) = 2
'''        End If
'''        If Rec!tnyBidderStatus = 2 Then                         '   Second Successful Bidder    '
'''            fgBidderList.Cell(flexcpChecked, mRowCount, 5) = 1
'''        Else
'''            fgBidderList.Cell(flexcpChecked, mRowCount, 5) = 2
'''        End If
        If Rec!tnyBidderStatus = 1 Then
            fgBidderList.Cell(flexcpChecked, mRowCount, 5) = vbChecked
        ElseIf Rec!tnyBidderStatus = 2 Then
            fgBidderList.Cell(flexcpChecked, mRowCount, 5) = vbChecked
        Else
            fgBidderList.Cell(flexcpChecked, mRowCount, 5) = 2
        End If
        mDepositAmount = IIf(IsNull(Rec!fltDepositAmount), "", Rec!fltDepositAmount)
        fgBidderList.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!numDepositID), "", Rec!numDepositID)
        fgBidderList.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intAddressID), "", Rec!intAddressID)
        mAuctionTypeID = Rec!intAuctionTypeID
        mAddressID = Rec!intAddressID
        mAuctionTypeID = Rec!intAuctionTypeID
        mAuctionID = Rec!intAuctionID
        Rec.MoveNext
        mRowCount = mRowCount + 1
        fgBidderList.Rows = fgBidderList.Rows + 1
    Wend
End Sub


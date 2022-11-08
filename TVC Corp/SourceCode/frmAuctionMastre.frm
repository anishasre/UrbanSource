VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAuctionMastre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define An Auction"
   ClientHeight    =   6045
   ClientLeft      =   975
   ClientTop       =   1560
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMasters 
      BackColor       =   &H00CBECEC&
      Height          =   1635
      Left            =   5460
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auction Demand"
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   75
      TabIndex        =   22
      Top             =   2145
      Width           =   11145
      Begin VB.CheckBox chkGenerateDemand 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generate Demand"
         Height          =   240
         Left            =   9045
         TabIndex        =   32
         Top             =   3090
         Width           =   1920
      End
      Begin VB.TextBox txtBidAmount 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   8385
         TabIndex        =   31
         Top             =   300
         Width           =   2070
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   2430
         TabIndex        =   10
         Top             =   3090
         Width           =   2550
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   9000
         TabIndex        =   12
         Top             =   3435
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   609
         _Version        =   393216
         Format          =   58130433
         CurrentDate     =   39693
      End
      Begin VB.TextBox txtCouncilDate 
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
         ForeColor       =   &H80000003&
         Height          =   330
         Left            =   7455
         TabIndex        =   11
         Top             =   3435
         Width           =   1530
      End
      Begin VB.ListBox lstAuctionNumbers 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   1380
         Left            =   165
         TabIndex        =   7
         Top             =   870
         Width           =   1920
      End
      Begin VB.TextBox txtBidderName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   3585
         TabIndex        =   26
         Top             =   285
         Width           =   3585
      End
      Begin VB.CommandButton cmdViewDemand 
         BackColor       =   &H8000000D&
         Caption         =   "&View Demand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   465
         MaskColor       =   &H80000002&
         TabIndex        =   9
         Top             =   2655
         Width           =   1380
      End
      Begin VB.TextBox txtAuctionNoToList 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   180
         TabIndex        =   8
         Top             =   2295
         Width           =   1905
      End
      Begin VB.CommandButton cmdGenerateDemand 
         Caption         =   "&Generate Demand"
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
         Height          =   390
         Left            =   9300
         TabIndex        =   13
         Top             =   3390
         Width           =   1665
      End
      Begin VSFlex8LCtl.VSFlexGrid fgDemandGrid 
         Height          =   2205
         Left            =   2415
         TabIndex        =   23
         Top             =   810
         Width           =   8580
         _cx             =   15134
         _cy             =   3889
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAuctionMastre.frx":0000
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bid Amount"
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   7335
         TabIndex        =   30
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   225
         Left            =   1575
         TabIndex        =   29
         Top             =   3090
         Width           =   765
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Council Approval Date"
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
         Left            =   5520
         TabIndex        =   28
         Top             =   3495
         Width           =   1860
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bidder Name"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2400
         TabIndex        =   25
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Number"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   510
         TabIndex        =   24
         Top             =   375
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auction Define"
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   60
      TabIndex        =   15
      Top             =   30
      Width           =   11160
      Begin VB.TextBox txtFileNo 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1890
         TabIndex        =   2
         Top             =   1065
         Width           =   3165
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5550
         TabIndex        =   5
         Top             =   1530
         Width           =   1080
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6675
         TabIndex        =   6
         Top             =   1530
         Width           =   1080
      End
      Begin VB.TextBox txtAuctionTitle 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1890
         TabIndex        =   1
         Top             =   735
         Width           =   3165
      End
      Begin VB.TextBox txtAuctionNumber 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   9180
         TabIndex        =   17
         Top             =   390
         Width           =   1290
      End
      Begin VB.TextBox txtAuctionType 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1890
         TabIndex        =   0
         Top             =   405
         Width           =   3165
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5085
         TabIndex        =   16
         Top             =   405
         Width           =   285
      End
      Begin VB.TextBox txtAuctionDate 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   1890
         TabIndex        =   3
         Top             =   1440
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3195
         TabIndex        =   4
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58130433
         CurrentDate     =   39676
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File No"
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   1215
         TabIndex        =   27
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Title"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   825
         TabIndex        =   21
         Top             =   765
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Number"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   7770
         TabIndex        =   20
         Top             =   420
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Type"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   810
         TabIndex        =   19
         Top             =   435
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Date"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   810
         TabIndex        =   18
         Top             =   1470
         Width           =   1035
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   -3525
      Top             =   -390
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   3
      Common_Dialog   =   0   'False
   End
End
Attribute VB_Name = "frmAuctionMastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTotalBidAmount As Double
Dim mAuctionType As Integer
Dim strQry As String
Dim mAuctionID As Integer
Dim mDepositID As Variant

Private Sub chkGenerateDemand_Click()
    If chkGenerateDemand.Value = 1 Then
        cmdGenerateDemand.Enabled = True
    Else
        cmdGenerateDemand.Enabled = False
    End If
End Sub

Private Sub cmdClear_Click()
    txtAuctionDate.Text = ""
    txtAuctionNumber.Text = ""
    txtAuctionTitle.Text = ""
    txtAuctionType.Text = ""
End Sub

Private Sub cmdGenerateDemand_Click()
    If txtCouncilDate.Text <> "" And fgDemandGrid.Cell(flexcpChecked, 1, 5) = vbChecked Then
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim aryIn As Variant
        Dim AryOut As Variant
        Dim mDemandOnBidDate As Double
        Dim mBalanceAmt As Double
        Dim mAmtAfterCouncilDecision As Double
        Dim mInstalmentTotalAmt As Double
        Dim mInstalmentRates As Double
    '''    Dim mFirstInstalment As Double
    '''    Dim mSecInstalment As Double
    '''    Dim mThirdInstalment As Double
        '       Demand amount details       '
        mAmtAfterCouncilDecision = (mTotalBidAmount * 40) / 100
        mDemandOnBidDate = (mTotalBidAmount * 25) / 100
        mBalanceAmt = mTotalBidAmount - (fgDemandGrid.TextMatrix(1, 2) + mDemandOnBidDate + mAmtAfterCouncilDecision)
        mInstalmentTotalAmt = (mBalanceAmt * 35) / 100
        mInstalmentRates = mInstalmentTotalAmt / 3
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        aryIn = Array(171, _
                107, _
                200, _
                1, _
                47, _
                gbFinancialYearID, _
                9, _
                txtCouncilDate.Text, _
                mAuctionType, _
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
                        (Val(mTotalBidAmount) * 25) / 100, _
                        txtRemarks.Text, _
                        0, _
                        txtCouncilDate.Text _
                       )
        objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        aryIn = Array(mAuctionID, _
                mDepositID, _
                AryOut(0, 0), _
                340100100, _
                47, _
                (Val(mTotalBidAmount) * 45) / 100, _
                txtCouncilDate.Text, _
                Null, _
                Null, _
                txtRemarks.Text, _
                0 _
            )
        objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
        
        '      ---      First Instalmets        ---     '
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        DTPicker2.Month = DTPicker2.Month + 1
        aryIn = Array(171, _
                107, _
                200, _
                1, _
                47, _
                gbFinancialYearID, _
                9, _
                DTPicker2.Value, _
                mAuctionType, _
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
                        mInstalmentRates, _
                        txtRemarks.Text, _
                        0, _
                        DTPicker2.Value _
                       )
        objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        aryIn = Array(mAuctionID, _
                mDepositID, _
                AryOut(0, 0), _
                340100100, _
                47, _
                mInstalmentRates, _
                DTPicker2.Value, _
                Null, _
                Null, _
                txtRemarks.Text, _
                0 _
            )
        objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
        '       ---         Second Instalment       ---     '
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        DTPicker2.Month = DTPicker2.Month + 1
        aryIn = Array(171, _
                107, _
                200, _
                1, _
                47, _
                gbFinancialYearID, _
                9, _
                DTPicker2.Value, _
                mAuctionType, _
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
                        mInstalmentRates, _
                        txtRemarks.Text, _
                        0, _
                        DTPicker2.Value _
                       )
        objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        aryIn = Array(mAuctionID, _
                mDepositID, _
                AryOut(0, 0), _
                340100100, _
                47, _
                mInstalmentRates, _
                DTPicker2.Value, _
                Null, _
                Null, _
                txtRemarks.Text, _
                0 _
            )
        objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
        '       ---         Third and Final Instalment      ---     '
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        DTPicker2.Month = DTPicker2.Month + 1
        aryIn = Array(171, _
                107, _
                200, _
                1, _
                47, _
                gbFinancialYearID, _
                9, _
                DTPicker2.Value, _
                mAuctionType, _
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
                        mInstalmentRates, _
                        txtRemarks.Text, _
                        0, _
                        DTPicker2.Value _
                       )
        objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
        mCnn.Close
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        aryIn = Array(mAuctionID, _
                mDepositID, _
                AryOut(0, 0), _
                340100100, _
                47, _
                mInstalmentRates, _
                DTPicker2.Value, _
                Null, _
                Null, _
                txtRemarks.Text, _
                0 _
            )
        objDB.ExecuteSP "spSaveAuctionDemand", aryIn, AryOut, True, mCnn, adCmdStoredProc
        MsgBox "Demand Gerated Successfully", vbInformation
    Else
        MsgBox "Please Enter the Council Decision Date and Click the respective Auction Number from The List Box", vbCritical
    End If
End Sub

Private Sub cmdSave_Click()
    If txtAuctionDate.Text <> "" And txtAuctionTitle.Text <> "" And txtFileNo.Text <> "" And txtAuctionType.Text <> "" Then
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.RecordSet
        Dim aryIn As Variant
        Dim AryOut As Variant
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        aryIn = Array(lstMasters.ItemData(lstMasters.ListIndex), _
                        txtAuctionTitle.Text, _
                        txtAuctionDate.Text, _
                        txtFileNo.Text, _
                        Null, _
                        gbFinancialYearID, _
                        171, _
                        0 _
                       )
        objDB.ExecuteSP "spSaveAuctionTitles", aryIn, AryOut, , mCnn
        txtAuctionNumber.Text = AryOut(0, 0)
        MsgBox "Saved Successfully", vbInformation
    Else
        MsgBox " Please Fill the AuctionType,Title,Date and FileNo", vbCritical
    End If
End Sub

Private Sub cmdViewDemand_Click()
    If txtAuctionNoToList.Text <> "" Then
''''        Dim objDB As New clsDB
''''        Dim mCnn As New ADODB.Connection
''''        Dim Rec As New ADODB.Recordset
''''        Dim mSQL As String
''''        Dim mRowCount As Integer
''''            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        strQry = "SELECT  smAuctionDeposit.intAddressID, smAddressBook.vchName, smAuctionDemand.vchDemandNo, smAuctionDemand.fltAmount,smAuctionDemand.vchAccountHeadCode,smAuctionDemand.intAuctionID,smAuctionDemand.numDepositID,smAuctionDemand.tnyPayStatus,"
        strQry = strQry + " smAuctionDemand.vchDemandNo, smAuctionDemand.dtDemandDate,smAuctionDeposit.vchAuctionNo,smAuctionDeposit.tnyBidderStatus,smAuctions.fltTotalBidAmount,smAuctionTitles.intAuctionTypeID"
        strQry = strQry + " FROM    smAuctionDeposit INNER JOIN "
        strQry = strQry + " smAuctionDemand ON smAuctionDeposit.numDepositID = smAuctionDemand.numDepositID INNER JOIN"
        strQry = strQry + " smAuctions ON smAuctionDeposit.intAuctionID = smAuctions.intAuctionID INNER JOIN"
        strQry = strQry + " smAddressBook ON smAuctionDeposit.intAddressID = smAddressBook.intAddressID INNER JOIN"
        strQry = strQry + " smAuctionTitles ON smAuctionTitles.intAuctionID = smAuctionDeposit.intAuctionID"
        strQry = strQry + " WHERE   (smAuctionDeposit.tnyBidderStatus = 1) And (smAuctionDeposit.vchAuctionNo = '" & txtAuctionNoToList.Text & "')"
''''        Rec.Open mSQL, mCnn
''''        If Not Rec.EOF Or Not Rec.BOF Then
''''            txtBidderName.Text = Rec!vchName
''''            fgDemandGrid.Rows = 2
''''            mRowCount = 1
''''            While Not Rec.EOF
''''                fgDemandGrid.TextMatrix(mRowCount, 0) = Rec!vchDemandNo
''''                fgDemandGrid.TextMatrix(mRowCount, 1) = Rec!vchAccountHeadCode
''''                fgDemandGrid.TextMatrix(mRowCount, 2) = Rec!fltAmount
''''                fgDemandGrid.TextMatrix(mRowCount, 3) = Rec!dtDemandDate
''''                fgDemandGrid.TextMatrix(mRowCount, 4) = ""
''''                fgDemandGrid.Cell(flexcpChecked, mRowCount, 5) = Rec!tnyBidderStatus
''''                mRowCount = mRowCount + 1
''''                fgDemandGrid.Rows = fgDemandGrid.Rows + 1
''''                Rec.MoveNext
''''                mTotalBidAmount = Rec!fltTotalBidAmount
''''                mAuctionType = Rec!intAuctionTypeID
''''            Wend
''''        Else
''''            MsgBox "No Record Exists", vbInformation
''''        End If
        Call FillDemandGrid
    Else
        MsgBox "Please insert the Auction Number", vbCritical
    End If
End Sub

Private Sub Command1_Click()
'    lstMasters.Left = 4530
'    lstMasters.Top = 870
    lstMasters.Visible = True
    Me.Refresh
    Dim mSQL As String
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "Select vchCollectionType,intCollectionTypeID from smMSTCollectionTypes Where tnyInAuctionList = 1 Order by intCollectionTypeID"
    Call PopulateList(lstMasters, mSQL, , , , True)
End Sub

Private Sub DTPicker1_CloseUp()
        txtAuctionDate.Text = DdMmmYy(DTPicker1.Value)
    End Sub
    
Private Sub DTPicker1_DropDown()
    If IsDate(txtAuctionDate) Then
        DTPicker1.Value = txtAuctionDate.Text
    End If
End Sub
Private Sub DTPicker2_CloseUp()
        txtCouncilDate.Text = DdMmmYy(DTPicker2.Value)
    End Sub
    
Private Sub DTPicker2_DropDown()
    If IsDate(txtCouncilDate) Then
        DTPicker2.Value = txtCouncilDate.Text
    End If
End Sub

Private Sub fgDemandGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Len(gbSearchStr) Then
        fgDemandGrid.TextMatrix(fgDemandGrid.Row, 1) = Token(gbSearchStr, " ")
    End If
    gbSearchStr = ""
    gbSearchID = -1
End Sub

Private Sub fgDemandGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType NOT IN (1) And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
    'frmSearchAccountHeads.Show vbModal
End Sub

Private Sub fgDemandGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgDemandGrid.Rows = fgDemandGrid.Rows + 1
    End If
End Sub

Private Sub Form_Load()
    XPC.InitSubClassing
    Dim mSQL As String
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "Select vchAuctionNo,intAuctionID from smAuctionTitles  Order by vchAuctionNo"
    Call PopulateList(lstAuctionNumbers, mSQL, , , , True)
    fgDemandGrid.ColComboList(1) = "|..."
    DTPicker1.Value = Date
    DTPicker2.Value = Date
End Sub

Private Sub FillGrid()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    mSQL = "Select * from smAuctionDemand"
End Sub

Private Sub lstAuctionNumbers_DblClick()
'''    Dim objDB As New clsDB
'''        Dim mCnn As New ADODB.Connection
'''        Dim Rec As New ADODB.Recordset
'''        Dim mSQL As String
'''        Dim mRowCount As Integer
'''            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        strQry = "SELECT  smAuctionDeposit.intAddressID, smAddressBook.vchName, smAuctionDemand.vchDemandNo, smAuctionDemand.fltAmount,smAuctionDemand.vchAccountHeadCode,smAuctionDemand.intAuctionID,smAuctionDemand.numDepositID,smAuctionDemand.tnyPayStatus,"
        strQry = strQry + " smAuctionDemand.vchDemandNo, smAuctionDemand.dtDemandDate,smAuctionDeposit.vchAuctionNo,smAuctionDeposit.tnyBidderStatus,smAuctions.fltTotalBidAmount,smAuctionTitles.intAuctionTypeID"
        strQry = strQry + " FROM    smAuctionDeposit INNER JOIN "
        strQry = strQry + " smAuctionDemand ON smAuctionDeposit.numDepositID = smAuctionDemand.numDepositID INNER JOIN"
        strQry = strQry + " smAuctions ON smAuctionDeposit.intAuctionID = smAuctions.intAuctionID INNER JOIN"
        strQry = strQry + " smAddressBook ON smAuctionDeposit.intAddressID = smAddressBook.intAddressID INNER JOIN"
        strQry = strQry + " smAuctionTitles ON smAuctionTitles.intAuctionID = smAuctionDeposit.intAuctionID"
        strQry = strQry + " WHERE   (smAuctionDeposit.tnyBidderStatus = 1) And (smAuctionDeposit.vchAuctionNo = '" & lstAuctionNumbers.Text & "')"
'''        Rec.Open mSQL, mCnn
'''        If Not Rec.EOF Or Not Rec.BOF Then
'''            txtBidderName.Text = Rec!vchName
'''            fgDemandGrid.Rows = 2
'''            mRowCount = 1
'''            While Not Rec.EOF
'''                fgDemandGrid.TextMatrix(mRowCount, 0) = Rec!vchDemandNo
'''                fgDemandGrid.TextMatrix(mRowCount, 1) = Rec!vchAccountHeadCode
'''                fgDemandGrid.TextMatrix(mRowCount, 2) = Rec!fltAmount
'''                fgDemandGrid.TextMatrix(mRowCount, 3) = Rec!dtDemandDate
'''                fgDemandGrid.TextMatrix(mRowCount, 4) = ""
'''                fgDemandGrid.Cell(flexcpChecked, mRowCount, 5) = Rec!tnyBidderStatus
'''                mRowCount = mRowCount + 1
'''                fgDemandGrid.Rows = fgDemandGrid.Rows + 1
'''                mTotalBidAmount = Rec!fltTotalBidAmount
'''                Rec.MoveNext
'''            Wend
'''        Else
'''            MsgBox "No Record Exists", vbInformation
'''            fgDemandGrid.Clear 1, 1
'''        End If
    Call FillDemandGrid
End Sub

Private Sub lstMasters_DblClick()
    txtAuctionType.Text = lstMasters.Text
    txtAuctionType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
    lstMasters.Visible = False
    txtAuctionTitle.SetFocus
End Sub

Private Sub lstMasters_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lstMasters.Visible = False
    End If
    txtAuctionType.Text = lstMasters.Text
    txtAuctionType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
    lstMasters.Visible = False
End Sub
Private Sub lstMasters_LostFocus()
    lstMasters.Visible = False
End Sub
Private Sub FillDemandGrid()
    Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.RecordSet
        Dim mSQL As String
        Dim mRowCount As Integer
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
'''        mSQL = "SELECT  smAuctionDeposit.intAddressID, smAddressBook.vchName, smAuctionDemand.vchDemandNo, smAuctionDemand.fltAmount,smAuctionDemand.vchAccountHeadCode,"
'''        mSQL = mSQL + " smAuctionDemand.vchDemandNo, smAuctionDemand.dtDemandDate,smAuctionDeposit.vchAuctionNo,smAuctionDeposit.tnyBidderStatus,smAuctions.fltTotalBidAmount,smAuctionTitles.intAuctionTypeID"
'''        mSQL = mSQL + " FROM    smAuctionDeposit INNER JOIN "
'''        mSQL = mSQL + " smAuctionDemand ON smAuctionDeposit.numDepositID = smAuctionDemand.numDepositID INNER JOIN"
'''        mSQL = mSQL + " smAuctions ON smAuctionDeposit.intAuctionID = smAuctions.intAuctionID INNER JOIN"
'''        mSQL = mSQL + " smAddressBook ON smAuctionDeposit.intAddressID = smAddressBook.intAddressID INNER JOIN"
'''        mSQL = mSQL + " smAuctionTitles ON smAuctionTitles.intAuctionID = smAuctionDeposit.intAuctionID"
'''        mSQL = mSQL + " WHERE   (smAuctionDeposit.tnyBidderStatus = 1) And (smAuctionDeposit.vchAuctionNo = '" & txtAuctionNoToList.Text & "')"
        Rec.Open strQry, mCnn
        If Not Rec.EOF Or Not Rec.BOF Then
            txtBidderName.Text = Rec!vchName
            fgDemandGrid.Rows = 2
            mRowCount = 1
            While Not Rec.EOF
                fgDemandGrid.TextMatrix(mRowCount, 0) = Rec!vchDemandNo
                fgDemandGrid.TextMatrix(mRowCount, 1) = Rec!vchAccountHeadCode
                fgDemandGrid.TextMatrix(mRowCount, 2) = Rec!fltAmount
                fgDemandGrid.TextMatrix(mRowCount, 3) = Rec!dtDemandDate
                fgDemandGrid.TextMatrix(mRowCount, 4) = ""
                If Rec!tnyPayStatus = 1 Then
                    fgDemandGrid.Cell(flexcpChecked, mRowCount, 5) = vbChecked
                Else
                    fgDemandGrid.Cell(flexcpChecked, mRowCount, 5) = 2
                End If
                'fgDem5andGrid.Cell(flexcpChecked, mRowCount, 5) = Rec!tnyPayStatus
                mRowCount = mRowCount + 1
                fgDemandGrid.Rows = fgDemandGrid.Rows + 1
                Rec.MoveNext
                'mTotalBidAmount = Rec!fltTotalBidAmount
                'mAuctionType = Rec!intAuctionTypeID
            Wend
            Rec.MoveFirst
            mTotalBidAmount = IIf(IsNull(Rec!fltTotalBidAmount), "", Rec!fltTotalBidAmount)
            mAuctionType = Rec!intAuctionTypeID
            mAuctionID = Rec!intAuctionID
            mDepositID = Rec!numDepositID
            txtBidAmount.Text = mTotalBidAmount
        Else
            MsgBox "No Record Exists", vbInformation
        End If
End Sub


VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmZonalDaily 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zonal Daily Collection"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   285
      Left            =   8730
      TabIndex        =   32
      Top             =   5580
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   9630
      TabIndex        =   21
      Top             =   5580
      Width           =   690
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4005
      TabIndex        =   19
      Top             =   5535
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Height          =   5370
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   10275
      Begin VB.Frame fraDate 
         Height          =   600
         Left            =   45
         TabIndex        =   22
         Top             =   180
         Width           =   7575
         Begin VB.ComboBox cmbYear 
            Height          =   315
            Left            =   720
            TabIndex        =   24
            Top             =   180
            Width           =   1545
         End
         Begin VB.ComboBox cmbMonth 
            Height          =   315
            Left            =   3105
            TabIndex        =   23
            Top             =   180
            Width           =   1545
         End
         Begin VB.Label lblYear 
            Caption         =   "Year"
            Height          =   240
            Left            =   180
            TabIndex        =   26
            Top             =   225
            Width           =   510
         End
         Begin VB.Label lblMonth 
            Caption         =   "Month"
            Height          =   240
            Left            =   2520
            TabIndex        =   25
            Top             =   225
            Width           =   555
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridCollectionStatus 
         Height          =   3885
         Left            =   180
         TabIndex        =   1
         Top             =   1170
         Width           =   7530
         _cx             =   13282
         _cy             =   6853
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalDaily.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label lblCollectionstatus 
         Caption         =   "Daily Collection Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   20
         Top             =   900
         Width           =   2400
      End
   End
   Begin VB.Frame fraChild 
      Caption         =   "Collection details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5370
      Left            =   90
      TabIndex        =   2
      Top             =   135
      Visible         =   0   'False
      Width           =   10320
      Begin VB.TextBox txtNonCashTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5895
         TabIndex        =   17
         Top             =   4995
         Width           =   1545
      End
      Begin VB.Frame fraInst 
         Height          =   870
         Left            =   45
         TabIndex        =   5
         Top             =   2160
         Width           =   10005
         Begin VB.TextBox txtBankName 
            Height          =   285
            Left            =   4455
            TabIndex        =   30
            Top             =   450
            Width           =   2850
         End
         Begin VB.CommandButton cmdDemandGenarate 
            Caption         =   "Genarate Demand"
            Height          =   330
            Left            =   8145
            TabIndex        =   29
            Top             =   495
            Width           =   1725
         End
         Begin VB.TextBox txtCashTotal 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6975
            TabIndex        =   13
            Top             =   135
            Width           =   1140
         End
         Begin VB.TextBox txtDemandDate 
            Height          =   285
            Left            =   4455
            TabIndex        =   11
            Top             =   135
            Width           =   1545
         End
         Begin VB.TextBox txtInstNo 
            Height          =   285
            Left            =   1125
            TabIndex        =   9
            Top             =   450
            Width           =   1230
         End
         Begin VB.TextBox txtInstType 
            Height          =   285
            Left            =   1125
            TabIndex        =   7
            Top             =   135
            Width           =   2565
         End
         Begin VB.Label lblBankName 
            Caption         =   "Name Of Bank:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3330
            TabIndex        =   31
            Top             =   450
            Width           =   1140
         End
         Begin VB.Label lblCashToatal 
            Caption         =   "CashTotal:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6075
            TabIndex        =   12
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblDate 
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            TabIndex        =   10
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lblInstNo 
            Caption         =   "InstrumentNo:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   8
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label lblInstType 
            Caption         =   "Instrument:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   6
            Top             =   180
            Width           =   825
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridNonCashDetails 
         Height          =   1635
         Left            =   45
         TabIndex        =   3
         Top             =   3285
         Width           =   9960
         _cx             =   17568
         _cy             =   2884
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalDaily.frx":00A3
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
      Begin VSFlex8LCtl.VSFlexGrid vsGridCashDetails 
         Height          =   1635
         Left            =   180
         TabIndex        =   14
         Top             =   450
         Width           =   10095
         _cx             =   17806
         _cy             =   2884
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Rows            =   50
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmZonalDaily.frx":01A4
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
      Begin VB.Label lblTot 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   900
         TabIndex        =   28
         Top             =   5040
         Width           =   1185
      End
      Begin VB.Label lblSelectedDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   8370
         TabIndex        =   27
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label lblAllTotal 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Top             =   5040
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "TotalAmount:"
         Height          =   240
         Left            =   4860
         TabIndex        =   16
         Top             =   4995
         Width           =   1005
      End
      Begin VB.Label lblNonCash 
         BackStyle       =   0  'Transparent
         Caption         =   "Non Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   3060
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   225
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmZonalDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub cmbMonth_Click()
      Call FillvsGridCollectionStatus
      'Call UpdateDemandStatus
    End Sub
    

Private Sub cmdBack_Click()
'    Dim mSQL As String
'    Dim Rec As New ADODB.Recordset
'    Dim objDB As New clsDB
    fraChild.Visible = False
    fraMain.Visible = True
    fraDate.Visible = True
    cmdBack.Visible = False
    cmdSend.Visible = False
'    Dim mCnn As New ADODB.Connection
'    mSQL = "select vchInstrumentNo from faIdemandTBL where dtDemandDate = '" & CheckDateInMMM(lblSelectedDate) & " '"
'    objDB.SetConnection mCnn
'     Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockBatchOptimistic, adCmdText
'
'      If Not (Rec.BOF And Rec.EOF) Then
'        txtInstNo.Text = Rec(0)
'        Else
          txtInstNo.Text = ""
          txtBankName.Text = ""
'      End If
End Sub

Private Sub cmdClose_Click()
'  fraChild.Visible = False
'  fraMain.Visible = False
'  fraDate.Visible = False
'  cmdClose.Visible = False
' cmdBack.Visible = False
   Unload frmZonalDaily
End Sub

Private Sub cmdDemandGenarate_Click()
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim count As Integer
    Dim mCount1 As Integer
    
    Dim numDemandID As Variant
    Dim intLBID As Integer
    Dim tnyExtAppID As Variant
    Dim tnyExtModuleID As Variant
    Dim tnyDemandType As Variant
    Dim intTransactionTypeID As Variant
    Dim intYearID As Variant
    Dim tnyPeriodID As Variant
    Dim dtDemandDate As Variant
    Dim numSubLedgerID As Variant
    Dim intKeyID As Variant
    Dim intKeyID2 As Variant
    Dim vchRemarks As Variant
    Dim tnyStatus As Variant
    Dim tnyArrearFlag As Variant
    Dim intVoucherID As Variant
    Dim dtVoucherDate As Variant
    
    Dim dtExpiryDate As Variant
    Dim intFinancialYearID As Variant
    Dim numSeatID As Variant
    Dim intSectionID As Variant
    Dim numUserID As Variant
    Dim numCounterID As Variant
    Dim vchAdminNote As Variant
    Dim vchDemandNo As Variant
    Dim numZoneID As Variant
    Dim intWardNo As Variant
    Dim intDoorNo As Variant
    Dim vchDoorNo2 As Variant
    Dim numForwardedSeatID As Variant
    Dim intInstrumentTypeID As Variant
    Dim vchInstrumentNo As Variant
    Dim dtInstrumentDate As Variant
    Dim vchDrawnFrom As Variant
    Dim vchDrawnPlace As Variant
    Dim dtDueDate As Variant
    Dim tnyAccuralType As Variant
    Dim numLocationID As Variant
    Dim tnySend As Variant
    Dim intFunctionID As Variant
    Dim intFunctionaryID As Variant
    Dim intSourceFundID As Variant
    Dim dtTransactionDate As Variant
    Dim intDemandMode As Variant
    Dim tnyAccrualType As Variant
    'Dim vchAccountHeadCode As Variant
    'Dim intAccountHeadID As Variant
    
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim intTransactionTID As Variant
    Dim mCount As Integer
    Dim DemandDate As Date
    Dim SourceFundID As Variant
    Dim FunctionID As Variant
    Dim countChek As Integer
    Dim intFunctionIDChek As Variant
    Dim intTransactionTypeIDChek As Variant
    Dim intSourceFundIDChek As Variant
    
     
    '    '------------Save To faIDemandChild---------'

    Dim tnySINo As Integer
    Dim intAccountHeadID As Integer
    Dim vchAccountHeadCode As String
    Dim fltAmount As Double
    Dim dtOnDate As Date
    Dim snyRate As Variant
    Dim objAc As New clsAccounts
    Dim arrInDemandChild As Variant

'--------To insert DemandAddress Table----------------------'

    Dim arrInDemandAdress As Variant
    Dim arrInDemandTbl As Variant
    Dim countAdress As Integer
    Dim RecDemandTbl As New ADODB.Recordset

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

    intWardNo = Null
    intDoorNo = Null
    vchDoorNo2 = Null
    vchName = Null
    vchInit1 = Null
    vchInit2 = Null
    vchInit3 = Null
    vchInit4 = Null
    vchHouseName = Null
    vchStreet = Null
    vchLocalPlace = Null
    vchMainPlace = Null
    vchPost = Null
    vchPin = Null
    vchPhone = Null
    tnyFlag = 0
    numZoneID = gbLocationID
    
'-----For Demand Generation And Save To DemandTables---------------'
        intLBID = gbLocalBodyID
        intDemandMode = 2
        tnyExtAppID = AppID.Saankhya
        tnyExtModuleID = 35
        tnyDemandType = 1
        dtDemandDate = Format(lblSelectedDate, "DD/MMM/YYYY")
        tnyPeriodID = 0
        intYearID = gbFinancialYearID
        numSubLedgerID = Null
        intKeyID = Null
        intKeyID2 = gbLocationID
        tnyStatus = 1
        intVoucherID = Null
        dtVoucherDate = Null
        tnyArrearFlag = Null
        dtExpiryDate = Null
       ' numDemandID = Null
        intFinancialYearID = gbFinancialYearID
        numSeatID = gbSeatID
        vchDemandNo = ""
        intSectionID = 99
        numUserID = gbUserID
        vchAdminNote = ""
        numCounterID = gbCounterID
        vchAdminNote = ""
        numZoneID = gbLocationID
        intWardNo = Null
        intDoorNo = Null
        vchDoorNo2 = ""
        numForwardedSeatID = 0
        tnyAccrualType = Null
        intInstrumentTypeID = val(txtInstType.Text)
        If txtInstNo.Text <> "" Then
            vchInstrumentNo = txtInstNo.Text
        Else
            MsgBox ("Enter Instrument No")
            Exit Sub
        End If
        
        dtInstrumentDate = Null
        vchDrawnFrom = ""
        vchDrawnPlace = ""
        dtDueDate = Null
        tnyAccuralType = 0
        numLocationID = gbLocationID
        tnySend = 0
        intFunctionaryID = 0
        dtTransactionDate = gbTransactionDate
      
        intYearID = gbFinancialYearID
        tnyStatus = 0
        tnyPeriodID = Null
        dtOnDate = lblSelectedDate
        snyRate = Null
        intVoucherID = Null
        dtVoucherDate = Null
        tnyArrearFlag = 0

'---ADDRESS--------------
        intWardNo = Null
        intDoorNo = Null
        vchDoorNo2 = Null
        vchName = Null
        vchInit1 = Null
        vchInit2 = Null
        vchInit3 = Null
        vchInit4 = Null
        vchHouseName = Null
        vchStreet = Null
        vchLocalPlace = Null
        vchMainPlace = Null
        vchPost = Null
        vchPin = Null
        vchPhone = Null
        tnyFlag = 0
        numZoneID = gbLocationID

     For count = 1 To vsGridCashDetails.Rows - 1
            If val(vsGridCashDetails.TextMatrix(count, 1)) > 0 Then
                intTransactionTypeID = val(vsGridCashDetails.TextMatrix(count, 8))
                intFunctionID = val(vsGridCashDetails.TextMatrix(count, 9))
                intSourceFundID = val(vsGridCashDetails.TextMatrix(count, 10))
                vchRemarks = vsGridCashDetails.TextMatrix(count, 4) + " - " + lblSelectedDate
                          
                tnySINo = count 'vsGridCashDetails.TextMatrix(count, 1)
                fltAmount = val(vsGridCashDetails.TextMatrix(count, 5))
                vchAccountHeadCode = vsGridCashDetails.TextMatrix(count, 14)
                intAccountHeadID = val(vsGridCashDetails.TextMatrix(count, 13))
        
                
'1----all not same----'
                If ((intTransactionTypeIDChek <> intTransactionTypeID) And (intFunctionIDChek <> intFunctionID) And (intSourceFundIDChek <> intSourceFundID)) Then
                  arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
'----Adress---
                    arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                    objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                End If

'2---------TransactionType Same---------
                
                If ((intTransactionTypeIDChek = intTransactionTypeID) And (intFunctionIDChek <> intFunctionID) And (intSourceFundIDChek <> intSourceFundID)) Then
                  arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
'----Adress---
                    arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                    objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                End If
'3------TransactionTyp & FunID Same----
                
                If ((intTransactionTypeIDChek = intTransactionTypeID) And (intFunctionIDChek = intFunctionID) And (intSourceFundIDChek <> intSourceFundID)) Then
                  arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
             
'----Adress---

                    arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                    objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                End If
                
'4---------Tran & SorceFund Same---
                 If ((intTransactionTypeIDChek = intTransactionTypeID) And (intFunctionIDChek <> intFunctionID) And (intSourceFundIDChek = intSourceFundID)) Then
                  arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
'----Adress---

                     arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                     objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                End If
                
'5---FunID NAd SourCe Same---
                 If ((intTransactionTypeIDChek <> intTransactionTypeID) And (intFunctionIDChek = intFunctionID) And (intSourceFundIDChek = intSourceFundID)) Then
                  arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc

'----Adress---

                    arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                    objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                    
                End If
'6--- SourCe Same---
                 If ((intTransactionTypeIDChek <> intTransactionTypeID) And (intFunctionIDChek <> intFunctionID) And (intSourceFundIDChek = intSourceFundID)) Then
                    arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
                                    tnyDemandType, intTransactionTypeID, intYearID, _
                                    tnyPeriodID, dtDemandDate, numSubLedgerID, intKeyID, _
                                    intKeyID2, vchRemarks, tnyStatus, intVoucherID, _
                                    dtVoucherDate, tnyArrearFlag, dtExpiryDate, numDemandID, _
                                    intFinancialYearID, numSeatID, intSectionID, _
                                    numUserID, numCounterID, vchAdminNote, vchDemandNo, _
                                    numZoneID, intWardNo, intDoorNo, vchDoorNo2, numForwardedSeatID, _
                                    dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, _
                                    vchDrawnFrom, vchDrawnPlace, tnyAccrualType, numLocationID, intFunctionaryID, _
                                    intFunctionID, intSourceFundID, dtTransactionDate, intDemandMode)
                    
                    objDB.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
'----Adress---

                    arrInDemandAdress = Array(arrOutPut(0, 0), _
                                intLBID, _
                                numZoneID, _
                                intWardNo, _
                                intDoorNo, _
                                vchDoorNo2, _
                                vchName, _
                                vchInit1, vchInit2, vchInit3, vchInit4, _
                                vchHouseName, _
                                vchStreet, _
                                vchLocalPlace, _
                                vchMainPlace, _
                                vchPost, _
                                vchPin, _
                                vchPhone)
                    objDB.ExecuteSP "spSaveIDemandAddress", arrInDemandAdress, , , mCnn, adCmdStoredProc
                    
                End If
'6---All Same---
              
                If ((intTransactionTypeIDChek <> intTransactionTypeID) And (intFunctionIDChek <> intFunctionID) And (intSourceFundIDChek = intSourceFundID)) Then
                  
                End If
                intTransactionTypeIDChek = intTransactionTypeID
                intFunctionIDChek = intFunctionID
                intSourceFundIDChek = intSourceFundID
                arrInDemandChild = Array(arrOutPut(0, 0), _
                          intLBID, _
                          tnySINo, _
                          val(intAccountHeadID), _
                          vchAccountHeadCode, _
                          fltAmount, _
                          vchRemarks, _
                          tnyStatus, _
                          dtOnDate, _
                          intYearID, _
                          tnyPeriodID, _
                          tnyArrearFlag, intTransactionTypeID)
                objDB.ExecuteSP "spSaveIDemandChild", arrInDemandChild, , , mCnn, adCmdStoredProc
  
             End If
    Next
Call UpdtaeDemandField
End Sub

Private Sub cmdSend_Click()
 If cmdDemandGenarate.Enabled = True Then
    MsgBox ("Genarate Demand ")
    Exit Sub
 End If
    '-------Send To DB_FinanceHO(Main Office) DEMANDTABLE------------'
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mCnnSvr As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecSvr As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim RecSvrChild As New ADODB.Recordset
        Dim RecSvrAdress As New ADODB.Recordset
        Dim RecAdress As New ADODB.Recordset
        Dim mSQL As String
        Dim mSqlChild As String
        Dim mSqlAdress As String
        Dim mDemandID As Variant
        Dim mCount As Integer

        mSQL = "Select * From faIDemandTbl  WHERE tnyExtModuleID=35 AND intDemandMode=2 AND dtDemandDate= '" & CheckDateInMMM(lblSelectedDate) & "'"

    objDB.SetConnection mCnn
    Rec.CursorLocation = adUseClient
    Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockBatchOptimistic, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
        objDB.CreateNewConnection mCnnSvr, SaankhyaHO
        mCnnSvr.BeginTrans
        ' Error GoTo ErrRollBack:
        RecSvr.CursorLocation = adUseServer
        RecSvr.Open "faIDemandTbl", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
       ' mDemandID = Rec!numDemandID
       While Not Rec.EOF
            RecSvr.ADDNEW
            mDemandID = Rec!numDemandID
            If mDemandID <> "" Then
                If Not (RecSvr.EOF And RecSvr.BOF) Then
                    RecSvr!numDemandID = Rec!numDemandID
                    RecSvr!intLBID = Rec!intLBID
                    RecSvr!tnyExtAppID = Rec!tnyExtAppID
                    RecSvr!tnyExtModuleID = Rec!tnyExtModuleID
                    RecSvr!tnyDemandType = Rec!tnyDemandType
                    RecSvr!intTransactionTypeID = Rec!intTransactionTypeID
                    RecSvr!intYearID = Rec!intYearID
                    RecSvr!tnyPeriodID = Rec!tnyPeriodID
                    RecSvr!dtDemandDate = Rec!dtDemandDate
                    RecSvr!numSubLedgerID = Rec!numSubLedgerID
                    RecSvr!intKeyID = Rec!intKeyID
                    RecSvr!intKeyID2 = Rec!intKeyID2
                    RecSvr!vchRemarks = Rec!vchRemarks
                    RecSvr!tnyStatus = 0
                    RecSvr!tnyArrearFlag = Rec!tnyArrearFlag
                    'RecSvr!intVoucherID = Rec!intVoucherID
                    'RecSvr!dtVoucherDate = Rec!dtVoucherDate
                    RecSvr!dtExpiryDate = Rec!dtExpiryDate
                    RecSvr!intFinancialYearID = Rec!intFinancialYearID
                    RecSvr!numSeatID = Rec!numSeatID
                    RecSvr!intSectionID = Rec!intSectionID
                    RecSvr!numUserID = Rec!numUserID
                    RecSvr!numCounterID = Rec!numCounterID
                    RecSvr!vchAdminNote = Rec!vchAdminNote
                    RecSvr!vchDemandNo = Rec!vchDemandNo
                    RecSvr!numZoneID = Rec!numZoneID
                    RecSvr!intWardNo = Rec!intWardNo
                    RecSvr!intDoorNo = Rec!intDoorNo
                    RecSvr!vchDoorNo2 = Rec!vchDoorNo2
                    RecSvr!numForwardedSeatID = Rec!numForwardedSeatID
                    RecSvr!intInstrumentTypeID = Rec!intInstrumentTypeID
                    RecSvr!vchInstrumentNo = Rec!vchInstrumentNo
                    RecSvr!dtInstrumentDate = Rec!dtInstrumentDate
                    RecSvr!vchDrawnFrom = Rec!vchDrawnFrom
                    RecSvr!vchDrawnPlace = Rec!vchDrawnPlace
                    RecSvr!dtDueDate = Rec!dtDueDate
                    RecSvr!tnyAccrualType = Rec!tnyAccrualType
                    RecSvr!numLocationID = Rec!numLocationID
                    RecSvr!dtTransactionDate = Rec!dtTransactionDate
                    RecSvr!intDemandMode = Rec!intDemandMode
                    RecSvr!intFunctionID = Rec!intFunctionID
                    RecSvr!intFunctionaryID = Rec!intFunctionaryID
                    RecSvr!intSourceFundID = Rec!intSourceFundID
                    'RecSvr.Update
                       
                    '--------DEMANDCHILD----------
                    
                    mSqlChild = "Select * From faIDemandChild Where numDemandID = " & mDemandID
                    RecChild.Open mSqlChild, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrChild.CursorLocation = adUseServer
                    RecSvrChild.Open "faIDemandChild", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
                    'If Not (RecSvrChild.BOF And RecSvrChild.EOF) Then
                        While Not RecChild.EOF
                            RecSvrChild.ADDNEW
                            If Not (RecSvrChild.BOF And RecSvrChild.EOF) Then
                            ' RecSvrChild.AddNew
                             RecSvrChild!numDemandID = RecChild!numDemandID
                             RecSvrChild!intLBID = RecChild!intLBID
                             RecSvrChild!tnySlNo = RecChild!tnySlNo
                             RecSvrChild!intAccountHeadID = RecChild!intAccountHeadID
                             RecSvrChild!vchAccountHeadCode = RecChild!vchAccountHeadCode
                             RecSvrChild!fltAmount = RecChild!fltAmount
                             RecSvrChild!intYearID = RecChild!intYearID
                             RecSvrChild!tnyPeriodID = RecChild!tnyPeriodID
                             RecSvrChild!tnyArrearFlag = RecChild!tnyArrearFlag
                             RecSvrChild!vchRemarks = RecChild!vchRemarks
                             RecSvrChild!tnyStatus = RecChild!tnyStatus
                             RecSvrChild!dtOnDate = RecChild!dtOnDate
                             RecSvrChild!snyRate = RecChild!snyRate
                             'RecSvr!intVoucherID = Rec!intVoucherID
                             'RecSvr!dtVoucherDate = Rec!dtVoucherDate
                             RecSvrChild!intTransactionTypeID = RecChild!intTransactionTypeID
                             RecSvr.Update
                             RecSvrChild.Update
                             RecChild.MoveNext
                            End If
                        Wend
                        
             '-----DEMANDADDRESS-------------------------------------------------
                        
    
                    mSqlAdress = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
                    RecAdress.Open mSqlAdress, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrAdress.CursorLocation = adUseServer
                    RecSvrAdress.Open "faIDemandAddress", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
                    While Not RecAdress.EOF
                        RecSvrAdress.ADDNEW
                        If Not (RecSvrAdress.BOF And RecSvrAdress.EOF) Then
                            RecSvrAdress!numDemandID = RecAdress!numDemandID
                            RecSvrAdress!numZoneID = RecAdress!numZoneID
                            RecSvrAdress!intWardNo = RecAdress!intWardNo
                            RecSvrAdress!intDoorNo = RecAdress!intDoorNo
                            RecSvrAdress!vchDoorNo2 = RecAdress!vchDoorNo2
                            RecSvrAdress!vchName = RecAdress!vchName
                            RecSvrAdress!vchInit1 = RecAdress!vchInit1
                            RecSvrAdress!vchInit2 = RecAdress!vchInit2
                            RecSvrAdress!vchInit3 = RecAdress!vchInit3
                            RecSvrAdress!vchInit4 = RecAdress!vchInit4
                            RecSvrAdress!vchHouseName = RecAdress!vchHouseName
                            RecSvrAdress!vchStreet = RecAdress!vchStreet
                            RecSvrAdress!vchLocalPlace = RecAdress!vchLocalPlace
                            RecSvrAdress!vchMainPlace = RecAdress!vchMainPlace
                            RecSvrAdress!vchPost = RecAdress!vchPost
                            RecSvrAdress!vchPin = RecAdress!vchPin
                            RecSvrAdress!vchPhone = RecAdress!vchPhone
                            
                            RecSvrAdress.Update
                            RecAdress.MoveNext
                        End If
                        mCnn.Execute "Update faIDemandTbl Set tnySend = 1 Where numDemandID = " & mDemandID
                    Wend
                    RecSvrAdress.Close
                    RecAdress.Close
                    RecSvrChild.Close
                    RecChild.Close
            End If
    End If
    Rec.MoveNext
    Wend
   ' Call SendToVoucherTbls
End If
    mCnnSvr.CommitTrans
    Call SendToVoucherTbls
    RecSvr.Close
    Rec.Close
    'mCnn.Execute "Update faIDemandTbl Set tnySend = 1 Where numDemandID = " & mDemandID
    vsGridCollectionStatus.Cell(flexcpChecked, vsGridCollectionStatus.Row, 4) = 1
    ' vsGridCollectionStatus.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
    mCnn.Close
    MsgBox "Successfully updated in Head Office !", vbInformation
    cmdSend.Enabled = False
    Exit Sub
    
    
ErrRollBack:
            MsgBox (Error$)
            mCnnSvr.RollbackTrans
            mCnnSvr.Close
    'Call send

End Sub

    Private Sub Form_Load()
        Call FormInitialize
    End Sub
    Private Sub FormInitialize()
        Dim mMonth As String
        fraMain.Width = 7900
        frmZonalDaily.Width = 8200
        mMonth = Month(gbTransactionDate)
        cmbMonth.Text = MonthName(mMonth)
        Call FillYear
        Call FillMonth
        Call FillvsGridCollectionStatus
'        If cmbMonth.Text = Null And cmbYear.Text = Null Then
'           vsGridCollectionStatus.Editable = flexEDNone
'        End If
    End Sub
    Private Sub FillMonth()
        Dim mCount As Integer
        For mCount = 1 To 12
            cmbMonth.AddItem (MonthName(mCount))
            cmbMonth.ItemData(cmbMonth.NewIndex) = mCount
        Next
    End Sub
    Private Sub FillYear()
         PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4))+'-'+Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear", , , , True
         cmbYear.ListIndex = cmbYear.ListCount - 1
    End Sub
    Private Sub FillvsGridCollectionStatus()
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim RecDemnadNo As New ADODB.Recordset
        Dim mCount As Integer
        Dim intMonth As Integer
        Dim intYear As Integer
        Dim arrInput As Variant
        Dim arrInput1 As Variant
        Dim mLoop As Integer
        Dim mRowCount As Integer
        Dim mRowCount1 As Integer
        
        mRowCount = 1
        If cmbMonth.ListIndex <> -1 Then
            intMonth = cmbMonth.ItemData(cmbMonth.ListIndex)
        End If
        intYear = cmbYear.ItemData(cmbYear.ListIndex)
        arrInput = Array(intMonth, intYear)
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objDB.ExecuteSP("spZonalDailyCollection", arrInput, , , mCnn, adCmdStoredProc)
            If Not Rec.EOF Then
               While Not Rec.EOF
                   vsGridCollectionStatus.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                   vsGridCollectionStatus.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    '-------Demand Sttus Updation in grid vsGridCollectionStatus--------------------'
                   If UpdateDemandStatus(vsGridCollectionStatus.TextMatrix(mRowCount, 1)) = True Then
                       vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 3) = vbChecked
                   Else
                       vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 3) = vbUnchecked
                   End If
                   If Rec!tnySend = 1 Then
                       vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 4) = vbChecked
                   Else
                       vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 4) = vbUnchecked
                   End If
    
                   mRowCount = mRowCount + 1
                   Rec.MoveNext
               Wend
           End If
 End Sub
Private Sub txtDemandDate_LostFocus()
    'txtDemandDate.Text = TimeValue(Now)
   ' txtDemandDate.Text = CheckDateInMMM(txtDemandDate.Text)
    
End Sub


Private Sub vsGridCollectionStatus_Click()
    Dim dtDate As Date
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec2 As New ADODB.Recordset
    Dim Rec As New ADODB.Recordset
    Dim RecDemnadNo As New ADODB.Recordset
    Dim arrInput As Variant
    Dim mRowCount As Integer
    Dim mRowCountDemand As Integer
    Dim mRowCountNCash As Integer
    Dim mCashTotal As Double
    Dim mNonCashTotal As Double
   '  frmZonalDaily.ControlBox = False
   ' If cmbMonth.Text <> Null And cmbYear.Text <> Null Then
   If vsGridCollectionStatus.TextMatrix(vsGridCollectionStatus.Row, 1) <> "" Then
        mRowCount = 0
        mRowCountNCash = 0
        mRowCountDemand = 0
        cmdBack.Visible = True
        fraChild.Visible = True
        fraMain.Visible = False
        fraDate.Visible = False
        cmdSend.Visible = True
        cmdClose.Visible = True
        fraMain.Width = 10275
        frmZonalDaily.Width = 10875
     ' If vsGridCollectionStatus.TextMatrix(vsGridCollectionStatus.Row, 3) = vbChecked Then
            dtDate = vsGridCollectionStatus.TextMatrix(vsGridCollectionStatus.Row, 1)
            lblSelectedDate.Caption = DdMmmYy(dtDate)
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            arrInput = Array(dtDate)
            
            '------------For Cash Details--------------'
            Set Rec = objDB.ExecuteSP("spGetZoneCash", arrInput, , , mCnn, adCmdStoredProc)
'            If vsGridCollectionStatus.Cell(flexcpChecked, vsGridCollectionStatus.Row, 3) = vbChecked Then
'            Else
            vsGridCashDetails.Clear 1, 1
'            End If
            If Not Rec.EOF Then
                While Not Rec.EOF
                    mRowCount = mRowCount + 1
                    vsGridCashDetails.TextMatrix(mRowCount, 1) = mRowCount
                    vsGridCashDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    vsGridCashDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                    vsGridCashDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    vsGridCashDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    mCashTotal = mCashTotal + Rec!fltAmount
                    vsGridCashDetails.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    vsGridCashDetails.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    vsGridCashDetails.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    vsGridCashDetails.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
                    'vsGridCashDetails.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    vsGridCashDetails.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                    vsGridCashDetails.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    vsGridCashDetails.TextMatrix(mRowCount, 14) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    Rec.MoveNext
                Wend
            End If
            
            
            Call UpdtaeDemandField
            If vsGridCashDetails.TextMatrix(vsGridCashDetails.Row, 6) <> "" Then
                cmdDemandGenarate.Enabled = False
            Else
                cmdDemandGenarate.Enabled = True
            End If
            If vsGridCollectionStatus.Cell(flexcpChecked, vsGridCollectionStatus.Row, 4) = vbChecked Then
                cmdSend.Enabled = False
            Else
                cmdSend.Enabled = True
            
            End If
            
'            Set RecDemnadNo = objdb.ExecuteSP("spSelectDemandNo", arrInput, , , mCnn, adCmdStoredProc)
'                If Not RecDemnadNo.EOF Then
'                    While Not RecDemnadNo.EOF
'                         mRowCountDemand = mRowCountDemand + 1
'                        If Not IsNull(RecDemnadNo!vchDemandNo) Then
'                            vsGridCashDetails.TextMatrix(mRowCountDemand, 6) = IIf(IsNull(RecDemnadNo!vchDemandNo), "", RecDemnadNo!vchDemandNo)
'                            cmdDemandGenarate.Enabled = False
'                        Else
'                            cmdDemandGenarate.Enabled = True
'                        End If
'                     RecDemnadNo.MoveNext
'                      Wend
'                Else
'                     cmdDemandGenarate.Enabled = True
'                End If
    
            txtCashTotal.Text = mCashTotal
            txtInstType.Text = "Directly Debited To Bank"
            txtInstType.Enabled = False
            txtDemandDate.Text = DdMmmYy(gbTransactionDate)
            txtDemandDate.Enabled = False
         
                    
            '-------------NonCashDetails-------------------'
            Set Rec2 = objDB.ExecuteSP("spGetZoneNonCash", arrInput, , , mCnn, adCmdStoredProc)
            vsGridNonCashDetails.Clear 1, 1
            If Not Rec2.EOF Then
                While Not Rec2.EOF
                    mRowCountNCash = mRowCountNCash + 1
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 1) = mRowCountNCash
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 2) = IIf(IsNull(Rec2!vchTransactionType), "", Rec2!vchTransactionType)
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 3) = IIf(IsNull(Rec2!vchInstrumentType), "", Rec2!vchInstrumentType)
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 4) = IIf(IsNull(Rec2!intVoucherNo), "", Rec2!intVoucherNo)
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 5) = IIf(IsNull(Rec2!fltAmount), "", Rec2!fltAmount)
                    mNonCashTotal = mNonCashTotal + Rec2!fltAmount
                    'vsGridNonCashDetails.TextMatrix(mRowCountNCash, 6) = IIf(IsNull(Rec2!vchDemandNo), "", Rec2!vchDemandNo)
                    vsGridNonCashDetails.TextMatrix(mRowCountNCash, 7) = IIf(IsNull(Rec2!intVoucherID), "", Rec2!intVoucherID)
                    Rec2.MoveNext
                Wend
            End If
            txtNonCashTotal.Text = mNonCashTotal
            lblTot.Caption = val(txtCashTotal.Text) + val(txtNonCashTotal.Text)
    Else
     vsGridCollectionStatus.Editable = flexEDNone
     fraChild.Visible = False
    fraMain.Visible = True
    End If
 End Sub
Private Function UpdateDemandStatus(dtDate As Date) As Boolean
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim RecDemnadNo As New ADODB.Recordset
    Dim mCount As Integer
    Dim arrInput As Variant
    Dim DemandDate As Date
    Dim mStatusFlag As Boolean
        arrInput = Array(dtDate)
        Set RecDemnadNo = objDB.ExecuteSP("spSelectDemandNo", arrInput, , , mCnn, adCmdStoredProc)
        If Not RecDemnadNo.EOF Then
            If Not IsNull(RecDemnadNo!vchDemandNo) Then
                UpdateDemandStatus = True
            Else
                UpdateDemandStatus = False
                MsgBox ("Demand Not Genarated For All")
                Exit Function
            End If
        Else
           UpdateDemandStatus = False
        End If
'    If mStatusFlag = True Then
'        vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 3) = vbChecked
'    Else
'        vsGridCollectionStatus.Cell(flexcpChecked, mRowCount, 3) = vbUnchecked
'    End If
End Function

Private Sub UpdtaeDemandField()
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim intTransactionTID As Variant
    Dim mCount As Integer
    Dim DemandDate As Date
    Dim SourceFundID As Variant
    Dim FunctionID As Variant
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset

     For mCount = 1 To vsGridCashDetails.Rows - 1
        intTransactionTID = val(vsGridCashDetails.TextMatrix(mCount, 8))
        DemandDate = Format(lblSelectedDate, "DD/MMM/YYYY")
        FunctionID = val(vsGridCashDetails.TextMatrix(mCount, 9))
        SourceFundID = val(vsGridCashDetails.TextMatrix(mCount, 10))
        arrIn = Array(DemandDate, intTransactionTID, FunctionID, SourceFundID)
        Set Rec = objDB.ExecuteSP("spGetZoneDemandStatus", arrIn, , , mCnn, adCmdStoredProc)
        If Not Rec.EOF Then
            vsGridCashDetails.TextMatrix(mCount, 13) = Rec!numDemandID
            vsGridCashDetails.TextMatrix(mCount, 6) = Rec!vchDemandNo
            txtInstNo.Text = Rec!vchInstrumentNo
            txtInstNo.Enabled = False
            Else
            txtInstNo.Enabled = True
        End If
         
    Next
    
    cmdDemandGenarate.Enabled = False
   
End Sub

Private Sub SendToVoucherTbls()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mCnnSvrV As New ADODB.Connection
    Dim Recv As New ADODB.Recordset
    Dim RecSvrV As New ADODB.Recordset
    Dim RecChild As New ADODB.Recordset
    Dim RecSvrChild As New ADODB.Recordset
    Dim RecAdress As New ADODB.Recordset
    Dim RecSvrAdress As New ADODB.Recordset
    Dim RecSvrTrans As New ADODB.Recordset
    Dim RecTrans As New ADODB.Recordset
    Dim RecTransChild As New ADODB.Recordset
    Dim RecSvrTransChild As New ADODB.Recordset
    Dim mSqlChild As String
    Dim mSqlTransChild As String
    Dim mSqlAdress As String
    Dim mSqlTrans As String
    Dim arrInput As Variant
    Dim mVocherID As Variant
    'Dim InDate As Date
   'InDate = vsGridCollectionStatus.TextMatrix(vsGridCollectionStatus.Row, 1)
    arrInput = Array(lblSelectedDate)
    Set Recv = objDB.ExecuteSP("spGetZoneNonCash", arrInput, , , mCnn, adCmdStoredProc)
    If Not Recv.EOF Then
        objDB.CreateNewConnection mCnnSvrV, SaankhyaHO
        mCnnSvrV.BeginTrans
        RecSvrV.CursorLocation = adUseServer
        RecSvrV.Open "faVouchers", mCnnSvrV, adOpenDynamic, adLockOptimistic, adCmdTable
        While Not Recv.EOF
            RecSvrV.ADDNEW
            mVocherID = Recv!intVoucherID
            If Recv!intVoucherID <> "" Then
                If Not (RecSvrV.EOF And RecSvrV.BOF) Then
                    RecSvrV!intVoucherID = Recv!intVoucherID
                    RecSvrV!intLocalBodyID = Recv!intLocalBodyID
                    RecSvrV!intTransactionID = Recv!intTransactionID
                    RecSvrV!intTransactionTypeID = Recv!intTransactionTypeID
                    RecSvrV!tnyVoucherTypeID = Recv!tnyVoucherTypeID
                    RecSvrV!intVoucherNo = Recv!intVoucherNo
                    RecSvrV!intBookNo = Recv!intBookNo
                    RecSvrV!dtDate = Recv!dtDate
                    RecSvrV!fltAmount = Recv!fltAmount
                    RecSvrV!intInstrumentTypeID = Recv!intInstrumentTypeID
                    RecSvrV!vchInstrumentNo = Recv!vchInstrumentNo
                    RecSvrV!dtInstrumentDate = Recv!dtInstrumentDate
                    RecSvrV!vchDescription = Recv!vchDescription
                    RecSvrV!numZoneID = Recv!numZoneID
                    RecSvrV!numWardId = Recv!numWardId
                    RecSvrV!intDoorNoP1 = Recv!intDoorNoP1
                    RecSvrV!vchDoorNoP2 = Recv!vchDoorNoP2
                    RecSvrV!vchDoorNoP3 = Recv!vchDoorNoP3
                    RecSvrV!intUserID = Recv!intUserID
                    RecSvrV!intCounterID = Recv!intCounterID
                    RecSvrV!numSubLedgerID = Recv!numSubLedgerID
                    RecSvrV!intKeyID1 = Recv!intKeyID1
                    RecSvrV!intKeyID2 = Recv!intKeyID2
                    RecSvrV!intExternalApplicationID = Recv!intExternalApplicationID
                    RecSvrV!intExternalModuleID = Recv!intExternalModuleID
                    RecSvrV!intFinancialYearID = Recv!intFinancialYearID
                    RecSvrV!tnyShiftID = Recv!tnyShiftID
                    RecSvrV!tnyPrintFlag = Recv!tnyPrintFlag
                    RecSvrV!tnyCancelFlag = Recv!tnyCancelFlag
                    RecSvrV!dtRealisationDate = Recv!dtRealisationDate
                    RecSvrV!vchRemarks = Recv!vchRemarks
                    RecSvrV!tnyStatus = Recv!tnyStatus
                    RecSvrV!vchBank = Recv!vchBank
                    RecSvrV!vchBankPlace = Recv!vchBankPlace
                    RecSvrV!intFundID = Recv!intFundID
                    RecSvrV!numSeatID = Recv!numSeatID
                    RecSvrV!intSessionID = Recv!intSessionID
                    RecSvrV!vchRefNo = Recv!vchRefNo
                    RecSvrV!fltRoundOff = Recv!fltRoundOff
                    RecSvrV!fltAdvAmtAdj = Recv!fltAdvAmtAdj
                    RecSvrV!numInwardNo = Recv!numInwardNo
                    RecSvrV!numLocationID = Recv!numLocationID
                    RecSvrV!tnyVoucherGroupID = Recv!tnyVoucherGroupID
                    RecSvrV!numLinkKeyID = Recv!numLinkKeyID
                    RecSvrV!dtTimeStamp = Recv!dtTimeStamp
                    RecSvrV!numTockenID = Recv!numTockenID
                    RecSvrV!tnyReconciled = Recv!tnyReconciled
                    RecSvrV!dtChequeRealiseDate = Recv!dtChequeRealiseDate
                    RecSvrV!vchVersionKey = Recv!vchVersionKey
                    'RecSvrV!tnyReversed = Recv!tnyReversed
                    'RecSvrV!dtValueDate = Recv!dtValueDate
                    
                    '----VOUCHER CHILD--------
                    
                    mSqlChild = "select * from faVoucherChild where intVoucherID=" & mVocherID
                    RecChild.Open mSqlChild, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrChild.CursorLocation = adUseServer
                    RecSvrChild.Open "faVoucherChild", mCnnSvrV, adOpenDynamic, adLockOptimistic, adCmdTable
                    While Not RecChild.EOF
                            RecSvrChild.ADDNEW
                            If Not (RecSvrChild.BOF And RecSvrChild.EOF) Then
                                RecSvrChild!intVoucherID = RecChild!intVoucherID
                                RecSvrChild!intLocalBodyID = RecChild!intLocalBodyID
                                RecSvrChild!intSlNo = RecChild!intSlNo
                                RecSvrChild!intAccountHeadID = RecChild!intAccountHeadID
                                RecSvrChild!tnyDebitOrCredit = RecChild!tnyDebitOrCredit
                                RecSvrChild!intYearID = RecChild!intYearID
                                RecSvrChild!tnyPeriodID = RecChild!tnyPeriodID
                                RecSvrChild!tnyArrearFlag = RecChild!tnyArrearFlag
                                RecSvrChild!numDemandID = RecChild!numDemandID
                                RecSvrChild!fltAmount = RecChild!fltAmount
                                RecSvrChild.Update
                                RecSvrV.Update
                                RecChild.MoveNext
                                
                            End If
                    Wend
                    
                    '---VocherAdress-----------
                    
                    mSqlAdress = "select * from faVoucherAddress where intVoucherID=" & mVocherID
                    RecAdress.Open mSqlAdress, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrAdress.CursorLocation = adUseServer
                    RecSvrAdress.Open "faVoucherAddress", mCnnSvrV, adOpenDynamic, adLockOptimistic, adCmdTable
                    While Not RecAdress.EOF
                            RecSvrAdress.ADDNEW
                            If Not (RecSvrAdress.BOF And RecSvrAdress.EOF) Then
                                RecSvrAdress!intVoucherID = RecAdress!intVoucherID
                                RecSvrAdress!intLocalBodyID = RecAdress!intLocalBodyID
                                RecSvrAdress!vchName = RecAdress!vchName
                                RecSvrAdress!vchHouseName = RecAdress!vchHouseName
                                RecSvrAdress!vchStreetName = RecAdress!vchStreetName
                                RecSvrAdress!vchMainPlace = RecAdress!vchMainPlace
                                RecSvrAdress!vchPostOffice = RecAdress!vchPostOffice
                                RecSvrAdress!vchDistrict = RecAdress!vchDistrict
                                RecSvrAdress!vchPinNumber = RecAdress!vchPinNumber
                                RecSvrAdress!vchInit1 = RecAdress!vchInit1
                                RecSvrAdress!vchInit2 = RecAdress!vchInit2
                                RecSvrAdress!vchInit3 = RecAdress!vchInit3
                                RecSvrAdress!vchInit4 = RecAdress!vchInit4
                                RecSvrAdress!vchPhone = RecAdress!vchPhone
                                RecSvrAdress!intWardNo = RecAdress!intWardNo
                                RecSvrAdress!intDoorNo = RecAdress!intDoorNo
                                RecSvrAdress!vchDoorNo2 = RecAdress!vchDoorNo2
                                RecSvrAdress.Update
                                RecAdress.MoveNext
                           End If
                    Wend
                    
                    
                    '-------faTransactions----------
                    
                    mSqlTrans = "select * from faTransactions where intVoucherID=" & mVocherID
                    RecTrans.Open mSqlTrans, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrTrans.CursorLocation = adUseServer
                    RecSvrTrans.Open "faTransactions", mCnnSvrV, adOpenDynamic, adLockOptimistic, adCmdTable
                    While Not RecTrans.EOF
                            RecSvrTrans.ADDNEW
                            If Not (RecSvrTrans.BOF And RecSvrTrans.EOF) Then
                                RecSvrTrans!intTransactionID = RecTrans!intTransactionID
                                RecSvrTrans!intLocalBodyID = RecTrans!intLocalBodyID
                                RecSvrTrans!intFinancialYearID = RecTrans!intFinancialYearID
                                RecSvrTrans!dtTransactionDate = RecTrans!dtTransactionDate
                                RecSvrTrans!intExternalApplicationID = RecTrans!intExternalApplicationID
                                RecSvrTrans!intFunctionID = RecTrans!intFunctionID
                                RecSvrTrans!intFunctionaryID = RecTrans!intFunctionaryID
                                RecSvrTrans!intFieldID = RecTrans!intFieldID
                                RecSvrTrans!intFundID = RecTrans!intFundID
                                RecSvrTrans!intBudgetCentreID = RecTrans!intBudgetCentreID
                                RecSvrTrans!vchNarration = RecTrans!vchNarration
                                RecSvrTrans!intTransactionTypeID = RecTrans!intTransactionTypeID
                                RecSvrTrans!intVoucherID = RecTrans!intVoucherID
                                RecSvrTrans!intProcessID = RecTrans!intProcessID
                                RecSvrTrans!intGroupID = RecTrans!intGroupID
                                RecSvrTrans!vchGroup = RecTrans!vchGroup
                                RecSvrTrans!numSubLedgerID = RecTrans!numSubLedgerID
                                RecSvrTrans!intKeyID = RecTrans!intKeyID
                                RecSvrTrans!numUserID = RecTrans!numUserID
                                RecSvrTrans!intVoucherNo = RecTrans!intVoucherNo
                                RecSvrTrans!tnyStatus = RecTrans!tnyStatus
                                RecSvrTrans!tnyVoucherGroupID = RecTrans!tnyVoucherGroupID
                               ' RecSvrTrans!tnyReversed = RecTrans!tnyReversed
                               ' RecSvrTrans!dtValueDate = RecTrans!dtValueDate
                                RecSvrTrans.Update
                                RecTrans.MoveNext
                            End If
                    Wend
                    
                    '-----------faTranscationChild-------------
                    mSqlTransChild = "select * from faTransactionChild where intTransactionID=" & mVocherID
                    RecTransChild.Open mSqlTransChild, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    RecSvrTransChild.CursorLocation = adUseServer
                    RecSvrTransChild.Open "faTransactionChild", mCnnSvrV, adOpenDynamic, adLockOptimistic, adCmdTable
                    While Not RecTransChild.EOF
                        RecSvrTransChild.ADDNEW
                        If Not (RecSvrTransChild.BOF And RecSvrTransChild.EOF) Then
                                RecSvrTransChild!intTransactionID = RecTransChild!intTransactionID
                                RecSvrTransChild!intSerialNo = RecTransChild!intSerialNo
                                RecSvrTransChild!intAccountHeadID = RecTransChild!intAccountHeadID
                                RecSvrTransChild!fltAmount = RecTransChild!fltAmount
                                RecSvrTransChild!tinDebitOrCreditFlag = RecTransChild!tinDebitOrCreditFlag
                                RecSvrTransChild!intByAccountHeadID = RecTransChild!intByAccountHeadID
                                RecSvrTransChild!vchNarration = RecTransChild!vchNarration
                                RecSvrTransChild!intFundID = RecTransChild!intFundID
                                RecSvrTransChild!fltOpeningBalance = RecTransChild!fltOpeningBalance
                                RecSvrTransChild!numTockenID = RecTransChild!numTockenID
                                RecSvrTransChild!dtReconcileDate = RecTransChild!dtReconcileDate
                                RecSvrTransChild.Update
                                RecTransChild.MoveNext
                        End If
                    Wend
                    RecChild.Close
                    RecSvrChild.Close
                    RecAdress.Close
                    RecSvrAdress.Close
                    RecSvrTrans.Close
                    RecTrans.Close
                    RecSvrTransChild.Close
                    RecTransChild.Close
                    
                End If
            End If
            Recv.MoveNext
        Wend
         mCnnSvrV.CommitTrans
    RecSvrV.Close
    Recv.Close
    
    End If
   
End Sub




    


VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListOfZonalDailyCollection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Collection to Head Office"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdView 
      Caption         =   "View Demand"
      Height          =   390
      Left            =   3950
      TabIndex        =   12
      Top             =   5970
      Width           =   1200
   End
   Begin VB.ComboBox cmbYear 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   675
      Width           =   1980
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   1275
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   390
      Left            =   5220
      TabIndex        =   8
      Top             =   5970
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   45
      TabIndex        =   7
      Top             =   1050
      Width           =   11370
   End
   Begin VB.ComboBox cmbMonth 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9315
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   675
      Width           =   1980
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4155
      Left            =   150
      TabIndex        =   2
      Top             =   1635
      Width           =   11130
      _cx             =   19632
      _cy             =   7329
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfZonalDailyCollection.frx":0000
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11415
      TabIndex        =   1
      Top             =   5865
      Width           =   11475
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   0
      Width           =   11475
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5580
      TabIndex        =   11
      Top             =   735
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8595
      TabIndex        =   5
      Top             =   735
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zonal"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1230
      TabIndex        =   4
      Top             =   705
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   735
      Width           =   945
   End
End
Attribute VB_Name = "frmListOfZonalDailyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '***************************************************************************************************'
    'Form to list collection of a Zonal for a particular period and send that details to the Main Office'
    '***************************************************************************************************'
    Private Sub FillMonth()
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
    End Sub
     
    Private Sub FillvsGrid()
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDB
        Dim mSQL            As String
        Dim mSQLDemand      As String
        Dim Rec             As New ADODB.Recordset
        Dim RecDemand       As New ADODB.Recordset
        Dim mRowCount       As Variant
        Dim mCount          As Double
        
        On Error GoTo err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mRowCount = ""
        mCount = 0
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        pbProgress.value = 0
        Me.MousePointer = vbHourglass
        'Added By sunil on 24-08-2011
        '===========================================
        mSQL = "Select dtDate,Sum(fltAmount) As Amount From faVouchers "
        mSQL = mSQL + "Inner Join faCounters ON faCounters.intCounterID = faVouchers.intCounterID"
        mSQL = mSQL + " Where Month(dtDate) =" & cmbMonth.ItemData(cmbMonth.ListIndex)
        mSQL = mSQL + " And intFinancialYearID = " & cmbYear.ItemData(cmbYear.ListIndex)
        mSQL = mSQL + "AND intSectionID = 99"
        mSQL = mSQL + " And tnyCancelFlag <> 1"
        mSQL = mSQL + " And intInstrumentTypeID = 1"
        mSQL = mSQL + " Group By dtDate"
        mSQL = mSQL + " Order By dtDate"
        '=============================================
        Rec.Open mSQL, mCnn
        While Not Rec.EOF
            mCount = mCount + 1
            Rec.MoveNext
        Wend
        If mCount <> 0 Then
            pbProgress.Max = mCount
        End If
        If Not (Rec.EOF And Rec.BOF) Then
            Rec.MoveFirst
        End If
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If vsGrid.TextMatrix(mRowCount, 0) <> "" Then
                mSQLDemand = "Select * From faIDemandTbl"
                mSQLDemand = mSQLDemand + " Where intTransactionTypeID = " & gbTransactionTypeZonalCollection
                mSQLDemand = mSQLDemand + " And dtDemandDate = '" & CheckDateInMMM(vsGrid.TextMatrix(mRowCount, 0)) & "'"
                mSQLDemand = mSQLDemand + " And tnyStatus <>9"
                RecDemand.Open mSQLDemand, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                If Not (RecDemand.BOF And RecDemand.EOF) Then
                    vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecDemand!vchDemandNo), "", RecDemand!vchDemandNo)
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecDemand!vchInstrumentNo), "", RecDemand!vchInstrumentNo)
                    If Not (IsNull(RecDemand!tnySend)) Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 4) = IIf(RecDemand!tnySend = 1, 1, 0)
                        vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 4) = &HC0FFC0
                    End If
                End If
                RecDemand.Close
            End If
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Format(Rec!Amount, "0.00")), "", Format(Rec!Amount, "0.00"))
            Rec.MoveNext
            If pbProgress.value < pbProgress.Max + 1 Then
                pbProgress.value = pbProgress.value + 1
            End If
            mRowCount = mRowCount + 1
        Wend
        Me.MousePointer = vbArrow
        Exit Sub
err:
        MsgBox err.Description
    End Sub
        
    Private Sub cmbMonth_Click()
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim objDB           As New clsDB
        Dim mSQL            As String
        
        '*********************************************************************************************'
        '               Procedure to show the List of Transactions for a particular month             '
        '*********************************************************************************************'
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If cmbMonth.ListIndex <> -1 Then
            If cmbMonth.ItemData(cmbMonth.ListIndex) <> 0 Then
                If cmbYear.ListIndex > -1 Then
                    Call FillvsGrid
                End If
'                mSQL = "Select dtDate,Sum(fltAmount) As Amount From faVouchers "
'                mSQL = mSQL + " Where Month(dtDate) =" & cmbMonth.ItemData(cmbMonth.ListIndex)
'                mSQL = mSQL + " And intFinancialYearID = 2009"
'                mSQL = mSQL + " And tnyCancelFlag <> 1"
'                mSQL = mSQL + " And intInstrumentTypeID = 1"
'                mSQL = mSQL + " Group By dtDate"
'                mSQL = mSQL + " Order By dtDate"
'                Rec.Open mSQL, mCnn
'                If Not (Rec.EOF And Rec.BOF) Then
'                    Call FillvsGrid(Rec)
'                End If
'                Rec.Close
            End If
        End If
    End Sub

    Private Sub cmbYear_Click()
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim objDB           As New clsDB
        Dim mSQL            As String
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If cmbMonth.ListIndex <> -1 Then
            If cmbMonth.ItemData(cmbMonth.ListIndex) <> 0 Then
                If cmbYear.ListIndex > -1 Then
                    Call FillvsGrid
                End If
            End If
        End If
    End Sub

    Private Sub cmdSend_Click()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mCnnSvr As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecSvr As New ADODB.Recordset
        Dim mSQL As String
        Dim mDemandID As Variant
        
        '*********************************************************************************************'
        '               Procedure to send the Transaction Details to Main Office                      '
        '*********************************************************************************************'
        'Note:-Inserting into Demand Table
        If vsGrid.Row <> 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
                mSQL = "Select * From faIDemandTbl Where intTransactionTypeID = " & gbTransactionTypeZonalCollection & " AND dtDemandDate = '" & CheckDateInMMM(vsGrid.TextMatrix(vsGrid.Row, 0)) & "'"
                objDB.SetConnection mCnn
                Rec.CursorLocation = adUseClient
                Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockBatchOptimistic, adCmdText
                If Not (Rec.BOF And Rec.EOF) Then
                    objDB.CreateNewConnection mCnnSvr, SaankhyaHO
                    mCnnSvr.BeginTrans
                    On Error GoTo ErrRollBack:
                    RecSvr.CursorLocation = adUseServer
                    
                    mDemandID = Rec!numDemandID
                    If mDemandID <> "" Then
                        mSQL = "Select * From faIDemandTbl Where numDemandID = " & mDemandID
'                        mSQL = mSQL + " And tnyStatus = 0"
                        RecSvr.Open mSQL, mCnnSvr
                        If Not (RecSvr.EOF And RecSvr.BOF) Then
                            If Not (IsNull(RecSvr!tnyStatus)) Then
                                If RecSvr!tnyStatus = 0 Then
                                    RecSvr.Close
                                    mCnnSvr.Execute "Delete From faIDemandTbl Where numDemandID = " & mDemandID
                                    mCnnSvr.Execute "Delete From faIDemandChild Where numDemandID = " & mDemandID
                                    mCnnSvr.Execute "Delete From faIDemandAddress Where numDemandID = " & mDemandID
                                ElseIf RecSvr!tnyStatus = 1 Then
                                    MsgBox "Sorry! Receipt Issued, Can't send to Head Office", vbInformation
                                    Exit Sub
                                End If
                            End If
                        End If
                        If RecSvr.State = 1 Then RecSvr.Close
                    End If
                    RecSvr.Open "faIDemandTbl", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
                    RecSvr.AddNew
                    mDemandID = Rec!numDemandID
                     
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
                    ' Added by Sunil on 08-08-2011
                    RecSvr!dtTransactionDate = Rec!dtTransactionDate
                    RecSvr!intDemandMode = Rec!intDemandMode
                    RecSvr!intFunctionID = Rec!intFunctionID
                    RecSvr!intFunctionaryID = Rec!intFunctionaryID
                    RecSvr!intSourceFundID = Rec!intSourceFundID
                    RecSvr.Update
                    
                    Rec.Close
                    RecSvr.Close
                    
                    'Note:-Inserting into DemandChild Table
                    mSQL = "Select * From faIDemandChild Where numDemandID = " & mDemandID
                    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    If Not (Rec.BOF And Rec.EOF) Then
                        RecSvr.CursorLocation = adUseServer
                        RecSvr.Open "faIDemandChild", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
                        While Not Rec.EOF
                            RecSvr.AddNew
                            RecSvr!numDemandID = Rec!numDemandID
                            RecSvr!intLBID = Rec!intLBID
                            RecSvr!tnySlNo = Rec!tnySlNo
                            RecSvr!intAccountHeadID = Rec!intAccountHeadID
                            RecSvr!vchAccountHeadCode = Rec!vchAccountHeadCode
                            RecSvr!fltAmount = Rec!fltAmount
                            RecSvr!intYearID = Rec!intYearID
                            RecSvr!tnyPeriodID = Rec!tnyPeriodID
                            RecSvr!tnyArrearFlag = Rec!tnyArrearFlag
                            RecSvr!vchRemarks = Rec!vchRemarks
                            RecSvr!tnyStatus = Rec!tnyStatus
                            RecSvr!dtOnDate = Rec!dtOnDate
                            RecSvr!snyRate = Rec!snyRate
                            'RecSvr!intVoucherID = Rec!intVoucherID
                            'RecSvr!dtVoucherDate = Rec!dtVoucherDate
                            RecSvr!intTransactionTypeID = Rec!intTransactionTypeID 'Added by sunil on 05-08-2011
                            RecSvr.Update
                            Rec.MoveNext
                        Wend
                    End If
                    Rec.Close
                    RecSvr.Close
                    
                    'Note:- Inserting into DemandAddress Table
                    mSQL = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
                    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
                    If Not (Rec.BOF And Rec.EOF) Then
                        RecSvr.CursorLocation = adUseServer
                        RecSvr.Open "faIDemandAddress", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
                        RecSvr.AddNew
                        
                        RecSvr!numDemandID = Rec!numDemandID
                        RecSvr!numZoneID = Rec!numZoneID
                        RecSvr!intWardNo = Rec!intWardNo
                        RecSvr!intDoorNo = Rec!intDoorNo
                        RecSvr!vchDoorNo2 = Rec!vchDoorNo2
                        RecSvr!vchName = Rec!vchName
                        RecSvr!vchInit1 = Rec!vchInit1
                        RecSvr!vchInit2 = Rec!vchInit2
                        RecSvr!vchInit3 = Rec!vchInit3
                        RecSvr!vchInit4 = Rec!vchInit4
                        RecSvr!vchHouseName = Rec!vchHouseName
                        RecSvr!vchStreet = Rec!vchStreet
                        RecSvr!vchLocalPlace = Rec!vchLocalPlace
                        RecSvr!vchMainPlace = Rec!vchMainPlace
                        RecSvr!vchPost = Rec!vchPost
                        RecSvr!vchPin = Rec!vchPin
                        RecSvr!vchPhone = Rec!vchPhone
            
                        RecSvr.Update
                        RecSvr.Close
                        Rec.Close
                        MsgBox "Successfully updated in Head Office !", vbInformation
                    End If
                    mCnnSvr.CommitTrans
                    mCnn.Execute "Update faIDemandTbl Set tnySend = 1 Where numDemandID = " & mDemandID
                    vsGrid.Cell(flexcpChecked, vsGrid.Row, 4) = 1
                    vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
                    mCnn.Close
                Else
                    MsgBox "Please generate the Demand !", vbInformation
                End If
            End If
        Else
            MsgBox "Please select any Transaction !", vbInformation
        End If
        Exit Sub
        
ErrRollBack:
        MsgBox (Error$)
        mCnnSvr.RollbackTrans
        mCnnSvr.Close
    End Sub

Private Sub Form_Activate()
    If cmbMonth.ListIndex <> -1 Then
        If cmbMonth.ItemData(cmbMonth.ListIndex) <> 0 Then
            If cmbYear.ListIndex > -1 Then
                Call FillvsGrid
            End If
        End If
    End If
End Sub

    Private Sub Form_Load()
        Call FillMonth
        PopulateList cmbYear, "Select Cast(intFinancialYearID as varchar(4))+'-'+Right(Cast(intFinancialYearID+1 as varchar(4)),2),intFinancialYearID  From faFinancialYear", , , , True
        cmbYear.ListIndex = cmbYear.ListCount - 1
        vsGrid.SelectionMode = flexSelectionByRow
    End Sub
    Private Sub VSGrid_DblClick()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim mSQL    As String
        Dim Rec     As New ADODB.Recordset
        
        If vsGrid.Row <> 0 Then
            Unload frmDemandInterface
            frmDemandInterface.Mode = 1
            mSQL = "Select * From faIDemandTbl"
            mSQL = mSQL + " Where intTransactionTypeID = " & gbTransactionTypeZonalCollection
            mSQL = mSQL + " And dtDemandDate = '" & CheckDateInMMM(vsGrid.TextMatrix(vsGrid.Row, 0)) & "'"
            mSQL = mSQL + " And tnyStatus <>9"
            objDB.SetConnection mCnn
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            If Not (Rec.BOF And Rec.EOF) Then
                vsGrid.TextMatrix(vsGrid.Row, 2) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                vsGrid.TextMatrix(vsGrid.Row, 3) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                MsgBox "Demand Already Generated on this Date!", vbInformation
                frmDemandInterface.Mode = 0
                Exit Sub
            End If
            Rec.Close
            frmDemandInterface.txtTransactionDate = gbTransactionDate
            frmDemandInterface.txtTransactionDate.Enabled = False
            frmDemandInterface.Show vbModal
           ' Load frmDemandInterface
           ' frmDemandInterface.ZOrder (0)
          '  MsgBox "afsdf"
        End If
    End Sub
Private Sub cmdView_Click() 'Added By sunil on 22-08-2011
        If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
           frmDemandInterface.Mode = 1
           frmDemandInterface.txtTransactionDate = gbTransactionDate
           frmDemandInterface.txtTransactionDate.Enabled = False
           frmDemandInterface.cmdSave.Enabled = False
           frmDemandInterface.cmdNew.Enabled = False
           frmDemandInterface.Show vbModal
        Else
           MsgBox "Generate a Demand first", vbInformation
        End If
End Sub



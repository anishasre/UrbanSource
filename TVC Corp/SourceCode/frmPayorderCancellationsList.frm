VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPayorderCancellationsList 
   BackColor       =   &H00EDF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Payorder Reversals"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   Icon            =   "frmPayorderCancellationsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EDF7F7&
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6465
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14415
      TabIndex        =   12
      Top             =   0
      Width           =   14415
   End
   Begin VB.TextBox txtPaymentVoucherNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9495
      TabIndex        =   11
      Top             =   585
      Width           =   1710
   End
   Begin VB.TextBox txtPayOrderNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5220
      TabIndex        =   9
      Top             =   585
      Width           =   1755
   End
   Begin VB.TextBox txtFromDate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   585
      TabIndex        =   5
      Top             =   555
      Width           =   1395
   End
   Begin VB.TextBox txtToDate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   555
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00EDF7F7&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6465
      Width           =   1140
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00EDF7F7&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6465
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EDF7F7&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6465
      Width           =   1140
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5250
      Left            =   45
      TabIndex        =   0
      Top             =   1080
      Width           =   14295
      _cx             =   25215
      _cy             =   9260
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorBkg    =   16777215
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPayorderCancellationsList.frx":000C
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
   Begin VB.Label lblTotalRecords 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1575
      TabIndex        =   14
      Top             =   6570
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   6570
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Voucher No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7470
      TabIndex        =   10
      Top             =   630
      Width           =   1965
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PayOrder No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3960
      TabIndex        =   8
      Top             =   645
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   7
      Top             =   585
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      Height          =   195
      Left            =   2025
      TabIndex        =   6
      Top             =   585
      Width           =   90
   End
End
Attribute VB_Name = "frmPayorderCancellationsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mUserID            As Variant
    Dim mPreviousYearMode As Integer
    Dim mPendingTransactionDate As Date
    Dim mPendingTaskReqID As Integer
    
    Private Sub cmdClear_Click()
        Call formInitialise
        Call FillGrid
    End Sub


    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        frmPayOrderCancellations.cmdPayorderSearch.Enabled = True
        frmPayOrderCancellations.Show vbModal
        Call FillGrid
    End Sub
    
    Private Sub cmdSearch_Click()
'        Call formInitialise
'        Call SearchData
        Call FillGrid
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        mUserID = gbUserID
        Call formInitialise
        Call FillGrid
    End Sub


    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
    End Sub

    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            
            If CDate(vsGrid.TextMatrix(vsGrid.Row, 2)) < gbStartingDate Then
                Call CheckPreviousYearTaskRequest(Trim(vsGrid.TextMatrix(vsGrid.Row, 3)))
                If (mPendingTransactionDate >= DateAdd("yyyy", -1, gbStartingDate) And mPendingTransactionDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                    frmPayOrderCancellations.PreviousYearMode = 1
                    frmPayOrderCancellations.PendingTaskReqID = mPendingTaskReqID
                    frmPayOrderCancellations.PreviousYearTransactionDate = mPendingTransactionDate
                Else
                    MsgBox "Only Previous Year's Pending Tasks can process here", vbInformation
                    Exit Sub
                End If
            End If
            
            frmPayOrderCancellations.cmdPayorderSearch.Enabled = False
            frmPayOrderCancellations.txtPayOrderNo.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            Call frmPayOrderCancellations.txtPayOrderNo_LostFocus
            frmPayOrderCancellations.Show vbModal
            Call FillGrid
        End If
    End Sub
    
    
    Private Sub CheckPreviousYearTaskRequest(mPONo As String)
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
        Dim Rec         As New ADODB.Recordset
        
        Dim mTaskID     As Integer
        
        'On Error GoTo Err
        'If mPendingTaskReqID > 0 Then
            
            If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSQL = "SELECT * FROM faPendingTaskRequest WHERE intTaskID = 8 AND vchInstrumentNo = '" & mPONo & "'"
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mPendingTransactionDate = Rec!dtTransactionDate
                    mPendingTaskReqID = Rec!intRequestID
                End If
                Rec.Close
            End If
        'End If
    
    End Sub
    
    Private Sub formInitialise()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        
        txtFromDate.Text = DdMmmYy(DateAdd("d", -30, gbTransactionDate))
        If CDate(txtFromDate.Text) < gbStartingDate Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        End If
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        vsGrid.Clear 1, 0
    End Sub
    Private Function PreviousYearMode(mPONo)
        Dim mSQL As String
        Dim objDB            As New clsDB
        Dim mCnn             As New ADODB.Connection
        Dim Rec              As New ADODB.Recordset
        mSQL = "Select * from faPendingTaskRequest Where intTaskId=8 And vchInstrumentNo=" & mPONo
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mPreviousYearMode = 1
        Else
            mPreviousYearMode = 0
        End If
    End Function
    Private Sub FillGrid()
    
        Dim mSQL             As String
        Dim objDB            As New clsDB
        Dim mCnn             As New ADODB.Connection
        Dim Rec              As New ADODB.Recordset
        Dim mRowCount        As Integer
        Dim mRecCnt          As Integer
       
        mSQL = "SELECT faReverseEntry.intRequestID,faReverseEntry.intCategoryID,faReverseEntry.tnyVoucherTypeID,faReverseEntry.numDemandNo, "
        mSQL = mSQL + " faReverseEntry.numDemandID , faReverseEntry.tnyStatus,faReverseEntry.dtRequestDate, faReverseEntry.intPaymentVoucherID, faReverseEntry.tnyPaid,faSeats.chvSeatTitle,A.vchUserName as Accountant,B.vchUserName as AO,C.vchUserName as Secretary,"
        mSQL = mSQL + " faReverseEntry.numRequestedUserID,faReverseEntry.numRequestedSeatID, faReverseEntry.numAuthorisedByAO,faReverseEntry.numAuthorisedBySec FROM faReverseEntry"
        mSQL = mSQL + " LEFT OUTER JOIN faSeats ON faReverseEntry.numRequestedSeatID = faSeats.numSeatID"
        mSQL = mSQL + " LEFT OUTER JOIN faUser A ON faReverseEntry.numRequestedUserID = A.numUserID"
        mSQL = mSQL + " LEFT OUTER JOIN faUser B ON faReverseEntry.numAuthorisedByAO = B.numUserID"
        mSQL = mSQL + " LEFT OUTER JOIN faUser C ON faReverseEntry.numAuthorisedBySec = C.numUserID"
        mSQL = mSQL + " Where faReverseEntry.intCategoryID = 70 And faReverseEntry.tnyStatus <> 3 "
        'mSQL = mSQL + " AND dtRequestDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "' "
        If (txtFromDate.Text) <> "" And (txtToDate.Text) <> "" Then
            mSQL = mSQL + "And dtRequestDate Between '" & Trim(txtFromDate.Text) & "' And '" & Trim(txtToDate.Text) & "'"
        End If
        If Trim(txtPayOrderNo.Text) <> "" Then
            mSQL = mSQL + "And faReverseEntry.numDemandNo LIKE '%" & Trim(txtPayOrderNo.Text) & "%'"
        End If
        If Trim(txtPaymentVoucherNo.Text) <> "" Then
            mSQL = mSQL + " And faReverseEntry.intPaymentVoucherID LIKE '%" & Trim(txtPaymentVoucherNo.Text) & "'"
        End If
        
        mSQL = mSQL + "ORDER BY faReverseEntry.dtRequestDate Desc,faReverseEntry.tnyStatus, faReverseEntry.intRequestID Desc"
        objDB.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        
        mRowCount = 1
        mRecCnt = 1
'        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
            If Not (Rec.BOF And Rec.EOF) Then
                While Not (Rec.EOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.TextMatrix(mRowCount, 0) = mRecCnt
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                    vsGrid.TextMatrix(mRowCount, 2) = DdMmmYy(IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate))
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!numDemandNo), "", Rec!numDemandNo)
                    vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!intPaymentVoucherID), "", Rec!intPaymentVoucherID)
                    If val(vsGrid.TextMatrix(mRowCount, 4)) < 1 Then
                        vsGrid.TextMatrix(mRowCount, 4) = ""
                    End If
                    If Rec!tnyPaid = 1 Then
                        vsGrid.TextMatrix(mRowCount, 5) = "Yes"
                    Else
                        vsGrid.TextMatrix(mRowCount, 5) = "No"
                    End If
                    If Rec!tnyStatus = 0 Then
                        vsGrid.TextMatrix(mRowCount, 6) = "Requested"
                    ElseIf Rec!tnyStatus = 1 Then
                        vsGrid.TextMatrix(mRowCount, 6) = "First Level"
                    ElseIf Rec!tnyStatus = 2 Then
                        vsGrid.TextMatrix(mRowCount, 6) = "Final Level"
                    ElseIf Rec!tnyStatus = 3 Then
                        vsGrid.TextMatrix(mRowCount, 6) = ""
                    ElseIf Rec!tnyStatus = 4 Then
                        vsGrid.TextMatrix(mRowCount, 6) = "Request Cancelled"
                    End If
                    vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!Accountant), "", (Rec!Accountant))
                    vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!AO), "", Rec!AO)
    '                If Rec!numAuthorisedBySec = 0 Then
                        vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!Secretary), "", Rec!Secretary)
    '                End If
                    
    '                vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                    vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!tnyStatus), "", (Rec!tnyStatus))
                    vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!tnyPaid), "", Rec!tnyPaid)
    '                vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!numRequestedUserID), "", Rec!numRequestedUserID)
                    vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!numRequestedSeatID), "", Rec!numRequestedSeatID)
                    vsGrid.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!numAuthorisedByAO), "", Rec!numAuthorisedByAO)
                    vsGrid.TextMatrix(mRowCount, 14) = IIf(IsNull(Rec!numAuthorisedBySec), "", Rec!numAuthorisedBySec)
                    If IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus) = 2 Then
'                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mRowCount) = &HD2AE9E
                        vsGrid.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 9) = &HD2AE9E
                    End If
                    Rec.MoveNext
        
                    mRowCount = mRowCount + 1
                    mRecCnt = mRecCnt + 1
            Wend
            If gbLBPanchayat Then
                vsGrid.ColHidden(8) = True
            End If
        End If
        Rec.Close
        lblTotalRecords.Caption = CStr(mRecCnt - 1)
    End Sub

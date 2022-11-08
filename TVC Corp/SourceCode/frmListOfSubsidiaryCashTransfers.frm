VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmListOfSubsidiaryCashTransfers 
   Caption         =   "Subsidiary Cash Book Details"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmListOfSubsidiaryCashTransfers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   15180
      TabIndex        =   2
      Top             =   8475
      Width           =   15240
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   450
         Left            =   2925
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   450
         Left            =   1320
         TabIndex        =   4
         Top             =   450
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   450
         Left            =   -285
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remit Backed"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9345
         TabIndex        =   15
         Top             =   90
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   210
         Left            =   11055
         TabIndex        =   14
         Top             =   105
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Making Disbursement"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6765
         TabIndex        =   13
         Top             =   90
         Width           =   1740
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   210
         Left            =   6495
         TabIndex        =   12
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblVerified 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   210
         Left            =   2355
         TabIndex        =   11
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblForward 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   210
         Left            =   4170
         TabIndex        =   10
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblTransferred 
         AutoSize        =   -1  'True
         Caption         =   "Transferred"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2625
         TabIndex        =   9
         Top             =   90
         Width           =   975
      End
      Begin VB.Label lblApprovedByClerk 
         AutoSize        =   -1  'True
         Caption         =   "Approved By Clerk"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4425
         TabIndex        =   8
         Top             =   90
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Process Completed"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11310
         TabIndex        =   7
         Top             =   90
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   210
         Left            =   9105
         TabIndex        =   6
         Top             =   105
         Width           =   240
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   7935
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   15225
      _cx             =   26855
      _cy             =   13996
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Rows            =   30
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfSubsidiaryCashTransfers.frx":1CCA
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
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "frmListOfSubsidiaryCashTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    'tnyStatus in faSubsidiaryCashBook
    '   0-Transferred, 1-Clerk Appproved, 2-Payment, 3-Remit Back, 4-Remit Back Approved
    '*********************************************************************************************'
    '                       Form to list all the Subsidiary Cash Books                            '
    '       In the case of Accounts Clerk, it lists only the Subsidiary Cash Books allotted to    '
    '                               that particular Employee                                      '
    '*********************************************************************************************'
    Private Function GetDemandID(numSubLedgerID As Variant)
        Dim mCnnDemand      As New ADODB.Connection
        Dim RecDemand       As New ADODB.Recordset
        Dim mSQLDemand      As String
        Dim objDemand       As New clsDB
        '*********************************************************************************************'
        '          Function to get the Demand ID for a particular Subsidiary Cash Book             '
        '*********************************************************************************************'
        On Error GoTo err
        objDemand.CreateNewConnection mCnnDemand, enuSourceString.Saankhya
        
        mSQLDemand = "Select numDemandID From faIDemandTbl"
        mSQLDemand = mSQLDemand + " Where intTransactionTypeID =" & 1211
        mSQLDemand = mSQLDemand + " And numSubLedgerID = " & numSubLedgerID
        mSQLDemand = mSQLDemand + " And tnyStatus <> 9"
        RecDemand.Open mSQLDemand, mCnnDemand
        If Not (RecDemand.EOF And RecDemand.BOF) Then
            GetDemandID = IIf(IsNull(RecDemand!numDemandID), "", RecDemand!numDemandID)
        End If
        RecDemand.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
        
    Private Function GetUserName(mUserID As Variant) As String
        Dim mCnnUserName    As New ADODB.Connection
        Dim RecUserName     As New ADODB.Recordset
        Dim objUserName     As New clsDB
        Dim mSQLUserName    As String
        
        '*********************************************************************************************'
        '               Function to get the User Name from DB_Masters                                 '
        '*********************************************************************************************'
        On Error GoTo err
        objUserName.CreateNewConnection mCnnUserName, enuSourceString.DBMaster
        
        mSQLUserName = "Select * From GM_User"
        mSQLUserName = mSQLUserName + " Where numUserID = " & mUserID
        RecUserName.Open mSQLUserName, mCnnUserName
        If Not (RecUserName.EOF And RecUserName.BOF) Then
            GetUserName = IIf(IsNull(RecUserName!vchEmpName), "", RecUserName!vchEmpName)
        End If
        RecUserName.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        '*********************************************************************************************'
        '                       Function to get the Seat Name from DB_Masters                         '
        '*********************************************************************************************'
        On Error GoTo err
        objSeatName.CreateNewConnection mCnnSeatName, enuSourceString.DBMaster
        
        mSQLSeatName = "Select * From GL_Seats"
        mSQLSeatName = mSQLSeatName + " Where numSeatID = " & mSeatID
        RecSeatName.Open mSQLSeatName, mCnnSeatName
        If Not (RecSeatName.EOF And RecSeatName.BOF) Then
            GetSeatName = IIf(IsNull(RecSeatName!chvSeatTitle), "", RecSeatName!chvSeatTitle)
        End If
        RecSeatName.Close
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Sub FillvsGrid()
        Dim mRowCount   As Double
        Dim Rec         As New ADODB.Recordset
        Dim RecCheck    As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objDbCheck  As New clsDB
        Dim mSQLCheck   As String
        Dim mSql        As String
        
        On Error GoTo err
        objDbCheck.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        
        mSql = "Select faSubsidiaryCashBook.*, faSubsidiaryAccountHeads.*, faFunctionaries.*, faFunctions.*, faVouchers.intVoucherNo,faSubsidiaryCashBook.dtDate [Date], faSubsidiaryCashBook.fltAmount As Amount, faSubsidiaryCashBook.tnyStatus As Status  From faSubsidiaryCashBook"
        mSql = mSql + " Left Join faSubsidiaryAccountHeads On faSubsidiaryCashBook.intSubsidiaryAccountHeadID = faSubsidiaryAccountHeads.intSubsidiaryAccountHeadID"
        mSql = mSql + " Left Join faFunctionaries On faSubsidiaryCashBook.intFunctionaryID = faFunctionaries.intFunctionaryID"
        mSql = mSql + " Left Join faFunctions On faSubsidiaryCashBook.intFunctionID = faFunctions.intFunctionID"
        mSql = mSql + " Left Join faVouchers On faVouchers.intVoucherID = faSubsidiaryCashBook.intVoucherID"
        mSql = mSql + " Where intTypeID = 50"
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            mSql = mSql + " And faSubsidiaryCashBook.numSeatID = " & gbSeatID
'        End If
        ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
            mSql = mSql + " And faSubsidiaryCashBook.numSeatID = " & gbSeatID
        End If
        Rec.Open mSql, mCnn
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode) + " " + IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!Date), "", Rec!Date)
            If Not (IsNull(Rec!numUserID)) Then
                vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                vsGrid.TextMatrix(mRowCount, 2) = GetUserName(Rec!numUserID)
            End If
            If Not (IsNull(Rec!numSeatID)) Then
                vsGrid.TextMatrix(mRowCount, 3) = GetSeatName(Rec!numSeatID)
            End If
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
            If Not (IsNull(Rec!Status)) Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    If Rec!Status = 4 Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 1
                    Else
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 0
                    End If
                ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    If Rec!Status > 0 Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 1
                    Else
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 0
                    End If
                ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
                    If Rec!Status > 0 Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 1
                    Else
                        vsGrid.Cell(flexcpChecked, mRowCount, 8) = 0
                    End If
                End If
                If Rec!Status = 0 Then
                    vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 16) = &H80000005
                ElseIf Rec!Status = 1 Then
                    vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 16) = &HC0E0FF
                ElseIf Rec!Status = 2 Then
                    vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 16) = &HC0FFC0
                ElseIf Rec!Status = 3 Then
                    vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 16) = &HC0C0FF
                ElseIf Rec!Status = 4 Then
                    vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 16) = &HE0E0E0
                End If
            End If
            
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intID), "", Rec!intID)
            vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!intTransferID), "", Rec!intTransferID)
            vsGrid.TextMatrix(mRowCount, 12) = GetDemandID(val(vsGrid.TextMatrix(mRowCount, 9)))
            vsGrid.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
            vsGrid.TextMatrix(mRowCount, 14) = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            vsGrid.TextMatrix(mRowCount, 15) = IIf(IsNull(Rec!intSubsidiaryAccountHeadID), "", Rec!intSubsidiaryAccountHeadID)
            vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!Status), "", Rec!Status)
            mRowCount = mRowCount + 1
            Rec.MoveNext
        Wend
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub ShowTransferForm()
        '*********************************************************************************************'
        '               Procedure to show the Form to Transfer a Subsidiary Cash Book                 '
        '*********************************************************************************************'
        frmTransferSubsidiaryCashBookMoney.SubSidiaryCashBookID = vsGrid.TextMatrix(vsGrid.Row, 9)
        frmTransferSubsidiaryCashBookMoney.Show vbModal
        Call FillvsGrid
    End Sub
    
    Private Sub ShowTransactionForm(intID As Integer)
        '*********************************************************************************************'
        '               Procedure to show the Form to disburse the Amount                             '
        '*********************************************************************************************'
        If intID > 0 Then
            frmSubsidiaryCashTransactions.intID = intID
        Else
            frmSubsidiaryCashTransactions.intID = vsGrid.TextMatrix(vsGrid.Row, 9)
        End If
        frmSubsidiaryCashTransactions.TransferID = vsGrid.TextMatrix(vsGrid.Row, 10)
        frmSubsidiaryCashTransactions.AmtReceived = val(vsGrid.TextMatrix(vsGrid.Row, 7))
        frmSubsidiaryCashTransactions.DemandID = vsGrid.TextMatrix(vsGrid.Row, 12)
        If vsGrid.TextMatrix(vsGrid.Row, 16) = 4 Then
            frmSubsidiaryCashTransactions.cmdPayment.Enabled = False
            frmSubsidiaryCashTransactions.cmdRemitBack.Enabled = False
            frmSubsidiaryCashTransactions.cmdSave.Enabled = False
        End If
        frmSubsidiaryCashTransactions.Show vbModal
        Call FillvsGrid
    End Sub
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        frmTransferSubsidiaryCashBookMoney.Show vbModal
        Call FillvsGrid
    End Sub

    Private Sub Form_Activate()
        Me.WindowState = 2
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        On Error GoTo err
        vsGrid.SelectionMode = flexSelectionByRow
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Call FillvsGrid
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdNew.Visible = False
        ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdNew.Visible = False
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub Form_Resize()
        If Me.WindowState <> 2 Then
            Me.Left = 0
            Me.Top = 0
            Me.Width = 15360
            Me.Height = 9450
        End If
    End Sub

    Private Sub vsGrid_Click()
        'If gbUserTypeID = 3 Then
'        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
'            If vsGrid.col = 8 Then
'                If vsGrid.TextMatrix(vsGrid.row, 0) <> "" Then
'                    vsGrid.Editable = flexEDKbdMouse
'                Else
'                    vsGrid.Editable = flexEDNone
'                End If
'            Else
'                vsGrid.Editable = flexEDNone
'            End If
'        End If
    End Sub
    
    Private Sub vsGrid_DblClick()
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim RecChild        As New ADODB.Recordset
        Dim RecCheck        As New ADODB.Recordset
        Dim mSqlChild       As String
        Dim mRowCount       As Double
        Dim mID             As Integer
        Dim mMAXID          As Variant
        
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If vsGrid.Row > 0 Then
            ' To get the ID of Subsidiary Cash Book in which disbursement occurs
            If vsGrid.TextMatrix(vsGrid.Row, 16) = 2 Or vsGrid.TextMatrix(vsGrid.Row, 16) = 3 Or vsGrid.TextMatrix(vsGrid.Row, 16) = 4 Then
                mSql = "Select intID From faSubsidiaryCashBook"
                mSql = mSql + " Where intTransferID =" & vsGrid.TextMatrix(vsGrid.Row, 10)
                mSql = mSql + " And intTypeID = 20"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mID = IIf(IsNull(Rec!intID), 0, Rec!intID)
                End If
                Rec.Close
            End If
            If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                If vsGrid.TextMatrix(vsGrid.Row, 16) = 0 Then
                    Call ShowTransferForm
                    Exit Sub
                Else
                    Call ShowTransactionForm(mID)
                End If
            ElseIf gbSeatGroupID = gbSeatGroupChiefCashier Then
                If vsGrid.TextMatrix(vsGrid.Row, 16) = 0 Then
                    Call ShowTransferForm
                    Exit Sub
                Else
                    Call ShowTransactionForm(mID)
                End If
            End If
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                If vsGrid.TextMatrix(vsGrid.Row, 16) < 3 Then
                    Call ShowTransferForm
                    Exit Sub
                Else
                    Call ShowTransactionForm(mID)
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsGrid.ToolTipText = ""
        If vsGrid.Col <> 8 Then
            vsGrid.ToolTipText = vsGrid.TextMatrix(vsGrid.Row, vsGrid.Col)
        End If
    End Sub

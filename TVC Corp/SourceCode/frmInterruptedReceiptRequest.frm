VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInterruptedReceiptRequest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Request For Interrupted Receipt"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReject 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reject"
      Height          =   345
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2565
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   390
      TabIndex        =   6
      Top             =   2655
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   40152
      MinDate         =   25569
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2565
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      Height          =   345
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2565
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "No"
      Height          =   345
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2910
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VSFlex8LCtl.VSFlexGrid vsRequests 
      Height          =   2310
      Left            =   45
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   6120
      _cx             =   10795
      _cy             =   4075
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
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
      FormatString    =   $"frmInterruptedReceiptRequest.frx":0000
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
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Yes"
      Height          =   345
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2910
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblRequest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to send request for Interrupted Receipt ?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   -2895
      TabIndex        =   0
      Top             =   2730
      Visible         =   0   'False
      Width           =   3420
   End
End
Attribute VB_Name = "frmInterruptedReceiptRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mStatus     As Variant
    Dim mCount      As Variant
    
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdNo_Click()
        Unload Me
    End Sub
'''    Private Sub cmdReject_Click()
'''        Dim mRowCount   As Integer
'''        For mRowCount = 1 To mCount
'''         If vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked Then
'''                frmReject.Mode = 3
'''                frmReject.RequestTypeID = vsRequests.TextMatrix(mRowCount, 5)   'CounterID
'''                frmReject.Show vbModal
'''                cmdReject.Enabled = False
'''                cmdSave.Enabled = False
'''            End If
'''        Next
'''    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mSanction   As Integer
        Dim mRowCount   As Integer
    
        '*********************************************************************************************'
        '               Procedure to Approve Interrupt Receipt Request                                '
        '*********************************************************************************************'
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If vsRequests.Rows > 1 Then
            If gbSeatGroupID = gbSeatGroupCashSuperintended Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then 'gbUserTypeID <> 3 Then
                For mRowCount = 1 To mCount
                    mSql = "Select * From faInterruptedReceiptBooks"
                    mSql = mSql + " Where intCounterID = " & val(vsRequests.TextMatrix(mRowCount, 5))
                    mSql = mSql + " And tnyClosed <> 1"
                    Rec.Open mSql, mCnn
                    If (Rec.EOF Or Rec.BOF) Then
                        If vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked Then
                            MsgBox "Please issue Interrupted Receipt Book for " & vsRequests.TextMatrix(mRowCount, 1) & "!", vbInformation
                            Exit Sub
                        End If
                    End If
                    Rec.Close
                    
                    If vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked Then
                        mSanction = 2
                    Else
                        mSanction = 1
                    End If
                    mSql = "Update faInterruptedRequests"
                    mSql = mSql + " Set tnyStatus =" & mSanction
                    mSql = mSql + " Where numUserID=" & vsRequests.TextMatrix(mRowCount, 4)
                    mSql = mSql + " And intCounterID = " & vsRequests.TextMatrix(mRowCount, 5)
                    mCnn.Execute mSql
                Next
            End If
            MsgBox "Successfully Saved", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdYes_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mAryIn  As Variant
        Dim mSql    As String
        Dim mDate   As Date
        Dim mLFAStat As Integer
        Dim mDiff   As Integer
        Dim mStat   As Integer
        '*********************************************************************************************'
        '               Procedure to send Interrupt Receipt Request                                   '
        '*********************************************************************************************'
        'On Error GoTo Err
        
        
        If dtpdate.Value > gbTransactionDate Or dtpdate.Value < dtpdate.MinDate Then
            MsgBox "Please check the date"
            dtpdate.Value = gbTransactionDate
            Exit Sub
        End If
        
        mDate = dtpdate.Value
        If mStatus <> 1 And mStatus <> 2 Then
             '-----------------LAST POSTING VALIDATION------------------
            If CDate(mDate) <= CDate(gbLastPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                Exit Sub
            End If
            '-------------------------------------------------------------
        End If
        If gbLBPanchayat = 1 Then
            If mStatus = "" Then
            If mDate < gbRPOnlinedate Then
                mSql = "Date must be greater than the ONLINE date" & vbCrLf
                mSql = mSql & "[" & DdMmmYy(CDate(gbRPOnlinedate)) & "]"
                MsgBox mSql, vbInformation
                Exit Sub
            End If
            End If
        End If

'        Call CheckInterruptReceiptRequestStatus

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If mStatus = "" Then
            If CDate(dtpdate.Value) < CDate(gbStartingDate) And CDate(dtpdate.Value) < CDate(gbEndingDate) Then
                mSql = "Select * From faBLSubmission Where intYearID=" & gbFinancialYearID - 1
                Rec.Open mSql, mCnn
                
                If Not (Rec.EOF And Rec.BOF) Then
                    mLFAStat = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
                Else
                    mLFAStat = 0
                End If
                Rec.Close
                If mLFAStat = 0 Or mLFAStat = 1 Or mLFAStat = 4 Then
'                    MsgBox "AFS is Submitted to LFA you can't do Interrupted Receipt in Previous year", vbApplicationModal
'                    Exit Sub
                    If gbLBPanchayat = 1 Then
                        If gbLocalBodyID = 975 Then
                            mSql = "Select DATEDIFF(day, '1/Apr/2020', getdate()) as Diff From faConfig "
                            Rec.Open mSql, mCnn
                            If Not (Rec.EOF And Rec.BOF) Then
                                mDiff = IIf(IsNull(Rec!diff), "", Rec!diff)
                            End If
                            Rec.Close
                            If mDiff > 246 Then  ' valavnnur 86 80 65 54 Then '10/jun/19 upto Jun 20
                                mStat = 1
                            End If
                         Else
                            mStat = 1
                        End If
                    Else
                     mSql = "Select DATEDIFF(day, '1/Apr/2020', getdate()) as Diff From faConfig "
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mDiff = IIf(IsNull(Rec!diff), "", Rec!diff)
                        End If
                        Rec.Close
                        If mDiff > 121 Then  '86 80 65 54 Then '10/jun/19 upto Jun 20
                            mStat = 1
                        End If
                    End If
                    If mStat = 1 Then
                        MsgBox "Previous year Process is Disabled", vbApplicationModal
                        Exit Sub
                    ElseIf MsgBox("Requested date is in Previous Financial Year.Do you want to Proceed?", vbYesNo, "Saankhya") = vbNo Then
                        Exit Sub
                    End If
                Else
                    If gbLBPanchayat = 1 Then
                        mStat = 1
                    Else
                    ''''''' Disable Previous Entry
                        mSql = "Select DATEDIFF(day, '1/Apr/2020', getdate()) as Diff From faConfig "
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mDiff = IIf(IsNull(Rec!diff), "", Rec!diff)
                        End If
                        Rec.Close
                        If mDiff > 100 Then  ' 80 65 54 Then '10/jun/19 upto Jun 20
                            mStat = 1
                        End If
                    End If
                    If mStat = 1 Then
                        MsgBox "Previous year Process is Disabled", vbApplicationModal
                        Exit Sub
                    ElseIf MsgBox("Requested date is in Previous Financial Year.Do you want to Proceed?", vbYesNo, "Saankhya") = vbNo Then
                        Exit Sub
                    End If
                End If
                
                
            End If
            mAryIn = Array(gbCounterID, _
                           gbUserID, _
                           1, _
                           gbTransactionDate, _
                           1, _
                           dtpdate.Value)
            objdb.ExecuteSP "spSaveInterruptedRequest", mAryIn, , , mCnn, adCmdStoredProc
            If gbLBPanchayat Then
                MsgBox "Request sent to Secretary", vbInformation
            ElseIf gbLBType = 4 Then
                MsgBox "Request sent to Nodal Officer", vbInformation
            ElseIf gbLBType = 3 Then
                MsgBox "Request sent to Accounts Supdt", vbInformation
            End If
        End If
        If mStatus = 1 Or mStatus = 2 Then
            mSql = "Delete From faInterruptedRequests"
            mSql = mSql + " Where intCounterID =" & gbCounterID
            mSql = mSql + " And numUserID =" & gbUserID
            mSql = mSql + " And intTypeID = 1"
            mCnn.Execute mSql
            MsgBox "Request Cancelled Successfully", vbInformation
        End If

        If frmMenu.RequestforInterruptedReceipt.Caption = "Cancel request for InterruptedReceipt" Then
            frmMenu.RequestforInterruptedReceipt.Caption = "Request for InterruptedReceipt"
        Else
            frmMenu.RequestforInterruptedReceipt.Caption = "Cancel request for InterruptedReceipt"
        End If
        Unload Me
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_Load()
        
        
        Unload frmReceiptsCounter
        
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mRowCount   As Integer
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open "Select isNull(max(dtDate),1)[dtDate] From faVouchers Where tnyVoucherGroupID = 4", mCnn
        If Rec!dtDate < "01/01/1970" Then
            dtpdate.MinDate = CDate("01/01/1970")
        Else
            'dtpDate.MinDate = CDate("01/Apr/" + CStr(gbFinancialYearID)) 'Rec!dtDate
            dtpdate.MinDate = CDate(DateAdd("yyyy", -1, gbStartingDate))
        End If
        dtpdate.MaxDate = gbTransactionDate
        Rec.Close
        mCnn.Close
        
        'Note :-  Operator and Cash Seat Group
        'If (gbUserTypeID = 3 And gbCounterSectionID = gbJSKSectionID) Then                       'Or gbSeatGroupID = gbSeatByDeveloper Then
        If ((gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier) And gbCounterSectionID = gbJSKSectionID) Then
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            
            mStatus = ""
            mSql = "Select tnyStatus From faInterruptedRequests"
            mSql = mSql + " Where numUserID =" & gbUserID
            mSql = mSql + " And intCounterID =" & gbCounterID
            mSql = mSql + " And intTypeID = 1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            End If
            Rec.Close
            
            'Note:- Form Resizing and Displyaing Message
            Me.Width = 4770
            Me.Height = 2445
            If mStatus <> "" Then
                If mStatus = 1 Or mStatus = 2 Then
                    lblRequest.Caption = "Do you want to cancel request for Interrupted Receipt ?"
                Else
                    lblRequest.Caption = "Do you want to send Interrupted Receipt request for the following Date?"
                    dtpdate.Visible = True
                    dtpdate.Left = 1620
                    dtpdate.Top = 845
                End If
                lblRequest.Visible = True
                lblRequest.Left = 630
                lblRequest.Top = 245
            Else
                lblRequest.Caption = "Do you want to send Interrupted Receipt request for the following Date?"
                dtpdate.Visible = True
                dtpdate.Left = 1620
                dtpdate.Top = 845
                lblRequest.Visible = True
                lblRequest.Left = 630
                lblRequest.Top = 245
            End If
            cmdYes.Visible = True
            cmdYes.Left = 1095
            cmdYes.Top = 1300
            cmdNo.Visible = True
            cmdNo.Left = 2355
            cmdNo.Top = 1300
        End If
        
        'Note:- Approving Officer
        'If gbUserTypeID <> 3 Then
        If gbSeatGroupID = gbSeatGroupCashSuperintended Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then 'gbUserTypeID <> 3 Then
            mCount = 0
            Me.Width = 6300
            Me.Height = 3570
            vsRequests.Visible = True
            vsRequests.Clear 1, 1
            vsRequests.Rows = 1
            cmdSave.Visible = True
'            cmdReject.Visible = True
            cmdCancel.Visible = True
            
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            
            mSql = "Select * From faInterruptedRequests"
            mSql = mSql + " Inner Join faUser On faInterruptedRequests.numUserID = faUser.numUserID"
            mSql = mSql + " Inner Join faCounters On faInterruptedRequests.intCounterID = faCounters.intCounterID"
            mSql = mSql + " Where intTypeID = 1"
            Rec.Open mSql, mCnn
            mRowCount = 1
            While Not Rec.EOF
                vsRequests.Rows = vsRequests.Rows + 1
                vsRequests.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                vsRequests.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                'vsRequests.TextMatrix(mRowCount, 2) = DdMmmYy(IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate))
                vsRequests.TextMatrix(mRowCount, 2) = DdMmmYy(IIf(IsNull(Rec!dtReceiptDate), "", Rec!dtReceiptDate))
                If (Rec!tnyStatus = 1) Then
                    vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbUnchecked
                ElseIf (Rec!tnyStatus = 2) Then
                    vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked
                End If
                vsRequests.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                vsRequests.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!intCounterID), "", Rec!intCounterID)
                mRowCount = mRowCount + 1
                mCount = mCount + 1
                Rec.MoveNext
            Wend
            Rec.Close
            mCnn.Close
        End If
        Call SetgbLastPostingDate
    End Sub

    Private Sub vsRequests_Click()
        If vsRequests.Col = 3 Then
            If vsRequests.TextMatrix(vsRequests.Row, 0) <> "" Then
                vsRequests.Editable = flexEDKbdMouse
            Else
                vsRequests.Editable = flexEDNone
                vsRequests.Cell(flexcpChecked, vsRequests.Row, 3) = vbUnchecked
            End If
        End If
    End Sub



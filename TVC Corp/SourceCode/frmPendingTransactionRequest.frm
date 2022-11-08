VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPendingTransactionRequest 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   345
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2595
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3750
      TabIndex        =   7
      Top             =   2955
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   345
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2985
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   345
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2985
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2595
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2595
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   3375
      TabIndex        =   0
      Top             =   2910
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   582
      _Version        =   393216
      Format          =   17629185
      CurrentDate     =   40152
      MinDate         =   25569
   End
   Begin VSFlex8LCtl.VSFlexGrid vsRequests 
      Height          =   2400
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   6120
      _cx             =   10795
      _cy             =   4233
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
      FormatString    =   $"frmPendingTransactionRequest.frx":0000
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
      Left            =   -2400
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   3420
   End
End
Attribute VB_Name = "frmPendingTransactionRequest"
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
'''            If vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked Then
'''                frmReject.Mode = 11
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
        Dim objDb       As New clsDB
        Dim mSanction   As Integer
        Dim mRowCount   As Integer
        
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'If gbUserTypeID = 2 Or gbUserTypeID = 4 Then
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            For mRowCount = 1 To mCount
                If vsRequests.Cell(flexcpChecked, mRowCount, 3) = vbChecked Then
                    mSanction = 2
                Else
                    mSanction = 1
                End If
                mSql = "Update faInterruptedRequests"
                mSql = mSql + " Set tnyStatus =" & mSanction
                mSql = mSql + " Where numUserID=" & vsRequests.TextMatrix(mRowCount, 4)
                mSql = mSql + " And intCounterID = " & vsRequests.TextMatrix(mRowCount, 5)
                mSql = mSql + " And intTypeID = 2"
                mCnn.Execute mSql
            Next
        End If
        MsgBox "Successfully Saved", vbInformation
        CheckPendingTransactionStatus
        Unload Me
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdYes_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb   As New clsDB
        Dim mAryIn  As Variant
        Dim mSql    As String
        
'        Call CheckInterruptReceiptRequestStatus
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If txtDate.Text <> "" Then
            If CDate(txtDate.Text) >= gbTransactionDate Then
                MsgBox "Please enter a valid date", vbInformation
                dtpDate.SetFocus
                Exit Sub
            End If
            mSql = "Select top 1 dtStartingDate as Mindate From faFinancialYear Order by intFinancialYearID Asc"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
               If CDate(txtDate.Text) < Rec!MinDate Then
                    MsgBox "Financial Year settings is not Present For this Date, Please Conrtact System Administrator", vbInformation
                    Exit Sub
               End If
            End If
            Rec.Close
        End If
        
        If mStatus = "" Then
            mAryIn = Array(gbCounterID, _
                           gbUserID, _
                           1, _
                           txtDate.Text, _
                           2)
            'objDb.ExecuteSP "spSaveInterruptedRequest", mAryIn, , , mCnn, adCmdStoredProc'NOTE:SP CHANGED
            MsgBox "Request sent to Nodal Officer", vbInformation
        End If
        If mStatus = 1 Or mStatus = 2 Then
            mSql = "Delete From faInterruptedRequests"
            mSql = mSql + " Where intCounterID =" & gbCounterID
            mSql = mSql + " And numUserID =" & gbUserID
            mSql = mSql + " And intTypeID = 2"
            mCnn.Execute mSql
            MsgBox "Request Cancelled Successfully", vbInformation
        End If
        If frmMenu.RequestforPendingTransactions.Caption = "Cancel Request for Enable Previous Year's Transaction" Then
            frmMenu.RequestforPendingTransactions.Caption = "Request for Pending Transactions"
            frmMenu.MenuDetails
        Else
            frmMenu.RequestforPendingTransactions.Caption = "Cancel request for Pending Transactions"
            
        End If
        Call CheckPendingTransactionStatus
        Unload Me
        
    End Sub
    
    Private Sub dtpDate_CloseUp()
        txtDate.Text = CheckDateInMMM(dtpDate.value)
    End Sub
    Private Sub CheckPendingTransactionStatus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objDb   As New clsDB
        Dim mStatus As Variant
        Dim mMenu   As Control
        Dim mRequestDate As String
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mStatus = ""
        mSql = "Select tnyStatus,dtRequestDate From faInterruptedRequests"
        mSql = mSql + " Where  intTypeID = 2 "
        'If gbUserTypeID = 3 Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            mSql = mSql + " And intCounterID =" & gbCounterID
            mSql = mSql + " And numUserID =" & gbUserID
        End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            mRequestDate = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
        End If
        Rec.Close
        mCnn.Close
        If mStatus <> "" Then
            If mStatus = 1 Then
                'If gbUserTypeID = 4 Or gbUserTypeID = 2 Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    'Timer1.Enabled = True
                    frmMenu.lblPreYear.Caption = "Pending Approval for Pevious Year's Transaction Request"
                Else
                    frmMenu.lblPreYear.Caption = "Pevious Year's Transaction Request is Pending for Approval"
                    For Each mMenu In frmMenu.Controls
                        If TypeOf mMenu Is Menu Then
                            Debug.Print mMenu.Name
                            If mMenu.Name = "RequestforPendingTransactions" Or mMenu.Name = "Utilities" Or mMenu.Name = "Exit" Or mMenu.Name = "LogOut" Then
                                mMenu.Enabled = True
                            Else
                                mMenu.Enabled = False
                            End If
                        End If
                    Next
                    frmMenu.RequestforPendingTransactions.Caption = "Cancel Request for Enable Previous Year's Transaction"
                End If
            ElseIf mStatus = 2 Then
                'If gbUserTypeID = 4 Or gbUserTypeID = 2 Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                   ' Timer1.Enabled = False
                Else
                    'mPreYearMode = True
                    'gbTransactionDate = mRequestDate
                    If Month(gbTransactionDate) < 4 Then
                        gbFinancialYearID = Year(gbTransactionDate) - 1
                    Else
                        gbFinancialYearID = Year(gbTransactionDate)
                    End If
                    frmMenu.RequestforPendingTransactions.Caption = "Cancel Request for Enable Previous Year's Transaction"
                    'Timer1.Enabled = True
                    frmMenu.lblPreYear.Caption = "Pevious Year Transaction Mode is Enabled"
                End If
            End If
        Else
            'Timer1.Enabled = False
            frmMenu.RequestforPendingTransactions.Caption = "Request for Enable Previous Year's Transaction"
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub Form_Load()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mRowCount   As Integer
        
        On Error GoTo err
        dtpDate.value = Date
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'Note :-  Operator and Cash Seat Group
        'If gbUserTypeID = 3 Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            mStatus = ""
            mSql = "Select tnyStatus From faInterruptedRequests"
            mSql = mSql + " Where numUserID =" & gbUserID
            mSql = mSql + " And intCounterID =" & gbCounterID
            mSql = mSql + " And intTypeID = 2"
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
                    Me.Caption = "Cancel request for Pending Transaction"
                    lblRequest.Caption = "Do you want to cancel request for Pending Transactions ?"
                Else
                    Me.Caption = "Request for Pending Transaction"
                    lblRequest.Caption = "Do you want to send request for Pending Transactions?"
                    txtDate.Visible = True
                    txtDate.Left = 1620
                    txtDate.Top = 850
                    dtpDate.Visible = True
                    dtpDate.Left = 2970
                    dtpDate.Top = 845
                End If
                lblRequest.Visible = True
                lblRequest.Left = 630
                lblRequest.Top = 245
            Else
                Me.Caption = "Request for Pending Transaction"
                lblRequest.Caption = "Do you want to send request for Pending Transactions?"
                txtDate.Visible = True
                txtDate.Left = 1620
                txtDate.Top = 850
                dtpDate.Visible = True
                dtpDate.Left = 2970
                dtpDate.Top = 845
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
        
        'Note:- Approving Office
        'If gbUserTypeID = 1 Or gbUserTypeID = 2 Or gbUserTypeID = 4 Then
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            mCount = 0
            Me.Width = 6300
            Me.Height = 3570
            Me.Caption = "Approval of request for Pending Transaction"
            vsRequests.Visible = True
            cmdSave.Visible = True
'            cmdReject.Visible = True
            cmdCancel.Visible = True
            
            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
            
            mSql = "Select * From faInterruptedRequests"
            mSql = mSql + " Inner Join faUser On faInterruptedRequests.numUserID = faUser.numUserID"
            mSql = mSql + " Inner Join faCounters On faInterruptedRequests.intCounterID = faCounters.intCounterID"
            mSql = mSql + " Where intTypeID = 2"
            Rec.Open mSql, mCnn
            mRowCount = 1
            While Not Rec.EOF
                vsRequests.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                vsRequests.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                vsRequests.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
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
        Exit Sub
err:
        MsgBox err.Description
    End Sub

Private Sub txtDate_LostFocus()
    txtDate.Text = CheckDateInMMM(txtDate.Text)
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

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViewPaymentOrder 
   Caption         =   "~List of  Payment Orders ~"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15780
   Icon            =   "frmViewPaymentOrder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   15615
      Begin VB.CheckBox chkListToFwd 
         Alignment       =   1  'Right Justify
         Caption         =   "List of Forwarded Payment Orders"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4380
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   3330
      End
      Begin VB.CheckBox chkListToVerify 
         Alignment       =   1  'Right Justify
         Caption         =   "List of Verified Payment Orders"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   31
         Top             =   1410
         Width           =   3210
      End
      Begin VB.TextBox txtAmount2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8685
         TabIndex        =   28
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtAmount1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7095
         TabIndex        =   27
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtPayOrderNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7095
         TabIndex        =   19
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "..."
         Height          =   270
         Left            =   10020
         TabIndex        =   17
         Top             =   690
         Width           =   285
      End
      Begin VB.CommandButton cmdForwardedSeat 
         Caption         =   "..."
         Height          =   285
         Left            =   3705
         TabIndex        =   16
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7095
         TabIndex        =   15
         Top             =   690
         Width           =   2895
      End
      Begin VB.TextBox txtGeneratedSeat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2505
         TabIndex        =   14
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtForwardedSeat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2505
         TabIndex        =   13
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdGeneratedSeat 
         Caption         =   "..."
         Height          =   285
         Left            =   3705
         TabIndex        =   12
         Top             =   690
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   13920
         TabIndex        =   3
         Top             =   690
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61079553
         CurrentDate     =   40197
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   13920
         TabIndex        =   1
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61079553
         CurrentDate     =   40197
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12720
         TabIndex        =   2
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12720
         TabIndex        =   0
         Top             =   360
         Width           =   1185
      End
      Begin VB.CheckBox chkListToApprove 
         Alignment       =   1  'Right Justify
         Caption         =   "List of Approved Payment Orders"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11010
         TabIndex        =   4
         Top             =   1365
         Width           =   3210
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8460
         TabIndex        =   29
         Top             =   825
         Width           =   225
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6375
         TabIndex        =   26
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label lblTransactionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5550
         TabIndex        =   24
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label lblPaymentOrderNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Order No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5475
         TabIndex        =   23
         Top             =   330
         Width           =   1590
      End
      Begin VB.Label lblForwardedSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded Seat"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1065
         TabIndex        =   11
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12000
         TabIndex        =   10
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11760
         TabIndex        =   9
         Top             =   345
         Width           =   915
      End
      Begin VB.Label lblPayOrderGeneratedSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Generated Seat"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1005
         TabIndex        =   8
         Top             =   720
         Width           =   1425
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6045
      Left            =   90
      TabIndex        =   6
      Top             =   390
      Width           =   15675
      _cx             =   27649
      _cy             =   10663
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmViewPaymentOrder.frx":1CCA
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   15720
      TabIndex        =   5
      Top             =   8235
      Width           =   15780
      Begin VB.CommandButton cmdViewPOReport 
         Caption         =   "View Payment Order &Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6195
         TabIndex        =   30
         Top             =   105
         Width           =   3030
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   105
         Width           =   1500
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "List of Payment Orders"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15750
   End
End
Attribute VB_Name = "frmViewPaymentOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private mViewMode As Integer
    Dim mPreviousYearMode As Integer

    '*********************************************************************************************'
    '                           Form to list all the Payment Orders                               '
    '*********************************************************************************************'
    Private Function GetSeatName(mSeatID As Variant)
        Dim mCnnSeatName    As New ADODB.Connection
        Dim RecSeatName     As New ADODB.Recordset
        Dim objSeatName     As New clsDB
        Dim mSQLSeatName    As String
        
        On Error GoTo err:
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
        MsgBox (Error$)
    End Function
    
    Private Sub FormIntialize()
        Dim mCrl As Control
        
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.Value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        
    End Sub
    Private Sub PreviousYearMode()
        If CDate(txtDateFrom.Text) < CDate(gbStartingDate) And CDate(txtDateTo.Text) < CDate(gbEndingDate) Then
            mPreviousYearMode = 1
        Else
            mPreviousYearMode = 0
        End If
    End Sub
    Private Sub FetchPaymentOrder()
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mArrIn      As Variant
        Dim mSQL        As String
        Dim mFromDate   As String
        Dim mToDate     As String
        Dim mStatus     As Variant
        Dim mFwdSeat     As Variant
        '*********************************************************************************************'
        '                           Procedure to list all the Payment Orders                          '
        '*********************************************************************************************'
        On Error GoTo err:
        mStatus = ""
        vsGrid.Rows = 1
        objDB.SetConnection mCnn
        
        If txtDateFrom.Text = "" Then
            mSQL = "Select dtStartingDate From  faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFromDate = IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
            End If
            Rec.Close
        Else
            mFromDate = CheckDateInMMM(txtDateFrom.Text)
        End If
        
        If txtDateTo.Text = "" Then
            mSQL = "Select dtEndingDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mToDate = IIf(IsNull(Rec!dtEndingDate), "", CheckDateInMMM(Rec!dtEndingDate))
                'mToDate = DateAdd("M", -1, mToDate)
            End If
            Rec.Close
        Else
            mToDate = CheckDateInMMM(txtDateTo.Text)
        End If
        If txtForwardedSeat.Tag = "" Then
            mFwdSeat = ""
        Else
            mFwdSeat = txtForwardedSeat.Tag
        End If
        If chkListToApprove.Value = vbChecked Then
            mStatus = 1
        Else
            If gbLBPanchayat Then
                mStatus = 0
                If chkListToVerify.Value = vbChecked Then
                    mFwdSeat = ""
                ElseIf chkListToFwd.Value = vbChecked And gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    mStatus = 5
                    mFwdSeat = ""
                Else
                    If gbSeatGroupID = gbSeatGroupAccountSectionClerk Or gbSeatGroupID = gbSeatGroupAccountsClerk Then
                        mStatus = 0
                        vsGrid.ColHidden(11) = False
                    ElseIf gbSeatGroupID = gbSeatGroupHeadClerk Or gbSeatGroupID = gbSeatGroupAssistantSecretary Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        mFwdSeat = gbSeatID
                        mStatus = 5
                    ElseIf gbSeatGroupID = gbSeatGroupAccountsSuperintended And gbLBType = 4 Then
                        mStatus = 3
                        mFwdSeat = gbSeatID
                    End If
                    
                End If
                
                If mPreviousYearMode = 1 Then
                    If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
                        mStatus = 0
                        vsGrid.ColHidden(11) = False
                    ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                        mStatus = "0,5"
                        vsGrid.ColHidden(11) = False
                    ElseIf gbSeatGroupID = gbSeatGroupHeadClerk Or gbSeatGroupID = gbSeatGroupAssistantSecretary Or gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        mFwdSeat = gbSeatID
                        mStatus = "5,3"
                    ElseIf gbSeatGroupID = gbSeatGroupAccountsSuperintended And gbLBType = 4 Then
                        mStatus = 3
                        mFwdSeat = gbSeatID
                    End If
                End If
                
            Else
                mStatus = 0
                mFwdSeat = ""
            End If
        End If
        If mPreviousYearMode = 1 Then
            mSQL = " Select *,faPayOrder.vchDescription As Descriptions,faPayOrder.tnyStatus As Status,faPayOrder.numSeatID As SeatID  From faPayOrder" & vbNewLine
            mSQL = mSQL + " Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID And faPayOrderChild.tnyCategoryFlag = 3" & vbNewLine
            mSQL = mSQL + " Inner Join faTransactionType ON faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID" & vbNewLine
            mSQL = mSQL + " Inner Join faUser On faUser.numUserID = faPayOrder.numUserID" & vbNewLine
            mSQL = mSQL + " Left Join faVouchers On faVouchers.intVoucherID = faPayOrder.intVoucherID" & vbNewLine
            mSQL = mSQL + " Where (tnyCancelled <> 1 Or tnyCancelled Is Null)" & vbNewLine
            mSQL = mSQL + " And isnull(numFwdSeatID,0)  Like    '" & IIf(txtForwardedSeat.Tag = "", "%", txtForwardedSeat.Tag) & "'" & vbNewLine
            mSQL = mSQL + " And faPayOrder.numSeatID   LIke '" & IIf(txtGeneratedSeat.Tag = "", "%", txtGeneratedSeat.Tag) & "'" & vbNewLine
            mSQL = mSQL + " And dtPayOrderDate Between '" & mFromDate & "' And '" & mToDate & "'" & vbNewLine
            mSQL = mSQL + " And faPayOrder.tnyStatus in ( " & mStatus & ")" & vbNewLine
            mSQL = mSQL + " And faPayOrder. intTransactionTypeID Like '" & IIf(txtTransactionType.Tag = "", "%", txtTransactionType.Tag) & "'" & vbNewLine
            mSQL = mSQL + " And faPayOrder. vchPayOrderNo Like '" & IIf(txtPayOrderNo.Text = "", "%", txtPayOrderNo.Text) & "'" & vbNewLine
            'mSQL = mSQL + " And numAmount  Between '" & IIf(txtAmount1.Text = "", "%", val(txtAmount1.Text)) & "' And  '" & val(txtAmount2.Text) & "'" & vbNewLine
            mSQL = mSQL + " And isNull(intModuleID,0)=96 "
            mSQL = mSQL + " Order By faPayOrder.vchPayOrderNo Desc"
            Set Rec = objDB.ExecuteSP(mSQL, , , False, mCnn, adCmdText)
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBPanchayat = 1 Then
                    mSQL = " Select *,faPayOrder.vchDescription As Descriptions,faPayOrder.tnyStatus As Status,faPayOrder.numSeatID As SeatID  From faPayOrder" & vbNewLine
                    mSQL = mSQL + " Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID And faPayOrderChild.tnyCategoryFlag = 3" & vbNewLine
                    mSQL = mSQL + " Inner Join faTransactionType ON faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID" & vbNewLine
                    mSQL = mSQL + " Inner Join faUser On faUser.numUserID = faPayOrder.numUserID" & vbNewLine
                    mSQL = mSQL + " Left Join faVouchers On faVouchers.intVoucherID = faPayOrder.intVoucherID" & vbNewLine
                    mSQL = mSQL + " Where (tnyCancelled <> 1 Or tnyCancelled Is Null)" & vbNewLine
                    mSQL = mSQL + " And isnull(numFwdSeatID,0)  Like    '" & IIf(mFwdSeat = "", "%", mFwdSeat) & "'" & vbNewLine
                    mSQL = mSQL + " And faPayOrder.numSeatID   LIke '" & IIf(txtGeneratedSeat.Tag = "", "%", txtGeneratedSeat.Tag) & "'" & vbNewLine
                    mSQL = mSQL + " And dtPayOrderDate Between '" & mFromDate & "' And '" & mToDate & "'" & vbNewLine
                    If chkListToApprove.Value = vbChecked Then
                        mSQL = mSQL + " And faPayOrder.tnyStatus in (1) "
                    ElseIf chkListToVerify.Value = vbChecked Then
                        mSQL = mSQL + " And faPayOrder.tnyStatus in (3,1) "
                    Else
                        mSQL = mSQL + " And faPayOrder.tnyStatus in (5,3) "
                    End If
                    mSQL = mSQL + " And faPayOrder. intTransactionTypeID Like '" & IIf(txtTransactionType.Tag = "", "%", txtTransactionType.Tag) & "'" & vbNewLine
                    mSQL = mSQL + " And faPayOrder. vchPayOrderNo Like '" & IIf(txtPayOrderNo.Text = "", "%", txtPayOrderNo.Text) & "'" & vbNewLine
                    mSQL = mSQL + " Order By faPayOrder.vchPayOrderNo Desc"
                    Set Rec = objDB.ExecuteSP(mSQL, , , False, mCnn, adCmdText)
                ElseIf chkListToApprove.Value = vbChecked Or chkListToVerify.Value = vbChecked And gbLBPanchayat = 1 Then
                    mSQL = " Select *,faPayOrder.vchDescription As Descriptions,faPayOrder.tnyStatus As Status,faPayOrder.numSeatID As SeatID  From faPayOrder" & vbNewLine
                    mSQL = mSQL + " Inner Join faPayOrderChild ON faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID And faPayOrderChild.tnyCategoryFlag = 3" & vbNewLine
                    mSQL = mSQL + " Inner Join faTransactionType ON faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID" & vbNewLine
                    mSQL = mSQL + " Inner Join faUser On faUser.numUserID = faPayOrder.numUserID" & vbNewLine
                    mSQL = mSQL + " Left Join faVouchers On faVouchers.intVoucherID = faPayOrder.intVoucherID" & vbNewLine
                    mSQL = mSQL + " Where (tnyCancelled <> 1 Or tnyCancelled Is Null)" & vbNewLine
                    mSQL = mSQL + " And isnull(numFwdSeatID,0)  Like    '" & IIf(mFwdSeat = "", "%", mFwdSeat) & "'" & vbNewLine
                    mSQL = mSQL + " And faPayOrder.numSeatID   LIke '" & IIf(txtGeneratedSeat.Tag = "", "%", txtGeneratedSeat.Tag) & "'" & vbNewLine
                    mSQL = mSQL + " And dtPayOrderDate Between '" & mFromDate & "' And '" & mToDate & "'" & vbNewLine
                    If chkListToApprove.Value = vbChecked Then
                        mSQL = mSQL + " And faPayOrder.tnyStatus in (1) "
                    ElseIf chkListToVerify.Value = vbChecked Then
                        mSQL = mSQL + " And faPayOrder.tnyStatus in (3,1) "
                    End If
                    mSQL = mSQL + " And faPayOrder. intTransactionTypeID Like '" & IIf(txtTransactionType.Tag = "", "%", txtTransactionType.Tag) & "'" & vbNewLine
                    mSQL = mSQL + " And faPayOrder. vchPayOrderNo Like '" & IIf(txtPayOrderNo.Text = "", "%", txtPayOrderNo.Text) & "'" & vbNewLine
                    mSQL = mSQL + " Order By faPayOrder.vchPayOrderNo Desc"
                    Set Rec = objDB.ExecuteSP(mSQL, , , False, mCnn, adCmdText)
     
                '''''------------------
        Else
            
                mArrIn = Array(mFromDate, _
                            mToDate, _
                            mStatus, _
                            IIf(mFwdSeat = "", "%", mFwdSeat), _
                            IIf(txtGeneratedSeat.Tag = "", "%", txtGeneratedSeat.Tag), _
                            IIf(txtTransactionType.Tag = "", "%", txtTransactionType.Tag), _
                            IIf(txtAmount1.Text = "", "%", val(txtAmount1.Text)), _
                            IIf(txtAmount2.Text = "", Null, val(txtAmount2.Text)), _
                            IIf(txtPayOrderNo.Text = "", "%", txtPayOrderNo.Text) _
                            )
                Set Rec = objDB.ExecuteSP("spGetPaymentOrders", mArrIn, , False, mCnn, adCmdStoredProc)
        End If
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(DdMmmYy(Rec!dtPayOrderDate)), "", DdMmmYy(Rec!dtPayOrderDate))
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!Descriptions), "", Rec!Descriptions) ' Description
                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!SeatID), "", GetSeatName(Rec!SeatID))
                vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                If Rec!Status = 1 Then
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 7) = vbChecked
                Else
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 7) = vbUnchecked
                End If
                vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 9) = IIf(IsNull(Rec!intModuleID), "", Rec!intModuleID)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 10) = IIf(IsNull(Rec!numFwdSeatID), "", GetSeatName(Rec!numFwdSeatID))
                If Rec!Status = 1 Or Rec!Status = 3 Then
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 11) = vbChecked
                Else
                    vsGrid.Cell(flexcpChecked, vsGrid.Rows - 1, 11) = vbUnchecked
                End If
                Rec.MoveNext
            Wend
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub




    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdForwardedSeat_Click()
        txtForwardedSeat.Text = ""
        txtForwardedSeat.Tag = ""
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtForwardedSeat.Text = gbSearchStr
            txtForwardedSeat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdGeneratedSeat_Click()
        txtGeneratedSeat.Text = ""
        txtGeneratedSeat.Tag = ""
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtGeneratedSeat.Text = gbSearchStr
            txtGeneratedSeat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdNew_Click()
        frmPaymentOrder.Visible = True
        frmPaymentOrder.ZOrder (0)
        'FetchPaymentOrder
    End Sub

    Private Sub cmdsearch_Click()
        Call FetchPaymentOrder
    End Sub

    Private Sub cmdTransactionType_Click()
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        frmSearchTransactionType.ModeOfTransaction = 2
        frmSearchTransactionType.Show vbModal
        If gbSearchID > 0 Then
            txtTransactionType.Text = Trim(gbSearchStr)
            txtTransactionType.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdView_Click()
'        Dim mRowCnt As Integer
'        For mRowCnt = 1 To vsGrid.Rows
'
'        Next
        If vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = vbChecked Then
            Dim aryIn As Variant
            aryIn = Array(vsGrid.TextMatrix(vsGrid.Row, 1))
            frmViewVoucher.ArrayIn = aryIn
            frmViewVoucher.FormName = "frmViewPaymentOrder"
            frmViewVoucher.Show vbModal
        Else
            MsgBox "Only approved Payment Orders have Journals", vbInformation
        End If
    End Sub

    Private Sub cmdViewPOReport_Click()
        Dim mPayOrderNo As String
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        
        '*********************************************************************************************'
        '                       Procedure to view the Paymenr Order Report                            '
        '*********************************************************************************************'
        If vsGrid.Row < 1 Then
            MsgBox "Please select a Payment Order to view the Report"
            Exit Sub
        End If
        
        mPayOrderNo = Trim(vsGrid.TextMatrix(vsGrid.Row, 1))
        arInput = Array(mPayOrderNo)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptPayOrder.rpt"
        
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub dtpDateFrom_CloseUp()
        txtDateFrom.Text = dtpDateFrom.Value
        txtDateFrom.SetFocus
    End Sub
    
    Private Sub dtpDateTo_CloseUp()
        txtDateTo.Text = dtpDateTo.Value
        txtDateTo.SetFocus
    End Sub

    Private Sub Form_Activate()
        'Me.Left = 0
        'Me.Top = 0
        Me.WindowState = 2
        '-----------------------------------------------------'
        '                   Form Load Code                    '
        '-----------------------------------------------------'
        
        Me.WindowState = 2
        txtDateFrom.Text = Date - 31
        If CDate(txtDateFrom.Text) < gbStartingDate Then
            txtDateFrom.Text = gbStartingDate
        End If
        txtDateTo.Text = Date
        txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
        txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
''''        If gbUserTypeID = 3 Then
''''            cmdNew.Visible = True
''''        ElseIf gbUserTypeID = 2 Or gbUserTypeID = 4 Then
''''            cmdNew.Visible = False
''''        ElseIf gbUserTypeID = 1 Then
''''            cmdNew.Visible = True
''''        End If

        'Replacing UserTypeID with SeatGroupID'
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
            cmdNew.Visible = True
        Else
            cmdNew.Visible = False
        End If
        If gbLBPanchayat Then
            chkListToVerify.Visible = True
            If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                chkListToFwd.Visible = True
            End If
        Else
            chkListToVerify.Visible = False
            chkListToFwd.Visible = False
        End If
        Call FetchPaymentOrder
        '
        '-----------------------------------------------------'
    End Sub

    Private Sub Form_Load()
''''        Dim mSql As String
''''
''''        Call FormIntialize
''''        Me.WindowState = 2
''''        txtDateFrom.Text = Date - 31
''''        txtDateTo.Text = Date
''''        txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
''''        txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
''''        If gbUserTypeID = 3 Then
''''            cmdNew.Visible = True
''''        ElseIf gbUserTypeID = 2 Or gbUserTypeID = 4 Then
''''            cmdNew.Visible = False
''''        ElseIf gbUserTypeID = 1 Then
''''            cmdNew.Visible = True
''''        End If
''''        Call FetchPaymentOrder
        
         If frmViewPaymentOrder.ViewMode = 50 Then
            cmdNew.Visible = False
            cmdView.Visible = False
            cmdViewPOReport.Visible = False
            Call FormIntialize
        End If
        Call FormIntialize
    End Sub
    
    Private Sub Form_Resize()
        'Me.WindowState = 2
        If Me.WindowState <> 2 Then
            Me.Left = 0
            Me.Top = 0
            Me.Width = 15360
            Me.Height = 9450
        End If
    End Sub
    


    Private Sub txtAmount1_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtAmount2_KeyPress(KeyAscii As Integer)
        If txtAmount1.Text = "" Then
           KeyAscii = 0
        Else
           If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
               KeyAscii = 0
           End If
        End If
    End Sub

    Private Sub txtDateFrom_LostFocus()
        If txtDateFrom.Text <> "" Then
            txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
            If CDate(txtDateFrom.Text) < CDate(gbStartingDate) Then
                If CDate(txtDateFrom.Text) < CDate(DateAdd("yyyy", -1, gbStartingDate)) Then
                    txtDateFrom.Text = DateAdd("yyyy", -1, gbStartingDate)
                    txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
                End If
                txtDateTo.Text = DateAdd("yyyy", -1, gbEndingDate)
                txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
            End If
        End If
        Call PreviousYearMode
    End Sub

    Private Sub txtDateTo_LostFocus()
        If txtDateTo.Text <> "" Then
            txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
        

            If CDate(txtDateFrom.Text) > CDate(txtDateTo.Text) Then
                MsgBox "Please Enter a Date less than Or Equal to To Date", vbInformation
                txtDateFrom.Text = txtDateTo.Text
                txtDateFrom.SetFocus
                Exit Sub
            End If
            
        End If
        Call PreviousYearMode
    End Sub

    Private Sub txtForwardedSeat_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then 'Delete Key
            txtForwardedSeat.Text = ""
            txtForwardedSeat.Tag = ""
        Else
            txtForwardedSeat.Locked = True
        End If
    End Sub
    
    Private Sub txtGeneratedSeat_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then 'Delete Key
            txtGeneratedSeat.Text = ""
            txtGeneratedSeat.Tag = ""
        Else
            txtGeneratedSeat.Locked = True
        End If
    End Sub
    
    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = 46 Then 'Delete Key
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
        Else
            txtTransactionType.Locked = True
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            Dim mSQL As String
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDB As New clsDB
            
            If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
                MsgBox "Connction not Present ", vbCritical
                Exit Sub
            End If
            mSQL = "Select tnyStatus From faReverseEntry Where numDemandNo = " & vsGrid.TextMatrix(vsGrid.Row, 1)
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If IsNull(Rec!tnyStatus) = False Then
                    If Rec!tnyStatus = 0 Then        '  Requested for Cancellation
                        MsgBox "This Payment order is Requested for Cancellation", vbInformation
                        Exit Sub
                    ElseIf Rec!tnyStatus = 1 Then
                        MsgBox "This Payment order Cancellation Approved First Level", vbInformation
                        Exit Sub
                    ElseIf Rec!tnyStatus = 2 Then
                        MsgBox "This Payment order Cancellation Approved Final Level", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            Rec.Close
            mCnn.Close
            'gbSearchID = Val(vsGrid.TextMatrix(vsGrid.Row, 8))
            'gbSearchStr = Trim(vsGrid.TextMatrix(vsGrid.Row, 8))
            If mPreviousYearMode = 1 Then
                frmPaymentOrder.PendingTask = 3
            End If
            frmPaymentOrder.FillPayOrder (val(vsGrid.TextMatrix(vsGrid.Row, 8)))
            frmPaymentOrder.ListLoaded = True  ' To inform this From ( frmViewPaymentOrder ) is loaded
            frmPaymentOrder.Visible = True
            'FetchPaymentOrder
        End If
    End Sub
 Public Property Let ViewMode(mData As Integer)
        mViewMode = mData
    End Property
    
    Public Property Get ViewMode() As Integer
        ViewMode = mViewMode
    End Property

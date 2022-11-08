VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListofBillRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmListofBillRegister"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   13350
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   13290
      TabIndex        =   0
      Top             =   0
      Width           =   13350
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   13290
      TabIndex        =   2
      Top             =   6630
      Width           =   13350
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   750
         Left            =   1170
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   3540
      End
      Begin VB.CommandButton cmdSeat 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8220
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtForward2Seat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6465
         TabIndex        =   12
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmdPaymentOrder 
         Caption         =   "&Generate Payment Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9255
         TabIndex        =   7
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10845
         TabIndex        =   3
         Top             =   390
         Width           =   1410
      End
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   8700
         Top             =   900
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   16
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Forward to Seat :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4860
         TabIndex        =   15
         Top             =   240
         Width           =   1500
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   4860
      Left            =   -30
      TabIndex        =   1
      Top             =   1785
      Width           =   13320
      _cx             =   23495
      _cy             =   8572
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListofBillRegister.frx":0000
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
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   15
      TabIndex        =   4
      Top             =   1035
      Width           =   13290
      Begin VB.CheckBox chkPaid 
         Caption         =   "Paid"
         Height          =   270
         Left            =   11775
         TabIndex        =   17
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   1770
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   1770
      End
      Begin VB.ComboBox cmbRegisterType 
         Height          =   315
         Left            =   7095
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label Label3 
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Register Type :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   6
         Top             =   270
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmListofBillRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        Private mCheckDemandID As Variant '1 = to display specified demand only.
        Private mCheckRegID As Variant    'sets the RegID of specified record.
        Private mCheckBillID As Variant  ' to display the bill details
        Private mCheckPaymentID As Variant
        Dim intCheckMode As Integer
        Dim PO As uPaymentOrder
        Dim POC As uPaymentOrderChild
        Dim POAdd As uPaymentOrderAddress
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim Rec1 As New ADODB.Recordset
        Dim mSLNo As Integer
        Dim mLoop As Integer
        Dim vchPayOrderNo As String
        Dim mintID As Variant
        Dim mSQL As String
        Dim mAccountHeadCode As Variant
        Dim mAccountHeadID As Variant
        Dim mFunctionID As Integer
        Dim mFunctionaryID As Integer
        Dim mPeriodID As Variant
        Dim mintVoucherNo As Variant
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mMonthID As Integer
        Dim mArrIn As Variant
        Dim mID As Variant
        Dim mInstrumentTypeID As Variant
        Dim mInstrumentNo As Variant
        Dim mInstrumentDate As Variant
        Dim mPaymentOrderID  As Variant
        Dim mIDAccountHead As Variant
        Dim mLoops   As Integer
       
     Private Sub cmbRegisterType_Click()
        Call fillGrid
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdPaymentOrder_Click()
        Dim mSQL        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
    
        'If vsGrid.TextMatrix(vsGrid.Row, 6) = "" Or vsGrid.TextMatrix(vsGrid.Row, 7) = "" Or vsGrid.TextMatrix(vsGrid.Row, 9) = "" Then
        If vsGrid.TextMatrix(vsGrid.Row, 6) = "" Or vsGrid.TextMatrix(vsGrid.Row, 7) = "" Then
            MsgBox "Enter the Bill details", vbInformation
            Exit Sub
        End If
'''        If objDb.SetConnection(mCnn) Then
'''         If VSGrid.TextMatrix(VSGrid.Row - 1, 11) <> "Selection" Then
'''            mSQL = " SELECT * From faBillRegisters Where intRegID = " & VSGrid.TextMatrix(VSGrid.Row, 1) & " AND tnyStatus <> 3 "
'''            mSQL = mSQL + " AND (dtDemandDueDate BETWEEN '1/Apr/2010' AND '" & Format((VSGrid.TextMatrix(VSGrid.Row - 1, 3)), "dd/mmm/yyyy") & "') "
'''
'''            Rec.Open mSQL, mCnn
'''            If Not (Rec.EOF And Rec.BOF) Then
'''                    MsgBox "Generate Payments for the previous months", vbInformation
'''                    VSGrid.TextMatrix(VSGrid.Row, 11) = vbUnchecked
'''                    Exit Sub
'''            End If
'''            Rec.Close
'''          End If
'''        End If
        If intCheckMode = 1 Then
            Call SavePaymentOrder
        'ElseIf intCheckMode = 2 The.n
            'Call SavePastData
        End If
        mCheckDemandID = 1
        Call fillGrid
    End Sub

    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID <> -1 Then
            txtForward2Seat.Text = gbSearchStr
            txtForward2Seat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        XPC.InitSubClassing
        Call PopulateList(cmbRegisterType, "Select vchRegType, intRegTypeID From faRegisterTypes Order By vchRegType", , True, True, True)
        txtFromDate.Text = DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbEndingDate)
       
       intCheckMode = 1
       cmdPaymentOrder.Enabled = False
       mIDAccountHead = 0
    End Sub
    Private Sub Form_Activate()
        
        Call fillGrid
     End Sub
    
    Private Sub fillGrid()
    Dim mSQL        As String
    Dim objDb       As New clsDB
    Dim mCnn        As New ADODB.Connection
    Dim Rec         As New ADODB.Recordset
    Dim mRowCnt     As Integer
    Dim mRecCnt     As Integer
    Dim mTypeID     As Integer
    Dim mLoop       As Integer
    
     If objDb.SetConnection(mCnn) Then
     
            mSQL = "SELECT faBillRegisters.*, faBillRegisters.*, faBillRegisters.intID,faBillRegisters.intRegID,faBillRegisters.vchBillNo,faBillRegisters.dtBillDate,faBillRegisters.dtBillDueDate,faBillRegisters.numAmount,faBillRegisters.numPaidAmount,faRegisterOfBills.vchRegName, faBillRegisters.dtDemandDueDate, faBillRegisters.intYearID, "
            mSQL = mSQL + " faBillRegisters.intMonthID,faBillRegisters.tnyStatus,faBillRegisters.vchRemarks,faBillRegisters.numPaidAmount,faBillRegisters.intPaymentOrderNo FROM faBillRegisters INNER JOIN "
            mSQL = mSQL + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID  "
            If mCheckDemandID = 1 Then
                mSQL = mSQL + " where faBillRegisters.intRegID= " & mCheckRegID & " and (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ')"
            End If
            If mCheckDemandID = 0 Then
                If cmbRegisterType.ListIndex < 1 Then
                    mSQL = mSQL + " WHERE  (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ') "
                Else
                    mTypeID = cmbRegisterType.ItemData(cmbRegisterType.ListIndex)
                    mSQL = mSQL + " WHERE  (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ')"
                    mSQL = mSQL + " and faRegisterOfBills.intRegTypeID = " & mTypeID & "  "
                End If
            End If
            If mCheckBillID = 1 Then
               ' msQl = msQl + " WHERE (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ')"
                mSQL = mSQL + " and faBillRegisters.tnyStatus = 2  "
            End If
            If mCheckPaymentID = 1 Then
                mSQL = mSQL + " and faBillRegisters.tnyStatus = 3  "
            End If
            If chkPaid.value = vbChecked Then
                mSQL = mSQL + " and faBillRegisters.tnyStatus = 3 "
            End If
            mSQL = mSQL + "ORDER BY faBillRegisters.intMonthID"
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            vsGrid.Rows = 1
            mRowCnt = 1
            mRecCnt = 1
            vsGrid.Clear 1, 1
            While Not (Rec.EOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = mRecCnt
                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
                vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
                vsGrid.TextMatrix(mRowCnt, 3) = DdMmmYy(IIf(IsNull(Rec!dtDemandDueDate), "", Rec!dtDemandDueDate))
                vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                If Rec!intMonthID = 1 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "January"
                ElseIf Rec!intMonthID = 2 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "February"
                ElseIf Rec!intMonthID = 3 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "March"
                ElseIf Rec!intMonthID = 4 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "April"
                ElseIf Rec!intMonthID = 5 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "May"
                ElseIf Rec!intMonthID = 6 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "June"
                ElseIf Rec!intMonthID = 7 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "July"
                ElseIf Rec!intMonthID = 8 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "August"
                ElseIf Rec!intMonthID = 9 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "September"
                ElseIf Rec!intMonthID = 10 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "October"
                ElseIf Rec!intMonthID = 11 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "November"
                ElseIf Rec!intMonthID = 12 Then
                    vsGrid.TextMatrix(mRowCnt, 5) = "December"
                End If

                vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!dtBillDate), "", (Rec!dtBillDate))
                vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!dtBillDueDate), "", Rec!dtBillDueDate)
                vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
                vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
                vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!dtPaymentOrderDate), "", (Rec!dtPaymentOrderDate))
                If Rec!tnyStatus = 3 Then
                    vsGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked

                End If

                For mLoop = 0 To vsGrid.Cols - 1
                    If vsGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked Then
                        vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HD2AE9E
                    End If
                Next mLoop

                vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intID), "", Rec!intID)
                vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
                vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                vsGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                Rec.MoveNext

                mRowCnt = mRowCnt + 1
                mRecCnt = mRecCnt + 1
            Wend
            mCheckDemandID = 0
            mCheckBillID = 0
            mCheckPaymentID = 0
            Rec.Close
     
'----------------------------------------------------------------------------------------------------------------------------------
''''''''''''         If mCheckDemandID = 1 Then
''''''''''''            msQl = "SELECT faBillRegisters.intID,faBillRegisters.intRegID,faBillRegisters.vchBillNo,faBillRegisters.dtBillDate,faBillRegisters.dtBillDueDate,faBillRegisters.numAmount,faBillRegisters.numPaidAmount,faRegisterOfBills.vchRegName, faBillRegisters.dtDemandDueDate, faBillRegisters.intYearID, "
''''''''''''            msQl = msQl + "faBillRegisters.intMonthID,faBillRegisters.tnyStatus,faBillRegisters.vchRemarks,faBillRegisters.numPaidAmount,faBillRegisters.intPaymentOrderNo FROM faBillRegisters INNER JOIN "
''''''''''''            msQl = msQl + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID  "
''''''''''''            msQl = msQl + " where faBillRegisters.intRegID= " & mCheckRegID & ""
''''''''''''            Rec.CursorLocation = adUseClient
''''''''''''            Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''''''''''''            VSGrid.Rows = 1
''''''''''''            mRowCnt = 1
''''''''''''            mRecCnt = 1
''''''''''''            VSGrid.Clear 1, 1
''''''''''''            While Not (Rec.EOF)
''''''''''''                VSGrid.Rows = VSGrid.Rows + 1
''''''''''''                VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
''''''''''''                VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtDemandDueDate), "", Rec!dtDemandDueDate)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
''''''''''''                If Rec!intMonthID = 1 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "January"
''''''''''''                ElseIf Rec!intMonthID = 2 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "February"
''''''''''''                ElseIf Rec!intMonthID = 3 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "March"
''''''''''''                ElseIf Rec!intMonthID = 4 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "April"
''''''''''''                ElseIf Rec!intMonthID = 5 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "May"
''''''''''''                ElseIf Rec!intMonthID = 6 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "June"
''''''''''''                ElseIf Rec!intMonthID = 7 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "July"
''''''''''''                ElseIf Rec!intMonthID = 8 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "August"
''''''''''''                ElseIf Rec!intMonthID = 9 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "September"
''''''''''''                ElseIf Rec!intMonthID = 10 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "October"
''''''''''''                ElseIf Rec!intMonthID = 11 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "November"
''''''''''''                ElseIf Rec!intMonthID = 12 Then
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 5) = "December"
''''''''''''                End If
''''''''''''
''''''''''''                VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!dtBillDate), "", Rec!dtBillDate)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!dtBillDueDate), "", Rec!dtBillDueDate)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
''''''''''''                'vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
''''''''''''               ' vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!dtPaymentOrderDate), "", Rec!dtPaymentOrderDate)
''''''''''''                If Rec!tnyStatus = 3 Then
''''''''''''                    VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked
''''''''''''
''''''''''''                End If
''''''''''''
''''''''''''                For mLoop = 0 To VSGrid.Cols - 1
''''''''''''                    If VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked Then
''''''''''''                        VSGrid.Cell(flexcpBackColor, VSGrid.Rows - 1, mLoop) = &HD2AE9E
''''''''''''                    End If
''''''''''''                Next mLoop
''''''''''''
''''''''''''                VSGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intID), "", Rec!intID)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
''''''''''''                VSGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
''''''''''''                Rec.MoveNext
''''''''''''
''''''''''''                mRowCnt = mRowCnt + 1
''''''''''''                mRecCnt = mRecCnt + 1
''''''''''''            Wend
''''''''''''            mCheckDemandID = 0
''''''''''''            Rec.Close
''''''''''''
''''''''''''        ElseIf mCheckDemandID = 0 Then
''''''''''''            If cmbRegisterType.ListIndex < 1 Then
''''''''''''                msQl = "SELECT faBillRegisters.intID,faBillRegisters.intRegID,faBillRegisters.vchBillNo,faBillRegisters.dtBillDate,faBillRegisters.dtBillDueDate,faBillRegisters.numAmount,faBillRegisters.numPaidAmount,faRegisterOfBills.vchRegName, faBillRegisters.dtDemandDueDate, faBillRegisters.intYearID, "
''''''''''''                msQl = msQl + "faBillRegisters.intMonthID,faBillRegisters.tnyStatus,faBillRegisters.vchRemarks,faBillRegisters.numPaidAmount,faBillRegisters.intPaymentOrderNo FROM faBillRegisters INNER JOIN "
''''''''''''                msQl = msQl + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID WHERE  (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ')"
''''''''''''                Rec.CursorLocation = adUseClient
''''''''''''                Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''''''''''''                mRowCnt = 1
''''''''''''                mRecCnt = 1
''''''''''''                VSGrid.Rows = 1
''''''''''''
''''''''''''                VSGrid.Clear 1, 1
''''''''''''
''''''''''''                While Not (Rec.EOF)
''''''''''''                    VSGrid.Rows = VSGrid.Rows + 1
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtDemandDueDate), "", CheckDateInMMM(Rec!dtDemandDueDate))
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
''''''''''''                    If Rec!intMonthID = 1 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "January"
''''''''''''                    ElseIf Rec!intMonthID = 2 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "February"
''''''''''''                    ElseIf Rec!intMonthID = 3 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "March"
''''''''''''                    ElseIf Rec!intMonthID = 4 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "April"
''''''''''''                    ElseIf Rec!intMonthID = 5 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "May"
''''''''''''                    ElseIf Rec!intMonthID = 6 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "June"
''''''''''''                    ElseIf Rec!intMonthID = 7 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "July"
''''''''''''                    ElseIf Rec!intMonthID = 8 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "August"
''''''''''''                    ElseIf Rec!intMonthID = 9 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "September"
''''''''''''                    ElseIf Rec!intMonthID = 10 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "October"
''''''''''''                    ElseIf Rec!intMonthID = 11 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "November"
''''''''''''                    ElseIf Rec!intMonthID = 12 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "December"
''''''''''''                    End If
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
''''''''''''                    'VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!dtBillDate), "", CheckDateInMMM(Rec!dtBillDate))
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!dtBillDueDate), "", Rec!dtBillDueDate)
''''''''''''                    'VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
''''''''''''
''''''''''''                    If Rec!tnyStatus = 3 Then
''''''''''''                        VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked
''''''''''''
''''''''''''                    End If
''''''''''''                     For mLoop = 0 To VSGrid.Cols - 1
''''''''''''                    If VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked Then
''''''''''''                        VSGrid.Cell(flexcpBackColor, VSGrid.Rows - 1, mLoop) = &HD2AE9E
''''''''''''                    End If
''''''''''''                Next mLoop
''''''''''''
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intID), "", Rec!intID)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
''''''''''''                    Rec.MoveNext
''''''''''''
''''''''''''                    mRowCnt = mRowCnt + 1
''''''''''''                    mRecCnt = mRecCnt + 1
''''''''''''                Wend
''''''''''''                Rec.Close
''''''''''''            Else
''''''''''''                mTypeID = cmbRegisterType.ItemData(cmbRegisterType.ListIndex)
''''''''''''
'''''''''''''                mSql = "SELECT faBillRegisters.intID,faBillRegisters.intRegID, faRegisterOfBills.vchRegName, faBillRegisters.dtDemandDueDate, faBillRegisters.intYearID, "
'''''''''''''                mSql = mSql + "faBillRegisters.intMonthID,faBillRegisters.tnyStatus FROM faBillRegisters INNER JOIN "
'''''''''''''                mSql = mSql + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID WHERE faRegisterOfBills.intRegTypeID = " & mTypeID & " "
''''''''''''
''''''''''''
''''''''''''                msQl = "SELECT faBillRegisters.intID,faBillRegisters.intRegID,faBillRegisters.vchBillNo,faBillRegisters.dtBillDate,faBillRegisters.dtBillDueDate,faBillRegisters.numAmount,faBillRegisters.numPaidAmount,faRegisterOfBills.vchRegName, faBillRegisters.dtDemandDueDate, faBillRegisters.intYearID, "
''''''''''''                msQl = msQl + "faBillRegisters.intMonthID,faBillRegisters.tnyStatus,faBillRegisters.vchRemarks,faBillRegisters.numPaidAmount,faBillRegisters.intPaymentOrderNo FROM faBillRegisters INNER JOIN "
''''''''''''                msQl = msQl + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID WHERE  (faBillRegisters.dtDemandDueDate BETWEEN '" & txtFromDate.Text & "' AND '" & txtToDate.Text & " ')"
''''''''''''                msQl = msQl + " and faRegisterOfBills.intRegTypeID = " & mTypeID & "  "
''''''''''''
''''''''''''
''''''''''''                Rec.CursorLocation = adUseClient
''''''''''''                Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''''''''''''                mRowCnt = 1
''''''''''''                mRecCnt = 1
''''''''''''                VSGrid.Rows = 1
''''''''''''
''''''''''''                VSGrid.Clear 1, 1
''''''''''''
''''''''''''                While Not (Rec.EOF)
''''''''''''                    VSGrid.Rows = VSGrid.Rows + 1
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtDemandDueDate), "", Rec!dtDemandDueDate)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
''''''''''''                    If Rec!intMonthID = 1 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "January"
''''''''''''                    ElseIf Rec!intMonthID = 2 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "February"
''''''''''''                    ElseIf Rec!intMonthID = 3 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "March"
''''''''''''                    ElseIf Rec!intMonthID = 4 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "April"
''''''''''''                    ElseIf Rec!intMonthID = 5 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "May"
''''''''''''                    ElseIf Rec!intMonthID = 6 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "June"
''''''''''''                    ElseIf Rec!intMonthID = 7 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "July"
''''''''''''                    ElseIf Rec!intMonthID = 8 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "August"
''''''''''''                    ElseIf Rec!intMonthID = 9 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "September"
''''''''''''                    ElseIf Rec!intMonthID = 10 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "October"
''''''''''''                    ElseIf Rec!intMonthID = 11 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "November"
''''''''''''                    ElseIf Rec!intMonthID = 12 Then
''''''''''''                        VSGrid.TextMatrix(mRowCnt, 5) = "December"
''''''''''''                    End If
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!dtBillDate), "", Rec!dtBillDate)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!dtBillDueDate), "", Rec!dtBillDueDate)
''''''''''''                    'VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
''''''''''''                    If Rec!tnyStatus = 3 Then
''''''''''''                        VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked
''''''''''''
''''''''''''                    End If
''''''''''''
''''''''''''                    For mLoop = 0 To VSGrid.Cols - 1
''''''''''''                        If VSGrid.Cell(flexcpChecked, mRowCnt, 11) = vbChecked Then
''''''''''''                            VSGrid.Cell(flexcpBackColor, VSGrid.Rows - 1, mLoop) = &HD2AE9E
''''''''''''                        End If
''''''''''''                    Next mLoop
''''''''''''
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intID), "", Rec!intID)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
''''''''''''                    VSGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
''''''''''''                    Rec.MoveNext
''''''''''''
''''''''''''                    mRowCnt = mRowCnt + 1
''''''''''''                    mRecCnt = mRecCnt + 1
''''''''''''                Wend
''''''''''''                Rec.Close
''''''''''''            End If
''''''''''''        End If 'demand

'-------------------------------------------------------------------------------------------------------------------------------------

    End If 'connection
    End Sub
    Public Property Let CheckDemandID(mData As Variant)
        mCheckDemandID = mData
    End Property
    Public Property Get CheckDemandID() As Variant
        CheckDemandID = mCheckDemandID
    End Property
     Public Property Let CheckBillID(mData As Variant)
        mCheckBillID = mData
    End Property
    Public Property Get CheckBillID() As Variant
        CheckBillID = mCheckBillID
    End Property
     Public Property Let CheckPaymentID(mData As Variant)
        mCheckPaymentID = mData
    End Property
    Public Property Get CheckPaymentID() As Variant
        CheckPaymentID = mCheckPaymentID
    End Property
    Public Property Let CheckRegID(mData As Variant)
        mCheckRegID = mData
    End Property

    Public Property Get CheckRegID() As Variant
        CheckRegID = mCheckRegID
    End Property
    Private Sub chkPaid_Click()
        Call fillGrid
    End Sub
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
    End Sub
    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        If txtFromDate.Text <> "" Then
            Call fillGrid
        End If
        If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
            MsgBox "Please Enter a Date Less than Or equal to ToDate", vbInformation
            txtFromDate.Text = ""
            txtFromDate.SetFocus
            Exit Sub
        End If
    End Sub

'    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'        If VSGrid.Cell(flexcpChecked, VSGrid.Row, 11) = 1 Then
'            Cancel = True
'        End If
'    End Sub

Private Sub VSGrid_Click()
If vsGrid.Row > 0 Then
    If vsGrid.Col = 11 Then
        If vsGrid.TextMatrix(vsGrid.Row, 14) = 3 Then
           vsGrid.Editable = flexEDNone
        Else
            vsGrid.Editable = flexEDKbdMouse
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 11) = 1 Then
                cmdPaymentOrder.Enabled = True
            Else
                cmdPaymentOrder.Enabled = False
            End If
        End If
    End If
Else
    Exit Sub
End If
End Sub

    Private Sub VSGrid_DblClick()
    Dim mCnn As New ADODB.Connection
    Dim Rec  As New ADODB.Recordset
    Dim objDb As New clsDB
    Dim mSQL As String

        If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then Exit Sub

' '----------IF BILL DETAILS ALREADY GENERATED--------------------------------------------------------

        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        If VSGrid.TextMatrix(VSGrid.Row, 13) = 2 Then
'            msQl = " SELECT  vchBillNo, dtBillDate, numAmount, numPaidAmount, intPaymentOrderNo, intPaymentVoucherNo,  "
'            msQl = msQl + " vchRemarks , tnyStatus, intID FROM  faBillRegisters WHERE ( tnyStatus = 2)"
'            msQl = msQl + "    AND intID = " & VSGrid.TextMatrix(VSGrid.Row, 11) & " "
'            Rec.Open msQl, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'
'                frmRecurringBillRegisters.cmdVerifyBill.Enabled = False
'                frmRecurringBillRegisters.txtPaymentOrderNo.Text = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
'                frmRecurringBillRegisters.txtPaymentVoucherNo.Text = IIf(IsNull(Rec!intPaymentVoucherNo), "", Rec!intPaymentVoucherNo)
'                frmRecurringBillRegisters.txtBillNo.Text = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
'                frmRecurringBillRegisters.txtBillDate.Text = IIf(IsNull(Rec!dtBillDate), "", Rec!dtBillDate)
'                frmRecurringBillRegisters.txtAmount.Text = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
'                frmRecurringBillRegisters.txtPaidAmt.Tag = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
'
'                frmRecurringBillRegisters.txtExtraAmt = val(frmRecurringBillRegisters.txtPaidAmt.Tag) - val(frmRecurringBillRegisters.txtAmount.Text)
'
'                frmRecurringBillRegisters.txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
'                frmRecurringBillRegisters.Frame1.Enabled = False
'                frmRecurringBillRegisters.Frame2.Enabled = False
'                frmRecurringBillRegisters.txtRegID = VSGrid.TextMatrix(VSGrid.Row, 1)
'                frmRecurringBillRegisters.txtRegID.Tag = VSGrid.TextMatrix(VSGrid.Row, 11)
'                frmRecurringBillRegisters.txtRegName = VSGrid.TextMatrix(VSGrid.Row, 2)
'                frmRecurringBillRegisters.txtDemandDueDate = VSGrid.TextMatrix(VSGrid.Row, 3)
'                frmRecurringBillRegisters.txtYear = VSGrid.TextMatrix(VSGrid.Row, 4)
'                frmRecurringBillRegisters.txtMonth = VSGrid.TextMatrix(VSGrid.Row, 5)
'                frmRecurringBillRegisters.txtBilDueDate = VSGrid.TextMatrix(VSGrid.Row, 3)
'                frmRecurringBillRegisters.Show vbModal
'                Exit Sub
'
'            End If
'            Rec.Close
'        End If
        
        'If VSGrid.TextMatrix(VSGrid.Row, 13) = 3 Then
            mSQL = " SELECT  vchBillNo, dtBillDate, numAmount, numPaidAmount, intPaymentOrderNo, intPaymentVoucherNo,  "
            mSQL = mSQL + " vchRemarks , tnyStatus, intID FROM  faBillRegisters WHERE ((tnyStatus=2)or "
            mSQL = mSQL + "   (tnyStatus = 3))  AND intID = " & vsGrid.TextMatrix(vsGrid.Row, 13) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                
                frmRecurringBillRegisters.cmdVerifyBill.Enabled = False
                frmRecurringBillRegisters.txtPaymentOrderNo.Text = IIf(IsNull(Rec!intPaymentOrderNo), "", Rec!intPaymentOrderNo)
                frmRecurringBillRegisters.txtPaymentVoucherNo.Text = IIf(IsNull(Rec!intPaymentVoucherNo), "", Rec!intPaymentVoucherNo)
                frmRecurringBillRegisters.txtBillNo.Text = IIf(IsNull(Rec!vchBillNo), "", Rec!vchBillNo)
                frmRecurringBillRegisters.txtBillDate.Text = IIf(IsNull(Rec!dtBillDate), "", Rec!dtBillDate)
                frmRecurringBillRegisters.txtAmount.Text = IIf(IsNull(Rec!numAmount), "", Rec!numAmount)
                frmRecurringBillRegisters.txtPaidAmt.Text = IIf(IsNull(Rec!numPaidAmount), "", Rec!numPaidAmount)
                
                frmRecurringBillRegisters.txtExtraAmt = val(frmRecurringBillRegisters.txtPaidAmt.Text) - val(frmRecurringBillRegisters.txtAmount.Text)
                
                frmRecurringBillRegisters.txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                frmRecurringBillRegisters.Frame1.Enabled = False
                frmRecurringBillRegisters.Frame2.Enabled = False
                frmRecurringBillRegisters.txtRegID = vsGrid.TextMatrix(vsGrid.Row, 1)
            
                frmRecurringBillRegisters.txtRegID.Tag = vsGrid.TextMatrix(vsGrid.Row, 13)
                frmRecurringBillRegisters.txtRegName = vsGrid.TextMatrix(vsGrid.Row, 2)
                frmRecurringBillRegisters.txtDemandDueDate = vsGrid.TextMatrix(vsGrid.Row, 3)
                frmRecurringBillRegisters.txtYear = vsGrid.TextMatrix(vsGrid.Row, 4)
                frmRecurringBillRegisters.txtMonth = vsGrid.TextMatrix(vsGrid.Row, 5)
                frmRecurringBillRegisters.txtBilDueDate = vsGrid.TextMatrix(vsGrid.Row, 3)
                frmRecurringBillRegisters.Show vbModal
                Exit Sub
    
            End If
            Rec.Close
        'End If
    
' '------------------------PAYMENT FOR PREVIOUS DATE CHECKING------------------------
'
'        If mID$(VSGrid.TextMatrix(VSGrid.Row, 3), 4, 2) < mID$(gbTransactionDate, 4, 2) Then
'         If mID$(VSGrid.TextMatrix(VSGrid.Row, 3), 7, 4) <= gbFinancialYearID Then
'            MsgBox "Payment cannot be generated for previous date", vbInformation
'            If MsgBox("Do you wish to enter the past Data", vbYesNo) = vbYes Then
'                intCheckMode = 2
'            Else
'                Exit Sub
'            End If
'          End If
'        End If
' '---------------------------------------------------------------------------------
        frmRecurringBillRegisters.txtRegID = vsGrid.TextMatrix(vsGrid.Row, 1)
       
        frmRecurringBillRegisters.txtRegID.Tag = vsGrid.TextMatrix(vsGrid.Row, 13)
        frmRecurringBillRegisters.txtRegName = vsGrid.TextMatrix(vsGrid.Row, 2)
        frmRecurringBillRegisters.txtDemandDueDate = vsGrid.TextMatrix(vsGrid.Row, 3)
        frmRecurringBillRegisters.txtYear = vsGrid.TextMatrix(vsGrid.Row, 4)
        frmRecurringBillRegisters.txtMonth = vsGrid.TextMatrix(vsGrid.Row, 5)
       
        frmRecurringBillRegisters.txtPaymentOrderNo.Enabled = False
        frmRecurringBillRegisters.txtPaymentVoucherNo.Enabled = False
        frmRecurringBillRegisters.Show vbModal

    End Sub
    Private Sub SavePaymentOrder()
    
        '-----------------------------------------------------------------------------------------'
        '                           To generate Voucher Number                                    '
        '-----------------------------------------------------------------------------------------'
        
        objDb.SetConnection mCnn
        mSQL = "SELECT faRegisterOfBills.intFunctionaryID, faRegisterOfBills.intFunctionID, faAccountHeads.vchAccountHeadCode,faAccountHeads.intAccountHeadID,faBillRegisters.tnyPreriodID"
        mSQL = mSQL + " FROM faBillRegisters INNER JOIN "
        mSQL = mSQL + " faRegisterOfBills ON faBillRegisters.intRegID = faRegisterOfBills.intRegID INNER JOIN"
        mSQL = mSQL + " faAccountHeads ON faRegisterOfBills.intExpenditureHeadID = faAccountHeads.intAccountHeadID"
        mSQL = mSQL + " Where faBillRegisters.intID = " & vsGrid.TextMatrix(vsGrid.Row, 13) & " "
        
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
       
        If Not (Rec.EOF And Rec.BOF) Then
            mAccountHeadCode = Rec!vchAccountHeadCode
            mAccountHeadID = Rec!intAccountHeadID
            mFunctionID = Rec!intFunctionID
            mFunctionaryID = Rec!intFunctionaryID
            mPeriodID = IIf(IsNull(Rec!tnyPreriodID), Null, Rec!tnyPreriodID)
        End If
        arrInput = Array(mAccountHeadCode, _
                                    20, _
                                    gbFinancialYearID, _
                                    mintVoucherNo _
                         )
        Rec.Close
      
        mSQL = "Declare @intVoucherNo Numeric " & vbNewLine
        mSQL = mSQL + " Exec spGetVoucherNo Null,20," & gbFinancialYearID & ",@intVoucherNo output" & vbNewLine
        mSQL = mSQL + " Select @intVoucherNo [numVoucherNo]"
        Rec.Open mSQL, mCnn
        mintVoucherNo = Rec!numVoucherNo
        
        'Rec.Close
        '-----------------------------------------------------------------------------------------'
        '                           To generate Payment Order Number                              '
        '-----------------------------------------------------------------------------------------'

        With PO
            .intPayOrderID = IIf(frmRecurringBillRegisters.txtPaymentOrderNo.Tag = "", Null, frmRecurringBillRegisters.txtPaymentOrderNo.Tag)
            .vchPayOrderNo = IIf(frmRecurringBillRegisters.txtPaymentVoucherNo.Text = "", Null, frmRecurringBillRegisters.txtPaymentVoucherNo.Text)
            .dtPayOrderDate = Trim(vsGrid.TextMatrix(vsGrid.Row, 7))
            .dtDueDate = Trim(vsGrid.TextMatrix(vsGrid.Row, 8))
            .intFunctionaryID = mFunctionaryID
            .intFunctionID = mFunctionID
            .intTransactionTypeID = Null
            .vchBillNo = Trim(vsGrid.TextMatrix(vsGrid.Row, 6))
            .numBillAmount = Trim(vsGrid.TextMatrix(vsGrid.Row, 12))
            .dtBillDate = Trim(vsGrid.TextMatrix(vsGrid.Row, 7))
            .intInstrumentTypeID = Null
            .intCashOrBankHeadID = Null
            .vchDescription = Trim(txtRemarks.Text)
            .vchTitle = Null
            .intSubLedgerTypeID = Null
            .intPayToSubLedgerID = Null
            .intSubsidiaryCashBookID = Null
            .intImplementingOfficerID = Null
            .numProjectNo = Null
            .intStockRegisterID = Null
            .vchStockRefNo = Null
            .intAssetTypeID = Null
            .intAssetID = Null
            .numFwdSeatID = txtForward2Seat.Tag
            .intLocalBodyID = gbLocalBodyID
            .intZonalID = gbLocationID
            .intFinancialYearID = gbFinancialYearID
            .numUserID = gbUserID
            .numSeatID = gbSeatID
            .numApprovingOfficerID = Null
            .numApprovingSeatID = Null
            .dtApprovingDate = Null
            .intSourceOfFundID = Null
            .intAllotmentID = Null
            .intAgreementID = Null
            .tnyCategoryID = Null
            .tnySectorID = Null
            .tnyIsFinalBill = Null
            .intVoucherID = Null
            .intVoucherNo = Null 'mintVoucherNo
            .dtVoucherDate = Null
            .tnyStatus = 0
            .intKeyID = Null 'Section ID stores from Pay Bill -Sthapana for Pay&Allowance
            .numKeyID = Null
            .dtKeyDate = Null
            .tnyCancelled = 0
            .intAppID = 115
            .intModuleID = Null

            arrInput = Array(.intPayOrderID, .vchPayOrderNo, .dtPayOrderDate, .dtDueDate, .intFunctionaryID, _
            .intFunctionID, .intTransactionTypeID, .vchBillNo, .numBillAmount, _
            .dtBillDate, .intInstrumentTypeID, .intCashOrBankHeadID, .vchDescription, _
            .vchTitle, .intSubLedgerTypeID, .intPayToSubLedgerID, .intSubsidiaryCashBookID, _
            .intImplementingOfficerID, .numProjectNo, .intStockRegisterID, .vchStockRefNo, _
            .intAssetTypeID, .intAssetID, .numFwdSeatID, .intLocalBodyID, _
            .intZonalID, .intFinancialYearID, .numUserID, .numSeatID, _
            .numApprovingOfficerID, .numApprovingSeatID, .dtApprovingDate, .intVoucherID, .intVoucherNo, .dtVoucherDate, _
            .tnyStatus, .intKeyID, .numKeyID, .dtKeyDate, .tnyCancelled, .intAppID, .intModuleID, .intSourceOfFundID, _
            .intAllotmentID, .intAgreementID, .tnyCategoryID, .tnySectorID, _
            .tnyIsFinalBill)

            objDb.ExecuteSP "spSavePayOrder", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        If IsNumeric(arrOutPut(0, 0)) Then
            mPaymentOrderID = arrOutPut(0, 0)
            vchPayOrderNo = arrOutPut(1, 0)
        End If
        Rec.Close
        '---------------------------------------------------------------------------------------------------------

        mSQL = "Delete From faPayOrderChild Where intPayOrderID = " & mPaymentOrderID
        mCnn.Execute mSQL
        With frmListofBillRegister.vsGrid
        For mLoop = 1 To .Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 11) = 1 And vsGrid.TextMatrix(mLoop, 14) = 2 Then
                   mSQL = "SELECT faRegisterOfBills.intExpenditureHeadID, faBillRegisters.tnyStatus, faBillRegisters.intID"
                   mSQL = mSQL + " FROM faRegisterOfBills INNER JOIN faBillRegisters ON faRegisterOfBills.intRegID = faBillRegisters.intRegID "
                   mSQL = mSQL + " WHERE  faBillRegisters.intID = " & vsGrid.TextMatrix(vsGrid.Row, 13) & " "
                   Rec.Open mSQL, mCnn
                   vsGrid.TextMatrix(vsGrid.Row, 13) = Rec!intID
                    Rec.Close
                    mSLNo = mSLNo + 1
                    With POC
                        .intPayOrderID = mPaymentOrderID
                        .intSlNo = mSLNo
                        .intAccountHeadID = mAccountHeadID
                        .vchAccountHeadCode = mAccountHeadCode
                        .numAmount = Trim(vsGrid.TextMatrix(vsGrid.Row, 12))
                        .tnyCategoryFlag = 1
                        .tnyDebitOrCreditFlag = 1
                        .vchDescription = Null
            
                        arrInput = Array(.intPayOrderID, _
                        .intSlNo, _
                        .intAccountHeadID, _
                        .vchAccountHeadCode, _
                        .numAmount, _
                        .tnyCategoryFlag, _
                        .tnyDebitOrCreditFlag, _
                        .vchDescription)
            
                        objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
                    End With
            
            '---------------------------------------------------------------------------------------------------'
            '                                    TO SAVE IN faBillRegister                                      '                         '
            '---------------------------------------------------------------------------------------------------'
                    
                    If Trim(vsGrid.TextMatrix(mLoop, 5)) = "January" Then
                        mMonthID = 1
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "February" Then
                        mMonthID = 2
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "March" Then
                        mMonthID = 3
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "April" Then
                        mMonthID = 4
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "May" Then
                        mMonthID = 5
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "June" Then
                        mMonthID = 6
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "July" Then
                        mMonthID = 7
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "August" Then
                        mMonthID = 8
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "September" Then
                        mMonthID = 9
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "October" Then
                        mMonthID = 10
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "November" Then
                        mMonthID = 11
                    ElseIf Trim(vsGrid.TextMatrix(mLoop, 5)) = "December" Then
                        mMonthID = 12
                    End If

               
                     mID = IIf(Trim(vsGrid.TextMatrix(mLoop, 13)) = "", -1, Trim(vsGrid.TextMatrix(mLoop, 13))) ' 13-mintVoucherNo
                     mArrIn = Array(mID, Trim(vsGrid.TextMatrix(mLoop, 3)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 1)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 4)), _
                                         mMonthID, _
                                         mPeriodID, _
                                         Trim(vsGrid.TextMatrix(mLoop, 6)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 7)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 8)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 12)), _
                                         Trim(vsGrid.TextMatrix(mLoop, 12)), _
                                         vchPayOrderNo, _
                                         mintVoucherNo, _
                                         Null, _
                                         Null, _
                                         Null, _
                                         Trim(vsGrid.TextMatrix(mLoop, 15)), _
                                         3, _
                                         gbTransactionDate _
                                         )

                     objDb.ExecuteSP "spSaveBillRegisters", mArrIn, , , mCnn, adCmdStoredProc
                     vsGrid.TextMatrix(mLoop, 14) = 3
                  End If
                  'mLoop = mLoop + 1
           Next mLoop
        End With

'-------------------------------------------------------------------------------------------------------------------'
'                                           PAYORDER ADDRESS                                                        '
'-------------------------------------------------------------------------------------------------------------------'

    With POAdd

            'ObjSubLed.SetSubLedgerDetails (val(txtName.Tag))
            .intPayOrderID = mPaymentOrderID
            .intSubsidiaryAccountHeadID = Null
            .intSubLegerTypeID = Null
            .vchSubLedgerCode = Null
            .vchName = Null
            .vchHouseName = Null
            .vchStreet = Null
            .vchLocalPlace = Null
            .vchMainPlace = Null
            .vchPost = Null
            .vchPinCode = Null
            .vchPhone = Null

            arrInput = Array(.intPayOrderID, _
            .intSubsidiaryAccountHeadID, _
            .intSubLegerTypeID, _
            .vchSubLedgerCode, _
            .vchName, _
            .vchHouseName, _
            .vchStreet, _
            .vchLocalPlace, _
            .vchMainPlace, _
            .vchPost, _
            .vchPinCode, _
            .vchPhone)
            objDb.ExecuteSP "spSavePayOrderAddress", arrInput, , , mCnn, adCmdStoredProc

        End With
'----------------------------------------------------------------------------------------------------------------
        MsgBox "Saved Payment!", vbInformation, "Saankhya"
        cmdPaymentOrder.Enabled = False
   End Sub
   
   
   
'''    Public Property Let CheckMode(mData As Integer)
'''        intCheckMode = mData
'''    End Property
'''    Public Property Get CheckMode() As Integer
'''        CheckMode = intCheckMode
'''    End Property
''    Private Sub SavePastData()
''        ObjDb.SetConnection mCnn
''        msQl = "SELECT tnyPreriodID From faBillRegisters Where intID = " & vsGrid.TextMatrix(vsGrid.Row, 11) & " "
''        Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''        If Not (Rec.EOF And Rec.BOF) Then
''             mPeriodID = IIf(IsNull(Rec!tnyPreriodID), Null, Rec!tnyPreriodID)
''        Else
''             mPeriodID = Null
''        End If
''        Rec.Close
''        If Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "January" Then
''            mMonthID = 1
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "February" Then
''            mMonthID = 2
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "March" Then
''            mMonthID = 3
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "April" Then
''            mMonthID = 4
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "May" Then
''            mMonthID = 5
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "June" Then
''            mMonthID = 6
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "July" Then
''            mMonthID = 7
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "August" Then
''            mMonthID = 8
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "September" Then
''            mMonthID = 9
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "October" Then
''            mMonthID = 10
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "November" Then
''            mMonthID = 11
''        ElseIf Trim(vsGrid.TextMatrix(vsGrid.Row, 5)) = "December" Then
''            mMonthID = 12
''        End If
''
''         mID = IIf(val(vsGrid.TextMatrix(vsGrid.Row, 11)) = "", -1, val(vsGrid.TextMatrix(vsGrid.Row, 11)))
''         mArrIn = Array(mID, val(vsGrid.TextMatrix(vsGrid.Row, 8)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 1)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 4)), _
''                             mMonthID, _
''                             mPeriodID, _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 6)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 7)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 8)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 9)), _
''                             val(vsGrid.TextMatrix(vsGrid.Row, 12)), _
''                             vchPayOrderNo, _
''                             mintVoucherNo, _
''                             mInstrumentTypeID, _
''                             mInstrumentNo, _
''                             mInstrumentDate, _
''                             Trim(txtRemarks.Text), _
''                             3 _
''                             )
''         ObjDb.ExecuteSP "spSaveBillRegisters", mArrIn, , , mCnn, adCmdStoredProc
''         MsgBox "Saved Payment!", vbInformation, "Saankhya"
''         cmdPaymentOrder.Enabled = False
''    End Sub
''



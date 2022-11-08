VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSearchDishonoredCheque 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Dishonored Cheque"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10395
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5565
      Left            =   30
      TabIndex        =   10
      Top             =   360
      Width           =   10365
      _cx             =   18283
      _cy             =   9816
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchDishonoredCheque.frx":0000
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8985
      TabIndex        =   9
      Top             =   1200
      Width           =   1305
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   810
      Width           =   5460
   End
   Begin VB.CommandButton cmdSearchBank 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10020
      TabIndex        =   4
      Top             =   810
      Width           =   270
   End
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9030
      TabIndex        =   3
      Top             =   450
      Width           =   1260
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7500
      TabIndex        =   2
      Top             =   450
      Width           =   1260
   End
   Begin VB.TextBox txtInstrumentNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   0
      Top             =   450
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8970
      TabIndex        =   13
      Top             =   5280
      Width           =   1305
   End
   Begin VB.CheckBox chkFalse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unable To Find the Cheque"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6180
      TabIndex        =   12
      Top             =   5370
      Width           =   2625
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      Picture         =   "frmSearchDishonoredCheque.frx":0107
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   5400
      Width           =   480
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   3525
         Left            =   150
         TabIndex        =   16
         Top             =   735
         Width           =   2235
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Identify the Cheque No - Double Click the Selected Cheque No From the Grid - Continue with Cheque Bounce Process"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   540
      TabIndex        =   14
      Top             =   5310
      Width           =   5025
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   60
      Picture         =   "frmSearchDishonoredCheque.frx":0411
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label lblAmountType 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Dishonoured Cheques"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   10335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   210
      X2              =   10410
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblBank 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
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
      Left            =   3945
      TabIndex        =   8
      Top             =   855
      Width           =   420
   End
   Begin VB.Label lblDatePeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Date Period"
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
      Left            =   5730
      TabIndex        =   7
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
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
      Left            =   8820
      TabIndex        =   6
      Top             =   435
      Width           =   180
   End
   Begin VB.Label lblInstrumentNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instrument No"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   450
      Width           =   1245
   End
End
Attribute VB_Name = "frmSearchDishonoredCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mType As Integer 'mType=1 for Remittance mType=2 for Return 3 for Bank Charge
    Public Sub FillGrid(mQry As String, qType As Integer)
        '---------String :Qry to Fill Bank Scroll details in the Grid
        '---------mType: For idetifying Type of Amount  ie for Remitance ,Returned or Bank Charge
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDB As New clsDb
            Dim mRowCnt As Integer
            Dim mAmt As Double
            QryType = qType
            If objDB.SetConnection(mCnn) Then
''                mSql = "Select intReconciliationID, intBankAccountHeadID, dtBankEntryDate, dtChequeDate, "
''                mSql = mSql + " vchChequeNo, vchParticulars, fltDrAmount, fltCrAmount"
''                mSql = mSql + " from faBankReconciliationEntries "
''                mSql = mSql + " Where dtChequeDate Between '" & CheckDateInMMM(txtFromDate.Text) & "' and '" & CheckDateInMMM(txtToDate.Text) & "'"
''                If txtInstrumentNo.Text <> "" Then
''                    mSql = mSql + " and vchChequeNo Like '%" & txtInstrumentNo.Text & "%'"
''                End If
''                If Val(txtBank.Tag) <> 0 Then
''                    mSql = mSql + " and intBankAccountHeadID = " & Val(txtBank.Tag)
''                End If
                mSql = mQry
                Rec.Open mSql, mCnn, adOpenStatic, adLockPessimistic
                mRowCnt = 1
                vsGrid.Rows = 2
                If Rec.RecordCount > 0 Then
                    While Not (Rec.EOF Or Rec.BOF)
                        If mType = 1 Then
                            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                            vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intBankAccountHeadID), "", Rec!intBankAccountHeadID)
                            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
                            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                            mAmt = IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                        ElseIf mType = 2 Then
                            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                            vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intBankAccountHeadID), "", Rec!intBankAccountHeadID)
                            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
                            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                            mAmt = IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                        Else
                            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                            vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intBankAccountHeadID), "", Rec!intBankAccountHeadID)
                            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
                            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                            mAmt = IIf(IsNull(Rec!fltDrAmount), Rec!fltDrAmount, Rec!fltCrAmount)
                        End If
                        vsGrid.TextMatrix(mRowCnt, 6) = mAmt
                        Rec.MoveNext
                        mRowCnt = mRowCnt + 1
                        vsGrid.Rows = vsGrid.Rows + 1
                    Wend
                 Else
                    MsgBox "No Item Found", vbApplicationModal
                    gbSearchCode = ""
                    gbSearchStr = ""
                 End If
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your System Administrator"
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub Form_Load()
        If mType = 1 Then
           lblAmountType.Caption = "Remittance Amount"
        ElseIf mType = 1 Then
            lblAmountType.Caption = "Return Amount"
        Else
            lblAmountType.Caption = "Bank Charge/Interest"
        End If
    End Sub


'''''    Private Sub cmdClose_Click()
'''''        If chkFalse.Value = 0 Then
'''''            MsgBox "Please Tick the Check Box and Continue with the Process", vbInformation
'''''            chkFalse.SetFocus
'''''            Exit Sub
'''''        End If
'''''        frmChequeBounceRequest.ChequeIdentifyStatus = 0
'''''        Unload Me
'''''    End Sub
'''''
'''''    Private Sub cmdSearch_Click()
'''''        Call FillGrid
'''''    End Sub
'''''
'''''    Private Sub cmdSearchBank_Click()
'''''        On Error GoTo Err:
'''''            Dim mSql As String
'''''            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & 2
'''''            frmSearchAccountHeads.SQLString = mSql
'''''            frmSearchAccountHeads.Show vbModal
'''''            txtBank.Text = gbSearchStr
'''''            txtBank.Tag = gbSearchID
'''''            txtBank.SetFocus
'''''            gbSearchID = -1
'''''            gbSearchStr = ""
'''''        Exit Sub
'''''Err:
'''''        MsgBox (Error$)
'''''    End Sub
''''
''''    Private Sub Form_Load()
''''        txtFromDate.Text = CheckDateInMMM(DateAdd("m", -1, Date))
''''        txtToDate.Text = CheckDateInMMM(Date)
''''    End Sub
''''
''''    Private Sub Form_Unload(Cancel As Integer)
'''''         If chkFalse.Value = 0 Then
'''''            frmChequeBounceRequest.ChequeIdentifyStatus = 2
'''''         End If
''''    End Sub
''''
''''    Private Sub txtFromDate_GotFocus()
''''        txtFromDate.SelStart = 0
''''        txtFromDate.SelLength = Len(txtFromDate)
''''    End Sub
''''
''''    Private Sub txtFromDate_LostFocus()
''''        If txtFromDate.Text <> "" Then
''''            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
''''        End If
''''    End Sub
''''
''''    Private Sub txtToDate_GotFocus()
''''        txtToDate.SelStart = 0
''''        txtToDate.SelLength = Len(txtToDate)
''''    End Sub
''''
''''    Private Sub txtToDate_LostFocus()
''''        If txtToDate.Text <> "" Then
''''            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
''''        End If
''''    End Sub

    Private Sub vsGrid_DblClick()
        On Error GoTo Err:
            If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
''''            frmChequeBounceRequest.txtInstrumentNo.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
''''            frmChequeBounceRequest.txtBank.Tag = vsGrid.TextMatrix(vsGrid.Row, 1)
''''            frmChequeBounceRequest.txtPerticulars.Text = vsGrid.TextMatrix(vsGrid.Row, 5)
''''            frmChequeBounceRequest.txtChequeDate.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
''''            frmChequeBounceRequest.txtBankEntryDate.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
''''            frmChequeBounceRequest.txtChequeTotal.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
''''            frmChequeBounceRequest.ChequeIdentifyStatus = 1
                gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 6) 'Amount
                gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 3) 'Check Date
            Unload Me
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Public Property Let QryType(mQryType As Integer)
        mType = mQryType
    End Property

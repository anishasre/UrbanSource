VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfReceiptCancellationRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Of Receipt Cancellation Request"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   12825
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   12825
      TabIndex        =   2
      Top             =   0
      Width           =   12825
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   12060
      Top             =   6585
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdNew 
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
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   6180
      Width           =   1245
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5295
      Left            =   75
      TabIndex        =   0
      Top             =   840
      Width           =   12705
      _cx             =   22410
      _cy             =   9340
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfReceiptCancellationRequest.frx":0000
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
End
Attribute VB_Name = "frmListOfReceiptCancellationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mVoucherID     As Variant
    Private Sub cmdNew_Click()
        If gbLBPanchayat = 1 Then
            frmReceiptCancellationRequestPreviousDate.Frame2.Visible = False
            frmReceiptCancellationRequestPreviousDate.Frame1.Left = 3800
            frmReceiptCancellationRequestPreviousDate.Frame1.Caption = "Authorize by Secretary"
        End If
        frmReceiptCancellationRequestPreviousDate.Show
        Me.Hide
    End Sub
    Private Sub Form_Activate()
        Me.Top = 500
        Me.Left = (frmMenu.Width - Me.Width) / 2
        Call FillGrid
    End Sub
    Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        XPC.InitSubClassing
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupSecretary Then
            cmdNew.Enabled = False
        End If
        If gbLBPanchayat = 1 Then
           vsGrid.ColHidden(9) = True
          vsGrid.TextMatrix(0, 8) = "Authorized by Secretary"
        End If
    End Sub
    Private Sub FillGrid()
    Dim mSQL    As String
    Dim objDB   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim mRowCnt As Integer
    Dim mRecCnt As Integer
    Dim mChk    As Integer
   
    If objDB.SetConnection(mCnn) Then
'''        mSql = "SELECT  faReceiptCancellationRequest.vchProceedingsNO,faReceiptCancellationRequest.intID, faReceiptCancellationRequest.dtProceedingsDate,"
'''        mSql = mSql + " faVouchers.intVoucherNo, faVouchers.dtDate,faVouchers.intTransactionTypeID,faReceiptCancellationRequest.vchReason, faUser.vchUserName, faReceiptCancellationRequest.dtRequestedDate,"
'''        mSql = mSql + " faReceiptCancellationRequest.tnyStatusAO,faReceiptCancellationRequest.tnyStatusSec,faReceiptCancellationRequest.intVoucherID,faReceiptCancellationRequest.numStationaryNo,faReceiptCancellationRequest.vchRemarks FROM faReceiptCancellationRequest INNER JOIN "
'''        mSql = mSql + " faUser ON faReceiptCancellationRequest.numRequestedBy = faUser.numUserID INNER JOIN"
'''        mSql = mSql + " faVouchers ON faReceiptCancellationRequest.intVoucherID = faVouchers.intVoucherID"
        
        mSQL = " SELECT  faReceiptCancellationRequest.vchProceedingsNO,faReceiptCancellationRequest.intID, faReceiptCancellationRequest.dtProceedingsDate,"
        mSQL = mSQL + " faVouchers.intVoucherNo, faVouchers.dtDate,faVouchers.intTransactionTypeID,faReceiptCancellationRequest.vchReason, faUser.vchUserName,"
        mSQL = mSQL + " faReceiptCancellationRequest.dtRequestedDate,"
        mSQL = mSQL + " faReceiptCancellationRequest.tnyStatusAO,faReceiptCancellationRequest.tnyStatusSec,faReceiptCancellationRequest.intVoucherID,"
        mSQL = mSQL + " faReceiptCancellationRequest.numStationaryNo , faReceiptCancellationRequest.vchRemarks"
        mSQL = mSQL + " From faReceiptCancellationRequest"
        mSQL = mSQL + " INNER JOIN faUser ON faReceiptCancellationRequest.numRequestedBy = faUser.numUserID"
        mSQL = mSQL + " INNER JOIN faVouchers ON faReceiptCancellationRequest.intVoucherID = faVouchers.intVoucherID"
        mSQL = mSQL + " WHERE faVouchers.dtDate IN   ("
        mSQL = mSQL + "     Select DISTINCT Top 2 dtDate From ("
        mSQL = mSQL + "     Select Convert(varchar(18),getdate(),106) as dtDate"
        mSQL = mSQL + "     Union All"
        mSQL = mSQL + "     Select dtDate From ("
        mSQL = mSQL + "     Select distinct Top 2  dtDate From faVouchers Order by dtDate Desc"
        mSQL = mSQL + "     ) A"
        mSQL = mSQL + "     )B"
        mSQL = mSQL + " )"

        
        
        
        Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        mRecCnt = 1
        vsGrid.Clear 1, 1
        While Not (Rec.EOF Or Rec.BOF)
          
            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intID), "", Rec!intID) 'mRecCnt
            vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchProceedingsNo), "", Rec!vchProceedingsNo)
            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtProceedingsDate), "", Rec!dtProceedingsDate)
            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
            vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
            vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!dtRequestedDate), "", Rec!dtRequestedDate)
            vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!numStationaryNo), "", Rec!numStationaryNo)
            vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
            vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
            If Rec!tnyStatusAO = 0 And Rec!tnyStatusSec = 0 Then
                 vsGrid.Cell(flexcpChecked, mRowCnt, 8) = vbUnchecked
                 vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbUnchecked
            End If
            If Rec!tnyStatusAO = 1 Then
                 vsGrid.Cell(flexcpChecked, mRowCnt, 8) = vbChecked
                 vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbUnchecked
            End If
            If Rec!tnyStatusSec = 1 Then
              If Rec!tnyStatusAO = 1 Then
                 vsGrid.Cell(flexcpChecked, mRowCnt, 8) = vbChecked
              Else
                vsGrid.Cell(flexcpChecked, mRowCnt, 8) = vbUnchecked
              End If
              vsGrid.Cell(flexcpChecked, mRowCnt, 9) = vbChecked
            End If
            vsGrid.TextMatrix(mRowCnt, 10) = Rec!intID
            Rec.MoveNext
            vsGrid.Rows = vsGrid.Rows + 1
            mRowCnt = mRowCnt + 1
            mRecCnt = mRecCnt + 1
        Wend
    Rec.Close
    End If
    End Sub
    Private Sub vsGrid_DblClick()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
       Dim objDB As New clsDB
        Dim mSQL As String
        If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then Exit Sub
        
        '***********************************'
        '   OPERATOR LOGIN                  '
        '***********************************'
        
            frmListOfReceiptCancellationRequest.Hide
            frmReceiptCancellationRequestPreviousDate.Show
            frmReceiptCancellationRequestPreviousDate.txtProceedingsNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 10)
            frmReceiptCancellationRequestPreviousDate.txtProceedingsNo = vsGrid.TextMatrix(vsGrid.Row, 1)
            frmReceiptCancellationRequestPreviousDate.txtProceedingsDate = vsGrid.TextMatrix(vsGrid.Row, 2)
            frmReceiptCancellationRequestPreviousDate.txtReceiptsNo = vsGrid.TextMatrix(vsGrid.Row, 3)
            frmReceiptCancellationRequestPreviousDate.txtReceiptsDate = CheckDateInMMM(CStr(vsGrid.TextMatrix(vsGrid.Row, 4)))
            frmReceiptCancellationRequestPreviousDate.cmbReason = vsGrid.TextMatrix(vsGrid.Row, 5)
            frmReceiptCancellationRequestPreviousDate.txtRequestedBy = vsGrid.TextMatrix(vsGrid.Row, 6)
            frmReceiptCancellationRequestPreviousDate.txtRequestedByDate = vsGrid.TextMatrix(vsGrid.Row, 7)
            frmReceiptCancellationRequestPreviousDate.mVoucherID = vsGrid.TextMatrix(vsGrid.Row, 11)
            frmReceiptCancellationRequestPreviousDate.txtStationaryNo = vsGrid.TextMatrix(vsGrid.Row, 12)
            frmReceiptCancellationRequestPreviousDate.txtRemarks = vsGrid.TextMatrix(vsGrid.Row, 13)
            frmReceiptCancellationRequestPreviousDate.mTransactionTypeID = vsGrid.TextMatrix(vsGrid.Row, 14)
            frmReceiptCancellationRequestPreviousDate.txtStationaryNo.Enabled = False
            
        '***********************************'
        '   IF ACCOUNTS OFFICER AUTHORIZED  '
        '***********************************'
            If gbLBPanchayat = 1 Then
                frmReceiptCancellationRequestPreviousDate.Frame1.Caption = "Authorize by Secretary"
            End If
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 8) = 1 Then
                frmReceiptCancellationRequestPreviousDate.cmdFirstAuthorize.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmdRequest.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmbReason.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRequestedBy.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRequestedByDate.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtProceedingsNo.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtProceedingsDate.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRemarks.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtStationaryNo.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmdSearchReceipt.Enabled = False
                'frmReceiptCancellationRequestPreviousDate.dtpProceedingsDate.Enabled = False
           End If
           
         '***********************************'
         '   IF SECRETARY AUTHORIZED  '
         '***********************************'
            If gbLBPanchayat = 1 Then
                frmReceiptCancellationRequestPreviousDate.Frame2.Visible = False
                frmReceiptCancellationRequestPreviousDate.Frame1.Left = 3800
            Else
            If vsGrid.Cell(flexcpChecked, vsGrid.Row, 9) = 1 Then
                frmReceiptCancellationRequestPreviousDate.cmdSecondtAuthorize.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmdRequest.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmbReason.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRequestedBy.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRequestedByDate.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtProceedingsNo.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtProceedingsDate.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtRemarks.Enabled = False
                frmReceiptCancellationRequestPreviousDate.txtStationaryNo.Enabled = False
                frmReceiptCancellationRequestPreviousDate.cmdSearchReceipt.Enabled = False
                'frmReceiptCancellationRequestPreviousDate.dtpProceedingsDate.Enabled = False
            End If
            End If
'-----------------------------------------------------------------------------------------------------------------------
         
          '*************************************************'
          '   TO DISPLAY ACCOUNTS OFFICER AUTHORIZED DATE   '
          '*************************************************'
    
    
    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateAO "
            mSQL = mSQL + " FROM  faReceiptCancellationRequest INNER JOIN "
            mSQL = mSQL + " faUser ON faReceiptCancellationRequest.numAuthorisedByAO = faUser.numUserID "
            mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedBy = Rec!vchUserName
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedDate = Rec!dtAuthorisationDateAO
           
            Else
                 frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedBy = gbUserName
                 frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedDate = gbDate
            End If
            Rec.Close
        End If
        If vsGrid.Cell(flexcpChecked, vsGrid.Row, 9) = 1 Then
            If objDB.SetConnection(mCnn) Then
                mSQL = " SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateSec "
                mSQL = mSQL + " FROM faReceiptCancellationRequest INNER JOIN faUser ON faReceiptCancellationRequest.numAuthorisedBySec = faUser.numUserID"
                mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedBy = Rec!vchUserName
                    frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedDate = Rec!dtAuthorisationDateSec
                End If
                Rec.Close
             End If
        End If
    End If
  '--------------------------------------------------------------------------------------------------------------------------
    
          '*************************************************'
          '   TO DISPLAY SECRETARY AUTHORIZED DATE   '
          '*************************************************'
    
    
    'If gbSeatGroupID = 10 Or gbSeatGroupID = 1 Or gbSeatGroupID = 2 Then
    If gbSeatGroupID = gbSeatGroupSecretary Then
        If objDB.SetConnection(mCnn) Then
            mSQL = " SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateSec "
            mSQL = mSQL + " FROM faReceiptCancellationRequest INNER JOIN faUser ON faReceiptCancellationRequest.numAuthorisedBySec = faUser.numUserID"
            mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedBy = Rec!vchUserName
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedDate = Rec!dtAuthorisationDateSec
            Else
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedBy = gbUserName
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedDate = gbDate
            End If
            Rec.Close
        End If
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateAO "
            mSQL = mSQL + " FROM  faReceiptCancellationRequest INNER JOIN "
            mSQL = mSQL + " faUser ON faReceiptCancellationRequest.numAuthorisedByAO = faUser.numUserID "
            mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedBy = Rec!vchUserName
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedDate = Rec!dtAuthorisationDateAO
            End If
        End If
    End If
  '-----------------------------------------------------------------------------------------------------------------------
  
  
  If gbSeatGroupID = gbSeatGroupCashier Then
        If objDB.SetConnection(mCnn) Then
            mSQL = " SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateSec "
            mSQL = mSQL + " FROM faReceiptCancellationRequest INNER JOIN faUser ON faReceiptCancellationRequest.numAuthorisedBySec = faUser.numUserID"
            mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedBy = Rec!vchUserName
                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedDate = Rec!dtAuthorisationDateSec
'            Else
'                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedBy = Rec!vchUsername
'                frmReceiptCancellationRequestPreviousDate.txtSecondtAuthorizedDate = Rec!dtAuthorisationDateSec
            End If
            Rec.Close
        End If
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT faUser.vchUserName, faReceiptCancellationRequest.dtAuthorisationDateAO "
            mSQL = mSQL + " FROM  faReceiptCancellationRequest INNER JOIN "
            mSQL = mSQL + " faUser ON faReceiptCancellationRequest.numAuthorisedByAO = faUser.numUserID "
            mSQL = mSQL + " where faReceiptCancellationRequest.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedBy = Rec!vchUserName
                frmReceiptCancellationRequestPreviousDate.txtFirstAuthorizedDate = Rec!dtAuthorisationDateAO
            End If
        End If
    End If
  
  
  
    End Sub



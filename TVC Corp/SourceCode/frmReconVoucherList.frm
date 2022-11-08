VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmReconVoucherList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE TO UNRECONCILED LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2970
      TabIndex        =   8
      Top             =   7575
      Width           =   3240
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   6945
      TabIndex        =   7
      Top             =   0
      Width           =   6945
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT UNRECONCILED TRANSACTIONS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   210
         Width           =   4545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   45
      TabIndex        =   0
      Top             =   615
      Width           =   4875
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   165
         Width           =   2085
      End
      Begin VB.CommandButton cmdYearUp 
         Caption         =   ">>"
         Height          =   345
         Left            =   1305
         TabIndex        =   2
         Top             =   143
         Width           =   525
      End
      Begin VB.CommandButton cmdYearDown 
         Caption         =   "<<"
         Height          =   345
         Left            =   30
         TabIndex        =   1
         Top             =   143
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "MONTH"
         Height          =   225
         Left            =   2115
         TabIndex        =   4
         Top             =   225
         Width           =   660
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         Caption         =   "2013-14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   3
         Top             =   210
         Width           =   600
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6315
      Left            =   30
      TabIndex        =   6
      Top             =   1185
      Width           =   13965
      _cx             =   24633
      _cy             =   11139
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
      FormatString    =   $"frmReconVoucherList.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      ShowComboButton =   0
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
Attribute VB_Name = "frmReconVoucherList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintBankAccountHeadID As Integer
Private mintYearID As Integer
Private mintMonthID As Integer
Private mdtLastDate As Variant
Private mintReconID As Variant
    
    
    Private Sub FillTransactions()
        Dim objDB As New clsDB
        Dim Rec As New Recordset
        Dim mCnn As New ADODB.Connection
        Dim msql As String
        Dim mDt1 As Date
        Dim mDt2 As Date
        Dim mYear As String
        vsGrid.Rows = 1
        If mintBankAccountHeadID < 1 Then
            Exit Sub
        End If
        
               
        If cmbMonth.Text <> "" Then
            If cmbMonth.ItemData(cmbMonth.ListIndex) < 4 Then
                mYear = lblYear.Tag + 1
            Else
                mYear = lblYear.Tag
            End If
            mDt1 = CDate("01-" + Left(cmbMonth.Text, 3) + "-" + mYear)
            If IsDate(mDt1) Then
                mDt2 = DateAdd("m", 1, mDt1)
                mDt2 = DateAdd("d", -1, mDt2)
            End If
            
            If Not IsDate(mDt2) Then
                Exit Sub
            End If
        End If
        
        msql = ""
        msql = msql + " SELECT dtDate, faVouchers.intVoucherNO, faTransactionChild.fltAmount," & vbCrLf
        msql = msql + " CASE WHEN intInstrumentTypeID = 2 THEN 'Treasury Chalan'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 3 THEN 'Postal Order'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 4 THEN 'Demand Draft'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 5 THEN 'Cheque'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 6 THEN 'Letter Of Authority'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 7 THEN 'Treasury Bill'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 8 THEN 'Bank Pay-in-Slip'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID = 9 THEN 'Directly Credited To Bank'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID =10 THEN 'Directly Debited To Bank'" & vbCrLf
        msql = msql + "      WHEN intInstrumentTypeID =11 THEN 'Card Payments'" & vbCrLf
        msql = msql + "      Else 'Cash'" & vbCrLf
        msql = msql + " END InstType," & vbCrLf
        msql = msql + " ISNULL(vchInstrumentNo,'') vchInstrumentNo, 0 Flag, faVouchers.intVoucherID, tinDebitOrCreditFlag, " & vbCrLf
        msql = msql + " faTransactionChild.intTransactionID , faTransactionChild.intSerialNo, dtInstrumentDate, tnyVoucherTypeID " & vbCrLf
        msql = msql + " FROM faTransactionChild" & vbCrLf
        msql = msql + " INNER JOIN faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID" & vbCrLf
        msql = msql + " INNER JOIN faVouchers ON faVouchers.intVoucherID = faTransactions.intVoucherID" & vbCrLf
        msql = msql + " LEFT JOIN faBankReconcileChild ON faBankReconcileChild.intTransactionID = faTransactionChild.intTransactionID" & vbCrLf
        msql = msql + " Where intAccountHeadID = " & mintBankAccountHeadID & vbCrLf
        msql = msql + "      AND dtDate Between '" & DdMmmYy(mDt1) & "' and '" & DdMmmYy(mDt2) & "'" & vbCrLf
        msql = msql + " AND intReconID IS NULL " & vbCrLf
        msql = msql + " ORDER BY dtDate, faVouchers.intVoucherID " & vbCrLf
        
        
        
        
        msql = ""
        msql = msql + " SELECT " & vbCrLf
        msql = msql + " dtDate, faVouchers.intVoucherNO, faTransactionChild.fltAmount," & vbCrLf
        msql = msql + "  CASE WHEN intInstrumentTypeID = 2 THEN 'Treasury Chalan'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 3 THEN 'Postal Order'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 4 THEN 'Demand Draft'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 5 THEN 'Cheque'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 6 THEN 'Letter Of Authority'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 7 THEN 'Treasury Bill'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 8 THEN 'Bank Pay-in-Slip'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID = 9 THEN 'Directly Credited To Bank'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID =10 THEN 'Directly Debited To Bank'" & vbCrLf
        msql = msql + "       WHEN intInstrumentTypeID =11 THEN 'Card Payments'" & vbCrLf
        msql = msql + "       Else 'Cash'" & vbCrLf
        msql = msql + "  END InstType," & vbCrLf
        msql = msql + "  ISNULL(faVouchers.vchInstrumentNo,'') vchInstrumentNo, 0 Flag, faVouchers.intVoucherID, tinDebitOrCreditFlag," & vbCrLf
        msql = msql + "  faTransactionChild.intTransactionID , faTransactionChild.intSerialNo, faVouchers.dtInstrumentDate, faVouchers.tnyVoucherTypeID" & vbCrLf
        msql = msql + "  From faTransactionChild" & vbCrLf
        msql = msql + "  INNER JOIN faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID" & vbCrLf
        msql = msql + "  INNER JOIN faVouchers ON faVouchers.intVoucherID = faTransactions.intVoucherID" & vbCrLf
        msql = msql + "  LEFT JOIN faBankReconcileChild ON faBankReconcileChild.intTransactionID = faTransactionChild.intTransactionID" & vbCrLf
        msql = msql + "  Where faTransactionChild.intAccountHeadID = " & mintBankAccountHeadID & vbCrLf
        msql = msql + "       AND dtDate Between '" & DdMmmYy(mDt1) & "' and '" & DdMmmYy(mDt2) & "'" & vbCrLf
        msql = msql + "      AND intReconID IS NULL" & vbCrLf
        msql = msql + "      AND ISNULL(tnyCancelFlag,0) <> 1 " & vbCrLf
        msql = msql + "  ORDER BY dtDate, faVouchers.intVoucherID" & vbCrLf
        
            
        
        objDB.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rec.EOF And Rec.BOF) Then
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = 11
            vsGrid.RowSel = vsGrid.Rows - 1
            msql = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = msql
            vsGrid.ColSel = 0
            vsGrid.RowSel = 0
        End If
        Rec.Close
    End Sub
    
    Private Sub SaveToUnreconcileList()
    
        Dim objDB As New clsDB
        Dim Rec As New Recordset
        Dim mCnn As New ADODB.Connection
        Dim msql As String
        Dim mDt1 As Date
        Dim mDt2 As Date
        
        Dim mLoop As Integer
        Dim mArrIn As Variant
        Dim mTypeID As Integer
        Dim mDrAmt As Variant
        Dim mCrAmt As Variant
        Dim mdtVoucherDate As Variant
        Dim mdtInstDate As Variant
        
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 5) = 1 Then
                'MsgBox vsGrid.TextMatrix(mLoop, 1)
                
                If val(vsGrid.TextMatrix(mLoop, 7)) = 1 Then ' tnyDebitOrCredit [Col:7]
                    mTypeID = 1
                    mDrAmt = val(vsGrid.TextMatrix(mLoop, 2))
                    mCrAmt = Null
                Else
                    mTypeID = 3
                    mDrAmt = Null
                    mCrAmt = val(vsGrid.TextMatrix(mLoop, 2))
                End If
                
                If IsDate(vsGrid.TextMatrix(mLoop, 0)) Then
                    mdtVoucherDate = CDate(vsGrid.TextMatrix(mLoop, 0))
                Else
                    mdtVoucherDate = Null
                End If
                
                If IsDate(vsGrid.TextMatrix(mLoop, 10)) Then
                    mdtInstDate = CDate(vsGrid.TextMatrix(mLoop, 10))
                Else
                    mdtInstDate = Null
                End If
                            
                mArrIn = Array(mintReconID, _
                                Null, _
                                mintBankAccountHeadID, _
                                mTypeID, _
                                mDrAmt, _
                                mCrAmt, _
                                val(vsGrid.TextMatrix(mLoop, 6)), _
                                vsGrid.TextMatrix(mLoop, 1), _
                                val(vsGrid.TextMatrix(mLoop, 8)), _
                                val(vsGrid.TextMatrix(mLoop, 9)), _
                                mdtVoucherDate, _
                                vsGrid.TextMatrix(mLoop, 4), _
                                mdtInstDate, _
                                val(vsGrid.TextMatrix(mLoop, 11)), _
                                Null, Null)
                
                objDB.ExecuteSP "spSaveBankReconcileChild", mArrIn
                
    
                '    STORED PROCEDURE :: spSaveBankReconcileChild
                '    PARAMETERS::
               
                            '@intReconID    [int],
                            '@intReconChdID     [Bigint]=Null,
                            '@intAccountHeadID  [int],
                            '@tnyTypeID     [int],
                            '@numDrAmount   [float],
                            '@numCrAmount   [float],
                            '@intVoucherID  [bigint],
                            '@vchVoucherNo  [numeric],
                            '@intTransactionID [bigint],
                            '@intSlNo   [int],
                            '@dtVoucherDate     [smalldatetime],
                            '@vchInstrumentNo   [varchar](50),
                            '@dtInstrumentDate  [smalldatetime],
                            '@tnyVoucherTypeID  [tinyint],
                            '@tnyFlag   [tinyint],
                            '@vchRemarks    [varchar](200))
     
           
            
                
            Else
                
            End If
        Next
        
    End Sub
    
    Private Sub cmbMonth_Click()
        Dim mIndexID As Integer
        If cmbMonth.ListIndex > -1 Then
            mIndexID = cmbMonth.ItemData(cmbMonth.ListIndex)
        Else
            Exit Sub
        End If
        
        '    If mintMonthID < 4 Then
        '        If Not (mIndexID <= mintMonthID Or (mIndexID > 3 And mIndexID < 13)) Then
        '            Exit Sub
        '        End If
        '    Else
        '        If Not (mIndexID > 3 And (mIndexID <= mintMonthID)) Then
        '            Exit Sub
        '        End If
        '    End If
    
    
        
        Dim mDate As Date
        
        If mIndexID < 4 Then
            mintYearID = val(lblYear.Tag) + 1
        Else
            mintYearID = val(lblYear.Tag)
        End If
        mDate = DateSerial(mintYearID, mIndexID, 1)
        mDate = DateAdd("m", 1, mDate)
        mDate = DateAdd("d", -1, mDate)
        If mDate <= mdtLastDate Then
            Call FillTransactions
        Else
            vsGrid.Rows = 1
        End If
    End Sub
    
    Private Sub cmdUpdate_Click()
        Call SaveToUnreconcileList
        Call FillTransactions
    End Sub
    
    Private Sub cmdYearDown_Click()
        mintYearID = mintYearID - 1
        lblYear.Caption = Trim(str(mintYearID)) + "-" + Right(Trim(str(mintYearID + 1)), 2)
        lblYear.Tag = mintYearID
    End Sub
    
    Private Sub cmdYearUp_Click()
        mintYearID = mintYearID + 1
        lblYear.Caption = Trim(str(mintYearID)) + "-" + Right(Trim(str(mintYearID + 1)), 2)
        lblYear.Tag = mintYearID
    End Sub
    
    Private Sub Form_Initialize()
        Me.Left = 0
        Me.Top = 1200
        Call FillMonth
        
        If IsDate(mdtLastDate) Then
            If Month(mdtLastDate) > 3 Then
                mintYearID = Year(mdtLastDate)
            Else
                mintYearID = Year(mdtLastDate) - 1
            End If
        End If
        
        If Not (mintYearID > 2000 And mintYearID <= gbFinancialYearID) Then
            mintYearID = gbFinancialYearID
        End If
        lblYear.Caption = Trim(str(mintYearID)) + "-" + Right(Trim(str(mintYearID + 1)), 2)
        lblYear.Tag = mintYearID
        cmdYearDown.Tag = mintYearID
        cmdYearUp.Tag = mintYearID
    End Sub
    
    Private Sub SetYearID()
        If IsDate(mdtLastDate) Then
            If Month(mdtLastDate) > 3 Then
                mintYearID = Year(mdtLastDate)
            Else
                mintYearID = Year(mdtLastDate) - 1
            End If
        End If
        
        If Not (mintYearID > 2000 And mintYearID <= gbFinancialYearID) Then
            mintYearID = gbFinancialYearID
        End If
        
        lblYear.Caption = Trim(str(mintYearID)) + "-" + Right(Trim(str(mintYearID + 1)), 2)
        lblYear.Tag = mintYearID
    End Sub
    
    Private Sub FillMonth()
        cmbMonth.Clear
        cmbMonth.AddItem ""
        cmbMonth.ItemData(cmbMonth.NewIndex) = -1
        
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
        
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        frmReconciliation.cmdRefresh.value = True
        Set frmReconVoucherList = Nothing
    End Sub
    
    Public Property Get BankAccountHeadID() As Variant
        BankAccountHeadID = mintBankAccountHeadID
    End Property
    
    Public Property Let BankAccountHeadID(mData As Variant)
        mintBankAccountHeadID = mData
    End Property
    
    Public Property Get LastDate() As Variant
        LastDate = mintYearID
    End Property
    
    Public Property Let LastDate(mData As Variant)
        mdtLastDate = mData
        If IsDate(mdtLastDate) Then
            Call SetYearID
            Dim mLoop As Integer
            For mLoop = 0 To cmbMonth.ListCount
                If cmbMonth.ItemData(mLoop) = Month(mdtLastDate) Then
                    cmbMonth.ListIndex = mLoop
                    Exit For
                End If
            Next
            
        End If
            
    End Property
    
    Public Property Get ReconID() As Variant
        ReconID = mintReconID
    End Property
    
    Public Property Let ReconID(mData As Variant)
        mintReconID = mData
    End Property
    
    

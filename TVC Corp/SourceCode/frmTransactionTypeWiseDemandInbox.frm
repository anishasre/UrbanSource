VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmTransactionTypeWiseDemandInbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TransactionTypeWise Demand Inbox"
   ClientHeight    =   6720
   ClientLeft      =   90
   ClientTop       =   2550
   ClientWidth     =   11385
   DrawMode        =   4  'Mask Not Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5730
      TabIndex        =   14
      Top             =   6210
      Width           =   1020
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4185
      TabIndex        =   12
      Top             =   930
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6810
      TabIndex        =   11
      Top             =   5670
      Width           =   1950
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4650
      TabIndex        =   10
      Top             =   6210
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11355
      TabIndex        =   9
      Top             =   6105
      Width           =   11385
   End
   Begin VB.TextBox txtDemandNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1230
      TabIndex        =   6
      Top             =   930
      Width           =   1695
   End
   Begin VB.TextBox txtTransactionType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Text            =   "txtTransactionType"
      Top             =   5670
      Width           =   4530
   End
   Begin VB.CommandButton cmdtransactionTypeSearch 
      Caption         =   "..."
      Height          =   285
      Left            =   5700
      TabIndex        =   2
      Top             =   5685
      Width           =   375
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11355
      _cx             =   20029
      _cy             =   7435
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransactionTypeWiseDemandInbox.frx":0000
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
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11355
      TabIndex        =   1
      Top             =   0
      Width           =   11385
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "TransactionType Wise Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   8775
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Demand Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2985
      TabIndex        =   13
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6270
      TabIndex        =   8
      Top             =   5700
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Demand No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trans.Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   4
      Top             =   5700
      Width           =   945
   End
End
Attribute VB_Name = "frmTransactionTypeWiseDemandInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mZonalID As Double ' Property Variable

Private Sub cmdClear_Click()
    txtTransactionType = ""
    txtTransactionType.Locked = True
    txtDemandNo.Text = frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 3) ', "", frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 4)))
    txtDemandNo.Locked = True
    Call FillGrid(CDate(frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 0)), IIf(val(txtTransactionType.Tag) = 0, "%", txtTransactionType.Tag))
End Sub

Private Sub cmdSearch_Click()
       Call FillGrid(CDate(txtDate.Text), IIf(val(txtTransactionType.Tag) = 0, "%", txtTransactionType.Tag))
End Sub

Private Sub cmdtransactionTypeSearch_Click()
   frmSearchTransactionType.Show vbModal
    If gbSearchID > 0 Then
        txtTransactionType.Tag = gbSearchID
        txtTransactionType.Text = gbSearchStr
        gbSearchID = -1
        gbSearchStr = ""
    End If
End Sub

Private Sub Form_Activate()
    Me.Height = 7230
    Me.Width = 11505
    Me.Left = 0
    Me.Top = 0
End Sub
Private Sub Form_Load()
 ' XPC.InitSubClassing
    txtTransactionType = ""
    txtTransactionType.Locked = True
    txtDemandNo.Text = frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 3) ', "", frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 4)))
    txtDemandNo.Tag = frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 4)
    txtDemandNo.Locked = True
    If IsDate(frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 0)) Then
        Call FillGrid(CDate(frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 0)), IIf(val(txtTransactionType.Tag) = 0, "%", txtTransactionType.Tag))
    End If
End Sub
Private Sub FillGrid(ByVal mDate As Date, mTransactionTypeID As String)
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim RecFin       As New ADODB.Recordset
        Dim arrInput    As Variant
        Dim objAcc  As New clsAccounts
        Dim mCnn As New ADODB.Connection
        Dim mCnnFin As New ADODB.Connection
        Dim mDt As Date
        Dim mRow As Integer
        Dim mSql As String
        Dim mTotal As Variant
        Dim mNumDemandID As Variant
        Dim mFlag As Variant
        Dim mReceiptNo1 As Variant
        Dim mReceiptNo2 As Variant
        Dim mAdvAmt As Variant
        mAdvAmt = 0
        mTotal = 0
        mRow = 1
      '  Dim t As Double
        T = 0
       'vsGrid.Rows = 1
       ' vsGrid.Rows = 20
        vsGrid.Clear
        vsGrid.TextMatrix(0, 1) = "Transaction Types"
        vsGrid.TextMatrix(0, 2) = "Amount"
        vsGrid.TextMatrix(0, 3) = "Receipt No"
        objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaHO
        objdb.CreateNewConnection mCnnFin, enuSourceString.Saankhya
        mDt = mDate
        txtDate = Format(mDt, "Dd-MMM-Yyyy")
        txtDate.Locked = True
        arrInput = Array(mDt, mTransactionTypeID)
        mSql = "Select numDemandID  from faIDemandTBL where vchDemandNO='" & txtDemandNo.Text & "'"
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            mNumDemandID = Rec!numDemandID
        End If
        mSql = ""
        Rec.Close

        mSql = mSql + " Select A.* From ("
        mSql = mSql + " Select faTransactionType.vchTransactionType,faTransactionType.intTransactionTypeID,sum(fltamount) as fltAmount,0 as flag  from faIDemandTBL"
        mSql = mSql + " Inner JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
        mSql = mSql + " Inner JOIN faTransactionType ON faTRansactionType.intTransactionTypeID=faIDemandChild.intTransactionTypeID"
        mSql = mSql + " where faIDemandTBL.numLocationID = " & mZonalID & " AND faIDemandTBL.dtDemandDate='" & Format(mDt, "dd/MMM/yyyy") & "' And faTRansactionType.intTransactionTypeID like'%'"
        mSql = mSql + " Group by faIDemandChild.intTransactionTypeID,faTransactionType.vchTransactionType,faTransactionType.intTransactionTypeID"
        mSql = mSql + " Union All"
        mSql = mSql + " Select faTransactionType.vchTransactionType+'(Adjusted)',faTransactionType.intTransactionTypeID,sum(fltamount) as fltAmount,1 as flag  from faIDemandTBL"
        mSql = mSql + " Inner JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
        mSql = mSql + " Inner JOIN faTransactionType ON faTRansactionType.intTransactionTypeID=faIDemandChild.intTransactionTypeID"
        mSql = mSql + " where faIDemandTBL.numLocationID =  " & mZonalID & "  AND faIDemandTBL.dtDemandDate='" & Format(mDt, "dd/MMM/yyyy") & "' And  faIDemandChild.tnystatus in(" & 10 & "," & 11 & ")"
        mSql = mSql + " Group by  faIDemandChild.intTransactionTypeID,faTransactionType.vchTransactionType,faTransactionType.intTransactionTypeID)A"
        mSql = mSql + " where A.inttransactionTypeID Like order by A.intTransactionTypeID"
                
        'CHANGED BY AIBY ON 26-Nov-2011
        mSql = " Select A.* From ("
        mSql = mSql + "     SELECT faTransactionType.vchTransactionType, faTransactionType.intTransactionTypeID, SUM(fltamount) AS fltAmount, 0 AS flag"
        mSql = mSql + "     From faIDemandTBL"
        mSql = mSql + "     INNER JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
        mSql = mSql + "     INNER JOIN faTransactionType ON faTRansactionType.intTransactionTypeID=faIDemandChild.intTransactionTypeID"
        mSql = mSql + "     WHERE faIDemandTBL.numLocationID = " & mZonalID & " AND faIDemandTBL.dtDemandDate = '" & Format(mDt, "dd/MMM/yyyy") & "' And faTRansactionType.intTransactionTypeID like '%'"
        mSql = mSql + "     GROUP BY faIDemandChild.intTransactionTypeID,faTransactionType.vchTransactionType,faTransactionType.intTransactionTypeID"
        mSql = mSql + " Union All"
        mSql = mSql + "     Select faTransactionType.vchTransactionType + '(Adjusted)', faTransactionType.intTransactionTypeID, SUM(fltamount) As fltAmount, 1 As flag"
        mSql = mSql + "     From faIDemandTBL"
        mSql = mSql + "     Inner JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
        mSql = mSql + "     Inner JOIN faTransactionType ON faTRansactionType.intTransactionTypeID=faIDemandChild.intTransactionTypeID"
        mSql = mSql + "     Where faIDemandTBL.numLocationID =  " & mZonalID & "  AND faIDemandTBL.dtDemandDate='" & Format(mDt, "dd/MMM/yyyy") & "' And  faIDemandChild.tnystatus in(10,11)"
        mSql = mSql + "     Group by  faIDemandChild.intTransactionTypeID, faTransactionType.vchTransactionType, faTransactionType.intTransactionTypeID"
        mSql = mSql + " )"
        mSql = mSql + " A Where A.inttransactionTypeID Like'" & mTransactionTypeID & "'"
        mSql = mSql + " Order by A.intTransactionTypeID"

        Rec.Open mSql, mCnn
       ' vsGrid.Clear
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                vsGrid.Row = mRow
                vsGrid.TextMatrix(mRow, 1) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                If Rec!flag = 1 Then
                  vsGrid.TextMatrix(mRow, 5) = 1
                  vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &H4DA4F2
                Else
                  vsGrid.TextMatrix(mRow, 5) = 0
                End If
                vsGrid.TextMatrix(mRow, 2) = IIf(IsNull(Rec!fltAmount), 0, Format(Rec!fltAmount, "0.00"))
                vsGrid.TextMatrix(mRow, 4) = IIf(IsNull(Rec!intTransactionTypeID), 0, Rec!intTransactionTypeID)
                
                mSql = "Select Distinct intVoucherNo,tnyCancelFlag,tnyVoucherTypeID from faVouchers"
                mSql = mSql + " Left Join faVoucherChild on faVouchers.intVoucherID=faVoucherChild.intVoucherID"
                mSql = mSql + " where faVoucherChild.numDemandID=" & mNumDemandID & "and faVouchers.intTransactionTypeID=" & Rec!intTransactionTypeID
                RecFin.Open mSql, mCnnFin, adOpenStatic, adLockOptimistic
                If Not (RecFin.BOF And RecFin.EOF) Then
                    '==================================================
                    If Rec!flag = 0 Then
                        While Not RecFin.EOF
                           If RecFin!tnyVoucherTypeID = 10 Then
                             mReceiptNo1 = RecFin!intVoucherNo
                           End If
                          RecFin.MoveNext
                         Wend
                    ElseIf Rec!flag = 1 Then
                        While Not RecFin.EOF
                           If RecFin!tnyVoucherTypeID = 40 Then
                             mReceiptNo2 = RecFin!intVoucherNo
                           End If
                           RecFin.MoveNext
                        Wend
                    End If
                    RecFin.MoveFirst
                    '====================================================
                    If vsGrid.TextMatrix(mRow, 1) = "Property Tax(Adjusted)" Or vsGrid.TextMatrix(mRow, 1) = "Rent on Building / Stalls(Adjusted)" Or vsGrid.TextMatrix(mRow, 1) = "Rent on Land/Bunks(Adjusted)" Then
                        If mReceiptNo2 <> "" Then
                             vsGrid.TextMatrix(mRow, 3) = mReceiptNo2
                             vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &H4DA4F2
                        End If
                    Else
                        If mReceiptNo1 <> "" Then
                            vsGrid.TextMatrix(mRow, 3) = mReceiptNo1
                            vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
                        End If
                    End If
                    '==========================================================
                    If mFlag = 1 Then
                           vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = vbRed
                    End If
                Else
                      vsGrid.TextMatrix(mRow, 3) = ""
                End If
                If Rec!flag = 0 Then
                    mTotal = mTotal + val(vsGrid.TextMatrix(mRow, 2))
                End If
                RecFin.Close
                mRow = mRow + 1
                Rec.MoveNext
            Wend
        End If
        txtTotal.Text = mTotal
        txtTotal.Locked = True
End Sub

Private Sub Form_Paint()
    'txtTransactionType = ""
    txtTransactionType.Locked = True
    txtDemandNo.Text = frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 3) ', "", frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 4)))
    txtDemandNo.Tag = frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 4)
    txtDemandNo.Locked = True
    If IsDate(frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 0)) Then
        Call FillGrid(CDate(frmDemandInbox.vsGrid.TextMatrix(frmDemandInbox.vsGrid.Row, 0)), IIf(val(txtTransactionType.Tag) = 0, "%", txtTransactionType.Tag))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmReceiptsCounter.ZonalCollection = 0
End Sub

Private Sub txtTransactionType_Change()
 If txtTransactionType.Text = "" Then
        txtTransactionType.Tag = 0
 End If
End Sub
Private Sub vsGrid_DblClick()
  Dim mDemandNo As String
  Dim objdb As New clsDB
  Dim mCnnFin As New ADODB.Connection
  Dim mCnn As New ADODB.Connection
  Dim Rec As New ADODB.Recordset
  Dim mSql As String
  Dim mReceiptNo1 As Variant
  Dim mReceiptNo2 As Variant
  objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaHO
  mSql = "Select numDemandID  from faIDemandTBL where vchDemandNO='" & txtDemandNo.Text & "'"
  Rec.Open mSql, mCnn
  If Not (Rec.BOF And Rec.EOF) Then
    mNumDemandID = Rec!numDemandID
  End If
  Rec.Close
  mSql = ""
  If vsGrid.TextMatrix(vsGrid.Row, 4) <> "" Then
  objdb.CreateNewConnection mCnnFin, enuSourceString.Saankhya
  mSql = "Select Distinct intVoucherNo,tnyCancelFlag,tnyVoucherTypeID from faVouchers"
  mSql = mSql + " Left Join faVoucherChild on faVouchers.intVoucherID=faVoucherChild.intVoucherID"
  mSql = mSql + " where faVoucherChild.numDemandID=" & mNumDemandID & "and faVouchers.intTransactionTypeID=" & frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4)
  Rec.Open mSql, mCnnFin, adOpenStatic, adLockOptimistic
  If Not (Rec.BOF And Rec.EOF) Then
        If vsGrid.TextMatrix(vsGrid.Row, 5) = 0 Then
                While Not Rec.EOF
                       If Rec!tnyVoucherTypeID = 10 Then
                             mReceiptNo1 = Rec!intVoucherNo
                             vsGrid.TextMatrix(vsGrid.Row, 3) = mReceiptNo1
                             vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
                             MsgBox "Receipt Already generated", vbInformation
                             Exit Sub
                       End If
                       Rec.MoveNext
                Wend
        ElseIf vsGrid.TextMatrix(vsGrid.Row, 5) = 1 Then
                  While Not Rec.EOF
                       If Rec!tnyVoucherTypeID = 40 Then
                             mReceiptNo2 = Rec!intVoucherNo
                             vsGrid.TextMatrix(vsGrid.Row, 3) = mReceiptNo2
                               vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &H4DA4F2
                             MsgBox "Jv Already  Saved", vbInformation
                             Exit Sub
                       End If
                       Rec.MoveNext
                Wend
        End If
        Rec.MoveFirst
        '==================================================
     If vsGrid.TextMatrix(mRow, 1) = "Property Tax(Adjusted)" Or vsGrid.TextMatrix(mRow, 1) = "Rent on Building / Stalls(Adjusted)" Or vsGrid.TextMatrix(mRow, 1) = "Rent on Land/Bunks(Adjusted)" Then
                    If mReceiptNo2 <> "" Then
                         vsGrid.TextMatrix(vsGrid.Row, 3) = mReceiptNo2
                         vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
                    Else
                        Call funJournal
                    End If
         Else
                    If mReceiptNo1 <> "" Then
                        vsGrid.TextMatrix(vsGrid.Row, 3) = mReceiptNo1
                        vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, , 4) = &HC0FFC0
                    Else
                       If vsGrid.TextMatrix(vsGrid.Row, 1) = "Property Tax" Or vsGrid.TextMatrix(vsGrid.Row, 1) = "Rent on Building / Stalls" Or vsGrid.TextMatrix(vsGrid.Row, 1) = "Rent on Land/Bunks" Then
                       
                            If Not (Rec.BOF And Rec.EOF) Then
                                While Not Rec.EOF
                                If Rec!tnyVoucherTypeID = 40 Then
                                    frmReceiptsCounter.ZonalCollection = 1
                                    frmReceiptsCounter.mZoneDate = txtDate.Text
                                    mDemandNo = txtDemandNo.Text
                                    frmReceiptsCounter.DataBaseHO = True
                                    frmReceiptsCounter.txtDemandPrefix = Token(mDemandNo, "-")
                                    frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                                    frmReceiptsCounter.txtDate.Text = txtDate.Text
                                    Call frmReceiptsCounter.DisplayTransactionWiseDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                                    frmReceiptsCounter.ZOrder (0)
                                End If
                                Rec.MoveNext
                             Wend
                           Else
                           MsgBox "Please Generate Jv", vbInformation
                           End If
                        Else
                           frmReceiptsCounter.ZonalCollection = 1
                           frmReceiptsCounter.mZoneDate = txtDate.Text
                           mDemandNo = txtDemandNo.Text
                           frmReceiptsCounter.DataBaseHO = True
                           frmReceiptsCounter.txtDemandPrefix = Token(mDemandNo, "-")
                           frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                            Call frmReceiptsCounter.DisplayTransactionWiseDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                           frmReceiptsCounter.ZOrder (0)
                    End If
                    End If
         End If
        
        '===================================================
        
    Else
        
        If vsGrid.TextMatrix(vsGrid.Row, 5) = 1 Then
                  Call funJournal
        Else
              If vsGrid.TextMatrix(vsGrid.Row, 1) = "Property Tax" Or vsGrid.TextMatrix(vsGrid.Row, 1) = "Rent on Building / Stalls" Or vsGrid.TextMatrix(vsGrid.Row, 1) = "Rent on Land/Bunks" Then
                       
                            If Not (Rec.BOF And Rec.EOF) Then
                                While Not Rec.EOF
                                If Rec!tnyVoucherTypeID = 40 Then
                                    frmReceiptsCounter.ZonalCollection = 1
                                    frmReceiptsCounter.mZoneDate = txtDate.Text
                                    mDemandNo = txtDemandNo.Text
                                    frmReceiptsCounter.DataBaseHO = True
                                    frmReceiptsCounter.txtDemandPrefix = Token(mDemandNo, "-")
                                    frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                                    Call frmReceiptsCounter.DisplayTransactionWiseDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                                    frmReceiptsCounter.ZOrder (0)
                                End If
                                Rec.MoveNext
                                Wend
                             Else
                                If vsGrid.TextMatrix(vsGrid.Row + 1, 1) = "Property Tax(Adjusted)" Or vsGrid.TextMatrix(vsGrid.Row + 1, 1) = "Rent on Building / Stalls(Adjusted)" Or vsGrid.TextMatrix(vsGrid.Row + 1, 1) = "Rent on Land/Bunks(Adjusted)" Then
                                    MsgBox "Please Jenerate Jv", vbInformation
                                Else
                                     frmReceiptsCounter.ZonalCollection = 1
                                     frmReceiptsCounter.mZoneDate = txtDate.Text
                                     mDemandNo = txtDemandNo.Text
                                     frmReceiptsCounter.DataBaseHO = True
                                     frmReceiptsCounter.txtDemandPrefix = Token(mDemandNo, "-")
                                     frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                                     Call frmReceiptsCounter.DisplayTransactionWiseDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                                     frmReceiptsCounter.ZOrder (0)
                                End If
                                 Exit Sub
                             End If
                Else
                      
                  frmReceiptsCounter.ZonalCollection = 1
                  frmReceiptsCounter.mZoneDate = txtDate.Text
                  mDemandNo = txtDemandNo.Text
                  frmReceiptsCounter.DataBaseHO = True
                  frmReceiptsCounter.txtDemandPrefix = Token(mDemandNo, "-")
                  frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                  Call frmReceiptsCounter.DisplayTransactionWiseDetails(val(vsGrid.TextMatrix(vsGrid.Row, 4)))
                  frmReceiptsCounter.ZOrder (0)
            End If
        End If
        End If
    End If
End Sub
Private Sub funJournal()
    Dim mjSQL As String
    Dim Recj As New ADODB.Recordset
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim mCnnFin As New ADODB.Connection
    Dim objAc As New clsAccounts
    Dim mRow As Variant
    Dim mSql As String
    mRow = 1
    
    frmJournalEntry.ZonalCollection = 1
    
    frmJournalEntry.optDebit.value = True
    frmJournalEntry.optDebit.Enabled = False
    frmJournalEntry.optCredit.Enabled = False
    If frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4) = 1 Then
        frmJournalEntry.cmbTransactionType.Clear
        frmJournalEntry.cmbTransactionType.AddItem "Property Tax"
        frmJournalEntry.cmbTransactionType.Tag = 1
        frmJournalEntry.cmbTransactionType.ListIndex = 0
        frmJournalEntry.txtAccountHeadCode.Text = gbAcHeadCodeAdvancePTax
        objAc.SetAccountCode gbAcHeadCodeAdvancePTax
        frmJournalEntry.txtAccountHead.Text = objAc.AccountHead
        frmJournalEntry.txtAccountHead.Enabled = False
        frmJournalEntry.txtAccountHeadCode.Enabled = False
        frmJournalEntry.txtNarration.Text = "Property Tax Adjusted   Zonal Collection  " & frmTransactionTypeWiseDemandInbox.txtDate
        frmJournalEntry.txtNarration.Enabled = False
        'frmJournalEntry.
    ElseIf frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4) = 4 Then
        frmJournalEntry.cmbTransactionType.Clear
        frmJournalEntry.cmbTransactionType.AddItem "Rent on Building / Stalls"
        frmJournalEntry.cmbTransactionType.Tag = 4
        frmJournalEntry.cmbTransactionType.ListIndex = 0
    
        frmJournalEntry.txtAccountHeadCode.Text = gbAcHeadCodeAdvanceBuilding
        objAc.SetAccountCode gbAcHeadCodeAdvancePTax
        frmJournalEntry.txtAccountHead.Text = objAc.AccountHead
        frmJournalEntry.txtAccountHead.Enabled = False
        frmJournalEntry.txtAccountHeadCode.Enabled = False
        frmJournalEntry.txtNarration.Text = "Rent on Building / Stalls Adjusted   Zonal Collection  " & frmTransactionTypeWiseDemandInbox.txtDate
        frmJournalEntry.txtNarration.Enabled = False
    ElseIf frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4) = 5 Then
    
      frmJournalEntry.cmbTransactionType.Clear
        frmJournalEntry.cmbTransactionType.AddItem " Rent on Land/Bunks"
        frmJournalEntry.cmbTransactionType.Tag = 5
        frmJournalEntry.cmbTransactionType.ListIndex = 1
    
        frmJournalEntry.txtAccountHeadCode.Text = gbAcHeadCodeAdvanceLand
        objAc.SetAccountCode gbAcHeadCodeAdvancePTax
        frmJournalEntry.txtAccountHead.Text = objAc.AccountHead
        frmJournalEntry.txtAccountHead.Enabled = False
        frmJournalEntry.txtAccountHeadCode.Enabled = False
        frmJournalEntry.txtNarration.Text = "Rent on Land/Bunks Adjusted   Zonal Collection  " & frmTransactionTypeWiseDemandInbox.txtDate
        frmJournalEntry.txtNarration.Enabled = False
    End If
    
   

   objdb.CreateNewConnection mCnnFin, enuSourceString.Saankhya
   mSql = mSql + "Select faTransactionType.intTransactionTypeID,faFunctions.vchFunction,faFunctions.intFunctionID,faFunctionaries.vchFunctionary,faFunctionaries.intFunctionaryID,suSourceOFFund.intSourceFundID,suSourceOFFund.vchSourceFundName from faTransactionType"
   mSql = mSql + " Inner join faFunctions on faFunctions.intFunctionID=faTransactionType.intFunctionID"
   mSql = mSql + " Inner join faFunctionaries on faFunctionaries.intFunctionaryID=faTransactionType.intFunctionaryID"
   mSql = mSql + " Inner Join suSourceofFund on suSourceOFFund.intSourceFundID=faTransactionType.intSourceFundID"
   mSql = mSql + " Where faTransactionType.intTransactionTypeID =" & frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4)
   objdb.CreateNewConnection mCnnFin, enuSourceString.Saankhya
   Rec.Open mSql, mCnnFin, adOpenStatic, adLockOptimistic
   If Not (Rec.BOF And Rec.EOF) Then
        frmJournalEntry.txtFunction.Text = Rec!vchFunction
        frmJournalEntry.txtFunction.Tag = Rec!intFunctionID
        frmJournalEntry.txtFunctionary.Text = Rec!vchFunctionary
        frmJournalEntry.txtFunctionary.Tag = Rec!intFunctionaryID
        frmJournalEntry.txtFund.Text = Rec!vchSourceFundName
        frmJournalEntry.txtFund.Tag = Rec!intSourceFundID
        
        '---------------------------------------
        
        frmJournalEntry.txtFunction.Enabled = False
        frmJournalEntry.txtFunctionary.Enabled = False
        frmJournalEntry.txtFund.Enabled = False
        frmJournalEntry.cmdFunction.Enabled = False
        frmJournalEntry.cmdFunctionary.Enabled = False
        frmJournalEntry.cmdFund.Enabled = False
        '-----------------------------------------
        frmJournalEntry.txtAmount.Enabled = False
        
        frmJournalEntry.cmdNew.Enabled = False
        
        frmJournalEntry.txtVoucherNo.Enabled = False
      
        
   End If
   
    objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaHO
    mjSQL = " Select faIDemandChild.intTransactionTypeID,faIDemandChild.intAccountHeadID as intAccountHeadID ,faIDemandChild.fltAmount as fltAmount from faIDemandTBL"
    mjSQL = mjSQL + " Inner JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
    mjSQL = mjSQL + "  where dtDemandDate='" & Format(CDate(frmTransactionTypeWiseDemandInbox.txtDate.Text), "dd/MMM/yyyy") & "'"
    mjSQL = mjSQL + " And faIDemandChild.intTransactionTypeID =" & frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4) & "and faIDemandChild.tnyStatus=" & 10
    Recj.Open mjSQL, mCnn, adOpenKeyset, adLockOptimistic
    If Not (Recj.BOF And Recj.EOF) Then
    While Not Recj.EOF
               objAc.SetAccountID Recj!intAccountHeadID
                If objAc.AccountHeadID > 0 Then
                      frmJournalEntry.vsGrid.TextMatrix(mRow, 1) = objAc.AccountCode
                      frmJournalEntry.vsGrid.TextMatrix(mRow, 2) = objAc.AccountHead
                      frmJournalEntry.vsGrid.TextMatrix(mRow, 4) = Recj!fltAmount
                 End If
                 mRow = mRow + 1
                 Recj.MoveNext
    Wend
     frmJournalEntry.vsGrid.Enabled = False
    End If
End Sub
Public Property Let ZonalID(mID As Double)
    mZonalID = mID
End Property

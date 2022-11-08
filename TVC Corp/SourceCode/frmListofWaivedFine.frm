VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmListofWaivedFine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Waived Fine From Receipt Counters"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approval"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6675
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12270
      _cx             =   21643
      _cy             =   11774
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      FormatString    =   $"frmListofWaivedFine.frx":0000
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
Attribute VB_Name = "frmListofWaivedFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub cmdApprove_Click()
        If gbUserID = 2 Then
            cmdApprove.Visible = True
        Else
            cmdApprove.Visible = False
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub lblClose_Click()
        FrameWaiveFine.Visible = False
    End Sub
    Private Sub FillGrid()
        Dim mCnn  As New ADODB.Connection
        Dim Rec   As New ADODB.Connection
        Dim objDb As New clsDB
        Dim mSql  As String
        
        'objDb.SetConnection mCnn, enuSourceString.Saankhya
        objDb.SetConnection mCnn
    
        mSql = " SELECT faVouchers.dtDate,faVouchers.intVoucherNo,faTransactionType.vchTransactionType,faFineWaiver.fltActualFine,faFineWaiver.fltChangedFine,"
        mSql = mSql + " faUser.vchUserName , faSeats.chvSeatTitle, faCounters.vchDescription, faVouchers.intVoucherID"
        mSql = mSql + " FROM faVouchers LEFT JOIN  faTransactionType"
        mSql = mSql + " ON faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
        mSql = mSql + " LEFT JOIN faFineWaiver "
        mSql = mSql + " ON faFineWaiver.intTransactionTypeID = faTransactionType.intTransactionTypeID "
        mSql = mSql + " LEFT JOIN faSeats "
        mSql = mSql + " on faSeats.numseatID = faVouchers.numSeatID "
        mSql = mSql + " LEFT JOIN faUser "
        mSql = mSql + " ON faUser.numUserID = faFineWaiver.numUserID "
        mSql = mSql + " LEFT JOIN faCounters "
        mSql = mSql + " ON faCounters.intCounterID = faVouchers.intCounterID "
        
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        vsGrid.Clear 1, 0
        
     If Not (Rec.BOF And Rec.EOF) Then
        vsGrid.Rows = Rec.RecordCount + 1
        vsGrid.Col = 0
        vsGrid.Row = 1
        vsGrid.ColSel = 1
        vsGrid.RowSel = vsGrid.Rows - 1
        vsGrid.Clip = mSql
    End If
    Rec.Close
    End Sub

    Private Sub Form_Load()
        Call FillGrid
        vsGrid.Clear 1, 0
    End Sub

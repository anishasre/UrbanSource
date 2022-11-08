VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmViewSystemJournals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View System Journals"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12720
   Icon            =   "frmViewSystemJournals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm1 
      Caption         =   "Filter Options"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   6840
      Width           =   12615
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAccountHeadCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPaymentOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearchAccountHeads 
         Caption         =   "..."
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtAccountHeads 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdSearchTransaction 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Order"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Voucher Number"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Account Heads"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.TextBox txtToDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9720
      TabIndex        =   3
      Top             =   615
      Width           =   1890
   End
   Begin VB.TextBox txtFromDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   2
      Top             =   600
      Width           =   1890
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   12720
      Top             =   8040
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   12735
      _cx             =   22463
      _cy             =   9975
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   25
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmViewSystemJournals.frx":1CCA
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System Generated Journals"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9360
      TabIndex        =   5
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "From :"
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
      Left            =   6240
      TabIndex        =   4
      Top             =   615
      Width           =   1020
   End
End
Attribute VB_Name = "frmViewSystemJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Sub cmdClear_Click()
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        txtAccountHeadCode.Text = ""
        txtAccountHeads.Text = ""
        txtAccountHeads.Tag = ""
        txtVoucherNo.Text = ""
        txtVoucherNo.Tag = ""
        txtPaymentOrder.Text = ""
        txtPaymentOrder.Tag = ""
    End Sub
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub
    Private Sub cmdSearchAccountHeads_Click()
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads "
        frmSearchAccountHeads.Show vbModal
           If gbSearchID <> -1 Then
               txtAccountHeadCode.Text = Token(gbSearchStr, " ") '
               txtAccountHeads.Text = gbSearchStr
               txtAccountHeads.Tag = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub
    Private Sub cmdSearchTransaction_Click()
        frmSearchTransactionType.StrQuery = " Select vchTransactionType,intTransactionTypeID from faTransactionType where isnull(tnyHidden,0)=0"
        frmSearchTransactionType.Show vbModal
        txtTransactionType.Text = Trim(gbSearchStr)
        txtTransactionType.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub

''    Private Sub cmdGo_Click()
''        Call FillGrid
''    End Sub

    Private Sub Form_Load()
        WindowsXPC1.InitSubClassing
        txtFromDate.Text = DdMmmYy(Date - 31) 'DdMmmYy(gbStartingDate)
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        Call FillGrid
    End Sub

    Private Sub txtFromDate_LostFocus()
        If Not IsDate(txtFromDate.Text) Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub

    Private Sub txtPaymentOrder_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub

    Private Sub txtToDate_LostFocus()
        If Not IsDate(txtToDate.Text) Then
            txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
            txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
'        If txtFromDate.Text <> "" Then
'            Call FillGrid
'        End If
    End Sub
    Private Sub FillGrid()
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecSub      As New ADODB.Recordset
        Dim mRowCnt     As Integer
        
     
        If objDB.SetConnection(mCnn) Then
            
                mSQL = " SELECT faVouchers.dtDate dtDate,faVouchers.intVoucherNo,faVouchers.tnyVoucherGroupID,faAccountHeads.vchAccountHead,faTransactionType.vchTransactionType,faVouchers.intKeyID2,"
                mSQL = mSQL + " faSeats.chvSeatTitle,faUser.vchUserName,faVouchers.intVoucherID,faVouchers.intExternalModuleID,faVouchers.intTransactionTypeID TrTypeID FROM faVouchers"
                mSQL = mSQL + " INNER JOIN faAccountHeads on faAccountHeads.intAccountHeadID=faVouchers.intKeyID1"
                mSQL = mSQL + " INNER JOIN  faTransactionType ON faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID"
                mSQL = mSQL + " INNER JOIN faUser on faUser.numUserID=faVouchers.intUserID"
                mSQL = mSQL + " INNER JOIN faSeats on faSeats.numSeatID=faVouchers.numSeatID"
                mSQL = mSQL + " Where faVouchers.intKeyID2 Is Not Null And faVouchers.tnyVoucherTypeID = 40"
                mSQL = mSQL + " and faVouchers.dtDate BETWEEN '" & txtFromDate.Text & " '  AND '" & txtToDate & " ' "
                
                If txtTransactionType.Text <> "" Then
                    mSQL = mSQL + " and faVouchers.intTransactionTypeID=" & txtTransactionType.Tag & ""
                End If
                If txtAccountHeadCode.Text <> "" Then
                    mSQL = mSQL + " and faVouchers.intKeyID1=" & txtAccountHeads.Tag & ""
                End If
                If txtVoucherNo.Text <> "" Then
                    mSQL = mSQL + " and faVouchers.intVoucherNo=" & Trim(txtVoucherNo.Text) & ""
                End If
                If txtPaymentOrder.Text <> "" Then
                    mSQL = mSQL + " and faVouchers.intKeyID2=" & Trim(txtPaymentOrder.Text) & ""
                End If
                mSQL = mSQL + " ORDER BY faVouchers.dtDate"
                
                Rec.CursorLocation = adUseClient
                Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                mRowCnt = 1
            
                VSGrid.Clear 1, 1
                VSGrid.Rows = 1
                While Not (Rec.EOF Or Rec.BOF)
                    VSGrid.Rows = VSGrid.Rows + 1
                    VSGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
                    VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                    VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    'VSGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                    VSGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                    VSGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                    'VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
                    VSGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    If Left((Rec!intKeyID2), 1) = 5 Then
                        VSGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                        VSGrid.TextMatrix(mRowCnt, 4) = "PO"
                        mSQL = "Select * from faPayOrder where vchPAyOrderNo= " & Rec!intKeyID2 & " "
                        RecSub.Open mSQL, mCnn
                        If Not (RecSub.EOF And RecSub.BOF) Then
                            VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(RecSub!dtPayOrderDate), "", CheckDateInMMM(RecSub!dtPayOrderDate))
                        End If
                        RecSub.Close
                    ElseIf Left((Rec!intKeyID2), 1) = 3 Then
                        VSGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
                        VSGrid.TextMatrix(mRowCnt, 4) = "CV"
                        mSQL = "Select * from faVouchers where intVoucherNo= " & Rec!intKeyID2 & " "
                        RecSub.Open mSQL, mCnn
                        If Not (RecSub.EOF And RecSub.BOF) Then
                            VSGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(RecSub!dtDate), "", CheckDateInMMM(RecSub!dtDate))
                        End If
                        RecSub.Close
                    ElseIf (Rec!TrTypeID) = 1211 Then
                        VSGrid.TextMatrix(mRowCnt, 4) = "Subsidary Cash Book"
                    End If
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    
                Wend
                Rec.Close
                
       End If
    End Sub
    Private Sub Form_Activate()
'        Me.Top = 500
'        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub

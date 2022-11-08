VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterruptedDateEditRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interrupted Receipt Date Edit Request"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   Icon            =   "frmInterruptedDateEditRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve Requests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4710
      TabIndex        =   24
      Top             =   6690
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Frame fraVouchers 
      Caption         =   "   Vouchers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      Left            =   375
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   11115
      Begin VB.TextBox txtBookNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   660
         TabIndex        =   22
         Top             =   5055
         Width           =   1320
      End
      Begin VB.CommandButton cmdSendRequest 
         Caption         =   "Send Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         TabIndex        =   6
         Top             =   5040
         Width           =   2130
      End
      Begin VB.TextBox txtReceiptDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5055
         Width           =   1320
      End
      Begin VB.TextBox txtChangeDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5730
         TabIndex        =   4
         Top             =   5055
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Vouchers"
         Top             =   60
         Width           =   1155
      End
      Begin VB.CommandButton cmdCloseVouchers 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10740
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   $"frmInterruptedDateEditRequest.frx":1CCA
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   15
         Width           =   11115
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridVouchers 
         Height          =   4665
         Left            =   30
         TabIndex        =   9
         Top             =   330
         Width           =   11040
         _cx             =   19473
         _cy             =   8229
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInterruptedDateEditRequest.frx":1D7A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         Begin VB.CheckBox chkSelect 
            Caption         =   "Check1"
            Height          =   195
            Left            =   10425
            TabIndex        =   14
            Top             =   45
            Width           =   180
         End
      End
      Begin MSComCtl2.DTPicker dtpChangeDate 
         Height          =   315
         Left            =   7065
         TabIndex        =   5
         Top             =   5055
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   39697
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
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
         Left            =   210
         TabIndex        =   23
         Top             =   5070
         Width           =   420
      End
      Begin VB.Label lblChangeTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change to "
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
         Left            =   4815
         TabIndex        =   17
         Top             =   5055
         Width           =   945
      End
      Begin VB.Label lblReceiptDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
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
         Left            =   2010
         TabIndex        =   15
         Top             =   5055
         Width           =   1125
      End
   End
   Begin VB.Frame fraRequests 
      Height          =   6705
      Left            =   15
      TabIndex        =   13
      Top             =   -75
      Width           =   11790
      Begin VB.Frame fraSearch 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   2685
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   6420
         Begin VB.ComboBox cmbDate 
            Height          =   315
            Left            =   2385
            TabIndex        =   25
            Text            =   "Combo1"
            Top             =   165
            Width           =   1635
         End
         Begin VB.ComboBox cmbBook 
            Height          =   315
            Left            =   570
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   165
            Width           =   1305
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2325
            TabIndex        =   1
            Top             =   -195
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.CommandButton cmdSearchVouchers 
            Caption         =   "Search Vouchers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4200
            TabIndex        =   3
            Top             =   150
            Width           =   2130
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   3675
            TabIndex        =   2
            Top             =   -195
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62324737
            CurrentDate     =   39697
         End
         Begin VB.Label lblBook 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Book"
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
            Left            =   105
            TabIndex        =   21
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   1980
            TabIndex        =   19
            Top             =   180
            Width           =   405
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   5685
         Left            =   15
         TabIndex        =   8
         Top             =   1005
         Width           =   11730
         _cx             =   20690
         _cy             =   10028
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInterruptedDateEditRequest.frx":1EB8
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label lblRequests 
         BackColor       =   &H0080C0FF&
         Caption         =   "  Requests"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   20
         Top             =   735
         Width           =   11715
      End
   End
End
Attribute VB_Name = "frmInterruptedDateEditRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub FillvsRequestsGrid()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRowCount   As Integer
        
        On Error GoTo Err
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select vchUserName,faCounters.vchDescription[Description],dtReceiptChangeDate,B.dtDate,B.intVoucherNo[StartVoucherNo],C.intVoucherNo[EndVoucherNo],B.intVoucherID[StartVoucherID],C.intVoucherID[EndVoucherID],A.tnyStatus[Status],A.numUserID[UserID],A.intCounterID[CounterID],A.dtRequestDate[RequestDate],A.intBookID "
            mSql = mSql + " From faInterruptedRequests A"
            mSql = mSql + " Left Join faVouchers B On A.intStartVoucherNo = B.intVoucherNo"
            mSql = mSql + " Left Join faVouchers C On A.intEndVoucherNo = C.intVoucherNo"
            mSql = mSql + " Left Join faCounters ON A.intCounterID = faCounters.intCounterID"
            mSql = mSql + " Left Join faUser On A.numUserID = faUser.numUserID"
            mSql = mSql + " Where A.intTypeID=4"
            Rec.Open mSql, mCnn
            While Not Rec.EOF
                vsGrid.AddItem ""
                vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!Description), "", Rec!Description)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!RequestDate), "", Rec!RequestDate)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtReceiptChangeDate), "", Rec!dtReceiptChangeDate)
                vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!StartVoucherNo), "", Rec!StartVoucherNo)
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!EndVoucherNo), "", Rec!EndVoucherNo)
                If Not (IsNull(Rec!Status)) Then
                    If Rec!Status = 1 Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 7) = vbChecked
                        vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 7) = &HC0E0FF
                    Else
                        vsGrid.Cell(flexcpChecked, mRowCount, 7) = vbUnchecked
                        vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 7) = &H80000005
                    End If
                End If
                vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!UserID), "", Rec!UserID)
                vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!CounterID), "", Rec!CounterID)
                vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!StartVoucherID), "", Rec!StartVoucherID)
                vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!EndVoucherID), "", Rec!EndVoucherID)
                vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!Status), "", Rec!Status)
                vsGrid.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!intBookID), "", Rec!intBookID)
                Rec.MoveNext
                mRowCount = mRowCount + 1
            Wend
            Rec.Close
        Else
            MsgBox "Connection to Finance does not exist, Pleas contact your System Administrator", vbInformation
            Exit Sub
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
    
    Private Sub FillvsVouchersGrid()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRowCount   As Integer
        Dim i           As Integer
        
        On Error GoTo Err
        vsGridVouchers.Clear 1, 1
        vsGridVouchers.Rows = 1
        mRowCount = 1
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select intVoucherNo,dtDate,fltAmount,vchTransactionType,faCounters.vchDescription[Description],faVouchers.intUserID[UserID],faVouchers.intCounterID[CounterID],intVoucherID From faVouchers"
            mSql = mSql + " Left Join faCounters ON faCounters.intCounterID=faVouchers.intCounterID"
            mSql = mSql + " Left Join faInterruptedReceiptBooks On faVouchers.intBookNo = faInterruptedReceiptBooks.intBookID"
            mSql = mSql + " Left Join faUser On faVouchers.intUserID = faUser.numUserID"
            mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
            mSql = mSql + " Where IsNull(tnyCancelFlag, 0) = 0"
            mSql = mSql + " And tnyVoucherTypeID IN (10)"
            mSql = mSql + " And tnyVoucherGroupID = 4"
            mSql = mSql + " And faVouchers.intBookNo = " & cmbBook.ItemData(cmbBook.ListIndex)
            'mSQL = mSQL + " And dtDate ='" & CheckDateInMMM(txtDate.Text) & "'"
            mSql = mSql + " And dtDate ='" & CheckDateInMMM(cmbDate.Text) & "'"
            mSql = mSql + " Order By intVoucherNo,dtDate"
            Rec.Open mSql, mCnn
            While Not Rec.EOF
                vsGridVouchers.AddItem ""
                vsGridVouchers.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                vsGridVouchers.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                vsGridVouchers.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsGridVouchers.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsGridVouchers.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!Description), "", Rec!Description)
                vsGridVouchers.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!UserID), "", Rec!UserID)
                vsGridVouchers.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!CounterID), "", Rec!CounterID)
                vsGridVouchers.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                vsGridVouchers.TextMatrix(mRowCount, 9) = 0
                Rec.MoveNext
                mRowCount = mRowCount + 1
            Wend
            Rec.Close
'            vsGridVouchers.AddItem ""
'            vsGridVouchers.RowHidden(mRowCount) = True
            
            For i = vsGridVouchers.Rows - 1 To 1 Step -1
                mSql = "Select * From faInterruptedRequests"
                mSql = mSql + " Where " & vsGridVouchers.TextMatrix(i, 0) & " Between intStartVoucherNo And intEndVoucherNo"
                mSql = mSql + " And tnyStatus = 0"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    vsGridVouchers.RemoveItem (i)
                    'i = i + 1
                    'vsGridVouchers.Rows = vsGridVouchers.Rows - 1
'                    vsGridVouchers.TextMatrix(i, 9) = 1
'                    vsGridVouchers.Cell(flexcpBackColor, i, 0, , 6) = &HC0E0FF
'                    vsGridVouchers.Cell(flexcpChecked, i, 5) = 2
                End If
                Rec.Close
            Next
        Else
            MsgBox "Connection to Finance does not exist, Pleas contact your System Administrator", vbInformation
            Exit Sub
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub chkSelect_Click()
        On Error GoTo Err
        If chkSelect.value = vbChecked Then
           If vsGridVouchers.Rows > 1 Then
               vsGridVouchers.Cell(flexcpChecked, 1, 5, vsGridVouchers.Rows - 1, 5) = True
           End If
        ElseIf chkSelect.value = vbUnchecked Then
           If vsGridVouchers.Rows > 1 Then
               vsGridVouchers.Cell(flexcpChecked, 1, 5, vsGridVouchers.Rows - 1, 5) = False
           End If
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
    
Private Sub cmbBook_Click()
    Dim mSql As String
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mCn As New ADODB.Connection
    If cmbBook.ListIndex > -1 Then
        If cmbBook.ItemData(cmbBook.ListIndex) > 0 Then
            mSql = "Select Distinct dtDate From faVouchers WHERE tnyVoucherGroupID =4 And intBookNo  = " & cmbBook.ItemData(cmbBook.ListIndex) & " Order by dtDate"
            objdb.SetConnection mCn
            Rec.Open mSql, mCn, adOpenForwardOnly, adLockReadOnly
            cmbDate.Clear
            cmbDate.AddItem ""
            While Not Rec.EOF
                cmbDate.AddItem DdMmmYy(Rec!dtDate)
                Rec.MoveNext
            Wend
        End If
    End If
End Sub

    Private Sub cmdApprove_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim i       As Integer
        Dim mStatus As Integer
        
        On Error GoTo Err
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            If (MsgBox("Do you want to Approve these Edit Requests?", vbYesNo) = vbYes) Then
                For i = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpChecked, i, 7) = 1 Then
                        mStatus = 1
                    Else
                        mStatus = 0
                    End If
                    
                    If vsGrid.TextMatrix(i, 12) = 0 Then
                        If mStatus = 1 Then
                            mSql = "Update faInterruptedRequests "
                            mSql = mSql + " Set tnyStatus = " & mStatus
                            mSql = mSql + " Where intStartVoucherNo = " & vsGrid.TextMatrix(i, 5)
                            mSql = mSql + " And intEndVoucherNo = " & vsGrid.TextMatrix(i, 6)
                            mSql = mSql + " And intBookID = " & val(vsGrid.TextMatrix(i, 13))
                            mCnn.Execute mSql
                            
                            mSql = "Update faVouchers"
                            mSql = mSql + " Set dtDate = '" & CheckDateInMMM(vsGrid.TextMatrix(i, 4)) & "'"
                            mSql = mSql + " Where intBookNo =  " & val(vsGrid.TextMatrix(i, 13)) & " And intVoucherNo Between " & vsGrid.TextMatrix(i, 5) & " And " & vsGrid.TextMatrix(i, 6)
                            mCnn.Execute mSql
                            
                            mSql = "Update faTransactions "  'changed by Poornima on 03/Oct/2011
                            mSql = mSql + " Set dtTransactionDate = '" & CheckDateInMMM(vsGrid.TextMatrix(i, 4)) & "'"
                            mSql = mSql + " Where intVoucherID IN (Select intVoucherID  From faVouchers where intBookNo =  " & val(vsGrid.TextMatrix(i, 13)) & " And intVoucherNo Between " & vsGrid.TextMatrix(i, 5) & " And " & vsGrid.TextMatrix(i, 6) & ")"

                            mCnn.Execute mSql
                            MsgBox "Successfully Saved", vbInformation
                            FillvsRequestsGrid
                        End If
                    End If
                Next
            End If
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub cmdCloseVouchers_Click()
        txtReceiptDate.Text = ""
        fraRequests.Enabled = True
        fraVouchers.Visible = False
        FillvsRequestsGrid
    End Sub

    Private Sub cmdSearchVouchers_Click()
        'If Trim(txtDate.Text) = "" Then
        '    MsgBox "Please enter the Date", vbInformation
        '    txtDate.SetFocus
        '    Exit Sub
        'End If
        
        If Not IsDate(cmbDate.Text) Then
            MsgBox "Please Select a Transaction Date", vbInformation
            cmbDate.SetFocus
            Exit Sub
        End If
        
        If cmbBook.ListIndex < 1 Then
            MsgBox "Please select the Book", vbInformation
            cmbBook.SetFocus
            Exit Sub
        End If
        
        If IsDate(cmbDate) Then
            txtReceiptDate.Text = CheckDateInMMM(cmbDate.Text)   'txtDate.Text
        Else
            txtReceiptDate.Text = ""
        End If
        
        txtBookNo.Text = cmbBook.Text
        txtBookNo.Tag = cmbBook.ItemData(cmbBook.ListIndex)
        fraRequests.Enabled = False
        fraVouchers.Visible = True
        FillvsVouchersGrid
    End Sub
    
    Private Sub cmdSendRequest_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mArrIn          As Variant
        Dim mStartVoucherNo As Variant
        Dim mEndVoucherNo   As Variant
        Dim i               As Integer
        Dim mSql            As String
        Dim mDate           As Variant
        
        On Error GoTo Err
        If Trim(txtChangeDate.Text) = "" Then
            MsgBox "Please enter the new date", vbInformation
            txtChangeDate.SetFocus
            Exit Sub
        End If
        For i = 1 To vsGridVouchers.Rows - 1
            If IsEmpty(mStartVoucherNo) Then
                If vsGridVouchers.Cell(flexcpChecked, i, 5) = 1 Then
                    mStartVoucherNo = vsGridVouchers.TextMatrix(i, 0)
                End If
            End If
            If vsGridVouchers.Cell(flexcpChecked, i, 5) = 1 Then
                mEndVoucherNo = vsGridVouchers.TextMatrix(i, 0)
            End If
            
        Next
        'mEndVoucherNo = vsGridVouchers.TextMatrix(vsGridVouchers.Rows - 1, 0)
        If mStartVoucherNo = "" And mEndVoucherNo = "" Then
            MsgBox "Please select the Voucher", vbInformation
            Exit Sub
        End If
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select Min(dtDate) as Date From faVouchers"
            mSql = mSql + " Where intBookNo = " & txtBookNo.Tag
            mSql = mSql + " And dtDate > '" & CheckDateInMMM(txtReceiptDate.Text) & "' "
            mSql = mSql + " And tnyVoucherGroupID = 4"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mDate = IIf(IsNull(Rec!Date), "", Rec!Date)
                If mDate <> "" Then
                    If Not (CDate(mDate) > CDate(txtChangeDate.Text)) Then
                        If CDate(txtChangeDate.Text) > gbTransactionDate Then
                            MsgBox "You can't change the Receipt Date to " & txtChangeDate.Text
                            Exit Sub
                        End If
                    End If
                Else
                    If CDate(txtChangeDate.Text) > gbTransactionDate Then
                        MsgBox "You can't change the Receipt Date to " & txtChangeDate.Text
                        Exit Sub
                    End If
                End If
            End If
            Rec.Close
            For i = 1 To vsGridVouchers.Rows - 2
                If vsGridVouchers.Cell(flexcpChecked, i, 5) = 1 Then
                    mSql = "Select * From faInterruptedRequests"
                    mSql = mSql + " Where " & vsGridVouchers.TextMatrix(i, 0) & " Between intStartVoucherNo And intEndVoucherNo"
                    mSql = mSql + " And tnyStatus = 0"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        MsgBox "You have already sent a request to change the date of Voucher (" & vsGridVouchers.TextMatrix(i, 0) & ")", vbInformation
                        Exit Sub
                    End If
                    Rec.Close
                End If
            Next
            '@intCounterID_1    [int],
            '@numUserID_2   [numeric],
            '@tnyStatus_3   [tinyint] = 1,
            '@dtRequestDate_4 [smalldatetime],
            '@intTypeID_5 int ,
            '@intReasonID_6 Int=  Null,
            '@vchRemarks_7 Varchar(100)=  Null,
            '@intVoucherNo_8  Numeric = Null,
            '@intVoucherID_9 Numeric=Null,
            '@intStartVoucherNo_10 Numeric=Null,
            '@intEndVoucherNo_11 Numeric=Null,
            '@dtReceiptDate_12   smalldatetime=Null,
            '@dtReceiptChangeDate_13 smalldatetime=Null
            mArrIn = Array(gbCounterID, _
                       gbUserID, _
                       0, _
                       Format(gbTransactionDate, "DD/MMM/yyyy"), _
                       4, _
                       Null, _
                       Null, _
                       Null, _
                       Null, _
                       mStartVoucherNo, _
                       mEndVoucherNo, _
                       Format(txtReceiptDate.Text, "DD/MMM/yyyy"), _
                       Format(txtChangeDate.Text, "DD/MMM/yyyy"), _
                       val(txtBookNo.Tag))
        'objdb.ExecuteSP "spSaveInterruptedRequest", mArrIn, , , mCnn, adCmdStoredProc'NOTE:SP CHANGED and THIS MODULE NOT IN USE
        MsgBox "Request sent to Nodal Officer", vbInformation
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub dtpChangeDate_CloseUp()
        txtChangeDate.Text = CheckDateInMMM(dtpChangeDate.value)
    End Sub

    Private Sub dtpDate_CloseUp()
        txtDate.Text = CheckDateInMMM(dtpDate.value)
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim mSql    As String
        cmbDate.Clear
        dtpChangeDate.value = gbTransactionDate
        dtpDate.value = gbTransactionDate
        mSql = "Select intBookNo,intBookID From faInterruptedReceiptBooks Where intCounterID = " & gbCounterID
        PopulateList cmbBook, mSql, , True, True, True, enuSourceString.Saankhya
        If gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            fraSearch.Visible = True
            vsGrid.Left = 15
            vsGrid.Top = 1005
        End If
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
            lblRequests.Left = 15
            lblRequests.Top = 95
            vsGrid.Left = 15
            vsGrid.Top = 360
            cmdApprove.Visible = True
        End If
        FillvsRequestsGrid
    End Sub

    Private Sub txtChangeDate_LostFocus()
        If Trim(txtChangeDate.Text) <> "" Then
            txtChangeDate.Text = CheckDateInMMM(txtChangeDate.Text)
        End If
    End Sub

    Private Sub txtDate_LostFocus()
        If Trim(txtDate.Text) <> "" Then
            txtDate.Text = CheckDateInMMM(txtDate.Text)
        End If
    End Sub

    Private Sub vsGrid_Click()
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
            If vsGrid.Rows > 1 Then
                If vsGrid.Row > 0 Then
                    If vsGrid.Col = 7 Then
                        vsGrid.Editable = flexEDKbdMouse
                    Else
                        vsGrid.Editable = flexEDNone
                    End If
                    If vsGrid.TextMatrix(vsGrid.Row, 12) = 1 Then
                        vsGrid.Cell(flexcpChecked, vsGrid.Row, 7) = 1
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub vsGridVouchers_Click()
        If vsGridVouchers.Rows > 1 Then
            If vsGridVouchers.Row > 0 Then
                If vsGridVouchers.Col = 5 Then
                    vsGridVouchers.Editable = flexEDKbdMouse
                Else
                    vsGridVouchers.Editable = flexEDNone
                End If
            End If
        End If
    End Sub

    Private Sub vsGridVouchers_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        On Error GoTo Err
        Dim mLoop As Long
        Dim mTopFlag As Boolean
        Dim mBottomFlag As Boolean
        If Row > 0 Then
            If vsGridVouchers.Rows > 1 Then
                If (vsGridVouchers.Cell(flexcpChecked, 1, 5) = 1) Then
                    mTopFlag = True
                    mBottomFlag = False
                End If
                If (vsGridVouchers.Cell(flexcpChecked, vsGridVouchers.Rows - 1, 5) = 1) Then
                    mTopFlag = False
                    mBottomFlag = True
                End If
            End If
            
            If (mTopFlag = False And mBottomFlag = False) And (Row <> 1 And Row <> vsGridVouchers.Rows - 1) Then
                Cancel = True
                Exit Sub
            End If
            
            If mTopFlag Then
                For mLoop = 1 To vsGridVouchers.Rows - 1
                    vsGridVouchers.Cell(flexcpChecked, mLoop, 5) = vbChecked
                Next
                For mLoop = Row + 1 To vsGridVouchers.Rows - 1
                    vsGridVouchers.Cell(flexcpChecked, mLoop, 5) = vbUnchecked
                Next
            End If
            
            If mBottomFlag Then
                For mLoop = 1 To Row - 1
                    vsGridVouchers.Cell(flexcpChecked, mLoop, 5) = vbUnchecked
                Next
                For mLoop = Row To vsGridVouchers.Rows - 1
                    vsGridVouchers.Cell(flexcpChecked, mLoop, 5) = vbChecked
                Next
            End If
            
            
            
'            If vsGridVouchers.Cell(flexcpChecked, Row, 5) = 1 Then
'                If (vsGridVouchers.Cell(flexcpChecked, Row + 1, 5) = 1) Then
'                   Cancel = True
'                End If
'            Else
'                If (vsGridVouchers.Cell(flexcpChecked, Row - 1, 5) = 2) Then
'                    For mLoop = 1 To Row - 1
'                        If vsGridVouchers.Cell(flexcpChecked, mLoop, 5) = 1 Then
'                            Cancel = True
'                            Exit For
'                        End If
'                    Next
'                    For mLoop = Row + 1 To vsGridVouchers.Rows - 1
'                        If vsGridVouchers.Cell(flexcpChecked, mLoop - 1, 5) = 1 Then
'                            Cancel = True
'                            Exit For
'                        End If
'                    Next
'                End If
'            End If
        End If
        Exit Sub
Err:
        'MsgBox Err.Description
    End Sub

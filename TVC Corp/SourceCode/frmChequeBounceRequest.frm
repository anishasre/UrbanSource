VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmChequeBounceRequest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Request for Cheque Bounce "
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frmChequeBounceRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "frmChequeBounceRequest.frx":1CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   1920
      Width           =   480
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   3525
         Left            =   150
         TabIndex        =   33
         Top             =   735
         Width           =   2235
      End
   End
   Begin VB.CheckBox chkFalse 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search From Vouchers"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   31
      Top             =   2010
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearchCheque 
      Caption         =   "Find Cheque From Bank Stmt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7890
      TabIndex        =   17
      Top             =   2070
      Width           =   2745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   10485
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5790
         TabIndex        =   26
         Top             =   1200
         Width           =   1260
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4260
         TabIndex        =   25
         Top             =   1200
         Width           =   1260
      End
      Begin VB.TextBox txtChequeTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   870
         Width           =   1395
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
         Left            =   7050
         TabIndex        =   22
         Top             =   210
         Width           =   270
      End
      Begin VB.TextBox txtBank 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   210
         Width           =   3990
      End
      Begin VB.TextBox txtBankEntryDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox txtChequeDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox txtPerticulars 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   570
         Width           =   5655
      End
      Begin VB.TextBox txtInstrumentNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1395
         TabIndex        =   11
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label lblChekFindStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------Unable to Locate the Cheque in Bank Scroll... Search the Cheque in Voucher Details--------"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   -150
         TabIndex        =   29
         Top             =   1470
         Visible         =   0   'False
         Width           =   10905
      End
      Begin VB.Label lblDatePeriod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Period"
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
         Left            =   3240
         TabIndex        =   28
         Top             =   1215
         Width           =   975
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
         Left            =   5580
         TabIndex        =   27
         Top             =   1215
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Left            =   7890
         TabIndex        =   24
         Top             =   870
         Width           =   1035
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
         Left            =   2550
         TabIndex        =   21
         Top             =   210
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Entry Date"
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
         Left            =   7530
         TabIndex        =   19
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Date"
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
         Left            =   7830
         TabIndex        =   16
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label lblBankDrawnFrom 
         BackStyle       =   0  'Transparent
         Caption         =   "Perticulars / Bank Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   14
         Top             =   585
         Width           =   1230
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
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   1245
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   7680
      Top             =   6840
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtSeat 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6510
      Width           =   2640
   End
   Begin VB.CommandButton cmdSeat 
      Caption         =   "..."
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   6495
      Width           =   300
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5940
      Width           =   3000
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5865
      Width           =   1635
   End
   Begin VB.CommandButton cmdRequest 
      Caption         =   "Request for Reverse Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      TabIndex        =   5
      Top             =   6270
      Width           =   2445
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search Voucher"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   0
      Top             =   2070
      Width           =   2535
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3120
      Left            =   150
      TabIndex        =   1
      Top             =   2580
      Width           =   10170
      _cx             =   17939
      _cy             =   5503
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChequeBounceRequest.frx":1FD4
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VSFlex8LCtl.VSFlexGrid vsBankStmt 
      Height          =   3345
      Left            =   0
      TabIndex        =   30
      Top             =   2580
      Visible         =   0   'False
      Width           =   10545
      _cx             =   18600
      _cy             =   5900
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
      FormatString    =   $"frmChequeBounceRequest.frx":2127
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Identify the Cheque No -Tick the CheckBox to Search Cheque in Voucher deatils"
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
      Left            =   420
      TabIndex        =   34
      Top             =   1920
      Width           =   2955
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarded Seat "
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
      Left            =   150
      TabIndex        =   9
      Top             =   6540
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   720
      TabIndex        =   8
      Top             =   5880
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   6450
      TabIndex        =   7
      Top             =   5865
      Width           =   1185
   End
End
Attribute VB_Name = "frmChequeBounceRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private intChequeIdentifyStatus     ' 1 = Identified ; 0 = Not Identified   '
    
    Private Function RequestValidation() As Boolean
        On Error GoTo Err:
            Dim mRowCnt As Integer
            Dim mCnt As Integer
            Dim mDiffer As Boolean
            Dim mCheckCnt As Integer
            
            If txtSeat.Tag = "" Then
                MsgBox "Please Select Seat", vbInformation
                cmdSeat.SetFocus
                RequestValidation = False
                Exit Function
            End If
            
            If txtRemarks.Text = "" Then
                MsgBox "Please Enter Remarks", vbInformation
                txtRemarks.SetFocus
                RequestValidation = False
                Exit Function
            End If
            mDiffer = False
            mCheckCnt = 0
                For mRowCnt = 1 To vsGrid.Rows - 1
                    For mCnt = 1 To vsGrid.Rows - 1
                        If vsGrid.TextMatrix(mCnt, 2) = "" Then GoTo lp:
                        If vsGrid.TextMatrix(mRowCnt, 2) = "" Then Exit For
                        If vsGrid.TextMatrix(mRowCnt, 2) <> vsGrid.TextMatrix(mCnt, 2) Then
                            mDiffer = True
                            GoTo diff:
                        End If
                    Next
lp:                     If vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                        mCheckCnt = mCheckCnt + 1
                    End If
                Next
            
            
diff:       If mDiffer Then
                MsgBox "Please Select any one Check Number from the Grid by Double Clicking", vbInformation
                vsGrid.SetFocus
                RequestValidation = False
                Exit Function
            End If
            
            If mCheckCnt = 0 Then
                MsgBox "Please Select any One Check Number From the Grid", vbInformation
                vsGrid.SetFocus
                RequestValidation = False
                Exit Function
            End If
            
            RequestValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub chkFalse_Click()
        If chkFalse.Value = vbChecked Then
            cmdSearch.Visible = True
        End If
    End Sub
    Private Sub cmdRequest_Click()
          On Error GoTo Err:
            Dim objDB       As New clsDb
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim arrIn       As Variant
            Dim arrOut      As Variant
            Dim mRequestID  As Integer
            Dim mSql        As String
            Dim mRowCnt     As Integer
            
            If RequestValidation = False Then Exit Sub
            
            If objDB.SetConnection(mCnn) Then
                arrIn = Array(-1, _
                            gbTransactionDate, _
                            Null, _
                            10, _
                            1, _
                            Trim(txtRemarks.Text), _
                            gbUserID, _
                            gbSeatID, _
                            Null, _
                            Null, _
                            txtSeat.Tag, _
                            gbFinancialYearID, _
                            0)
        
                objDB.ExecuteSP "spSaveReverseEntry", arrIn, arrOut, , mCnn, adCmdStoredProc
                
                If Not IsNumeric(arrOut) Then
                    mRequestID = arrOut(0, 0)
                End If
                
                For mRowCnt = 1 To vsGrid.Rows - 1
                    arrIn = ""
                    If vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                        arrIn = Array(mRequestID, Val(vsGrid.TextMatrix(mRowCnt, 7)))
                        objDB.ExecuteSP "spSaveReverseEntryChild", arrIn, , , mCnn, adCmdStoredProc
                    End If
                Next
                
                MsgBox "Reverse Entry requested to Higher Authority", vbInformation
                cmdRequest.Enabled = False
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearch_Click()
        Call SearchCheque
        cmdRequest.Enabled = True
    End Sub

    Private Sub cmdSearchBank_Click()
        On Error GoTo Err:
            Dim mSql As String
            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & 2
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.Show vbModal
            txtBank.Text = gbSearchStr
            txtBank.Tag = gbSearchID
            txtBank.SetFocus
            gbSearchID = -1
            gbSearchStr = ""
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function GetDrawnFromBank() As String
        On Error GoTo Err:
        
            Dim mLetter As String
            Dim mWord   As String
            Dim mLength As Integer
            Dim mStart  As Integer
            
            mWord = ""
            mLength = Len(txtPerticulars.Text)
            For mStart = 1 To mLength
                If mID(txtPerticulars.Text, mStart, 1) <> " " Then
                    mWord = mWord + mID(txtPerticulars.Text, mStart, 1) + "%"
                End If
            Next
            GetDrawnFromBank = mWord
        Exit Function
Err:
            MsgBox (Error$)
    End Function
    
    Private Function SearchCheque(Optional mString As String) As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDB As New clsDb
            Dim mRowCnt As Integer
            
            Dim mDrawnBank As String
            Dim mTotalAmt As Variant
            
            vsGrid.Clear 1, 1
            
            If txtFromDate.Text <> "" Then
                If txtToDate.Text = "" Then
                    txtToDate.Text = txtFromDate.Text
                End If
            Else
                If txtToDate <> "" Then
                    txtFromDate.Text = txtToDate.Text
                End If
            End If
            
            
            If txtPerticulars.Text = "" Then
                mDrawnBank = "%"
            Else
                mDrawnBank = GetDrawnFromBank
            End If
            
            mTotalAmt = 0
            
            If objDB.SetConnection(mCnn) Then
                mSql = "Select *, faVouchers.intVoucherID as VoucherID  from faVouchers "
                mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID "
                
                If ChequeIdentifyStatus = 0 Then
                    mSql = mSql + " Where intInstrumentTypeID = 5 and vchInstrumentNo Like '%" & Trim(txtInstrumentNo.Text) & "%' "
                    mSql = mSql + " and ( vchBank Like '" & mDrawnBank & "' or vchBank is null)"
                    If Val(txtBank.Tag) <> 0 Then
                        mSql = mSql + " and intKEYID1 = " & Val(txtBank.Tag)
                    End If
                    If txtFromDate.Text <> "" And txtToDate.Text <> "" Then
                        mSql = mSql + " and dtDate between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "'"
                    End If
                    If mString <> "" Then
                        mSql = mString
                    End If
                Else
                    mSql = mSql + " Where intInstrumentTypeID = 5 and vchInstrumentNo Like '%" & Trim(txtInstrumentNo.Text) & "%' "
                End If
                '==========================='
                '   Needs Verfication-Cijith'
                If mString <> "" Then
                    mSql = mString
                End If
                '==========================='
                Rec.Open mSql, mCnn
                vsGrid.Clear 1, 1
                vsGrid.Rows = 2
                mRowCnt = 1
                If (Rec.EOF Or Rec.BOF) Then
                    lblChekFindStatus.Visible = True
                    lblChekFindStatus.Caption = "------Unable to Find the Cheque in Voucher Details------"
                Else
                    lblChekFindStatus.Visible = False
                End If
                
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked
                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    mTotalAmt = mTotalAmt + Val(vsGrid.TextMatrix(mRowCnt, 5))
                    mRowCnt = mRowCnt + 1
                    vsGrid.Rows = vsGrid.Rows + 1
                    Rec.MoveNext
                Wend
                
                txtTotal.Text = Format(mTotalAmt, "0.00")
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub FormInitialize()
        txtBank.Text = ""
        txtBankEntryDate.Text = ""
        txtChequeDate.Text = ""
        txtChequeTotal.Text = ""
        txtInstrumentNo.Text = ""
        txtPerticulars.Text = ""
        txtRemarks.Text = ""
        txtSeat.Text = ""
        txtTotal.Text = ""
        vsGrid.Clear 1, 1
    End Sub
    
    Private Sub cmdSearchCheque_Click()
'        frmSearchDishonoredCheque.Show vbModal
        Call FillBankStmtGrid
''''        If Val(txtBank.Tag) <> 0 Then
''''            Dim objAcc As New clsAccounts
''''            objAcc.SetAccounts (Val(txtBank.Tag))
''''            txtBank.Text = objAcc.AccountHead
''''        End If
''''        If ChequeIdentifyStatus = 0 Then
''''            lblDatePeriod.Visible = True
''''            lblTo.Visible = True
''''            txtFromDate.Visible = True
''''            txtToDate.Visible = True
''''            cmdSearch.Enabled = True
''''
''''            txtToDate.Text = CheckDateInMMM(Date)
''''            txtFromDate.Text = CheckDateInMMM(DateAdd("m", -1, Date))
''''
''''            lblChekFindStatus.Visible = True
''''            lblChekFindStatus.Caption = "--------Unable to Locate the Cheque in Bank Scroll... Search the Cheque in Voucher Details--------"
''''        ElseIf ChequeIdentifyStatus = 1 Then
''''            lblDatePeriod.Visible = False
''''            lblTo.Visible = False
''''            txtFromDate.Visible = False
''''            txtToDate.Visible = False
''''
''''            cmdSearchBank.Visible = False
''''            cmdSearch.Enabled = True
''''
''''            lblChekFindStatus.Visible = True
''''            lblChekFindStatus.Caption = "--------Cheque Identified in Bank Scroll, Search the Voucher Entries for the Corresponding Cheque--------"
''''        Else
''''            cmdSearch.Enabled = False
''''            lblChekFindStatus.Visible = True
''''            lblChekFindStatus.Caption = "--------Please Select the Cheque from Bank Scroll--------"
''''        End If
    End Sub

    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID = -1 Then
            Exit Sub
        Else
            txtSeat.Text = gbSearchStr
            txtSeat.Tag = gbSearchID
        End If
    End Sub
    
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        cmdSearch.Enabled = False
        chkFalse.Visible = False
        txtFromDate.Text = CheckDateInMMM(DateAdd("m", -1, Date))
        txtToDate.Text = CheckDateInMMM(Date)
    End Sub
    


    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub

    Private Sub txtFromDate_LostFocus()
        If txtFromDate.Text <> "" Then
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    
    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub

    Private Sub txtToDate_LostFocus()
        If txtToDate.Text <> "" Then
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        End If
    End Sub

    Private Sub vsBankStmt_DblClick()
         On Error GoTo Err:
            If vsBankStmt.TextMatrix(vsBankStmt.Row, 0) = "" Then Exit Sub
                txtInstrumentNo.Text = vsBankStmt.TextMatrix(vsBankStmt.Row, 4)
                txtBank.Tag = vsBankStmt.TextMatrix(vsBankStmt.Row, 1)
                txtPerticulars.Text = vsBankStmt.TextMatrix(vsBankStmt.Row, 5)
                txtChequeDate.Text = vsBankStmt.TextMatrix(vsBankStmt.Row, 3)
                txtBankEntryDate.Text = vsBankStmt.TextMatrix(vsBankStmt.Row, 2)
                txtChequeTotal.Text = vsBankStmt.TextMatrix(vsBankStmt.Row, 6)
                ChequeIdentifyStatus = 1
                vsBankStmt.Visible = False
                vsGrid.Visible = True
                chkFalse.Visible = True
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub vsGrid_Click()
        On Error GoTo Err:
            Dim mRowCnt As Integer
            Dim mTotal As Double
            mTotal = 0
            For mRowCnt = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mRowCnt, 6) = vbChecked Then
                    mTotal = mTotal + vsGrid.TextMatrix(mRowCnt, 5)
                End If
            Next
            txtTotal.Text = Format(mTotal, "0.00")
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub vsGrid_DblClick()
        'If ChequeIdentifyStatus = 0 Then
            Dim mSql As String
            mSql = "Select *, faVouchers.intVoucherID as VoucherID  from faVouchers "
            mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID "
            mSql = mSql + " Where intInstrumentTypeID = 5 and vchInstrumentNo = '" & Trim(vsGrid.TextMatrix(vsGrid.Row, 2)) & "'"
            Call SearchCheque(mSql)
        'End If
    End Sub
    
    Public Property Let ChequeIdentifyStatus(mData As Integer)
        intChequeIdentifyStatus = mData
    End Property

    Public Property Get ChequeIdentifyStatus() As Integer
        ChequeIdentifyStatus = intChequeIdentifyStatus
    End Property
    Private Sub FillBankStmtGrid()
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objDB As New clsDb
            Dim mRowCnt As Integer
            Dim mAmt As Double
            vsBankStmt.Visible = True
            vsGrid.Visible = False
            If objDB.SetConnection(mCnn) Then
                mSql = "Select intReconciliationID, intBankAccountHeadID, dtBankEntryDate, dtChequeDate, "
                mSql = mSql + " vchChequeNo, vchParticulars, fltDrAmount, fltCrAmount"
                mSql = mSql + " from faBankReconciliationEntries "
                mSql = mSql + " Where dtChequeDate Between '" & CheckDateInMMM(txtFromDate.Text) & "' and '" & CheckDateInMMM(txtToDate.Text) & "'"
                If txtInstrumentNo.Text <> "" Then
                    mSql = mSql + " and vchChequeNo Like '%" & txtInstrumentNo.Text & "%'"
                End If
                If Val(txtBank.Tag) <> 0 Then
                    mSql = mSql + " and intBankAccountHeadID = " & Val(txtBank.Tag)
                End If
                
                Rec.Open mSql, mCnn, adOpenStatic, adLockPessimistic
                mRowCnt = 1
                vsGrid.Rows = 2
                If Rec.RecordCount <> 0 Then
                    While Not (Rec.EOF Or Rec.BOF)
                        vsBankStmt.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                        vsBankStmt.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!intBankAccountHeadID), "", Rec!intBankAccountHeadID)
                        vsBankStmt.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                        vsBankStmt.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!dtChequeDate), "", Rec!dtChequeDate)
                        vsBankStmt.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                        vsBankStmt.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                        If IsNull(Rec!fltDrAmount) Then
                            mAmt = Rec!fltCrAmount
                        Else
                            mAmt = Rec!fltDrAmount
                        End If
                        vsBankStmt.TextMatrix(mRowCnt, 6) = mAmt
                        Rec.MoveNext
                        mRowCnt = mRowCnt + 1
                        vsBankStmt.Rows = vsBankStmt.Rows + 1
                    Wend
                Else
                    If MsgBox(" Do you Want to Search the Cheque in Bank Scroll Again", vbYesNo) = vbYes Then
                        lblChekFindStatus.Visible = False
                        chkFalse.Visible = False
                    Else
                        lblChekFindStatus.Visible = True
                        lblChekFindStatus.Caption = "--Unable to Locate the Cheque in Bank Scroll...Change your Search Criteria/Check in Vouchers--"
                        ChequeIdentifyStatus = 0
                        chkFalse.Visible = True
                    End If
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your System Administrator"
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    

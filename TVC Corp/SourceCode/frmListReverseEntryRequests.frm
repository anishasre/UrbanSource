VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListReverseEntryRequests 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Entry List"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14985
   Icon            =   "frmListReverseEntryRequests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14985
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14925
      TabIndex        =   1
      Top             =   6960
      Width           =   14985
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   14
         Top             =   405
         Width           =   1095
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC2 
         Left            =   14805
         Top             =   720
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   4
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         Left            =   8190
         TabIndex        =   12
         Top             =   180
         Width           =   1590
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Request"
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
         Left            =   4905
         TabIndex        =   8
         Top             =   180
         Width           =   1590
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel Request"
         Enabled         =   0   'False
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
         Left            =   6555
         TabIndex        =   7
         Top             =   180
         Width           =   1590
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   14805
         Top             =   1755
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.Label lblRev 
         BackColor       =   &H009696FF&
         Caption         =   "Reversed"
         Height          =   195
         Left            =   14175
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFirst 
         BackColor       =   &H00FFC8FF&
         Caption         =   "First Level Approval"
         Height          =   195
         Left            =   13455
         TabIndex        =   10
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.Frame fmeSearch 
         Height          =   465
         Left            =   3555
         TabIndex        =   15
         Top             =   0
         Width           =   3390
         Begin VB.OptionButton optRev 
            BackColor       =   &H009696FF&
            Caption         =   "Reversed"
            Height          =   195
            Left            =   2340
            TabIndex        =   19
            Top             =   180
            Width           =   1005
         End
         Begin VB.OptionButton optFirst 
            BackColor       =   &H00FFC8FF&
            Caption         =   "First"
            Height          =   195
            Left            =   1575
            TabIndex        =   18
            Top             =   180
            Width           =   690
         End
         Begin VB.OptionButton optRequested 
            Caption         =   "Request"
            Height          =   195
            Left            =   630
            TabIndex        =   17
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.TextBox txtDateFrom 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   765
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "01-04-2009"
         Top             =   90
         Width           =   1230
      End
      Begin VB.TextBox txtDateTo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2265
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "01-04-2009"
         Top             =   90
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   13815
         Picture         =   "frmListReverseEntryRequests.frx":1CCA
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1110
      End
      Begin VB.Label lblRevStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmListReverseEntryRequests.frx":2FFD
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7200
         TabIndex        =   9
         Top             =   45
         Width           =   6570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   135
         Width           =   405
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridForCheque 
      Height          =   6330
      Left            =   0
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   12915
      _cx             =   22781
      _cy             =   11165
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmListReverseEntryRequests.frx":3091
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
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6300
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   14985
      _cx             =   26432
      _cy             =   11112
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
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmListReverseEntryRequests.frx":31A6
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
End
Attribute VB_Name = "frmListReverseEntryRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


    Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
    Const GWL_HWNDPARENT = (-8)
'    Dim frm As frmReceiptsCounter
    Dim oldOwner As Long
    
    Dim aryIn               As Variant 'To pass parameter to another form
    Private mLoadMode       As Integer ' 1 = Normal ; 2 = Check Return'
    Public mReceiptMode     As Boolean '1 =Through Receipt Screen
    Private mReceiptID      As Variant 'set from Receipt form
    
''       To vrify demand is Saved or not 1=Saved   set this variable from DemandInterface Form
''      Public mRevDemand         As Boolean
''      Public mDemandNo           As Variant
    Private Sub FillGrid()
        On Error GoTo err:
            Dim objDB       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim Recv         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim mSQL        As String
            Dim mRowCount   As Integer
            Dim mCnt        As Integer
            Dim mInstrNo    As String
            Dim mVoucherNo  As String
            Dim mAmt        As String
            If Not (IsDate(txtDateFrom.Text)) Then
                MsgBox "Please Enter Valid Date", vbInformation
                txtDateFrom.SetFocus
                Exit Sub
            End If
            If Not (IsDate(txtDateTo.Text)) Then
                MsgBox "Please Enter Valid Date", vbInformation
                txtDateTo.SetFocus
                Exit Sub
            End If
            If objDB.SetConnection(mCnn) Then
                If LoadMode = 1 Then
                    mSQL = "Select intVoucherNo,dtDate,fltAmount, faReasons.vchReason,vchUserName,chvSeatTitle,faReverseEntry.tnyStatus,faVouchers.intVoucherID,faReverseEntry.intRequestID, faVouchers.vchInstrumentNo "
                    mSQL = mSQL + ",Replace (faReverseEntry.vchRemarks,char(13),' ') vchRemarks,faReasons.intReasonID,faReverseEntry.numDemandNo,faReverseEntry.tnyStatus,faReasons.intCategory"
                    mSQL = mSQL + " From faReverseEntry Inner join faReverseEntryChild"
                    mSQL = mSQL + " ON faReverseEntryChild.intRequestID=faReverseEntry.intRequestID"
                    mSQL = mSQL + " Left Join faReasons On faReasons.intReasonID = faReverseEntry.intReasonID "
                    mSQL = mSQL + " Inner Join faUser On faUser.numUserID=faReverseEntry.numRequestedUserid"
                    mSQL = mSQL + " Inner Join faVouchers On faVouchers.intVoucherID=faReverseEntryChild.intVoucherID"
                    mSQL = mSQL + " Left JOIN faSeats ON faReverseEntry.numRequestedSeatID = faSeats.numSeatID "
                    mSQL = mSQL + " Where faReverseEntry.tnyStatus <> 4 and dtRequestDate between '" & txtDateFrom.Text & "' and '" & txtDateTo.Text & "'"
                    If optRequested.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 0"
                    ElseIf optFirst.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 1"
                    ElseIf optRev.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 2"
                    End If
                    mSQL = mSQL + " and faReverseEntry.intReasonID <> 500 Order By faReverseEntry.intRequestID "
                    Rec.Open mSQL, mCnn, adOpenStatic, adLockPessimistic
                    vsGrid.Clear 1, 1
                    If Not (Rec.EOF Or Rec.BOF) Then
                        vsGrid.Rows = Rec.RecordCount + 1
                        vsGrid.Col = 0
                        vsGrid.Row = 1
                        vsGrid.ColSel = 16
                        vsGrid.RowSel = vsGrid.Rows - 1
                        mSQL = Rec.GetString(, , vbTab, Chr(13))
                        vsGrid.Clip = mSQL
                        vsGrid.Row = 1
                        vsGrid.Col = 0
                        vsGrid.CellBackColor = &HE0E0E0
                    End If
                    Rec.Close
                ElseIf LoadMode = 2 Then
                    mSQL = " Select  faVouchers.vchInstrumentNo, Sum(faVouchers.fltAmount) as Amount , faReasons.vchReason,vchUserName,chvSeatTitle, "
                    mSQL = mSQL + " faReverseEntry.tnyStatus,faReverseEntry.intRequestID  "
                    mSQL = mSQL + ",Replace (faReverseEntry.vchRemarks,char(13),' ') vchRemarks,faReverseEntry.tnyStatus"
                    mSQL = mSQL + " From faReverseEntry "
                    mSQL = mSQL + " Inner join faReverseEntryChild ON faReverseEntryChild.intRequestID=faReverseEntry.intRequestID  "
                    mSQL = mSQL + " Left Join faReasons On faReasons.intReasonID = faReverseEntry.intReasonID  "
                    mSQL = mSQL + " Inner Join faUser On faUser.numUserID=faReverseEntry.numRequestedUserid  "
                    mSQL = mSQL + " Inner Join faVouchers On faVouchers.intVoucherID=faReverseEntryChild.intVoucherID  "
                    mSQL = mSQL + " Left JOIN faSeats ON faReverseEntry.numForwardedSeatID = faSeats.numSeatID "
                    mSQL = mSQL + " Where faReverseEntry.tnyStatus <> 4 and dtRequestDate between '" & txtDateFrom.Text & "' and '" & txtDateTo.Text & "'"
                    If optRequested.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 0"
                    ElseIf optFirst.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 1"
                    ElseIf optRev.Value = True Then
                        mSQL = mSQL + " And faReverseEntry.tnyStatus = 2"
                    End If
                    mSQL = mSQL + "     and faReverseEntry.intReasonID = 500  "
                    mSQL = mSQL + " Group By faReasons.vchReason,vchUserName,chvSeatTitle, "
                    mSQL = mSQL + "     faReverseEntry.tnyStatus,faReverseEntry.intRequestID, faVouchers.vchInstrumentNo,faReverseEntry.vchRemarks "
                    mSQL = mSQL + " Order By faReverseEntry.intRequestID "
                    Rec.Open mSQL, mCnn, adOpenStatic, adLockPessimistic
                    vsGridForCheque.Clear 1, 1
                    If Not (Rec.EOF Or Rec.BOF) Then
                        vsGridForCheque.Rows = Rec.RecordCount + 1
                        vsGridForCheque.Col = 0
                        vsGridForCheque.Row = 1
                        vsGridForCheque.ColSel = 8
                        vsGridForCheque.RowSel = vsGridForCheque.Rows - 1
                        mSQL = Rec.GetString(, , vbTab, Chr(13))
                        vsGridForCheque.Clip = mSQL
                        vsGridForCheque.Row = 1
                        vsGridForCheque.Col = 0
                        vsGridForCheque.Visible = True
                    End If
                    Rec.Close
                End If
                For mCnt = 1 To vsGridForCheque.Rows - 1
                     vsGridForCheque.Cell(flexcpBackColor, mCnt, 0, mCnt, 8) = &HFFE0FF = &HE0E0E0
                    If vsGridForCheque.TextMatrix(mCnt, 7) <> "" Then
                        If vsGridForCheque.TextMatrix(mCnt, 8) = 2 Then
                            vsGridForCheque.Cell(flexcpBackColor, mCnt, 0, mCnt, 8) = &H9696FF
                        End If
                    End If
                Next
                For mCnt = 1 To vsGrid.Rows - 1
                    vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 16) = &HFFE0FF = &HE0E0E0
                    If vsGrid.TextMatrix(mCnt, 14) <> "" Then
                        If val(vsGrid.TextMatrix(mCnt, 13)) = 1 Then
                            vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 16) = &HFFE0FF
                        ElseIf val(vsGrid.TextMatrix(mCnt, 13)) = 2 Then
                            vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 16) = &H9696FF
                            If vsGrid.TextMatrix(mCnt, 14) = 0 Or vsGrid.TextMatrix(mCnt, 14) = 3 Then
                                vsGrid.TextMatrix(mCnt, 15) = "No Receipt"
                            End If
                            mSQL = ""
                            mSQL = "Select intVoucherNo,tnyVoucherTypeID From faVouchers Where numLinkKeyID=" & vsGrid.TextMatrix(mCnt, 7)
                            Recv.Open mSQL, mCnn
                            If Not (Recv.EOF) Then
                                While Not (Recv.EOF Or Recv.BOF)
                                    If Recv!tnyVoucherTypeID = 10 Then
                                        vsGrid.TextMatrix(mCnt, 15) = IIf(IsNull(Recv!intVoucherNo), "", Recv!intVoucherNo)
                                    Else
                                        vsGrid.TextMatrix(mCnt, 16) = IIf(IsNull(Recv!intVoucherNo), "", Recv!intVoucherNo)
                                    End If
                                    Recv.MoveNext
                                Wend
                            End If
                            Recv.Close
                        End If
                    End If
                Next
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdCancel_Click()
        Dim objDB   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim mSQL    As String
        Dim mRequestID  As Integer
        Dim mStatus     As Integer
            If mLoadMode = 2 Then
                If vsGridForCheque.TextMatrix(vsGridForCheque.Row, 1) <> "" Then
                    If gbSeatName = vsGridForCheque.TextMatrix(vsGridForCheque.Row, 4) Then
                            mRequestID = vsGridForCheque.TextMatrix(vsGridForCheque.Row, 6)
                    Else
                        MsgBox "This Request not Generated in this Login (Seat)", vbInformation
                        Exit Sub
                    End If
                Else
                    MsgBox "Data does not Exists To Cancel", vbInformation
                    Exit Sub
                End If
                mStatus = CheckReverseRequestExist(vsGridForCheque.TextMatrix(vsGridForCheque.Row, 8))
            Else
                If vsGrid.Row = -1 Then
                    Exit Sub
                Else
                    If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
                        If gbSeatName = vsGrid.TextMatrix(vsGrid.Row, 5) Then
                                mRequestID = vsGrid.TextMatrix(vsGrid.Row, 8)
                        Else
                            MsgBox "This Request not Generated in this Login (Seat)", vbInformation
                            Exit Sub
                        End If
                    Else
                        MsgBox "Data does not Exists To Cancel", vbInformation
                        Exit Sub
                    End If
                End If
                mStatus = CheckReverseRequestExist(vsGrid.TextMatrix(vsGrid.Row, 7))
            End If
            If mStatus = 1 Then
                MsgBox "Approved Request can't be allowed to Cancel", vbInformation
                Exit Sub
            ElseIf mStatus = 2 Then
                MsgBox "This Request Already Reversed", vbInformation
                Exit Sub
            End If
            If MsgBox("Are you Sure To Proceed", vbCritical + vbYesNo) = vbYes Then
                objDB.SetConnection mCnn
                mSQL = "Update faReverseEntry Set tnyStatus=4 Where intRequestID=" & mRequestID
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                Exit Sub
            End If
            Call FillGrid
    End Sub
    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdEdit_Click()
        If vsGrid.Row = -1 Then Exit Sub
    End Sub

    Private Sub cmdNew_Click()
''        Dim mCrl As Control
''        Dim mStatus As Integer
''        If LoadMode = 1 Then
''            If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupChiefCashier Then
''                cmdNew.Visible = True
''                If RequsetValidation = False Then Exit Sub
''                mStatus = CheckReverseRequestExist(val(txtVoucherNo.Tag))
''                If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Then
''                    MsgBox "Request Already Exists", vbInformation
''                    Exit Sub
''                End If
''                    aryIn = Array(val(txtVoucherNo.Tag), val(txtReason.Tag), txtRemarks.Text, val(txtSeat.Tag))
''                    If txtVoucherNo.Tag = "" Then
''                        lblMsgBox.Visible = True
''                        lblMsgBox.Caption = "Please Select A Voucher to do Verification"
''                        Exit Sub
''                    End If
''                    If lblVoucherType.Tag = 10 Then
''                        If cmdSearchReason.Tag = 1 Then
''                            Select Case val(txtReason.Tag)
''                                Case 4  'Amount
''                                    With frmDemandInterface
''                                        MsgBox "Please Edit Amount in the Required Field", vbInformation
''                                        .Reverse = 1
''                                        .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                                        On Error Resume Next
''                                        For Each mCrl In frmDemandInterface.Controls
''                                            If TypeOf mCrl Is ComboBox Then
''                                                mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is CommandButton Then
''                                                    mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is TextBox Then
''                                                    mCrl.Enabled = False
''                                            End If
''                                        Next
''                                        .cmdSave.Enabled = True
''                                        .cmdCancel.Enabled = True
''                                        .Show vbModal
''                                    End With
''                                Case 5  'Account Head
''                                    MsgBox "Please Edit Account Head in the Required Field", vbInformation
''                                    With frmDemandInterface
''                                        .Reverse = 1
''                                        .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                                        On Error Resume Next
''                                        For Each mCrl In frmDemandInterface.Controls
''                                            If TypeOf mCrl Is ComboBox Then
''                                                mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is CommandButton Then
''                                                    mCrl.Enabled = False
''                                            ElseIf TypeOf mCrl Is TextBox Then
''                                                    mCrl.Enabled = False
''                                            End If
''                                        Next
''                                        .cmdSave.Enabled = True
''                                        .cmdCancel.Enabled = True
''                                        .Show vbModal
''                                    End With
''                                Case 6  'Transaction Type
''                                        MsgBox "Please Select Correct Transaction Type", vbInformation
''                                        With frmDemandInterface
''                                            .Reverse = 1
''                                            .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                                            On Error Resume Next
''                                            For Each mCrl In frmDemandInterface.Controls
''                                                If TypeOf mCrl Is ComboBox Then
''                                                    mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is CommandButton Then
''                                                        mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is TextBox Then
''                                                        mCrl.Enabled = False
''                                                End If
''                                            Next
''                                            .vsGrid.Editable = flexEDNone
''                                            .cmdSave.Enabled = True
''                                            .cmdCancel.Enabled = True
''                                            .cmbTransactionType.Enabled = True
''                                            .Show vbModal
''                                        End With
''                                Case 7  'Wrong demand
''                                        MsgBox "Please Enter Data to correct the Voucher in the Required Field", vbInformation
''                                        With frmDemandInterface
''                                            .Reverse = 1
''                                            .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                                            On Error Resume Next
''                                            For Each mCrl In frmDemandInterface.Controls
''                                                If TypeOf mCrl Is ComboBox Then
''                                                    mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is CommandButton Then
''                                                        mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is TextBox Then
''                                                        mCrl.Enabled = False
''                                                End If
''                                            Next
''                                            .vsGrid.Editable = flexEDNone
''                                            .cmdSave.Enabled = True
''                                            .cmdCancel.Enabled = True
''                                            .cmbTransactionType.Enabled = True
''                                            .txtWardNo.Enabled = True
''                                            .cmbZone.Enabled = True
''                                            .Show vbModal
''                                        End With
''                                Case 8 'Particulars
''                                        MsgBox "Please Enter Correct Details of the Voucher", vbInformation
''                                        With frmDemandInterface
''                                            .Reverse = 1
''                                            .ReverseDemandDetails (val(txtVoucherNo.Tag))
''                                            On Error Resume Next
''                                            For Each mCrl In frmDemandInterface.Controls
''                                                If TypeOf mCrl Is ComboBox Then
''                                                    mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is CommandButton Then
''                                                        mCrl.Enabled = False
''                                                ElseIf TypeOf mCrl Is TextBox Then
''                                                        mCrl.Enabled = False
''                                                End If
''                                            Next
''                                            .vsGrid.Editable = flexEDNone
''                                            .cmdSave.Enabled = True
''                                            .cmdCancel.Enabled = True
''                                            .txtName.Enabled = True
''                                            .txtHouseName.Enabled = True
''                                            .txtPhone.Enabled = True
''                                            .txtDrawnFrom.Enabled = True
''                                            .txtDrawnPlace.Enabled = True
''                                            .Show vbModal
''                                        End With
''                                End Select
''                                If mRevDemand = True Then
''                                    Call SaveRequest
''                                Else
''                                    MsgBox "Request Failed"
''                                    Exit Sub
''                                End If
''                        Else
''                            frmViewVoucher.MultipleVouchers = False
''                            frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
''                            frmViewVoucher.FormName = "frmListReverseEntryRequest"
''                            frmViewVoucher.Show vbModal
''                            If VerifyStatus = 1 Then
''                                Call SaveRequest
''                            Else
''                                MsgBox "Reverse Request Failed", vbInformation
''                                Exit Sub
''                            End If
''                        End If
''                    Else
''                        frmViewVoucher.MultipleVouchers = False
''                        frmViewVoucher.ArrayIn = Array(txtVoucherNo.Tag)
''                        frmViewVoucher.FormName = "frmListReverseEntryRequest"
''                        frmViewVoucher.Show vbModal
''                        If VerifyStatus = 1 Then
''                            Call SaveRequest
''                        Else
''                            MsgBox "Reverse Request Failed", vbInformation
''                            Exit Sub
''                        End If
''                    End If
''            Else
''                cmdNew.Visible = False
''            End If
''        ElseIf LoadMode = 2 Then
''            'frmChequeBounceRequest.Show vbModal
''            frmSearchCheque.Show vbModal
''        End If

        If LoadMode = 1 Then
            frmReverseRequest.Show vbModal
        ElseIf LoadMode = 2 Then
            frmSearchCheque.Show vbModal
        End If
        Call FillGrid
    End Sub

    Private Sub cmdPrint_Click()
        Dim arInput As Variant
        If vsGrid.Row = -1 Then Exit Sub
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 15)) Then
            arInput = Array(CStr(GetVoucherID(vsGrid.TextMatrix(vsGrid.Row, 15))))
            frmViewVoucher.FormName = "PaymentVoucher"
            frmViewVoucher.ArrayIn = arInput
            frmViewVoucher.Show vbModal
        Else
           MsgBox "New Receipt Does not Exists", vbInformation
           Exit Sub
        End If
    End Sub
    Private Function GetVoucherID(ByVal mVNo) As Variant
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            If objDB.SetConnection(mCnn) Then
                mSQL = " Select intVoucherID from faVouchers Where intVoucherNo=" & mVNo
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    GetVoucherID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                End If
            End If
    End Function
    Private Sub cmdView_Click()
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        If vsGrid.Row = -1 Then Exit Sub
        If vsGrid.TextMatrix(vsGrid.Row, 7) <> "" Then
            arInput = Array(vsGrid.TextMatrix(vsGrid.Row, 7), "")
            frmNewViewer.rptFileName = App.Path & "\Reports\rptReverseVoucherDetails.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show
        End If
    End Sub
    Private Sub Form_Paint()
        Me.Top = 0
        Me.Left = (Screen.Width - Me.Width) / 2
        Call FillGrid
    End Sub

    Private Sub Form_Load()
        txtDateFrom.Text = CheckDateInMMM(DateAdd("m", -1, Date))
        txtDateTo.Text = CheckDateInMMM(Date)
        optAll.Value = True
        Call FillGrid
        WindowsXPC1.InitIDESubClassing
        If gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdNew.Enabled = True
            cmdCancel.Enabled = True
        Else
            cmdNew.Enabled = False
            cmdCancel.Enabled = False
        End If
        If mLoadMode = 1 Then
            Me.Caption = "Reverse Entry List"
            vsGrid.Visible = True
            vsGridForCheque.Visible = False
        Else
            Me.Caption = "Cheque Return List"
            vsGrid.Visible = False
            vsGridForCheque.Visible = True
            lblFirst.Visible = False
            optFirst.Visible = False
            cmdView.Visible = False
            cmdPrint.Visible = False
        End If
        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
            optFirst.Visible = False
            lblFirst.Visible = False
        End If
    End Sub



    Private Sub optAll_Click()
        Call FillGrid
    End Sub

    Private Sub optFirst_Click()
        Call FillGrid
    End Sub

    Private Sub optRequested_Click()
        Call FillGrid
    End Sub

    Private Sub optRev_Click()
        Call FillGrid
    End Sub

    Private Sub txtDateFrom_LostFocus()
        If IsDate(txtDateFrom.Text) Then
            txtDateFrom.Text = CheckDateInMMM(txtDateFrom.Text)
            Call FillGrid
        Else
            MsgBox "Please Enter Valid Date", vbInformation
            Exit Sub
        End If
    End Sub

    Private Sub txtDateTo_LostFocus()
        If IsDate(txtDateTo.Text) Then
            txtDateTo.Text = CheckDateInMMM(txtDateTo.Text)
            Call FillGrid
        Else
            MsgBox "Please Enter Valid Date", vbInformation
            Exit Sub
        End If
    End Sub

    Private Sub vsGrid_Click()
        If vsGrid.Row = -1 Then Exit Sub
        If mLoadMode = 1 Then
           vsGrid.Row = vsGrid.MouseRow
           If vsGrid.Row > 0 Then
                If gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    cmdCancel.Visible = True
                    cmdCancel.Enabled = True
                    cmdNew.Visible = True
                End If
                If vsGrid.TextMatrix(vsGrid.Row, 13) <> "" Then
                    If vsGrid.TextMatrix(vsGrid.Row, 13) = 0 Then
                        lblRevStatus.Caption = "Need intermediary Approval"
                    ElseIf vsGrid.TextMatrix(vsGrid.Row, 13) = 1 Then
                        lblRevStatus.Caption = "InterMediary Approval Done. Final Approval For Reverse is Pending"
                    ElseIf vsGrid.TextMatrix(vsGrid.Row, 13) = 3 Then
                        lblRevStatus.Caption = "This Voucher is Approved for Reverse. Requested user Should Save this Voucher"
                    ElseIf vsGrid.TextMatrix(vsGrid.Row, 13) = 2 Then
                        lblRevStatus.Caption = "This Voucher is Reversed.Press View button to See the Details"
                    Else
                        lblRevStatus.Caption = "For rectification entries you can search and find any Voucher Postings, view and verify the details and Request for Reverse the Accounting Entry"
                    End If
                End If
           End If
        End If
    End Sub

    Private Sub vsGrid_DblClick()
        Dim mStatus     As Integer
        Dim mSQL        As String
        Dim objDB       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mDemandNo   As String
        Dim mCrl        As Control
        Dim mVrID       As Double
        Dim mReqID      As Double
        On Error GoTo err:
            If vsGrid.Row = -1 Then Exit Sub
                mVrID = val(vsGrid.TextMatrix(vsGrid.Row, 7))
                mReqID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                If vsGrid.Row = -1 Then Exit Sub
                mStatus = CheckReverseRequestExist(mVrID)
                
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    If mStatus = 1 Then
                        MsgBox "Already Approved", vbApplicationModal
                    ElseIf mStatus = 2 Then
                        MsgBox "This Voucher Reversed", vbApplicationModal
                    Else
                        frmReverseApproval.Request = mReqID
                        frmReverseApproval.Show vbModal
                    End If
                ElseIf gbSeatGroupID = gbSeatGroupSecretary Then
                    If mStatus = 1 Then
                        frmReverseApproval.Request = mReqID
                        frmReverseApproval.Show vbModal
                    ElseIf mStatus = 0 Then
                        MsgBox "Intermediary Approval is Pending", vbApplicationModal
                    ElseIf mStatus = 2 Then
                        MsgBox "Already Approved", vbApplicationModal
                    End If
                ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                    If mStatus = 0 Then
                        If vsGrid.TextMatrix(vsGrid.Row, 5) = gbSeatName Then
                            frmReverseRequest.EditDetails (mReqID)
                            frmReverseRequest.Show vbModal
                        Else
                            MsgBox "This Request is not done in the Current login, Not Allowed to Edit", vbInformation
                        End If
                    ElseIf mStatus = 3 Then
                        If val(vsGrid.TextMatrix(vsGrid.Row, 11)) = 508 Then
                            If gbSeatName = vsGrid.TextMatrix(vsGrid.Row, 5) Then
                                objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                                mSQL = "Select vchDemandNo From faIDemandTBL Where numDemandID=" & vsGrid.TextMatrix(vsGrid.Row, 12)
                                Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
                                If Not (Rec.EOF And Rec.BOF) Then
                                    mDemandNo = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                                    If mDemandNo <> "" Then
                                        frmReceiptsCounter.mReverseMode = True
                                         On Error Resume Next
                                        For Each mCrl In frmReceiptsCounter.Controls
                                            If TypeOf mCrl Is ComboBox Then
                                                mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is CommandButton Then
                                                    mCrl.Enabled = False
                                            ElseIf TypeOf mCrl Is TextBox Then
                                                    mCrl.Enabled = False
                                            End If
                                        Next
                                        On Error GoTo 0
                                        frmReceiptsCounter.txtDemandPrefix.Text = Token(mDemandNo, "-")
                                        frmReceiptsCounter.txtDemandNo.Text = mDemandNo
                                        frmReceiptsCounter.cmdSave.Enabled = True
                                        frmReceiptsCounter.cmdCancel.Enabled = True
                                        frmReceiptsCounter.txtDemandNo_LostFocus
    '                                    frmReceiptsCounter.Show , vbModal
                                        oldOwner = SetOwner(frmReceiptsCounter.hwnd, Me.hwnd)
                                        If mReceiptMode Then
                                            Call ReverseForWrongDemand(mReqID, mVrID, mReceiptID)
                                        Else
                                            MsgBox "Transaction Failed", vbInformation + vbCritical
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "This Request is not done in the Current login", vbInformation
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Public Property Let LoadMode(mData As Integer)
        mLoadMode = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadMode
    End Property
    Public Property Let ReceiptID(mData As Integer)
        mReceiptID = mData
    End Property
    
    Public Property Get ReceiptID() As Integer
        ReceiptID = mReceiptID
    End Property

    Private Sub vsGridForCheque_Click()
        If mLoadMode = 2 Then
                vsGridForCheque.Row = vsGridForCheque.MouseRow
                If vsGridForCheque.Row > -1 Then
                    cmdCancel.Visible = True
                    cmdCancel.Enabled = True
                End If
        End If
    End Sub
    Private Sub vsGridForCheque_DblClick()
         On Error GoTo err:
            'If gbUserTypeID = 1 Or gbUserTypeID = 2 Or gbUserTypeID = 4 Then
             If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
             '(gbSeatGroupID = gbSeatGroupAccountsOfficer And (gbLBType = 3 Or gbLBType = 4))
             
                If vsGridForCheque.Cell(flexcpChecked, vsGridForCheque.Row, 6) = vbChecked Then
                    frmReverseEntryRequest.UserType = 2
                Else
                    frmReverseEntryRequest.UserType = 1
                End If
                frmReverseEntryRequest.RequestID = val(vsGridForCheque.TextMatrix(vsGridForCheque.Row, 6))
                frmReverseEntryRequest.Show vbModal
                Call FillGrid
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub vsGridForCheque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If mLoadMode = 2 Then
                vsGridForCheque.Row = vsGridForCheque.MouseRow
                If vsGridForCheque.Row > -1 Then
                    cmdCancel.Visible = True
                    cmdCancel.Enabled = True
                    'Call PopupMenu(mnuPopup)
                End If
            End If
        End If
    End Sub
    
''    Private Sub CancelReverseEntryRequest_Click()
''        Dim objDb   As New clsDb
''        Dim mCnn    As New ADODB.Connection
''        Dim mSql    As String
''        Dim mRequestID  As Integer
''            If mLoadMode = 2 Then
''                If gbUserName = vsGridForCheque.TextMatrix(vsGridForCheque.Row, 3) Then
''                        mRequestID = vsGridForCheque.TextMatrix(vsGridForCheque.Row, 6)
''                Else
''                    MsgBox "Requested User doesn't Match", vbInformation
''                    Exit Sub
''                End If
''            Else
''                If gbUserName = vsGrid.TextMatrix(vsGrid.Row, 4) Then
''                        mRequestID = vsGrid.TextMatrix(vsGrid.Row, 8)
''                Else
''                    MsgBox "Requested User doesn't Match", vbInformation
''                    Exit Sub
''                End If
''            End If
''            objDb.SetConnection mCnn
''            mSql = "Update faReverseEntry Set tnyStatus=3 Where intRequestID=" & mRequestID
''            objDb.ExecuteSP mSql, , , , mCnn, adCmdText
''            Call FillGrid
''    End Sub

    Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If mLoadMode = 1 Then
                vsGrid.Row = vsGrid.MouseRow
                If vsGrid.Row > -1 Then
                    cmdCancel.Visible = True
                    cmdCancel.Enabled = True
                End If
            End If
        End If
    End Sub

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'----------------------
'---------------------- Redefined Reverse Entry Modification
'----------------------
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
    Private Sub ReverseForWrongDemand(ByVal mReqID As Double, ByVal mVrID As Double, ByVal mReceptID As Double)
        Dim objRev  As New clsReverseProcess
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim mRevVrID    As Double
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mRevVrID = objRev.ReverseTransaction(mVrID, mCnn)
            If mRevVrID > 0 Then
                mCnn.Execute "Update faVouchers set tnysync=Null,intExternalModuleID=55 Where intVoucherID=" & mReceptID
                mCnn.Execute "Update faReverseEntry set tnysync=Null,tnyStatus=2 Where intVoucherID=" & mReqID
            End If
        End If
    End Sub
    
    Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDB As New clsDB
            If objDB.SetConnection(mCnn) Then
                mSQL = " Select tnyStatus from faReverseEntry "
                mSQL = mSQL + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSQL = mSQL + " Where intVoucherID =  " & VchID
                mSQL = mSQL + " And tnyStatus<>4"
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then      'Requested
                        CheckReverseRequestExist = 0
                    ElseIf Rec!tnyStatus = 1 Then  ' Approved
                        CheckReverseRequestExist = 1
                    ElseIf Rec!tnyStatus = 2 Then   'Reversed
                        CheckReverseRequestExist = 2
                    ElseIf Rec!tnyStatus = 3 Then
                        CheckReverseRequestExist = 3
                    Else 'Cancelled Status=4
                        CheckReverseRequestExist = 4
                    End If
                    Exit Function
                Else
                    CheckReverseRequestExist = 5  'NOT EXISTS IN THE TABLE
                End If
                
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact your System Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Public Property Let ArrayIn(mData As Variant)
        aryIn = mData
    End Property
    Public Property Get ArrayIn() As Variant
        ArrayIn = aryIn
    End Property
    Function SetOwner(ByVal HwndtoUse, ByVal HwndofOwner) As Long
        SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
    End Function
     
''Private Sub cmdShowForm_Click()
'    ' show the form
''    Set frm = frmReceiptsCounter
''    frm.Show
''    ' make it owned by the current form
''    oldOwner = SetOwner(frm.hWnd, Me.hWnd)
''End Sub
 
''Private Sub cmdUnloadForm_Click()
''    ' restore original owner
''    SetOwner frm.hWnd, oldOwner
''    ' unload the form
''    Unload frm
''End Sub



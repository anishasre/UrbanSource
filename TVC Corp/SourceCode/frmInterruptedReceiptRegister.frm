VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmInterruptedReceiptRegister 
   BackColor       =   &H00E0E0E0&
   Caption         =   "INNTERRUPTED RECEIPT REGISTER"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterruptedReceiptRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   11820
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      Height          =   360
      Left            =   7425
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8235
      Width           =   660
   End
   Begin VB.TextBox dtSessionDate 
      Height          =   360
      Left            =   135
      TabIndex        =   38
      Text            =   "0"
      Top             =   8775
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   600
      TabIndex        =   46
      Top             =   7740
      Width           =   810
   End
   Begin VB.CommandButton cmdReGenerateVrNo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Re-Generate Voucher Number"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3735
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   630
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.TextBox txtBookStatus 
      Height          =   360
      Left            =   6315
      TabIndex        =   43
      Top             =   7740
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2130
      TabIndex        =   40
      Top             =   7740
      Width           =   1575
   End
   Begin VB.CommandButton cmdInsertSuffix 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insert Sufix No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9615
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7530
      Width           =   1665
   End
   Begin VB.Frame fraInsertSuffix 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   8985
      TabIndex        =   36
      Top             =   7605
      Width           =   2835
      Begin VB.TextBox txtSelInsertSuffix 
         Height          =   330
         Left            =   180
         TabIndex        =   51
         Top             =   585
         Width           =   1770
      End
      Begin VB.CommandButton cmdSendInsertSuffix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insert"
         Height          =   360
         Left            =   1980
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Receipt"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   270
         TabIndex        =   39
         Top             =   330
         Width           =   1065
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   8685
      Top             =   9000
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9585
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5430
      Width           =   1665
   End
   Begin VB.Frame fraChangeDate 
      BackColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   9180
      TabIndex        =   26
      Top             =   5490
      Width           =   2415
      Begin VB.TextBox txtChangeRptTO 
         Height          =   330
         Left            =   1290
         TabIndex        =   33
         Top             =   555
         Width           =   945
      End
      Begin VB.CommandButton cmdSendChangeDateRequest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send"
         Height          =   360
         Left            =   885
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtChangeDateReason 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   165
         TabIndex        =   28
         Top             =   1095
         Width           =   2070
      End
      Begin VB.TextBox txtChangeRptFrom 
         Height          =   330
         Left            =   180
         TabIndex        =   27
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt - To"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1290
         TabIndex        =   34
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   31
         Top             =   885
         Width           =   525
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rpt - From"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   345
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdEditReceipt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Receipt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9570
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Frame fraEditReceipt 
      BackColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   9180
      TabIndex        =   19
      Top             =   3420
      Width           =   2445
      Begin VB.TextBox txtSelEditReceiptNo 
         Height          =   330
         Left            =   405
         TabIndex        =   22
         Top             =   540
         Width           =   1605
      End
      Begin VB.TextBox txtEditReason 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   390
         TabIndex        =   21
         Top             =   1095
         Width           =   1620
      End
      Begin VB.CommandButton cmdSendEditRequest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send"
         Height          =   360
         Left            =   840
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Receipt"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   24
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   405
         TabIndex        =   23
         Top             =   870
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdCancelReceipt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel Receipt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1290
      Width           =   1665
   End
   Begin VB.Frame fraCancelReceipt 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   9180
      TabIndex        =   12
      Top             =   1365
      Width           =   2445
      Begin VB.CommandButton cmdSendCancellationRequest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send"
         Height          =   360
         Left            =   780
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1455
         Width           =   900
      End
      Begin VB.TextBox txtCancellationReason 
         Height          =   330
         Left            =   420
         TabIndex        =   17
         Top             =   1110
         Width           =   1620
      End
      Begin VB.TextBox txtSelCancellationReceiptNo 
         Height          =   360
         Left            =   420
         TabIndex        =   15
         Top             =   555
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   420
         TabIndex        =   16
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Receipt"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   420
         TabIndex        =   14
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   11760
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      Begin VB.CommandButton cmdSearchCounter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   315
         Left            =   11265
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   195
         Width           =   330
      End
      Begin VB.CommandButton cmdGenerateRptNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Generate Receipt No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1185
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   555
         Width           =   2070
      End
      Begin VB.CommandButton cmdSearchBook 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   315
         Left            =   2940
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   195
         Width           =   330
      End
      Begin VB.TextBox txtBook 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued Book"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lblTransactionDate 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10080
         TabIndex        =   6
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   9030
         TabIndex        =   5
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label lblUser 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8850
         TabIndex        =   4
         Top             =   870
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8460
         TabIndex        =   3
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblCounter 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10080
         TabIndex        =   2
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Counter"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9210
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6330
      Left            =   45
      TabIndex        =   11
      Top             =   1305
      Width           =   9105
      _cx             =   16060
      _cy             =   11165
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInterruptedReceiptRegister.frx":1CCA
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
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for Approval"
      Height          =   330
      Left            =   405
      TabIndex        =   49
      Top             =   8235
      Width           =   1725
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   90
      TabIndex        =   48
      Top             =   8235
      Width           =   285
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   240
      Left            =   90
      TabIndex        =   47
      Top             =   7740
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Status:"
      Height          =   240
      Left            =   5355
      TabIndex        =   44
      Top             =   7740
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      Height          =   240
      Left            =   1620
      TabIndex        =   41
      Top             =   7740
      Width           =   480
   End
End
Attribute VB_Name = "frmInterruptedReceiptRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mMasterTypeID As Integer
    Dim mBCount  As Integer
    Dim mReceiptNo   As String
    Dim mReceiptNoFirst   As String
    Dim mStatus As Integer
    Dim mCurrentUserSession As Boolean
    Dim mInterruptedModeFlag As Boolean
    Dim mIRMode As Boolean
    Public mIStatus As Variant
    Public mYearID As Integer
    Dim mBookIssue As Boolean
    Dim mPreviousYearID As Variant
    Dim mUserRequested As Variant
    Dim mIRRequestedDate As Variant
    Dim mSessionDate As Variant
    
    Private Sub FillGrid()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mRowCnt As Integer
        Dim mCount As Integer
        Dim mLoop As Integer
        Dim mArrayIn As Variant
        Dim mdtReceiptDate As Date
        Dim mREceiptNoSuff As Long
        Dim mSessionDt As Date
        mCount = 0
        txtTotal.Text = ""
        txtBookStatus.Text = ""
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mArrayIn = Array(val(txtBook.Tag))
            Rec.CursorLocation = adUseClient
            Set Rec = objdb.ExecuteSP("spGetIRRegisterDetails", mArrayIn, , , mCnn)
            
            vsGrid.Clear 1, 1
            If Not (Rec.EOF And Rec.BOF) Then
                mRowCnt = 1
                vsGrid.Rows = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    
                    'NOTE: CHECKING CURRENT SESSION OR NOT
                    '    :  mSESSIONDATE is a MODULE LEVEL VARIABLE SET FROM CHECKIRMODE
                    If IsNull(mSessionDate) Then
                        mSessionDt = gbTransactionDate 'NOTE: NO CURRENT SESSION SO SETTING SESSIONDATE AS TRANSACTION DATE
                    Else
                        mSessionDt = mSessionDate ' SESSION IS ALREADY STARTED
                    End If
                    
                    If IsDate(Rec!dtDataEntry) Then 'NOTE:: REGISTER DATAENTRY DATE
                        If Rec!dtDataEntry = mSessionDt Then
                            vsGrid.TextMatrix(mRowCnt, 7) = 1  'Current Session
                            If IsNull(mSessionDate) Then 'NOTE::IF SESSION VARIABLE NOT SET
                                Call SetSessionDate      '      UPDATE THE SESSION DATE TO TABLE - IR REQUEST
                            End If
                        Else
                            vsGrid.TextMatrix(mRowCnt, 7) = 0  'NOT IN CURRENT SESSION
                        End If
                    End If
                    
                   
                
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!intID), "", Rec!intID)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!ReceiptNO), "", Rec!ReceiptNO)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!dtVoucherDate), "", Rec!dtVoucherDate)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!Amt), "", Rec!Amt)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    If Not IsNull(Rec!tnyCancelled) And Rec!tnyCancelled <> 0 Then
                        vsGrid.TextMatrix(mRowCnt, 5) = "CANCELLED"
'                    ElseIf Rec!intTypeID = 2 And Rec!tnyFlag = 3 Then
'                        vsGrid.TextMatrix(mRowCnt, 5) = "EDITED"
'                    ElseIf Rec!intTypeID = 3 And Rec!tnyFlag = 2 Then
'                        vsGrid.TextMatrix(mRowCnt, 5) = "DATE EDITED"
'                    ElseIf Rec!intTypeID = 4 And Rec!tnyFlag = 2 Then
'                        vsGrid.TextMatrix(mRowCnt, 5) = "SUFFIX INSERTED"
                    End If
                    If Rec!intTypeID = 1 And Rec!tnyFlag = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 5) = "Requested for Cancellation"
                    ElseIf Rec!intTypeID = 2 And Rec!tnyFlag = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 5) = "Requested for Editing"
                    ElseIf Rec!intTypeID = 3 And Rec!tnyFlag = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 5) = "Requested for Date Editing"
                    ElseIf Rec!intTypeID = 4 And Rec!tnyFlag = 1 Then
                         vsGrid.TextMatrix(mRowCnt, 5) = "Requested for Suffix"
                    End If
                    
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!tnyClosed), "", Rec!tnyClosed)
                    If val(vsGrid.TextMatrix(1, 6)) = 1 Then
                        txtBookStatus.Text = "CLOSED"
                    Else
                        txtBookStatus.Text = "OPEN"
                    End If
                    If IsNull(Rec!vchSuffix) Then
                         mCount = mCount + 1
                    End If
                    
                    
                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyFlag), 0, Rec!tnyFlag)
                    vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!tnyCancelled), 0, Rec!tnyCancelled)
                    vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intTypeID), 0, Rec!intTypeID)
                    If val(vsGrid.TextMatrix(mRowCnt, 9)) = 1 Then  'Request for Approval
                            For mLoop = 0 To vsGrid.Cols - 1
                                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HC0E0FF
                            Next mLoop
                    ElseIf val(vsGrid.TextMatrix(mRowCnt, 9)) = 2 And val(vsGrid.TextMatrix(mRowCnt, 11)) = 2 And val(vsGrid.TextMatrix(mRowCnt, 7)) = 0 Then 'Edit Approved
                            For mLoop = 0 To vsGrid.Cols - 1
                                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HE0E0E0
                            Next mLoop
                    End If
                    vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!vchSuffix), 0, 1)
                    vsGrid.TextMatrix(mRowCnt, 14) = IIf(IsNull(Rec!tnyVerified), 0, Rec!tnyVerified)
                    vsGrid.TextMatrix(mRowCnt, 15) = IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)
                    vsGrid.TextMatrix(mRowCnt, 16) = IIf(IsNull(Rec!intTranstypeID), 0, Rec!intTranstypeID)
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                Wend
                Call CalculateTotal
                cmdGenerateRptNo.Enabled = False
                'txtCount.Text = val(mCount)
                
            Else
                cmdGenerateRptNo.Enabled = True
                txtBookStatus.Text = "OPEN"
                'txtCount.Text = "0"
                If mBookIssue = False And gbSeatGroupID <> gbSeatGroupAccountsOfficer Then
                    cmdGenerateRptNo.Enabled = False
                End If
            End If
            Rec.Close
        End If
    End Sub
    Private Sub SetSessionDate()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        If objdb.SetConnection(mCnn) Then
            mSql = "Update faInterruptedRequests set dtReceiptChangeDate ='" & DdMmmYy(gbTransactionDate) & "' Where numUserID =" & gbUserID & "  And intCounterID =" & gbCounterID & "  And intTypeID = 1  And tnyStatus = 2"
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            mSessionDate = gbTransactionDate
        End If
    End Sub
    Private Sub CalculateTotal()
        Dim mTotal As Double
        Dim mLoop As Integer
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        
        mTotal = 0
        For mLoop = 1 To vsGrid.Rows - 1
            mTotal = mTotal + val(vsGrid.TextMatrix(mLoop, 3))
        Next
        txtTotal.Text = mTotal
        
        objdb.SetConnection mCnn
        mSql = " Select count(A.mCount) mTCount From "
        mSql = mSql + " (Select count(*) mCount from faInterruptedRegister Where intBookID =" & val(txtBook.Tag) & "   Group by intReceiptNo)A"  'And isnull(tnyStatus,0)=0
        
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
                txtCount.Text = IIf(IsNull(Rec!mTCount), 0, Rec!mTCount)
        Else
            txtCount.Text = "0"
        End If
        Rec.Close
    End Sub
    Private Sub FormInitialize()  'to set the counter,Date,User
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim RecChild     As New ADODB.Recordset
        Dim mSql    As String
        Dim mSqlChild    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        Dim mdtDate As Date
        
        mMasterTypeID = -1
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mBookIssue = True
        mStatus = ""
        mSql = "Select tnyStatus,dtReceiptDate, dtRequestDate, faInterruptedRequests.intCounterID, vchDescription From faInterruptedRequests "
        mSql = mSql + " INNER JOIN faCounters ON faCounters.intCounterID = faInterruptedRequests.intCounterID"
        'mSql = mSql + " Where numUserID = " & gbUserID
        'mSql = mSql + " Where tnyStatus = 2"
        mSql = mSql + " And tnyStatus = 2"
        
        mSql = mSql + " And intTypeID = 1 "
        mSql = mSql + " And faInterruptedRequests.intCounterID=" & gbCounterID
        
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            
            'mdtDate = Rec!dtRequestDate
            mdtDate = Rec!dtReceiptDate 'NOTE: INTERRUPTED TRANSACTION DATE FROM REQUEST
            lblCounter.Caption = Rec!vchDescription
            lblCounter.Tag = Rec!intCounterID
            lblTransactionDate.Caption = Format(mdtDate, "dd-mmm-yyyy")
            lblUser.Caption = gbUserName
            
            Call CheckfinancialYear
            
            mSql = "SELECT * From faInterruptedReceiptBooks WHERE tnyClosed=0 And intFinancialYearID=" & mYearID & " And intCounterID = " & Rec!intCounterID
            If Rec.State = 1 Then
                Rec.Close
            End If
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtBook.Text = Rec!intBookNo
                txtBook.Tag = Rec!intBookID
                Call FillGrid
                Call CheckInterruptedBookStatus
            Else
                 mSqlChild = "SELECT * From faInterruptedReceiptBooks WHERE intFinancialYearID=" & mYearID & " And intCounterID = " & lblCounter.Tag
                 RecChild.Open mSqlChild, mCnn
                 If Not (RecChild.EOF And RecChild.BOF) Then
                    txtBook.Text = RecChild!intBookNo
                    txtBook.Tag = RecChild!intBookID
                    Call FillGrid
                    Call CheckInterruptedBookStatus
                Else
                    MsgBox "No Books Issued for the Requested FinancialYear", vbInformation
                    mBookIssue = False
                    Exit Sub
                End If
                RecChild.Close
            End If
        Else
            lblCounter.Caption = ""
            lblTransactionDate.Caption = ""
            lblUser.Caption = ""
        End If
        Rec.Close
        mCnn.Close
    End Sub
    Private Sub cmdCancelReceipt_Click()
    
        'NOTE:-val(vsGrid.TextMatrix(vsGrid.Row, 9)) :tnyFlag 1-Request,2-Approve
        
        'NOTE: ONLY CHIEF CASHIER OR CASHIER OR SECRETARY/ACCOUNTS OFFICERs
        '       ARE ONLY PERMITED TO DO ANY OPERATION OVER THIS FUNCTIONALITY
        If Not (gbSeatGroupID = gbSeatGroupAccountsOfficer _
            Or gbSeatGroupID = gbSeatGroupCashier _
            Or gbSeatGroupID = gbSeatGroupChiefCashier) Then
            
            cmdSendCancellationRequest.Enabled = False
            Exit Sub
        End If
        
        If vsGrid.Row > 0 Then
            vsGrid.ColHidden(12) = True

            cmdSendCancellationRequest.Enabled = True
            txtSelCancellationReceiptNo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 1 Then   'And vsGrid.TextMatrix(vsGrid.Row, 11) = 1
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    cmdSendCancellationRequest.Caption = "Approve"
                Else
                    cmdSendCancellationRequest.Caption = "Send"
                End If
            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 0 Then
                cmdSendCancellationRequest.Caption = "Send"
            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 1 Then  'UNDO
                cmdSendCancellationRequest.Caption = "UNDO"
            Else
                cmdSendCancellationRequest.Enabled = False
            End If
            
        End If
    End Sub
    
    Private Sub cmdChangeDate_Click()
        'NOTE:-val(vsGrid.TextMatrix(vsGrid.Row, 9)) :tnyFlag 1-Request,2-Approve
        
        'NOTE: ONLY CHIEF CASHIER OR CASHIER OR SECRETARY/ACCOUNTS OFFICERs
        '       ARE ONLY PERMITED TO DO ANY OPERATION OVER THIS FUNCTIONALITY
        If Not (gbSeatGroupID = gbSeatGroupAccountsOfficer _
            Or gbSeatGroupID = gbSeatGroupCashier _
            Or gbSeatGroupID = gbSeatGroupChiefCashier) Then
   
            cmdSendChangeDateRequest.Enabled = False
            Exit Sub
        End If
         
       
       
       If vsGrid.Row > 0 Then
            If gbSeatGroupID <> gbSeatGroupAccountsOfficer Then
                vsGrid.ColHidden(12) = False
                vsGrid.Editable = flexEDKbdMouse
            End If
            If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
                txtChangeRptFrom.Text = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 2))
                cmdSendChangeDateRequest.Enabled = True
            Else
                cmdSendChangeDateRequest.Enabled = False
                Exit Sub
            End If
            txtChangeDateReason.Tag = val(vsGrid.TextMatrix(vsGrid.Row, 1))
            If gbSeatGroupAccountsOfficer = gbSeatGroupID Then
                RefillvsGrid_Changedate
            End If
            If vsGrid.TextMatrix(vsGrid.Row, 9) = 0 Then
                cmdSendChangeDateRequest.Caption = "Send"
            End If
        End If
        txtChangeRptFrom.SetFocus
    End Sub
    Private Sub RefillvsGrid_Changedate()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim i       As Integer
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'mSQL = "Select * from faInterruptedRequests Where intTypeID=5 And tnyStatus=0 And intBookID= " & val(txtBook.Tag) & " "
        
        mSql = "Select * from faInterruptedRequests Where intTypeID=5 And tnyStatus=0 And intBookID= " & val(txtBook.Tag) & " "
        mSql = mSql + "And intStartVoucherNo=" & val(txtChangeDateReason.Tag) & ""
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            For i = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(i, 9) = 1 And vsGrid.TextMatrix(i, 11) = 3 Then
                     vsGrid.Cell(flexcpChecked, i, 12) = 1
                End If
            Next
            txtChangeRptTO.Text = DdMmmYy(IIf(IsNull(Rec!dtReceiptChangeDate), 0, Rec!dtReceiptChangeDate))
            cmdSendChangeDateRequest.Caption = "Approve"
        Else
            cmdSendChangeDateRequest.Enabled = False
        End If
        mCnn.Close
    End Sub

    Private Sub cmdClear_Click()
        Call fnCleartext
''''        Dim ctrl    As Control
''''
''''        For Each ctrl In Me.Controls
''''            If TypeOf ctrl Is TextBox Then
''''                ctrl.Text = ""
''''                ctrl.Tag = ""
''''            ElseIf TypeOf ctrl Is OptionButton Then
''''                ctrl.value = False
''''            ElseIf TypeOf ctrl Is ComboBox Then
''''                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
''''                ctrl.Tag = ""
''''            End If
''''        Next
    End Sub

    Private Sub cmdInsertSuffix_Click()
        'NOTE:-val(vsGrid.TextMatrix(vsGrid.Row, 9)) :tnyFlag 1-Request,2-Approve
        
        'NOTE: ONLY CHIEF CASHIER OR CASHIER OR SECRETARY/ACCOUNTS OFFICERs
        '       ARE ONLY PERMITED TO DO ANY OPERATION OVER THIS FUNCTIONALITY
        If Not (gbSeatGroupID = gbSeatGroupAccountsOfficer _
            Or gbSeatGroupID = gbSeatGroupCashier _
            Or gbSeatGroupID = gbSeatGroupChiefCashier) Then
            
            cmdSendInsertSuffix.Enabled = False
            Exit Sub
        End If
         
         If vsGrid.Row > 0 Then
            vsGrid.ColHidden(12) = True
            txtSelInsertSuffix.Text = val(vsGrid.TextMatrix(vsGrid.Row, 1))
            
            If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 1 Then
                MsgBox "The Receipt is cancelled", vbInformation
                txtSelInsertSuffix.Text = ""
                Exit Sub
            Else
                cmdSendInsertSuffix.Enabled = True
                If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 1 Then
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        cmdSendCancellationRequest.Caption = "Approve"
                    End If
                ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 0 Then
                    cmdSendCancellationRequest.Caption = "Send"
                Else
                    cmdSendCancellationRequest.Enabled = False
                End If
            End If
        End If
    End Sub
    
    Private Sub cmdEditReceipt_Click()
        
        'NOTE:-val(vsGrid.TextMatrix(vsGrid.Row, 9)) :tnyFlag 1-Edit Request,2-Approve Edit Request,3- Receipt Edited
        
        'NOTE: ONLY CHIEF CASHIER OR CASHIER OR SECRETARY/ACCOUNTS OFFICERs
        '       ARE ONLY PERMITED TO DO ANY OPERATION OVER THIS FUNCTIONALITY
        If Not (gbSeatGroupID = gbSeatGroupAccountsOfficer _
            Or gbSeatGroupID = gbSeatGroupCashier _
            Or gbSeatGroupID = gbSeatGroupChiefCashier) Then
            
            cmdSendEditRequest.Enabled = False
            Exit Sub
        End If
            
        
        If vsGrid.Row > 0 Then
            vsGrid.ColHidden(12) = True
            txtSelEditReceiptNo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            
            If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 1 Then 'val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 1- Receipt tnyCancelled
                MsgBox "The Receipt is cancelled", vbInformation
                txtSelEditReceiptNo.Text = ""
                Exit Sub
            ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 16)) = 1 Then
                If gbFetchDemandFromWeb = 1 Then
                    MsgBox "IR with Property Tax (Integrated) Editing is not possible", vbInformation
                    txtSelEditReceiptNo.Text = ""
                    Exit Sub
                End If
            Else
                cmdSendEditRequest.Enabled = True
                If vsGrid.TextMatrix(vsGrid.Row, 9) = 1 Then
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        cmdSendEditRequest.Caption = "Approve"
                    End If
                ElseIf vsGrid.TextMatrix(vsGrid.Row, 9) = 2 Then
                    cmdSendEditRequest.Caption = "Edit"
                Else
                    cmdSendEditRequest.Caption = "Send"
                End If
            End If
        End If
    End Sub
    
    Private Sub cmdGenerateRptNo_Click()
        Call FillNewReceiptNo
    End Sub
    Private Sub FillNewReceiptNo()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim RecChild     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mCount  As Integer
        Dim mArrIN  As Variant
        Dim mPrefix      As String
        Dim mCounterID   As Integer
        Dim mFinYearID   As Integer
        Dim mBookNo      As Long
        Dim mMaxRecNo    As String
        Dim mVoucherNo   As String
        Dim mReceiptFrom As Long
        Dim mCounterNo As Integer
        Dim i As Integer
        Dim mMAXID As Integer
        'Dim mCounterID As Integer
        
        mSql = "Select faInterruptedReceiptBooks.*,faCounters.intCounterNo as intCounterNo from faInterruptedReceiptBooks "
        mSql = mSql + " INNER JOIN faCounters ON faCounters.intCounterid=faInterruptedReceiptBooks.intCounterid "
        mSql = mSql + " Where intBookID = " & val(txtBook.Tag)

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mCounterNo = Rec!intCounterNo
             mCounterID = Rec!intCounterID
             mCount = Rec!intCount
             mMaxRecNo = Rec!numReceiptNoFrom + 1 - 1
             mReceiptFrom = Rec!numReceiptNoFrom
             mBookNo = Rec!intBookNo
             mPrefix = "9" + Right("00000" + LTrim(str(mBookNo)), 5) + "1"
             mMaxRecNo = str(mPrefix) + Right("00000" + str(mMaxRecNo), 5)
             mVoucherNo = "9" + Right("00000" + LTrim(str(mBookNo)), 5) + "1" + Right("00000" + LTrim(str(mReceiptFrom)), 5)
    'mReceiptNo = "9" + Right("0000" + LTrim(str(mBookNo)), 4) + "1" + Right("00000" + LTrim(str(mReceiptFrom)), 5)
        End If
        Rec.Close
        
        If mCounterID = gbCounterID Then
            mSql = " Select max(intID) mMaxID from faInterruptedRegister"
            RecChild.Open mSql, mCnn
            If Not (RecChild.EOF And RecChild.BOF) Then
                mMAXID = IIf(IsNull(RecChild!mMAXID), 0, RecChild!mMAXID)
            End If
            RecChild.Close
            mStatus = 3
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            If mVoucherNo > 0 Then
                
                '
                'NEED TO CHECK WHETHER RECEIPT NUMBER IS GENERATED FOR THE BOOK ID SELECTED
                '
                '
                '
                For i = 1 To mCount
                        mMAXID = mMAXID + 1
                        vsGrid.Rows = vsGrid.Rows + 1
                        vsGrid.TextMatrix(i, 0) = mMAXID
                        vsGrid.TextMatrix(i, 1) = val(mVoucherNo)
                        vsGrid.TextMatrix(i, 2) = ""
                        vsGrid.TextMatrix(i, 3) = ""
                        vsGrid.TextMatrix(i, 4) = ""
                        vsGrid.TextMatrix(i, 5) = ""
                        vsGrid.TextMatrix(i, 6) = 0
                        vsGrid.TextMatrix(i, 7) = 1 '""  For Current Session
                        vsGrid.TextMatrix(i, 8) = 3
                        vsGrid.TextMatrix(i, 9) = ""
                        vsGrid.TextMatrix(i, 10) = 0
                        
                        mSql = " INSERT INTO faInterruptedRegister"
                        mSql = mSql + " (intID, intBookID, intReceiptNo, vchSuffix, intSLNo, tnyCancelled, tnyStatus, intVoucherID, dtVoucherDate, fltAmount, intUserID, dtDataEntry, tnyFlag , tnyVerified)" & vbNewLine
                        mSql = mSql + " VALUES (" & mMAXID & "," & val(txtBook.Tag) & "," & val(mVoucherNo) & ",null,null,null," & mStatus & ",null,null,null," & gbUserID & ",'" & DdMmmYy(gbTransactionDate) & "',0 , 1) " & vbNewLine
                        
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                        mVoucherNo = val(mVoucherNo) + 1
                Next
                txtTotal.Text = "0"
                txtBookStatus = "OPEN"
                cmdGenerateRptNo.Enabled = False
            Else
                MsgBox "Generation Failed, Try again" & mCounterNo, vbInformation
                Exit Sub
            End If
        Else
            MsgBox "This Book is Issued to Counter Number :" & mCounterNo, vbInformation
            Exit Sub
        End If
        mCnn.Close
    End Sub
    
    Private Sub cmdReGenerateVrNo_Click()
        'Call fnReGenerateVrNo(mReceiptNoFirst, mBCount, mStatus) '(mReceiptNo, mBCount, mStatus)
        Call ReGenVoucherNo(mReceiptNoFirst, mBCount, mStatus)
    End Sub
    Private Function CheckSuffix(mID As Variant) As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mCancel As Variant
        'mSql = "Select vchSuffix  from faInterruptedRegister Where intReceiptNo =" & mReceiptNo & ""'And isnull(tnyCancelled,0)<>1 " 'And isnull(tnyCancelled,0)<>1 "
        mSql = "Select vchSuffix,tnyCancelled  from faInterruptedRegister Where intID =" & mID & " "
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mCancel = IIf(IsNull(Rec!tnyCancelled), 0, Rec!tnyCancelled)
            If (IsNull(Rec!vchSuffix)) Then
                CheckSuffix = False
            Else
                CheckSuffix = True
            End If
        End If
        Rec.Close
        mCnn.Close
    End Function
    
    Private Sub ReGenVoucherNo(ByVal mReceiptNo As String, ByVal mBCount As Integer, ByVal mStatus As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mVoucherFlag As Boolean
        Dim mMAXID As Long
        Dim mLoop As Integer
        Dim mID As Long
        Dim mSuffixFlag As Integer
        Dim mVNo As Variant
        Dim mInArr As Variant
        
        Me.MousePointer = vbHourglass
        
        objdb.SetConnection mCnn
    
        '[1]
        'CHECK VERIFIED FLAG IS OPEN OR NOT
        mSql = "SELECT *, ISNULL(tnyVerified,0) tnyVerified FROM faInterruptedRegister Where intBookID =" & val(txtBook.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            'NOTE: IF NOT VERIFIED DELETE RECORDS FROM REGISTER
            If Rec!tnyVerified <> 1 Then
                mSql = "DELETE FROM faInterruptedRegister WHERE intBOOKID = " & val(txtBook.Tag)
                mCnn.Execute mSql
                ':: BLOCKED CODE
                'mInArr = Array(val(txtBook.Tag))
                'objdb.ExecuteSP "spDeleteBookFromIRRegister", mInArr, , , mCnn, adCmdStoredProc
                '::
            Else
                MsgBox "This Book is already verified and linked with the Register!", vbInformation
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            Rec.Close
        End If
        If Rec.State Then Rec.Close
        
        '[2]
        'NOTE: FIND THE MAX ID ( RISK IS THERE BECAUSE RECORD IS NOT LOCKING HERE)
        mSql = "SELECT ISNULL(Max(intID) + 1,1)  MaxID FROM faInterruptedRegister "
        Rec.Open mSql, mCnn
        mMAXID = Rec!MaxID
        
        '[3]
        'NOTE: ADD NEW RECORDS TILL THE RECEITP COUNT AS PER BOOK ADDED
        For mLoop = 1 To mBCount
            mSql = " INSERT INTO faInterruptedRegister"
            mSql = mSql + " (intID, intBookID, intReceiptNo, vchSuffix, intSLNo, tnyCancelled, tnyStatus, intVoucherID, dtVoucherDate, fltAmount, intUserID, dtDataEntry, tnyFlag)" & vbNewLine
            mSql = mSql + " VALUES (" & mMAXID & "," & val(txtBook.Tag) & "," & val(mReceiptNo) & ",null,null,null," & 0 & ",null,null,null," & gbUserID & ",'" & DdMmmYy(gbTransactionDate) & "',0) " & vbNewLine
            mCnn.Execute mSql
            mReceiptNo = mReceiptNo + 1
            mMAXID = mMAXID + 1
        Next
        Rec.Close
        
        
        'NOTE: GET LIST FROM THE REGISTER IN AN ORDER
        mSql = "SELECT * FROM faInterruptedRegister WHERE intBOOKID = " & val(txtBook.Tag) & " ORDER BY intID"
        RecChild.Open mSql, mCnn
        
        'NOTE: FETCH AND LINK VOUCHERS
        mSql = "SELECT * FROM faVouchers WHERE intBookNo = " & val(txtBook.Tag)
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn
        mSuffixFlag = 0
        
        'NOTE: LINK VOUCHERS WITH REGISTER
        While Not Rec.EOF
Step2:
            'NOTE:IDENTIFY CURRENT VOUCHER IS SUFIX ENABLED OR NOT
            '     AND FIND VOUCHER NO TOO
            If IsNull(Rec!vchDoorNoP3) Then
                mSuffixFlag = 0
                mVNo = -1
            Else
                mSuffixFlag = mSuffixFlag + 1
                mVNo = Rec!intVoucherNo
            End If
            
            
            mID = RecChild!intID 'NOTE: REGISTER's CURRENT ID IS STORED IN A VARIABLE TO UPDATE VOUCHER DETAILS
            mReceiptNo = RecChild!intReceiptNo 'NOTE: RECEIPT NO. IS STORING FOR SUFFIX NUMBER IF INSERT
            
            'NOTE: UPDATING VOUCHER DETAILS IN REGISTERS
            mSql = "UPDATE faInterruptedRegister SET intVoucherID=" & Rec!intVoucherID
            If mSuffixFlag > 0 Then
                mSql = mSql + ", vchSuffix = 'A'"
            End If
            mSql = mSql + ", dtVoucherDate =  '" & DdMmmYy(Rec!dtDate) & "'"
            mSql = mSql + ", fltAmount = " & Rec!fltAmount
            mSql = mSql + ", tnyStatus = " & mStatus
            mSql = mSql + ", dtDataEntry = '" & DdMmmYy(Rec!dtTimeStamp) & "'"
            mSql = mSql + " WHERE intID = " & mID
            mCnn.Execute mSql
STEP1:
            Rec.MoveNext ':: MOVING TO NEXT VOUCHER
            If Not Rec.EOF Then
                'NOTE:: CHECKING IN VOUCHER SUFFIX AND VOUCHER NUMBER IS SAME WITH THE PREVIOUS RECEIPT
                If Not IsNull(Rec!vchDoorNoP3) And Rec!intVoucherNo = mVNo Then
                    mSuffixFlag = mSuffixFlag + 1
                    mSql = " INSERT INTO faInterruptedRegister"
                    mSql = mSql + " (intID, intBookID, intReceiptNo, vchSuffix, intSLNo, tnyCancelled, tnyStatus, intVoucherID, dtVoucherDate, fltAmount, intUserID, dtDataEntry, tnyFlag)" & vbNewLine
                    mSql = mSql + " VALUES (" & mMAXID & "," & val(txtBook.Tag) & "," & val(mReceiptNo) & ", '" & Chr(64 + mSuffixFlag) & "' , null,null," & 0 & "," & Rec!intVoucherID & ",'" & DdMmmYy(Rec!dtDate) & "'," & Rec!fltAmount & "," & gbUserID & ",'" & DdMmmYy(Rec!dtTimeStamp) & "',0) " & vbNewLine
                    mCnn.Execute mSql
                    mMAXID = mMAXID + 1
                    GoTo STEP1:
                Else
                    mSuffixFlag = 0       ':: NOTE: RESETTING SUFFIX
                    If Not RecChild.EOF Then
                        RecChild.MoveNext ':: MOVEING TO REGISTER's NEXT RECORD
                        GoTo Step2:       ':: LOOP TILL END OF VOUCHERS
                    Else
                        mSql = " Voucher Count is more than the Book Leaves!" & vbCrLf
                        mSql = mSql + " Error: Can't proceed further to link the voucher with Register"
                        MsgBox mSql, vbInformation
                        Exit Sub
                    End If
                End If
            Else
                'NOTE: EOF OF VOUCHERS:: LINKING VOUCHERS WITH REGISTER IS COMPLETED
                '    : NOW UPDATING VOUCHERS WITH NEW NUMBER FORMAT FROM REGISTER
                mSql = "UPDATE faVouchers SET intVoucherNo = faInterruptedRegister.intReceiptNo, vchDoorNoP3 = vchSuffix FROM faVouchers "
                mSql = mSql + " INNER JOIN faInterruptedRegister ON faVouchers.intVoucherID = faInterruptedRegister.intVoucherID"
                mSql = mSql + " WHERE intBookID = " & val(txtBook.Tag)
                mCnn.Execute mSql
            End If
        Wend
        Rec.Close
        
        'NOTE: MARKING THE BOOK AS VERIFIED IN REGISTER
        mSql = "Update faInterruptedRegister SET tnyVerified = 1 WHERE intBookID = " & val(txtBook.Tag)
        mCnn.Execute mSql
        
        Call FillGrid ':: REFILLING GRID
        Me.MousePointer = vbDefault

    End Sub
    
    Private Function fnReGenerateVrNo(ByVal mReceiptNo As String, ByVal mBCount As Integer, ByVal mStatus As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim i, mID     As Integer
        Dim mNewCount, mMAXID As Integer
        Dim mTCount, mCount As Integer
        
        mSql = " Select count(A.mCount) mTCount From "
        mSql = mSql + " (Select count(*) mCount from faInterruptedRegister Where intBookID =" & val(txtBook.Tag) & "   Group by intReceiptNo)A" 'And isnull(tnyStatus,0)=0
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mTCount = Rec!mTCount  'total Count for receipt inserted(suffix considered as one)
        End If
        Rec.Close
        
        mSql = "Select count(intReceiptNo)ReceiptCount,min(intID) intID  from faInterruptedRegister Where intBookID =" & val(txtBook.Tag) & ""
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mID = Rec!intID
             mCount = Rec!ReceiptCount
            If mTCount = mBCount Then
                 For i = 1 To mCount
                    If CheckSuffix(mID) = False Then
                        mSql = " Update faVouchers set intVoucherNo=" & mReceiptNo & " "
                        mSql = mSql + " From faVouchers "
                        mSql = mSql + " inner join faInterruptedRegister on faInterruptedRegister.intVoucherID=faVouchers.intVoucherID"
                        mSql = mSql + " Where faInterruptedRegister.intID = " & mID & "  And faInterruptedRegister.intBookID = " & val(txtBook.Tag) & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mSql = "Update faInterruptedRegister set intReceiptNo=" & mReceiptNo & ", tnyStatus=" & mStatus & " Where intBookID =" & val(txtBook.Tag) & " And intID=" & mID & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mReceiptNo = val(mReceiptNo) + 1
                    
                    Else
                        mSql = " Update faVouchers set intVoucherNo=" & mReceiptNo & " "
                        mSql = mSql + " From faVouchers "
                        mSql = mSql + " inner join faInterruptedRegister on faInterruptedRegister.intVoucherID=faVouchers.intVoucherID"
                        mSql = mSql + " Where faInterruptedRegister.intID = " & mID & "  And faInterruptedRegister.intBookID = " & val(txtBook.Tag) & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mSql = "Update faInterruptedRegister set intReceiptNo=" & mReceiptNo & ", tnyStatus=" & mStatus & " Where intBookID =" & val(txtBook.Tag) & " And intID=" & mID & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    End If
                    mID = mID + 1
                 Next
                 MsgBox "Re-generated successffully", vbInformation
                 cmdReGenerateVrNo.Visible = False
                 Call FillGrid
            Else
                For i = 1 To mCount
                
                    If CheckSuffix(mID) = False Then
                        mSql = " Update faVouchers set intVoucherNo=" & mReceiptNo & " "
                        mSql = mSql + " From faVouchers "
                        mSql = mSql + " inner join faInterruptedRegister on faInterruptedRegister.intVoucherID=faVouchers.intVoucherID"
                        mSql = mSql + " Where faInterruptedRegister.intID = " & mID & "  And faInterruptedRegister.intBookID = " & val(txtBook.Tag) & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mSql = "Update faInterruptedRegister set intReceiptNo=" & mReceiptNo & ", tnyStatus=" & mStatus & " Where intBookID =" & val(txtBook.Tag) & " And intID=" & mID & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mReceiptNo = val(mReceiptNo) + 1
                    
                    Else
                        mSql = " Update faVouchers set intVoucherNo=" & mReceiptNo & " "
                        mSql = mSql + " From faVouchers "
                        mSql = mSql + " inner join faInterruptedRegister on faInterruptedRegister.intVoucherID=faVouchers.intVoucherID"
                        mSql = mSql + " Where faInterruptedRegister.intID = " & mID & "  And faInterruptedRegister.intBookID = " & val(txtBook.Tag) & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
                        mSql = "Update faInterruptedRegister set intReceiptNo=" & mReceiptNo & ", tnyStatus=" & mStatus & " Where intBookID =" & val(txtBook.Tag) & " And intID=" & mID & " "
                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    End If
                    mID = mID + 1
                Next
                mNewCount = mBCount - val(mTCount)
                mSql = " Select max(intID) mMaxID from faInterruptedRegister" 'Where intBookID =" & val(txtBook.Tag) & " " & vbNewLine
                RecChild.Open mSql, mCnn
                If Not (RecChild.EOF And RecChild.BOF) Then
                    mMAXID = IIf(IsNull(RecChild!mMAXID), 0, RecChild!mMAXID)
                End If
                RecChild.Close
                For i = 1 To mNewCount
                    mMAXID = mMAXID + 1
                    mStatus = 3
                    mSql = " INSERT INTO faInterruptedRegister"
                    mSql = mSql + " (intID, intBookID, intReceiptNo, vchSuffix, intSLNo, tnyCancelled, tnyStatus, intVoucherID, dtVoucherDate, fltAmount, intUserID, dtDataEntry, tnyFlag)" & vbNewLine
                    mSql = mSql + " VALUES (" & mMAXID & "," & val(txtBook.Tag) & "," & mReceiptNo & ",null,null,null," & mStatus & ",null,null,null," & gbUserID & ",'" & DdMmmYy(gbTransactionDate) & "',0) " & vbNewLine
                    
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    mReceiptNo = val(mReceiptNo) + 1
                    
                Next
                cmdReGenerateVrNo.Visible = False
                Call FillGrid
            End If
        End If
        Rec.Close
        mCnn.Close
    End Function
    Private Sub cmdSearchBook_Click()
        Dim mSql As String
        mMasterTypeID = 2
        Call CheckfinancialYear
        
        'mSql = "SELECT  intBookID, intBookNo From faInterruptedReceiptBooks WHERE intFinancialYearID=" & mYearID & " "
        'And intCounterID = " & val(lblCounter.Tag)
        
        mSql = " SELECT  intBookID, convert(varchar(50),intBookNo) + '('+  convert(varchar(50),numReceiptNoFrom )+ '-' + convert(varchar(50),numReceiptNoTo) + ')' "
        mSql = mSql + " From faInterruptedReceiptBooks WHERE intFinancialYearID IN (" & gbFinancialYearID & "," & gbFinancialYearID - 1 & " )"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = mSql
        frmSearchMasters.Show vbModal
        
        If gbSearchStr <> "" Then
            txtBook.Text = gbSearchStr
            txtBook.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            mMasterTypeID = -1
            'Call FillGrid
            Call CheckInterruptedBookStatus
        End If
        
    End Sub
    
    Private Sub cmdSearchCounter_Click()
        Dim mSql As String
        mMasterTypeID = 1
        mSql = "Select  faCounters.intCounterID, vchDescription From faInterruptedRequests INNER JOIN"
        mSql = mSql + " faCounters ON faCounters.intCounterID = faInterruptedRequests.intCounterID"
        mSql = mSql + " Where tnyStatus = 2 And intTypeID = 1 AND numUserID = " & gbUserID
        
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = mSql
        frmSearchMasters.Show vbModal
        
        If gbSearchStr <> "" Then
            lblCounter.Caption = gbSearchStr
            lblCounter.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            mMasterTypeID = -1
            Call SetCounter
            Call FillGrid
            Call CheckInterruptedBookStatus
        End If
    End Sub
    Private Sub SetCounter()
     Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim dtRequestDate As Date
        
        Call CheckfinancialYear
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT  intBookID, intBookNo From faInterruptedReceiptBooks WHERE intFinancialYearID=" & mYearID & " And intCounterID = " & val(lblCounter.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtBook.Text = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
            txtBook.Tag = IIf(IsNull(Rec!intBookID), 0, Rec!intBookID)
        End If
        Rec.Close
        mCnn.Close
    End Sub
    
    Private Sub SetOpenBook()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim dtRequestDate As Date
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT  intBookID, intBookNo From faInterruptedReceiptBooks WHERE intFinancialYearID = " & mYearID & " And intCounterID = " & val(lblCounter.Tag) & " AND ISNULL(tnyClosed,0) = 0 "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtBook.Text = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
            txtBook.Tag = IIf(IsNull(Rec!intBookID), 0, Rec!intBookID)
        End If
        Rec.Close
        mCnn.Close
    End Sub
    Private Sub cmdSendCancellationRequest_Click()  'tnyFlag=1 intTypeID=1 -Request   tnyFlag=2 intTypeID=1 -Approve
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mSerialNo As Long
        Dim dtReceiptDate As String
        Dim mVoucherNo As String
        Dim mSuffix As String
        Dim mVrSuffix As String
        
            
        If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
            dtReceiptDate = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 2))
        Else
             dtReceiptDate = ""
        End If
        
        '-----------------LAST POSTING VALIDATION------------------
        If dtReceiptDate <> "" Then
            If CDate(dtReceiptDate) <= CDate(gbLastPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                txtSelCancellationReceiptNo.Text = ""
                'cmdSendCancellationRequest.Enabled = False
                Exit Sub
            End If
        End If
        '-------------------------------------------------------------
        
        
        
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 1 Then
                    mSerialNo = mID$(vsGrid.TextMatrix(vsGrid.Row, 1), 6, 5)
                    mSql = "Update faInterruptedRegister set tnyCancelled=1 ,tnyFlag=0,intTypeID=1 Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    
                    mSql = " INSERT INTO faInterruptedCancelledReceipts"
                    mSql = mSql + "(intBookID, intBookNo, intSerialNo, numUserID, dtReceiptDate, intReceiptNo, vchRemarks,dtRequestDate, dtApprovalDate, numApprovingOfficer, tnyStatus,intVoucherID)" & vbNewLine
                    mSql = mSql + " VALUES (" & val(txtBook.Tag) & "," & val(txtBook.Text) & "," & mSerialNo & "," & gbUserID & ", '" & dtReceiptDate & "' , " & val(txtSelCancellationReceiptNo.Text) & " ,'" & txtCancellationReason.Text & "' , " & vbNewLine
                    mSql = mSql + " '" & DdMmmYy(gbTransactionDate) & "','" & DdMmmYy(gbTransactionDate) & "'," & gbUserID & ",1," & val(vsGrid.TextMatrix(vsGrid.Row, 15)) & " ) "
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    
                    mSql = " Update faVouchers  SET tnyCancelFlag = 1,tnyStatus = 4   WHERE intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15)) & " "
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    mSql = " Update faTransactions SET tnyStatus = 4 WHERE intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            End If
            
            If vsGrid.TextMatrix(vsGrid.Row, 10) = 1 Then 'UNDO
                mSql = "Update faInterruptedRegister set tnyCancelled=null ,tnyFlag=0,intTypeID=1  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Delete from faInterruptedCancelledReceipts Where intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Update faVouchers  SET tnyCancelFlag = 0 , tnyStatus  = 0  WHERe intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Update faTransactions SET tnyStatus = 0 WHERE intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                vsGrid.TextMatrix(vsGrid.Row, 9) = 0
                vsGrid.TextMatrix(vsGrid.Row, 10) = 0
            End If
            '**********************************************************************************************************************
               Call UpdateVoucherIndex(val(vsGrid.TextMatrix(vsGrid.Row, 15)))    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
            '**********************************************************************************************************************
            
        ElseIf vsGrid.TextMatrix(vsGrid.Row, 7) = 1 Then 'CheckCurrentSession = True Then
        
            mSerialNo = mID$(vsGrid.TextMatrix(vsGrid.Row, 1), 6, 5)
            If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 0 Then  'SEND
                mSql = "Update faInterruptedRegister set tnyCancelled=1 ,tnyFlag=0,intTypeID=1  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " INSERT INTO faInterruptedCancelledReceipts"
                mSql = mSql + "(intBookID, intBookNo, intSerialNo, numUserID, dtReceiptDate, intReceiptNo, vchRemarks,dtRequestDate, dtApprovalDate, numApprovingOfficer, tnyStatus,intVoucherID)" & vbNewLine
                mSql = mSql + " VALUES (" & val(txtBook.Tag) & "," & val(txtBook.Text) & "," & mSerialNo & "," & gbUserID & ", '" & dtReceiptDate & "' , " & val(txtSelCancellationReceiptNo.Text) & " ,'" & txtCancellationReason.Text & "' , " & vbNewLine
                mSql = mSql + " '" & DdMmmYy(gbTransactionDate) & "','" & DdMmmYy(gbTransactionDate) & "'," & gbUserID & ",1," & val(vsGrid.TextMatrix(vsGrid.Row, 15)) & " ) "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Update faVouchers  SET tnyCancelFlag = 1 ,tnyStatus = 4   WHERe intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                mSql = " Update faTransactions SET tnyStatus = 4 WHERE intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                vsGrid.TextMatrix(vsGrid.Row, 9) = 1
            End If
            
            If vsGrid.TextMatrix(vsGrid.Row, 10) = 1 Then 'UNDO
                mSql = "Update faInterruptedRegister set tnyCancelled=null ,tnyFlag=0,intTypeID=1  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Delete from faInterruptedCancelledReceipts Where intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Update faVouchers  SET tnyCancelFlag = 0 , tnyStatus  = 0  WHERe intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                mSql = " Update faTransactions SET tnyStatus = 0 WHERE intVoucherID =" & val(vsGrid.TextMatrix(vsGrid.Row, 15))
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                
                vsGrid.TextMatrix(vsGrid.Row, 9) = 0
                vsGrid.TextMatrix(vsGrid.Row, 10) = 0
            End If
            
             '**********************************************************************************************************************
                Call UpdateVoucherIndex(val(vsGrid.TextMatrix(vsGrid.Row, 15)))    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
             '**********************************************************************************************************************
                        
        Else    'APPROVE
            MsgBox "Request send for Approval", vbInformation
            mSql = "Update faInterruptedRegister set tnyFlag=1,intTypeID=1  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0))
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
       End If
       txtSelCancellationReceiptNo.Text = ""
       txtCancellationReason.Text = ""
       Call FillGrid
    End Sub
    Private Sub cmdSendChangeDateRequest_Click()
        Dim mCnn       As New ADODB.Connection
        Dim Rec        As New ADODB.Recordset
        Dim mSql       As String
        Dim objdb      As New clsDB
        Dim mSerialNo, i As Integer
        Dim mArrIN     As Variant
        Dim mStartVrNo As Variant
        Dim mEndVrNo   As Variant
        Dim mCheck      As Boolean
        mCheck = False
        
        For i = 1 To vsGrid.Rows - 1
            
            If vsGrid.Cell(flexcpChecked, i, 12) = 1 Then
                mCheck = True
                Exit For
            End If
        Next
        If mCheck = False Then
            MsgBox "PLEASE SELECT THE CHECK BOX ", vbApplicationModal
            Exit Sub
        End If
        For i = 1 To vsGrid.Rows - 1
               
            If IsEmpty(mStartVrNo) Then
                If vsGrid.Cell(flexcpChecked, i, 12) = 1 Then
                    mStartVrNo = vsGrid.TextMatrix(i, 1)
                End If
            End If
            If vsGrid.Cell(flexcpChecked, i, 12) = 1 Then
                mEndVrNo = vsGrid.TextMatrix(i, 1)
            End If
            
        Next
        If CheckDateValidation = False Then
            MsgBox "Please check the Date Entered", vbInformation
            Exit Sub
        Else
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                For i = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpChecked, i, 12) = 1 Then 'vsGrid.TextMatrix(vsGrid.Row, 9) = 1 Then  'Request
                        mSql = "Update faInterruptedRegister "
                        mSql = mSql + " Set tnyFlag=0,intTypeID=3,dtVoucherDate= '" & DdMmmYy(txtChangeRptTO.Text) & "' "
                        mSql = mSql + " Where intID =" & val(vsGrid.TextMatrix(i, 0)) & " "
                        mCnn.Execute mSql
                        
                        mSql = "Update faVouchers"
                        mSql = mSql + " Set dtDate = '" & DdMmmYy(txtChangeRptTO.Text) & "'"
                        mSql = mSql + "  WHERE intVoucherID =" & val(vsGrid.TextMatrix(i, 15)) & " "
                        mCnn.Execute mSql
                        
                        mSql = "Update faTransactions "
                        mSql = mSql + " Set dtTransactionDate = '" & DdMmmYy(txtChangeRptTO.Text) & "'"
                        mSql = mSql + " WHERE intVoucherID =" & val(vsGrid.TextMatrix(i, 15)) & " "
                        mCnn.Execute mSql
                        
                        '**********************************************************************************************************************
                           Call UpdateVoucherIndex(val(vsGrid.TextMatrix(i, 15)))     'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
                        '**********************************************************************************************************************
                
                     End If
                 Next
                 mSql = "Update faInterruptedRequests "
                 mSql = mSql + " Set tnyStatus = 2"
                 mSql = mSql + " Where intTypeID=5 "
                 mSql = mSql + " And intBookID = " & val(txtBook.Tag)
                 mCnn.Execute mSql
                 MsgBox "Successfully Saved", vbInformation
            Else
                If mStartVrNo <> "" And mEndVrNo <> "" Then
                    mArrIN = Array(gbCounterID, _
                           gbUserID, _
                           0, _
                           Format(gbTransactionDate, "DD/MMM/yyyy"), _
                           5, _
                           Format(txtChangeRptFrom.Text, "DD/MMM/yyyy"), _
                           Null, _
                           Null, _
                           Null, _
                           Null, _
                           val(mStartVrNo), _
                           val(mEndVrNo), _
                           Format(txtChangeRptTO.Text, "DD/MMM/yyyy"), _
                           val(txtBook.Tag))
                     objdb.ExecuteSP "spSaveInterruptedRequest", mArrIN, , , mCnn, adCmdStoredProc
                     MsgBox "Request send for Approval", vbInformation
                     mSql = "Update faInterruptedRegister set tnyFlag=1,intTypeID=3  Where intReceiptNo between " & val(mStartVrNo) & " And " & val(mEndVrNo) & " "
                     objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                 End If
            End If
        txtChangeRptTO.Text = ""
        txtEditReason.Text = ""
        Call FillGrid
        End If
        mCnn.Close
    End Sub
    Private Function CheckBookFinancialYear() As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim dtRequestDate As Date
        
        Call CheckfinancialYear
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT  intFinancialYearID From faInterruptedReceiptBooks WHERE intBookID= " & val(txtBook.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
           'If Rec!intFinancialYearID = gbFinancialYearID Then
           If Rec!intFinancialYearID = mYearID Then
                CheckBookFinancialYear = True
           Else
                CheckBookFinancialYear = False
           End If
        End If
        Rec.Close
        mCnn.Close
    End Function
    Private Function CheckDateValidation() As Boolean
    
    
        Dim mD1 As Variant
        Dim mD2 As Variant
        Dim mLoop, X As Integer
        Dim mNewDate As Variant
        Dim mCount As Integer
        Dim mPreDate As Variant
        
        
        mCount = 0
        For mLoop = 1 To vsGrid.Rows - 1
            
            If vsGrid.Cell(flexcpChecked, mLoop, 12) = 1 Then
                If mCount = 0 Then
                    mPreDate = CDate(vsGrid.TextMatrix(mLoop, 2))
                Else
                    mPreDate = Null
                End If
                mCount = mCount + 1
                
                If mCount > 1 Then
                    GoTo mD2:
                End If
                
                If mLoop > 1 Then
                    For X = mLoop - 1 To 1 Step -1
                        If IsDate(vsGrid.TextMatrix(X, 2)) Then
                            mD1 = CDate(vsGrid.TextMatrix(X, 2))
                            Exit For
                        End If
                    
                    Next X
                    If Not IsDate(mD1) Then
                        GoTo SetDate:
                    End If

                Else
SetDate:
                    
                    If mPreviousYearID Then
                        mD1 = CDate(DateAdd("YYYY", -1, gbStartingDate))
                    Else
                        If CheckBookFinancialYear = False Then '
                            mD1 = CDate(DateAdd("YYYY", -1, gbStartingDate)) '
                        Else '
                            mD1 = CDate(gbStartingDate)
                        End If '
                    End If
                End If
mD2:
                If (mLoop + 1) <= (vsGrid.Rows - 1) Then
                    If IsDate(vsGrid.TextMatrix(mLoop + 1, 2)) Then
                        mD2 = CDate(vsGrid.TextMatrix(mLoop + 1, 2))
                    Else
                        mD2 = Null
                    End If
                Else
                    mD2 = Null
                End If
            End If
            
            
        Next
        
        If Not IsDate(mD2) Then
            If mPreviousYearID Then
                mD2 = CDate(DateAdd("YYYY", -1, gbEndingDate))
            Else
                If CheckBookFinancialYear = False Then   '
                    mD2 = CDate(DateAdd("YYYY", -1, gbEndingDate)) '
                Else '
                    mD2 = CDate(gbEndingDate)
                End If '
            End If
        End If
        
        If Not IsDate(mD1) Then
            CheckDateValidation = False
            Exit Function
        End If
        
        If IsDate(CDate(txtChangeRptTO.Text)) Then
            mNewDate = CDate(txtChangeRptTO.Text)
        Else
            CheckDateValidation = False
            Exit Function
        End If
        
        If mNewDate <= mD2 Then
            If Not (mNewDate >= mD1) Then
                CheckDateValidation = False
                Exit Function
            Else
                CheckDateValidation = True
                Exit Function
            End If
        Else
            CheckDateValidation = False
            Exit Function
        End If
        
        
        
        '        Dim mDate As Date
        '        Dim mNextDate As Date
        '        If vsGrid.Row > 0 Then
        '        mDate = CDate(txtChangeRptTO.Text)
        '        '          If vsGrid.Row <> vsGrid.Rows Then
        '            mNextDate = CDate(vsGrid.TextMatrix(vsGrid.Row + 1, 2))
        '            If mDate < mNextDate Then
        '              MsgBox "You can't change the Receipt Date to " & txtChangeRptTO.Text
        '              CheckDateValidation = False
        '              Exit Function
        '            End If
        '        '          End If
        '        If CDate(txtChangeRptTO.Text) > CDate(gbTransactionDate) Then
        '              MsgBox "You can't change the Receipt Date to " & txtChangeRptTO.Text
        '              CheckDateValidation = False
        '             Exit Function
        '        End If
        '        If Trim(txtChangeRptTO.Text) = "" Then
        '            MsgBox "Please enter the new date", vbInformation
        '            CheckDateValidation = False
        '            Exit Function
        '        End If
        '        CheckDateValidation = True
        '        End If
    End Function
    
    Private Sub cmdSendEditRequest_Click()
        'NOTE:-val(vsGrid.TextMatrix(vsGrid.Row, 9)) :tnyFlag 1-Edit Request,2-Approve Edit Request,3- Receipt Edited
        
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mVoucherID As Long
        Dim aryIn As Variant
        Dim dtReceiptDate As String
        
        If vsGrid.TextMatrix(vsGrid.Row, 2) <> "" Then
            dtReceiptDate = DdMmmYy(vsGrid.TextMatrix(vsGrid.Row, 2))
        Else
             dtReceiptDate = ""
        End If
        If dtReceiptDate = "" Then
             MsgBox "Interrupt Receipt details not entered", vbInformation
             Exit Sub
        End If
              
        '-----------------LAST POSTING VALIDATION------------------
        If CDate(dtReceiptDate) <= CDate(gbLastPostingDate) Then
            MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
            txtSelEditReceiptNo.Text = ""
            'cmdSendEditRequest.Enabled = False
            Exit Sub
        End If
        '-------------------------------------------------------------
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select intVoucherID from faInterruptedRegister Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mVoucherID = IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)
        End If
        
        If mVoucherID = 0 Then
            MsgBox "Interrupted Receipt Not Saved", vbInformation
            Exit Sub
        End If
        
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 1 Then
                mSql = "Update faInterruptedRegister set tnyFlag=2,intTypeID=2  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            End If
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 7)) = 1 Then 'CheckCurrentSession = True Then
             If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 0 Then    'SEND
                If mVoucherID <> 0 Then
                    Call LoadEditReceipt(mVoucherID, 2)
                End If
             End If
        ElseIf val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 2 Then
                If mVoucherID <> 0 Then
                    Call LoadEditReceipt(mVoucherID, 3)
                End If
        Else    'REQUEST FOR APPROVAL
            MsgBox "Request send for Approval", vbInformation
            mSql = "Update faInterruptedRegister set tnyFlag=1,intTypeID=2  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        End If
        txtSelEditReceiptNo.Text = ""
        txtEditReason.Text = ""
        Call FillGrid
        Rec.Close
        mCnn.Close
End Sub
     Public Sub LoadEditReceipt(mVoucherID As Long, mtnyFlag As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        
        frmReceiptsCounter.InterruptedRegister = 2
        frmReceiptsCounter.InterruptEditMode = True

        frmReceiptsCounter.IRBookID = val(txtBook.Tag)
        
        If IsDate(mIRRequestedDate) Then
            frmReceiptsCounter.InterruptedRegisterReceiptDate = DdMmmYy(CDate(mIRRequestedDate))  'lblTransactionDate.Caption
        Else
            MsgBox "Transaction Date for Interrupted Receipt is not set", vbInformation
            Exit Sub
        End If
        
        If mPreviousYearID = True Then
             frmReceiptsCounter.mPreviousYearMode = 1
        Else
             frmReceiptsCounter.mPreviousYearMode = 0
        End If
        
'        frmReceiptsCounter.DisplayReceiptDetails (mVoucherID)
        frmReceiptsCounter.DisplayReceiptDetailsIREdit (mVoucherID)
        frmReceiptsCounter.Show
        '**********************************************************************************************************************
        'Call UpdateVoucherIndex(mVoucherID)    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
        '**********************************************************************************************************************
                        
        
        
'        objDb.CreateNewConnection mcnn, enuSourceString.Saankhya
'        mSql = "Update faInterruptedRegister set tnyFlag=0,intTypeID=2  Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
'        objDb.ExecuteSP mSql, , , , mcnn, adCmdText
'        mcnn.Close
    End Sub
    Private Sub cmdSendInsertSuffix_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mSuffix As Variant
        Dim mVoucherNo As String
        Dim mMAXID, i As Integer
        
        If vsGrid.Row <= 0 Then
            Exit Sub
        End If
        If Trim(txtSelInsertSuffix.Text) = "" Then
            Exit Sub
        End If
        
        If val(vsGrid.TextMatrix(vsGrid.Row, 6)) = 1 Then
            Exit Sub
        End If
        mVoucherNo = Token(val(txtSelInsertSuffix.Text), " ")

        'If val(vsGrid.TextMatrix(vsGrid.Row, 7)) = 1 Then 'CheckCurrentSession = True
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSql = " Select max(intID) mMaxID From faInterruptedRegister"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mMAXID = IIf(IsNull(Rec!mMAXID), 0, Rec!mMAXID)
            End If
            Rec.Close
            mSql = " Select max(vchSuffix) vchSuffix From faInterruptedRegister Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mSuffix = IIf(IsNull(Rec!vchSuffix), "A", Rec!vchSuffix)
            End If
            Rec.Close
            If mSuffix = "A" Then
                 mSql = "Update faInterruptedRegister set intTypeID=4,vchSuffix='" & mSuffix & "' Where intID =" & val(vsGrid.TextMatrix(vsGrid.Row, 0)) & " "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            End If
            
            
            If mSuffix <> "Z" Then
                mSuffix = Chr$(Asc(mSuffix) + 1) 'Char(CODE(mSuffix) + 1)
            Else
                MsgBox "Suffix Limit Exceeds", vbInformation, "Saankhya"
                Exit Sub
            End If
            mStatus = 3
            mMAXID = mMAXID + 1
            
            'MsgBox vsGrid.Rows
            
            vsGrid.Rows = vsGrid.Rows + 1
            i = val(vsGrid.Rows - 1)
            vsGrid.TextMatrix(i, 0) = ""
            vsGrid.TextMatrix(i, 1) = mVoucherNo + " " + mSuffix
            vsGrid.TextMatrix(i, 2) = ""
            vsGrid.TextMatrix(i, 3) = ""
            vsGrid.TextMatrix(i, 4) = ""
            vsGrid.TextMatrix(i, 5) = ""
            vsGrid.TextMatrix(i, 6) = 0
            vsGrid.TextMatrix(i, 7) = 1 '""  For Current Session
            vsGrid.TextMatrix(i, 8) = 3
            vsGrid.TextMatrix(i, 9) = ""
            vsGrid.TextMatrix(i, 10) = 0
            vsGrid.TextMatrix(i, 11) = 4
            
            mSql = " INSERT INTO faInterruptedRegister"
            mSql = mSql + " (intID, intBookID, intReceiptNo, vchSuffix, intSLNo, tnyCancelled, tnyStatus, intVoucherID, dtVoucherDate, fltAmount, intUserID, dtDataEntry, tnyFlag,intTypeID,tnyVerified)" & vbNewLine
            mSql = mSql + " VALUES (" & mMAXID & "," & val(txtBook.Tag) & "," & val(mVoucherNo) & ",'" & mSuffix & "' ,null,null," & mStatus & ",null,null,null," & gbUserID & ",'" & DdMmmYy(gbTransactionDate) & "',0,4,1) " & vbNewLine
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            
            Call FillGrid
'        Else
'            MsgBox "Requested Receipt is not in current session", vbInformation
'            Exit Sub
'        End If
        mCnn.Close
    End Sub
    
    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
        
        '[1] READ IR REQUEST DETAILS
        Call CheckInterruptReceiptRequestStatus(gbCounterID)
        
        '[2]SET COUNTER IF ITs CASH COUNTER
        If gbCounterSectionID = 99 Then
            lblCounter.Caption = gbCounterName
            lblCounter.Tag = gbCounterID
            
            '[3]CHECK WHETHER THERE IS ANY OPEN BOOK (CASH COUNTER)
            Call SetOpenBook
        End If
        
        If mInterruptedModeFlag Then
            Call CheckIRMode 'NOTE:: THIS WILL SET IR TRANSACTION DATE AND SESSION DATE
            Call CheckInterruptedBookStatus
        Else 'NOT INTERRUPTED RECEIPT MODE
            Call CheckInterruptedBookStatus
        End If
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSearchCounter.Enabled = False
        End If
    End Sub
    Private Sub CheckInterruptedBookStatus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mCount  As Integer
        Dim mOpenFlag As Integer
        Dim mPrefix      As String
        Dim mCounterID   As Integer
        Dim mFinYearID   As Integer
        Dim mBookNo      As Long
        Dim mMaxRecNo    As String
        Dim mCountOld As Integer
        Dim mLenOld As Integer
        Dim mReceiptFrom As Long
        Dim mFlag As Integer
        Dim i As Integer
        Dim mTCount As Integer
        Dim mRecCount As Integer
        Dim mSuffixFlag As Integer
        Dim mVerifiedFlag As Integer
        
        mCount = 0
        mStatus = 0
        mFlag = 0
        cmdReGenerateVrNo.Visible = False
        
        '' For old records 11 digit
        mSql = "Select count(len(intReceiptNo)) VrLenCont,len(intReceiptNo) VrLen  From faInterruptedRegister "
        mSql = mSql + " Where intBookID=" & val(txtBook.Tag) & " and isnull(vchSuffix,0) in ('0','A') group by len(intReceiptNo) "
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
         mCountOld = Rec!VrLenCont
         mLenOld = Rec!VrLen
        End If
        Rec.Close
        '[1] FIND FIRST RECEIPT NO
        mSql = "Select * from faInterruptedReceiptBooks Where intBookID =" & val(txtBook.Tag) & ""
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mBCount = Rec!intCount
             If Rec!tnyClosed Then
                mOpenFlag = 0
             Else
                mOpenFlag = 1
             End If
             
             mMaxRecNo = Rec!numReceiptNoFrom + 1 - 1
             mReceiptFrom = Rec!numReceiptNoFrom
             mBookNo = Rec!intBookNo
             mPrefix = "9" + Right("00000" + LTrim(str(mBookNo)), 5) + "1"
             mMaxRecNo = str(mPrefix) + Right("00000" + str(mMaxRecNo), 5)
             If mCountOld = mBCount Then
              If mLenOld = 11 Then
                 mReceiptNo = "9" + Right("0000" + LTrim(str(mBookNo)), 4) + "1" + Right("00000" + LTrim(str(mReceiptFrom)), 5)
              Else
                 mReceiptNo = "9" + Right("00000" + LTrim(str(mBookNo)), 5) + "1" + Right("00000" + LTrim(str(mReceiptFrom)), 5)
               End If
            Else
            mReceiptNo = "9" + Right("00000" + LTrim(str(mBookNo)), 5) + "1" + Right("00000" + LTrim(str(mReceiptFrom)), 5)
            End If
             mReceiptNoFirst = mReceiptNo
        End If
        Rec.Close
        
        
        '[2] FIND THE COUNT OF RECEIPT PORTED
        'mSql = "Select count(*) mcount from faInterruptedRegister Where intBookID =" & val(txtBook.Tag) & " And isnull(tnyStatus,0)=0 "
        mSql = " Select count(A.mCount) mTCount From "
        mSql = mSql + " (Select count(*) mCount from faInterruptedRegister Where intBookID =" & val(txtBook.Tag) & "   Group by intReceiptNo)A"  'And isnull(tnyStatus,0)=0
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
                mTCount = IIf(IsNull(Rec!mTCount), 0, Rec!mTCount)
        End If
        Rec.Close
        
        '[3] COMPARING RECEIPT's NUMBER FORMAT
        mSql = "Select *, ISNULL(tnyVerified,0) tnyVerified from faInterruptedRegister Where intBookID =" & val(txtBook.Tag)
        mSql = mSql + " Order by intReceiptNo"
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mRecCount = Rec.RecordCount
            mSuffixFlag = 0
            While (Not Rec.EOF)
                'NOTE: CHECK VERIFIED FLAG AND SET ZERO IF NOT VERIFIED
                If Rec!tnyVerified = 0 Then
                   mVerifiedFlag = 0
                End If
                
                If Rec!intReceiptNo = mReceiptNo Then
                    mFlag = 1
                Else
                   'NOTE:THE ABOVE CONDITION WILL NOT MATCH IF RECEIPT IS SUFFIX ENABLED AND
                   '     HAVE ONLY ONE RECEIPT SUCH AS 912120045-A NEXT RECEIPT IS 912120046-A
                   '     THEN YOU NEED CHECK THE SAME WITH THE HELP OF SUFFIX FLAG VALUE
                   If (Rec!intReceiptNo = mReceiptNo + 1) And mSuffixFlag > 0 Then
                       mFlag = 1
                       mReceiptNo = mReceiptNo + 1
                       mSuffixFlag = 0
                   Else
                       mFlag = 0
                       GoTo skipwhile:
                   End If
                End If
                 
                 If IsNull(Rec!vchSuffix) Then
                     mReceiptNo = mReceiptNo + 1
                     mSuffixFlag = 0
                 Else
                    mSuffixFlag = mSuffixFlag + 1
                 End If
                 
                 If Not Rec.EOF Then Rec.MoveNext
            Wend
            
        Else
            'NOTE:CACHIER OR CHIEF CASHIER
            If gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                cmdGenerateRptNo.Enabled = True
            End If
            Exit Sub
        End If
        Rec.Close
        
        
        
skipwhile:
        
        'NOTE: mFlag 0:Wrong Number Format Or some error
        '    : mOpenFlag 0: Book is Closed
        If mFlag = 0 And mOpenFlag = 0 Then
            If MsgBox(" Number Format is Not Correct!!!Do you want to Regularize the Number Format?", vbYesNo, "Saankhya") = vbYes Then
               SetReGenerateButton
               mStatus = 1
            End If
        ElseIf mFlag = 0 And mOpenFlag = 1 Then 'OPEN BOOK /WRONG NUMBER FORMAT
           MsgBox "Number Generated is not in correct Format.Please Regularise to Correct Number Format", vbOKOnly
           SetReGenerateButton
           mStatus = 2
        ElseIf mFlag = 1 And mVerifiedFlag = 0 Then
            mSql = "UPDATE faInterruptedRegister SET tnyVerified = 1 WHERE intBOOKID = " & val(txtBook.Tag)
            mCnn.Execute mSql
        End If
        
        'NOTE: DIFFERENCE IN TOTAL RECEIPT COUNT AND TOTAL RECEIPT GENERATED
        If mBCount <> mTCount Then
           SetReGenerateButton
        End If
        
        Call FillGrid
        mCnn.Close
    End Sub
    
    Private Sub SetReGenerateButton()
        'NOTE:CACHIER OR CHIEF CASHIER
        If gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            cmdReGenerateVrNo.Visible = True
        End If
    End Sub
    
    Private Sub FraDisable()
        cmdCancelReceipt.Enabled = False
        cmdEditReceipt.Enabled = False
        cmdChangeDate.Enabled = False
        cmdInsertSuffix.Enabled = False
        cmdSendCancellationRequest.Enabled = False
        cmdSendEditRequest.Enabled = False
        cmdSendChangeDateRequest.Enabled = False
        cmdInsertSuffix.Enabled = False
        cmdSendInsertSuffix.Enabled = False
    End Sub
    Private Sub fnCleartext()
        txtSelCancellationReceiptNo.Text = ""
        txtCancellationReason.Text = ""
        txtSelEditReceiptNo.Text = ""
        txtEditReason.Text = ""
        txtChangeRptFrom.Text = ""
        txtChangeRptTO.Text = ""
        txtChangeDateReason.Text = ""
        txtSelInsertSuffix.Text = ""
    End Sub
    Private Sub FraEnable()
        cmdCancelReceipt.Enabled = True
        cmdEditReceipt.Enabled = True
        cmdChangeDate.Enabled = True
        If vsGrid.Row > 0 Then
            '            If CheckCurrentSession = True Then  'And val(vsGrid.TextMatrix(vsGrid.Row, 7)) = 1
            cmdInsertSuffix.Enabled = True
            '            Else
            '                cmdInsertSuffix.Enabled = False
            '            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        WindowsXPC1.InitSubClassing
        Call CheckIR
        cmdGenerateRptNo.Enabled = False
        Me.Height = 9555
        Me.Width = 15375
        Call FraDisable
        Call CheckIRMode
        Call CheckCurrentSession
        Call SetgbLastPostingDate
    End Sub
    
    Private Sub CheckIR()           'Function to check whether data is ported to IR Register table.
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        
        mSql = "Select count(*) countIR from faInterruptedRegister"
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!countIR = 0 Then
              objdb.ExecuteSP "spInsertIRRegister", , , , mCnn, adCmdStoredProc
            End If
        End If
        Rec.Close
        mCnn.Close
    End Sub
    
    Private Sub CheckIRMode()       'Function to check Interrupted Receipt Mode
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mIRMode = False
        mSql = "Select tnyStatus,dtRequestDate,dtReceiptDate,dtReceiptChangeDate From faInterruptedRequests"
        mSql = mSql + " Where intCounterID =" & gbCounterID
        mSql = mSql + " And intTypeID = 1 And tnyStatus = 2"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then 'NOTE:: IR MODE IS TRUE (REQUEST APPROVED)
            mIRMode = True
            If IsDate(Rec!dtReceiptChangeDate) Then
                'NOTE:: SESSION DATE IS SET or SESSION STARTED
                If Rec!dtReceiptChangeDate <> gbTransactionDate Then
                    mIRMode = False
                    mSessionDate = Null
                    mSql = "DELETE FROM faInterruptedRequests Where intCounterID = " & gbCounterID & "  And intTypeID = 1  And tnyStatus = 2"
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                Else
                    mSessionDate = Rec!dtReceiptChangeDate
                End If
            Else 'NOTE: SESSION NOT STARTED
                'mSql = "Update faInterruptedRequests set dtReceiptChangeDate ='" & DdMmmYy(gbTransactionDate) & "' Where numUserID =" & gbUserID & "  And intCounterID =" & gbCounterID & "  And intTypeID = 1  And tnyStatus = 2"
                'objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                mSessionDate = Null
            End If
        Else
            mSessionDate = Null
        End If
        Rec.Close
        mCnn.Close
    End Sub
    
    Private Sub txtChangeRptFrom_LostFocus()
        If txtChangeRptFrom.Text <> "" Then
            txtChangeRptFrom.Text = Format(txtChangeRptFrom.Text, "dd/mmm/yyyy")
        End If
                
        '-----------------LAST POSTING VALIDATION------------------
        If CDate(txtChangeRptFrom.Text) <= CDate(gbLastPostingDate) Then
            MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
            txtChangeRptFrom.Text = ""
            Exit Sub
        End If
        '-------------------------------------------------------------
    End Sub
   
    Private Sub txtChangeRptTO_LostFocus()
         If txtChangeRptTO.Text <> "" Then
            txtChangeRptTO.Text = CheckDateInMMM(txtChangeRptTO.Text)
        End If

    End Sub

    Private Sub vsGrid_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim Rec1     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        
        If vsGrid.Row > 0 Then
            'If gbSeatGroupID <> gbSeatGroupAccountsOfficer Then
                If vsGrid.ColHidden(12) = False Then
                    If vsGrid.Cell(flexcpChecked, vsGrid.Row, 12) = 1 Then
                         If vsGrid.TextMatrix(vsGrid.Row, 10) = 1 Then
                            MsgBox "The Receipt is cancelled", vbInformation
                            'txtChangeRptFrom.Text = ""
                            vsGrid.Cell(flexcpChecked, vsGrid.Row, 12) = 0
                            Exit Sub
                         End If
                    End If
                Else
                    Call FraEnable
                    Call fnCleartext
                End If
            'End If
        End If
        'Call CheckLastPostingDate
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            'objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)
            If val(vsGrid.TextMatrix(vsGrid.Row, 0)) < 1 Then
                MsgBox "Receipt Number not Generated", vbApplicationModal
                Exit Sub
            Else
                mSql = "select *, faInterruptedRegister.intBookID as intBookID, faInterruptedRegister.intReceiptNo as intReceiptNo,"
                mSql = mSql + " faInterruptedReceiptBooks.intFinancialYearID AS intFinancialYearID from"
                mSql = mSql + " faInterruptedRegister INNER JOIN faInterruptedReceiptBooks ON "
                mSql = mSql + " faInterruptedRegister.intBookID = faInterruptedReceiptBooks.intBookID "
                mSql = mSql + " where faInterruptedRegister.intID= " & vsGrid.TextMatrix(vsGrid.Row, 0) & ""
                mSql = mSql + " and faInterruptedRegister.intBookID =" & txtBook.Tag
                Rec.Open mSql, mCnn
            End If
            'objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            If Not (Rec.EOF And Rec.BOF) Then
            'If Rec!intFinancialYearID <> gbFinancialYearID Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                Rec1.Open "select count(intYearid) as cnt from faBLSubmission where tnyStatus=2 and intYearid=" & Rec!intFinancialYearID, mCnn
                    If Rec1!Cnt > 0 Then
                        MsgBox "AFS Submitted to LFA for the Financialyear " & Rec!intFinancialYearID & " So you can't edit", vbInformation
                        Exit Sub
                    End If
            End If
    End Sub
    
    Private Sub vsGrid_DblClick()
    
        'NOTE:
        '     vsGrid.TextMatrix(vsGrid.Row, 8) >> INTERRUPTEDREGISTER[tnySTATUS]
        '     STATUS 0: INITIAL STAGE (ONLY PORTED FROM BOOK)
        '            1:
        '            2: REGENERATED - WHEN THE NUMBER FORMAT WAS WRONG
        '            3: NEW BOOK - WITH NEWLY GENERATED RECORDS IN REGISTER
        '            4: WHEN RECORD RECEIPTS THROUGH THE REGISTER THE STATUS WILL BE 4
        '     vsGrid.TextMatrix(vsGrid.Row, 10) >> CANCELLED FLAG
        '            1: CANCELLED and 0: NOT CANCELLED
        '     vsGrid.TextMatrix(vsGrid.Row, 1) >> RECEIPT NO
        
        If vsGrid.Row > 0 Then
            'NOTE: IF VERIFIED EVOCKING RECEIPT ENTRY SCREEN
             If CheckRequestDate = True Then
                Exit Sub
             Else
                If val(vsGrid.TextMatrix(vsGrid.Row, 14)) = 1 Then
                    Call LoadReceipt
                Else
                    MsgBox "Interuppted Register For the Selected Book is not Verified", vbInformation
                End If
             End If
        End If
    End Sub
    
    Private Function CheckCurrentSession() As Boolean      'Function to check the currentSession
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatusDate As Date
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select dtReceiptChangeDate From faInterruptedRequests"
            mSql = mSql + " Where intCounterID =" & gbCounterID
            mSql = mSql + " And intTypeID = 1  And tnyStatus = 2"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then 'NOTE:: IR MODE IS TRUE (REQUEST APPROVED)
                
                If IsDate(Rec!dtReceiptChangeDate) Then
                    'NOTE:: SESSION DATE IS SET or SESSION STARTED
                    If Rec!dtReceiptChangeDate <> gbTransactionDate Then
                        mSessionDate = Null
                        CheckCurrentSession = False
                    Else
                        mSessionDate = Rec!dtReceiptChangeDate
                        CheckCurrentSession = True
                    End If
                Else 'NOTE: SESSION NOT STARTED
                    mSql = "Update faInterruptedRequests set dtReceiptChangeDate ='" & DdMmmYy(gbTransactionDate) & "' Where numUserID =" & gbUserID & "  And intCounterID =" & gbCounterID & "  And intTypeID = 1  And tnyStatus = 2"
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    mSessionDate = gbTransactionDate
                    CheckCurrentSession = True
                End If
                
            Else
                mSessionDate = Null
            End If
            Rec.Close
            mCnn.Close
        End If
    End Function
    
    Public Sub LoadReceipt()
        frmReceiptsCounter.InterruptedRegister = 1
        frmReceiptsCounter.IRBookID = val(txtBook.Tag)
        
        If IsDate(mIRRequestedDate) Then
            frmReceiptsCounter.InterruptedRegisterReceiptDate = DdMmmYy(CDate(mIRRequestedDate))  'lblTransactionDate.Caption
        Else
            MsgBox "Transaction Date for Interrupted Receipt is not set", vbInformation
            Exit Sub
        End If
        
        If mPreviousYearID = True Then
        
            If CheckBookFinancialYear = False Then
                MsgBox "Please Check the Financial Year of the Book", vbInformation
                Exit Sub
            Else
                frmReceiptsCounter.mPreviousYearMode = 1
            End If
        Else
            If CheckBookFinancialYear = False Then
                MsgBox "Please Check the Financial Year of the Book", vbInformation
                Exit Sub
            Else
                frmReceiptsCounter.mPreviousYearMode = 0
            End If
        End If
        
        frmReceiptsCounter.Show
    End Sub
    
    Public Sub CheckNewReceiptStatus(Optional mVrSuffix As String)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mFirstReceiptNo As Double
        Dim mTypeID As Integer
        Dim mVoucherNo As String
        Dim mSuffix As String
    
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
          
           'mVrSuffix = val(vsGrid.TextMatrix(vsGrid.Row, 1))
           mVoucherNo = Token(mVrSuffix, " ")
           mSuffix = mVrSuffix
    
               If mIRMode = True Then
               mSql = "sELECT min(intreceiptNo) intreceiptNo FROM faInterruptedRegister "
               mSql = mSql + " Where  intBookID=" & val(txtBook.Tag)
               mSql = mSql + " And isnull(tnyCancelled,0)=0"
               mSql = mSql + " And isnull(intVoucherID,0)=0"
               Rec.Open mSql, mCnn
               If Not (Rec.EOF And Rec.BOF) Then
                       mFirstReceiptNo = IIf(IsNull(Rec!intReceiptNo), 0, Rec!intReceiptNo)
               End If
               
               If mFirstReceiptNo = 0 Then
                   Unload frmReceiptsCounter
                   Exit Sub
               Else
                   mTypeID = val(vsGrid.TextMatrix(vsGrid.Row, 11))
                   If mTypeID = 4 Then
                       frmReceiptsCounter.InterruptedRegister = 3
                       frmReceiptsCounter.InterruptedRegisterReceiptNo = mVoucherNo
                       If mSuffix = "" Then
                            frmReceiptsCounter.txtIntruptNoSuffix.Text = "A"
                       Else
                            frmReceiptsCounter.txtIntruptNoSuffix.Text = mSuffix
                       End If
                   Else
                       frmReceiptsCounter.InterruptedRegister = 1
                       frmReceiptsCounter.InterruptedRegisterReceiptNo = mFirstReceiptNo
                       frmReceiptsCounter.txtReceiptNo.Text = mFirstReceiptNo
            End If
                   frmReceiptsCounter.InterruptedRegisterReceiptDate = lblTransactionDate.Caption
                   frmReceiptsCounter.Show
               End If
               If mPreviousYearID = True Then
                    frmReceiptsCounter.mPreviousYearMode = 1
               Else
                    frmReceiptsCounter.mPreviousYearMode = 0
               End If
           End If
        End If
    End Sub
     
     Public Sub CheckInterruptReceiptRequestStatus(mCounterID As Long)
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        
        'NOTE:-  mIStatus : -1= Not Requested, 0= Request Not Approved, 1= Request Approved
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mIStatus = ""
        mSql = "Select tnyStatus,dtReceiptDate,numUserID From faInterruptedRequests"
        mSql = mSql + " Where intCounterID =" & mCounterID
        mSql = mSql + " And numUserID =" & gbUserID
        mSql = mSql + " And intTypeID = 1" 'TypeID = IR REQUEST
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mIStatus = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
            mUserRequested = Rec!numUserID
            mIRRequestedDate = Rec!dtReceiptDate
        Else
            mIStatus = -1 ' NOT REQUESTED
            mUserRequested = Null
            mIRRequestedDate = Null 'gbTransactionDate
        End If
        Rec.Close
        mCnn.Close
        
        If gbSeatGroupID = gbSeatGroupCashier Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            If mIStatus > -1 Then
                If mIStatus = 0 Then ' REQUEST NOT APPROVED
                    mInterruptedModeFlag = False
                End If
                If mIStatus = 1 Then ' REQEUST APPROVED
                    mInterruptedModeFlag = True
                End If
            Else
                mInterruptedModeFlag = False
            End If
        End If
        
        If IsDate(mIRRequestedDate) Then
            If CDate(mIRRequestedDate) < CDate(gbStartingDate) And CDate(mIRRequestedDate) < CDate(gbEndingDate) Then
                mPreviousYearID = True
                mYearID = gbFinancialYearID - 1
            Else
                mPreviousYearID = False
                mYearID = gbFinancialYearID
            End If
        End If
        
    End Sub
    
    Public Sub CheckfinancialYear()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim dtRequestDate As Variant 'Date

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select tnyStatus,dtRequestDate,dtReceiptDate From faInterruptedRequests"
        'mSql = mSql + " And numUserID =" & gbUserID
        mSql = mSql + " Where intCounterID =" & gbCounterID
        mSql = mSql + " And intTypeID = 1 AND numUserID=" & gbUserID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            'dtRequestDate = IIf(IsNull(Rec!dtRequestDate), "", CDate(Rec!dtRequestDate))
            dtRequestDate = IIf(IsNull(Rec!dtReceiptDate), "", CDate(Rec!dtReceiptDate))
        Else
            dtRequestDate = Null
        End If
        Rec.Close
        mCnn.Close
        
        ''''mYearID = Year(CDate(dtRequestDate))
        
        If IsDate(dtRequestDate) Then
            If CDate(dtRequestDate) < CDate(gbStartingDate) And CDate(dtRequestDate) < CDate(gbEndingDate) Then
                mPreviousYearID = True
                mYearID = gbFinancialYearID - 1
            Else
                mPreviousYearID = False
                mYearID = gbFinancialYearID
            End If
        Else
             mPreviousYearID = False
             mYearID = gbFinancialYearID
        End If
    End Sub
    Public Property Let PreviousYear(ByVal mData As Integer)
        mPreviousYearID = mData
    End Property

    Private Function CheckRequestDate() As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim mRecDate As Variant
        'Dim mFlag As Boolean
        
        CheckRequestDate = False
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = " SELECT MAX(dtVoucherDate) dtVoucherDate From faInterruptedRegister WHERE intBookID=" & val(txtBook.Tag)
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            If Rec!dtVoucherDate <> "" Then
                mRecDate = CDate(Rec!dtVoucherDate)
            End If
        End If
         If IsDate(mRecDate) And IsDate(mIRRequestedDate) Then
            If CDate(mIRRequestedDate) < CDate(mRecDate) Then
                MsgBox "The Requested Date cannot be less than previous date", vbInformation
                CheckRequestDate = True
            End If
         End If
        Rec.Close
        mCnn.Close
        
    End Function


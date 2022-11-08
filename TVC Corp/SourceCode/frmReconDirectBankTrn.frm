VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmReconDirectBankTrn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direct Bank Transactions"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00F5F8F8&
      Height          =   6285
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   7020
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Height          =   330
         Left            =   3000
         TabIndex        =   12
         Top             =   2775
         Width           =   510
      End
      Begin VB.TextBox txtInstNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2025
         Width           =   1395
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1590
         TabIndex        =   10
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox txtVoucherDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1590
         TabIndex        =   9
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox txtDescription 
         Height          =   345
         Left            =   360
         TabIndex        =   8
         Top             =   2775
         Width           =   2610
      End
      Begin VB.PictureBox Picture2 
         Height          =   75
         Left            =   285
         ScaleHeight     =   15
         ScaleWidth      =   6375
         TabIndex        =   7
         Top             =   3255
         Width           =   6435
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "[Update To List]"
         Height          =   420
         Left            =   4485
         TabIndex        =   6
         Top             =   5730
         Width           =   1965
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F5F8F8&
         Height          =   690
         Left            =   210
         TabIndex        =   3
         Top             =   285
         Width           =   6585
         Begin VB.OptionButton optReceipt 
            BackColor       =   &H00F5F8F8&
            Caption         =   "Directly Credited By Bank [RECEIPT]"
            Height          =   195
            Left            =   150
            TabIndex        =   5
            Top             =   255
            Width           =   2940
         End
         Begin VB.OptionButton optPayment 
            BackColor       =   &H00F5F8F8&
            Caption         =   "Directly Debited By Bank [PAYMENT]"
            Height          =   195
            Left            =   3525
            TabIndex        =   4
            Top             =   270
            Width           =   3015
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2055
         Left            =   345
         TabIndex        =   13
         Top             =   3660
         Width           =   6375
         _cx             =   11245
         _cy             =   3625
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReconDirectBankTrn.frx":0000
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   17
         Top             =   1395
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   16
         Top             =   1740
         Width           =   750
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REF. NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   15
         Top             =   2100
         Width           =   660
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   375
         TabIndex        =   14
         Top             =   2535
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   2955
         Y1              =   2385
         Y2              =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7020
      TabIndex        =   0
      Top             =   0
      Width           =   7020
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECT BANK TRANSACTIONS WHICH IS NOT INCLUDED IN  CURRENT MONTH'S ACCOUNTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   4665
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmReconDirectBankTrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mdtLastDate As Variant
Private mintBankAccountHeadID As Variant
Private mintReconID As Variant

Private Sub FormInitialize()
    'Initialize
    Dim mCrl As Control
    For Each mCrl In Me.Controls
         If TypeOf mCrl Is TextBox Then
            mCrl.Text = ""
            mCrl.Tag = ""
        ElseIf TypeOf mCrl Is OptionButton Then
            mCrl.value = False
        ElseIf TypeOf mCrl Is ComboBox Then
            If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
        ElseIf TypeOf mCrl Is ComboBox Then
            mCrl.ListIndex = -1
        End If
    Next
    
End Sub

Private Sub cmdAdd_Click()

    Dim mStr As String
    
    If Not IsDate(txtVoucherDate.Text) Then
        MsgBox "Enter a valid Voucher Date", vbInformation
        txtVoucherDate.SetFocus
        Exit Sub
    End If
    If val(Trim(txtAmount.Text)) <= 0 Then
        txtAmount.SetFocus
        MsgBox "Enter the Voucher Amount", vbInformation
        Exit Sub
    End If
    If Trim(txtDescription.Text) = "" Then
        txtDescription.SetFocus
        Exit Sub
    End If
    
    
    mStr = txtVoucherDate.Text & vbTab
    If optReceipt.value = True Then
        mStr = mStr + "Cr" + vbTab
    ElseIf optPayment.value = True Then
        mStr = mStr + "Dr" + vbTab
    Else
        MsgBox "Specify, Directly Credited OR Directly Debited", vbInformation
        Exit Sub
    End If
    mStr = mStr & vbNull & vbTab
    mStr = mStr + Format(val(txtAmount.Text), "0.00") + vbTab
    mStr = mStr & vbNull & vbTab
    mStr = mStr + Trim(txtInstNo.Text) + vbTab
    mStr = mStr & vbNull & vbTab
    mStr = mStr + Trim(txtDescription.Text) + vbTab
    mStr = mStr & vbNull & vbTab
    
    vsGrid.Rows = vsGrid.Rows + 1
    vsGrid.Col = 0
    vsGrid.Row = vsGrid.Rows - 1
    vsGrid.ColSel = vsGrid.Cols - 1
    vsGrid.RowSel = vsGrid.Rows - 1
    vsGrid.Clip = mStr
    vsGrid.ColSel = 0
        
    Call FormInitialize
    
End Sub

Private Sub cmdUpdate_Click()
    Dim mLoop As Integer
    Dim mArrIn As Variant
    Dim mTypeID As Integer
    Dim mDrAmt As Variant
    Dim mCrAmt As Variant
    Dim mdtVoucherDate As Variant
    Dim mdtInstDate As Variant
    Dim mtnyVoucherTypeID As Variant
    Dim objDb As New clsDB
    '
    'GRID COLUMNS
    '0:Date|1:Type|2:V.No|3:Amt|4:InstType|5:Inst.No|6:InstDate|7:Desc.|8:Inst.TypeID
    '
    For mLoop = 1 To vsGrid.Rows - 1
        If IsDate(vsGrid.TextMatrix(mLoop, 0)) And _
                            val(vsGrid.TextMatrix(mLoop, 3)) > 0 Then
            
            
            
            If Trim(vsGrid.TextMatrix(mLoop, 1)) = "Cr" Then ' RECEIPT - CREDITED BY BANK
                mTypeID = 2
                mDrAmt = val(vsGrid.TextMatrix(mLoop, 3))
                mCrAmt = Null
                mtnyVoucherTypeID = 50
            Else ' PAYMENT - DEBITED BY BANK
                mTypeID = 4
                mDrAmt = Null
                mCrAmt = val(vsGrid.TextMatrix(mLoop, 3))
                mtnyVoucherTypeID = 60
            End If
            
            If IsDate(vsGrid.TextMatrix(mLoop, 0)) Then
                mdtVoucherDate = CDate(vsGrid.TextMatrix(mLoop, 0))
            Else
                mdtVoucherDate = Null
            End If
            
            mdtInstDate = Null
            
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
            
                mArrIn = Array(mintReconID, _
                                Null, _
                                mintBankAccountHeadID, _
                                mTypeID, _
                                mDrAmt, _
                                mCrAmt, _
                                Null, _
                                Null, _
                                Null, _
                                Null, _
                                mdtVoucherDate, _
                                vsGrid.TextMatrix(mLoop, 5), _
                                mdtInstDate, _
                                mtnyVoucherTypeID, _
                                Null, Trim(vsGrid.TextMatrix(mLoop, 7)) _
                                )
                
                objDb.ExecuteSP "spSaveBankReconcileChild", mArrIn
                
            
            
        End If
    Next
    MsgBox "Saved the transactions in List", vbInformation
    vsGrid.Rows = 1
    Unload Me
    
End Sub

Private Sub Form_Activate()
    Me.Left = 2250
    Me.Top = 3250
End Sub

Private Sub Form_Load()
    Call FormInitialize
    vsGrid.Rows = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mStr As String
    If vsGrid.Rows > 1 Then
        mStr = ""
        mStr = mStr + "The Newly added Transactions are not UPDATED " & vbCrLf
        mStr = mStr + "to the RECONCILIATION list!" & vbCrLf
        mStr = mStr + " " & vbCrLf
        mStr = mStr + "Do you want to close the form with out  " & vbCrLf
        mStr = mStr + "saving the items to the list?" & vbCrLf
        
        If MsgBox(mStr, vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
        End If
    End If
    frmReconciliation.cmdRefresh.value = True
    
End Sub

Private Sub txtAmount_LostFocus()
    If val(txtAmount.Text) < 0 Then txtAmount = 0
    txtAmount.Text = Format(val(txtAmount.Text), "0.00")
End Sub

Private Sub txtVoucherDate_LostFocus()
    txtVoucherDate.Text = CheckDateInMMM(txtVoucherDate.Text)
    If Trim(txtVoucherDate) <> "" Then
        txtVoucherDate.Text = CheckDateInMMM(txtVoucherDate.Text)
        If Not IsDate(txtVoucherDate.Text) Then
            txtVoucherDate.Text = ""
        Else
            If mdtLastDate < CDate(txtVoucherDate) Then
                MsgBox "You are doing reconciliation as on " & DdMmmYy(CDate(mdtLastDate)), vbInformation
                txtVoucherDate.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub

Public Property Get LastDate() As Variant
    LastDate = mdtLastDate
End Property
Public Property Let LastDate(mData As Variant)
    mdtLastDate = mData
End Property

Public Property Get BankAccountHeadID() As Variant
    BankAccountHeadID = mintBankAccountHeadID
End Property
Public Property Let BankAccountHeadID(mData As Variant)
    mintBankAccountHeadID = mData
End Property
 
Public Property Get ReconID() As Variant
    ReconID = mintReconID
End Property
Public Property Let ReconID(mData As Variant)
    mintReconID = mData
End Property



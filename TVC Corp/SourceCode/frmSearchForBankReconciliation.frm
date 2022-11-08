VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchForBankReconciliation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S e a r c h   F o r   B a n k   R e c o n c i l i a t i o n "
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11055
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Uncheck All"
      Height          =   225
      Left            =   8385
      TabIndex        =   20
      Top             =   1215
      Width           =   2115
   End
   Begin VB.CommandButton cmdManualReconcile 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Manual Reconcile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5670
      Width           =   2145
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "---"
      Height          =   315
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3900
      Width           =   315
   End
   Begin VB.TextBox txtBalanceAmt 
      Height          =   330
      Left            =   5145
      TabIndex        =   7
      Top             =   3900
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Height          =   345
      Left            =   7335
      TabIndex        =   9
      Top             =   3900
      Width           =   1575
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3480
      Top             =   5940
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid fgVoucherStatement 
      Height          =   2355
      Left            =   45
      TabIndex        =   6
      Top             =   1500
      Width           =   10890
      _cx             =   19209
      _cy             =   4154
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorFixed  =   12640511
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchForBankReconciliation.frx":0000
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   1
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Searching Area"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   10830
      Begin VB.CheckBox chkMonth 
         BackColor       =   &H00C0FFFF&
         Caption         =   "List This Month Only"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   8340
         TabIndex        =   19
         Top             =   840
         Width           =   2130
      End
      Begin VB.CommandButton cmdSearch2 
         Caption         =   "---"
         Height          =   315
         Left            =   6705
         TabIndex        =   2
         Top             =   240
         Width           =   315
      End
      Begin VB.CheckBox chkDeepSearching 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tick Here For Deep Searching"
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   795
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   7440
         TabIndex        =   18
         Top             =   120
         Width           =   3045
         Begin VB.OptionButton optAmount 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Amount"
            Height          =   225
            Left            =   1350
            TabIndex        =   4
            Top             =   180
            Width           =   945
         End
         Begin VB.OptionButton optCheque 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cheque"
            Height          =   225
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   4680
         TabIndex        =   1
         Top             =   240
         Width           =   1965
      End
      Begin VB.TextBox txtInstrumentNo 
         Height          =   375
         Left            =   1830
         TabIndex        =   0
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   225
         Left            =   3900
         TabIndex        =   17
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Number"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Width           =   1605
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgBankStatement 
      Height          =   1335
      Left            =   30
      TabIndex        =   10
      Top             =   4335
      Width           =   10890
      _cx             =   19209
      _cy             =   2355
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorFixed  =   12640511
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
      Rows            =   3
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchForBankReconciliation.frx":00F0
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   1
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Amount"
      Height          =   225
      Left            =   3735
      TabIndex        =   16
      Top             =   3930
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B a n k   S t a t e m e n t"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   4065
      Width           =   1890
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V o u c h e r   S t a t e m e n t"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   150
      TabIndex        =   14
      Top             =   1245
      Width           =   2280
   End
End
Attribute VB_Name = "frmSearchForBankReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public mInstNO As String
    Public mAmt As Variant
    Dim mLoop1 As Long
    Dim mLoop2 As Long
    Public Flag As Integer
    Dim mDumyTotal As Variant
    
    Private mvarVoucherID As Variant  'Added By Aiby :: Local Variable to Set Property
    
Private Sub FormInitialize()
    txtAmount.Text = ""
    txtInstrumentNo.Text = ""
    txtBalanceAmt.Text = ""
    txtTotal.Text = ""
    mAmt = ""
    mInstNO = ""
End Sub

Private Sub chkMonth_Click()
    Call SearchVoucher
End Sub

Private Sub cmdManualReconcile_Click()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim mRowCount As Integer
    Dim mVchrID As Long
    Dim mFlag As Boolean
    Dim mNumCount As Integer
    Dim mCount As Integer
        
        objDB.SetConnection mCnn
    
    mFlag = False
    mNumCount = 0
    
    '   Checking atleast one row is selected from both grids    '
    '---------------------------------------------------------------'
    For mRowCount = 1 To fgVoucherStatement.Rows - 1
        If fgVoucherStatement.Cell(flexcpChecked, mRowCount, 4) = vbChecked Then
            mFlag = True
        End If
    Next
    For mRowCount = 1 To fgBankStatement.Rows - 1
        If fgBankStatement.Cell(flexcpChecked, mRowCount, 4) = vbChecked Then
            mFlag = True
        End If
    Next
    If mFlag = False Then
        MsgBox "Please Check any one of the Flowing from the 2 Grid", vbInformation
        Exit Sub
    End If
    
    '   Checking More than One row Selection in Bank Statements '
    '---------------------------------------------------------------'
    For mRowCount = 1 To fgBankStatement.Rows - 1
        If fgBankStatement.Cell(flexcpChecked, mRowCount, 4) = vbChecked Then
            mNumCount = mNumCount + 1
            mCount = mRowCount
        End If
    Next
    
    If mNumCount > 1 Then
        mSQL = "        Multiple Row Selection in Bank Statements" & vbNewLine
        mSQL = mSQL + "                        Confusion Exists !!! " & vbNewLine
        mSQL = mSQL + " Please Select only ONE Row from Bank Statements"
        MsgBox mSQL, vbCritical
        Exit Sub
    End If
    '---------------------------------------------------------------'
    
    '   Updating faVouchers for the Manual Reconciliation Process   '
    For mRowCount = 1 To fgVoucherStatement.Rows - 1
        If fgVoucherStatement.Cell(flexcpChecked, mRowCount, 4) = vbChecked Then
            mVchrID = fgVoucherStatement.TextMatrix(mRowCount, 7)
            mSQL = "Update faVouchers Set tnyReconciled = 1, dtRealisationDate = " & fgBankStatement.TextMatrix(mCount, 0) & " , numTockenID = " & fgBankStatement.TextMatrix(mCount, 1) & " Where intVoucherID = " & mVchrID
            mCnn.Execute mSQL
        End If
    Next
    '   Updating faBankReconciliationEntries for the Manual Reconciliation Process   '
    For mRowCount = 1 To fgBankStatement.Rows - 1
        If fgBankStatement.Cell(flexcpChecked, mRowCount, 4) = vbChecked Then
            mSQL = "Update faBankReconciliationEntries Set tnyReconciled = 1, intVoucherNo = " & fgVoucherStatement.TextMatrix(1, 1) & " Where intReconciliationID = " & fgBankStatement.TextMatrix(1, 1)
            mCnn.Execute mSQL
        End If
    Next
    
    MsgBox "Reconciled!!!", vbInformation
    Call FormInitialize
    Flag = 1
    Unload Me
End Sub

Private Sub cmdSearch2_Click()
    Call SearchVoucher
End Sub

Private Sub fgVoucherStatement_DblClick()
    Dim No As Variant
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim Rec As New ADODB.Recordset
        objDB.SetConnection mCnn
    mSQL = " Select * from faVouchers Where intVoucheriD = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 7)
    Rec.Open mSQL, mCnn
    No = InputBox(IIf(IsNull(Rec!vchDescription), "No Remarks Entered", Rec!vchDescription), "Do You Want to Change the Instrument NO ? ? ?")
    'If No <> "" Then
    If No > 0 And No <> "" Then
        mSQL = "Update faVouchers Set vchInstrumentNo = '" & No & "' Where intVoucherID = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 7)
        mCnn.Execute mSQL
        MsgBox "Updated Instrument Number", vbInformation
        fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 5) = No
    ElseIf No = 0 Then
        mSQL = "Update faVouchers Set vchInstrumentNo = Null Where intVoucherID = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 7)
        mCnn.Execute mSQL
        MsgBox "Updated Instrument Number", vbInformation
        fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 5) = ""
    End If
    Rec.Close
    mCnn.Close
End Sub

Private Sub Form_Activate()
    Me.Left = 2000
    Me.Top = 2000
    Me.Height = 6630
    Me.Width = 11145
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitSubClassing
    txtInstrumentNo.Text = mInstNO
    txtAmount.Text = mAmt
    Flag = 2
    Call SearchVouchers
    mDumyTotal = Val(txtTotal.Text)
End Sub

Private Sub cmdSearch_Click()
    Dim mSQL As String
    Dim Rec As New ADODB.Recordset
    Dim RecBank As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim mInstrumentNo As Long
    Dim mBal As Double
        
        objDB.SetConnection mCnn
    
    If Val(txtBalanceAmt.Text) = 0 Then
        Exit Sub
    End If
    
    mSQL = "Select * from faVouchers "
    mSQL = mSQL + " Where fltAmount = " & txtBalanceAmt.Text
    mSQL = mSQL + " and DatePart(mm,dtDate) = '" & Month(fgVoucherStatement.TextMatrix(1, 0)) & "'"

    Rec.Open mSQL, mCnn
    
    If Rec.EOF Or Rec.BOF Then
        Exit Sub
    Else
        If IsNull(Rec!vchInstrumentNo) Then
            mBal = Val(Rec!fltAmount)
        Else
            mInstrumentNo = Val(Rec!vchInstrumentNo)
        End If
    End If
    
    If Val(mBal) = 0 Then
        mSQL = "Select * from faBankReconciliationEntries "
        mSQL = mSQL + " Where vchChequeNo = '" & mInstrumentNo & "'"
        mSQL = mSQL + " and DatePart(mm,dtBankEntryDate) = '" & Month(fgVoucherStatement.TextMatrix(1, 0)) & "'"
    Else
        mSQL = "Select * from faBankReconciliationEntries "
        mSQL = mSQL + " Where (fltDrAmount = " & Val(mBal) & " Or fltCrAmount = " & Val(mBal) & ")"
        mSQL = mSQL + " and DatePart(mm,dtBankEntryDate) = '" & Month(fgVoucherStatement.TextMatrix(1, 0)) & "'"
    End If
    RecBank.Open mSQL, mCnn
    
    'If RecBank.EOF Or RecBank.BOF Then
        While Not (Rec.EOF Or Rec.BOF)
            fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
            mLoop1 = mLoop1 + 1
            fgVoucherStatement.TextMatrix(mLoop1, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            fgVoucherStatement.TextMatrix(mLoop1, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            fgVoucherStatement.TextMatrix(mLoop1, 2) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            fgVoucherStatement.TextMatrix(mLoop1, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
            If IsNull(Rec!tnyReconciled) Then
                fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbUnchecked
            Else
                fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbChecked
            End If
            fgVoucherStatement.TextMatrix(mLoop1, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Rec!tnyVoucherTypeID = 10 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "R"
            If Rec!tnyVoucherTypeID = 20 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "P"
            If Rec!tnyVoucherTypeID = 30 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "C"
            If Rec!tnyVoucherTypeID = 40 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "J"
            fgVoucherStatement.TextMatrix(mLoop1, 7) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            Rec.MoveNext
        Wend
    'End If
    
'''''    If RecBank.EOF Or RecBank.BOF Then
'''''        While Not (Rec.EOF Or Rec.BOF)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''            fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
'''''            mRowCount = mRowCount + 1
'''''            Rec.MoveNext
'''''        Wend
'''''    Else
'''''        While Not (Rec.EOF Or Rec.BOF)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'''''            fgVoucherStatement.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'''''            fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
'''''            mRowCount = mRowCount + 1
'''''            Rec.MoveNext
'''''        Wend
'''''
'''''        While Not (RecBank.EOF Or RecBank.BOF)
'''''            fgBankStatement.TextMatrix(mLoop, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
'''''            fgBankStatement.TextMatrix(mLoop, 1) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
'''''            fgBankStatement.TextMatrix(mLoop, 2) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
'''''            fgBankStatement.TextMatrix(mLoop, 3) = IIf(IsNull(Rec!fltDrAmount), IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount), IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount))
'''''            fgBankStatement.Rows = fgVoucherStatement.Rows + 1
'''''            mLoop = mLoop + 1
'''''            Rec.MoveNext
'''''            Wend
'''''    End If
    
End Sub


Private Sub SearchVouchers()

    '----------------------------------------------------'
    ' Aiby
    '----------------------------------------------------'
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
    
    Dim mTotal As Double
    Dim mInstrumentNo As Variant
    Dim mSqlBank As String
    Dim RecBank As New ADODB.Recordset
        
    objDB.SetConnection mCnn
    
    mSQL = "Select * from faVouchers "
    If chkMonth.Value = 0 Then
        mSQL = mSQL + " Where DatePart(mm,dtDate) between '4' and '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    ElseIf chkMonth.Value = 1 Then
        mSQL = mSQL + " Where DatePart(mm,dtDate) = '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    End If
    
    If txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchInstrumentNo = '" & txtInstrumentNo.Text & "' And fltAmount = " & txtAmount.Text
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchInstrumentNo Like '" & txtInstrumentNo.Text & "%' And fltAmount Like " & txtAmount.Text & "%"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchInstrumentNo Like '" & txtInstrumentNo.Text & "%'"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchInstrumentNo = '" & txtInstrumentNo.Text & "'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And fltAmount Like '" & txtAmount.Text & "%'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And fltAmount = " & txtAmount.Text
    End If
    mSQL = mSQL + " And ( intKeyID1 = " & frmBankReconcilationProcess.mSearchID & " or intKeyID1 is Null )"
    
    If Not IsEmpty(mvarVoucherID) Then
        mSQL = ""
        mSQL = mSQL + " Select faVouchers.intVoucherID, tnyVoucherTypeID, intVoucherNo, vchInstrumentNo, dtDate, faVoucherChild.fltAmount, tnyReconciled, numTockenID, vchDescription from faVouchers"
        mSQL = mSQL + " Inner Join faVoucherChild On faVoucherChild.intVoucherID = faVouchers.intVoucherID"
        mSQL = mSQL + " Where faVouchers.intVoucherID = " & mvarVoucherID
    End If
    
    Rec.Open mSQL, mCnn
    fgVoucherStatement.Rows = 2
    mLoop1 = 1
    fgVoucherStatement.AutoResize = True
    While Not (Rec.EOF Or Rec.BOF)
        fgVoucherStatement.TextMatrix(mLoop1, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
        fgVoucherStatement.TextMatrix(mLoop1, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        fgVoucherStatement.TextMatrix(mLoop1, 2) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
        fgVoucherStatement.TextMatrix(mLoop1, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        If IsNull(Rec!tnyReconciled) Then
            fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbUnchecked
        Else
            fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbChecked
        End If
        fgVoucherStatement.TextMatrix(mLoop1, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
        If Rec!tnyVoucherTypeID = 10 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "R"
        If Rec!tnyVoucherTypeID = 20 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "P"
        If Rec!tnyVoucherTypeID = 30 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "C"
        If Rec!tnyVoucherTypeID = 40 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "J"
        fgVoucherStatement.TextMatrix(mLoop1, 7) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
        fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
        mLoop1 = mLoop1 + 1
        Rec.MoveNext
    Wend
    Rec.Close
    
    'If fgVoucherStatement.TextMatrix(1, 0) = "" Then Exit Sub
    
    mSQL = "Select * from faBankReconciliationEntries  "
    If chkMonth.Value = 0 Then
        mSQL = mSQL + " Where DatePart(mm,dtBankEntryDate) Between 4 and " & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex)
    ElseIf chkMonth.Value = 1 Then
        mSQL = mSQL + " Where DatePart(mm,dtBankEntryDate) = '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    End If
    If txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchChequeNo = '" & txtInstrumentNo.Text & "'"
        mSQL = mSQL + " And (fltDrAmount = " & Val(txtAmount.Text) & " Or fltCrAmount = " & Val(txtAmount.Text) & ")"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchChequeNo Like '" & txtInstrumentNo.Text & "%'"
        mSQL = mSQL + " And (fltDrAmount Like " & Val(txtAmount.Text) & "% Or fltCrAmount Like " & Val(txtAmount.Text) & "%)"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchChequeNo Like '" & txtInstrumentNo.Text & "%'"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchChequeNo = '" & txtInstrumentNo.Text & "'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And (fltDrAmount Like '" & Val(txtAmount.Text) & "%' Or fltCrAmount Like '" & Val(txtAmount.Text) & "%')"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And (fltDrAmount = " & Val(txtAmount.Text) & " Or fltCrAmount = " & Val(txtAmount.Text) & ")"
    End If
        
    Rec.Open mSQL, mCnn
    fgBankStatement.Rows = 2
    mLoop2 = 1
    While Not (Rec.EOF Or Rec.BOF)
        fgBankStatement.TextMatrix(mLoop2, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
        fgBankStatement.TextMatrix(mLoop2, 1) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
        fgBankStatement.TextMatrix(mLoop2, 2) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
        fgBankStatement.TextMatrix(mLoop2, 3) = IIf(IsNull(Rec!fltDrAmount), IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount), IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount))
        If IsNull(Rec!tnyReconciled) Then
            fgBankStatement.Cell(flexcpChecked, mLoop2, 4) = vbUnchecked
        Else
            fgBankStatement.Cell(flexcpChecked, mLoop2, 4) = vbChecked
        End If
        fgBankStatement.TextMatrix(mLoop2, 5) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
        
        If IsNull(Rec!fltDrAmount) Then
            fgBankStatement.TextMatrix(mLoop2, 6) = "C"
        End If
        If IsNull(Rec!fltCrAmount) Then
            fgBankStatement.TextMatrix(mLoop2, 6) = "D"
        End If
        
        '------------------------------------------------------------------------------------------------'
        'Added By Aiby on 24-Jan-2008 : Reason: While Porting Treasury Accounts Amount will not be Null  '
        '------------------------------------------------------------------------------------------------'
        If IsNumeric(Rec!fltDrAmount) Then
            If Rec!fltDrAmount > 0 Then
                fgBankStatement.TextMatrix(mLoop2, 6) = "D"
            Else
                fgBankStatement.TextMatrix(mLoop2, 6) = "C"
            End If
        End If
        '------------------------------------------------------------------------------------------------'
        
        
        fgBankStatement.Rows = fgBankStatement.Rows + 1
        mLoop2 = mLoop2 + 1
        Rec.MoveNext
    Wend
    Rec.Close
    
    mTotal = 0
    For mLoop1 = 1 To fgVoucherStatement.Rows - 2
        mTotal = mTotal + fgVoucherStatement.TextMatrix(mLoop1, 3)
    Next
    txtTotal.Text = mTotal
    txtBalanceAmt.Text = Val(fgBankStatement.TextMatrix(1, 3)) - Val(mTotal)
    If Val(txtBalanceAmt.Text) < 0 Then txtBalanceAmt.Text = Val(txtBalanceAmt.Text) * -1
End Sub



Private Sub SearchVoucher()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
    
    Dim mTotal As Double
    Dim mInstrumentNo As Variant
    Dim mSqlBank As String
    Dim RecBank As New ADODB.Recordset
        
        objDB.SetConnection mCnn
    
    mSQL = "Select * from faVouchers "
    If chkMonth.Value = 0 Then
        mSQL = mSQL + " Where DatePart(mm,dtDate) between '4' and '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    ElseIf chkMonth.Value = 1 Then
        mSQL = mSQL + " Where DatePart(mm,dtDate) = '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    End If
    If txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchInstrumentNo = '" & txtInstrumentNo.Text & "' And fltAmount = " & txtAmount.Text
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchInstrumentNo Like '" & txtInstrumentNo.Text & "%' And fltAmount Like " & txtAmount.Text & "%"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchInstrumentNo Like '" & txtInstrumentNo.Text & "%'"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchInstrumentNo = '" & txtInstrumentNo.Text & "'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And fltAmount Like '" & txtAmount.Text & "%'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And fltAmount = " & txtAmount.Text
    End If
    mSQL = mSQL + " And ( intKeyID1 = " & frmBankReconcilationProcess.mSearchID & " or intKeyID1 is Null )"
    Rec.Open mSQL, mCnn
    fgVoucherStatement.Rows = 2
    mLoop1 = 1
    fgVoucherStatement.AutoResize = True
    While Not (Rec.EOF Or Rec.BOF)
        fgVoucherStatement.TextMatrix(mLoop1, 0) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
        fgVoucherStatement.TextMatrix(mLoop1, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        fgVoucherStatement.TextMatrix(mLoop1, 2) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
        fgVoucherStatement.TextMatrix(mLoop1, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        If IsNull(Rec!tnyReconciled) Then
            fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbUnchecked
        Else
            fgVoucherStatement.Cell(flexcpChecked, mLoop1, 4) = vbChecked
        End If
        fgVoucherStatement.TextMatrix(mLoop1, 5) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
        If Rec!tnyVoucherTypeID = 10 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "R"
        If Rec!tnyVoucherTypeID = 20 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "P"
        If Rec!tnyVoucherTypeID = 30 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "C"
        If Rec!tnyVoucherTypeID = 40 Then fgVoucherStatement.TextMatrix(mLoop1, 6) = "J"
        fgVoucherStatement.TextMatrix(mLoop1, 7) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
        fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
        mLoop1 = mLoop1 + 1
        Rec.MoveNext
    Wend
    Rec.Close
    
    'If fgVoucherStatement.TextMatrix(1, 0) = "" Then Exit Sub
    
    mSQL = "Select * from faBankReconciliationEntries  "
    If chkMonth.Value = 0 Then
        mSQL = mSQL + " Where DatePart(mm,dtBankEntryDate) Between 4 and " & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex)
    ElseIf chkMonth.Value = 1 Then
        mSQL = mSQL + " Where DatePart(mm,dtBankEntryDate) = '" & frmBankReconcilationProcess.cmbMonth.ItemData(frmBankReconcilationProcess.cmbMonth.ListIndex) & "'"
    End If
    If txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchChequeNo = '" & txtInstrumentNo.Text & "'"
        mSQL = mSQL + " And (fltDrAmount = " & Val(txtAmount.Text) & " Or fltCrAmount = " & Val(txtAmount.Text) & ")"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchChequeNo Like '" & txtInstrumentNo.Text & "%'"
        mSQL = mSQL + " And (fltDrAmount Like " & Val(txtAmount.Text) & "% Or fltCrAmount Like " & Val(txtAmount.Text) & "%)"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And vchChequeNo Like '" & txtInstrumentNo.Text & "%'"
    ElseIf txtInstrumentNo.Text <> "" And txtAmount.Text = "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And vchChequeNo = '" & txtInstrumentNo.Text & "'"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 1 Then
        mSQL = mSQL + " And (fltDrAmount Like '" & Val(txtAmount.Text) & "%' Or fltCrAmount Like '" & Val(txtAmount.Text) & "%')"
    ElseIf txtInstrumentNo.Text = "" And txtAmount.Text <> "" And chkDeepSearching.Value = 0 Then
        mSQL = mSQL + " And (fltDrAmount = " & Val(txtAmount.Text) & " Or fltCrAmount = " & Val(txtAmount.Text) & ")"
    End If
        
    Rec.Open mSQL, mCnn
    fgBankStatement.Rows = 2
    mLoop2 = 1
    While Not (Rec.EOF Or Rec.BOF)
        fgBankStatement.TextMatrix(mLoop2, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
        fgBankStatement.TextMatrix(mLoop2, 1) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
        fgBankStatement.TextMatrix(mLoop2, 2) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
        fgBankStatement.TextMatrix(mLoop2, 3) = IIf(IsNull(Rec!fltDrAmount), IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount), IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount))
        If IsNull(Rec!tnyReconciled) Then
            fgBankStatement.Cell(flexcpChecked, mLoop2, 4) = vbUnchecked
        Else
            fgBankStatement.Cell(flexcpChecked, mLoop2, 4) = vbChecked
        End If
        fgBankStatement.TextMatrix(mLoop2, 5) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
        
        If IsNull(Rec!fltDrAmount) Then
            fgBankStatement.TextMatrix(mLoop2, 6) = "C"
        End If
        If IsNull(Rec!fltCrAmount) Then
            fgBankStatement.TextMatrix(mLoop2, 6) = "D"
        End If
        
        '------------------------------------------------------------------------------------------------'
        'Added By Aiby on 24-Jan-2008 : Reason: While Porting Treasury Accounts Amount will not be Null  '
        '------------------------------------------------------------------------------------------------'
        If IsNumeric(Rec!fltDrAmount) Then
            If Rec!fltDrAmount > 0 Then
                fgBankStatement.TextMatrix(mLoop2, 6) = "D"
            Else
                fgBankStatement.TextMatrix(mLoop2, 6) = "C"
            End If
        End If
        '------------------------------------------------------------------------------------------------'
        
        
        fgBankStatement.Rows = fgBankStatement.Rows + 1
        mLoop2 = mLoop2 + 1
        Rec.MoveNext
    Wend
    Rec.Close
    
    mTotal = 0
    For mLoop1 = 1 To fgVoucherStatement.Rows - 2
        mTotal = mTotal + fgVoucherStatement.TextMatrix(mLoop1, 3)
    Next
    txtTotal.Text = mTotal
    txtBalanceAmt.Text = Val(fgBankStatement.TextMatrix(1, 3)) - Val(mTotal)
    If Val(txtBalanceAmt.Text) < 0 Then txtBalanceAmt.Text = Val(txtBalanceAmt.Text) * -1
End Sub

Private Sub SearchBankStatements()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim mRowCount As Long
        objDB.SetConnection mCnn
    mSQL = "Select * from faBankReconciliationEntries Where vchChequeNo = '" & txtInstrumentNo.Text & "'"
    Rec.Open mSQL, mCnn
    fgBankStatement.Rows = 2
    mRowCount = 1
    While Not (Rec.EOF Or Rec.BOF)
        fgBankStatement.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
        fgBankStatement.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
        fgBankStatement.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
        'fgBankStatement.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        fgBankStatement.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltDrAmount), IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount), IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount))
        'fgBankStatement.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
        fgBankStatement.Rows = fgVoucherStatement.Rows + 1
        mRowCount = mRowCount + 1
        Rec.MoveNext
    Wend
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormInitialize
    If Flag <> 1 Then
        Flag = 2
    End If
End Sub

Private Sub optAmount_Click()
    If optCheque.Value = True Then
        txtAmount.Text = ""
        txtInstrumentNo.Text = mInstNO
        Call SearchVoucher
    End If
    If optAmount.Value = True Then
        txtInstrumentNo.Text = ""
        If mDumyTotal = "" Then Exit Sub
        txtAmount.Text = Val(mDumyTotal)
        Call SearchVoucher
    End If
End Sub

Private Sub optCheque_Click()
    If optCheque.Value = True Then
        txtAmount.Text = ""
        If mInstNO = "" Then Exit Sub
        txtInstrumentNo.Text = mInstNO
        Call SearchVoucher
    End If
    If optAmount.Value = True Then
        txtInstrumentNo.Text = ""
        txtAmount.Text = Val(mDumyTotal)
        Call SearchVoucher
    End If
End Sub

Private Sub txtAmount_LostFocus()
    If txtAmount.Text <> "" Then
        'txtInstrumentNo.Text = ""
        Call SearchVoucher
    End If
End Sub

Private Sub txtInstrumentNo_LostFocus()
    If txtInstrumentNo.Text <> "" Then
        Call SearchVoucher
        'Call SearchBankStatements
    End If
End Sub

Private Sub SearchVoucherWithAmount()
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim mSQL As String
    Dim mRowCount As Long
        objDB.SetConnection mCnn
    mSQL = "Select * from faVouchers "
    mSQL = mSQL + " Where fltAmount = " & txtBalanceAmt.Text
    mSQL = mSQL + " and DateName(mm,dtDate) = '" & fgVoucherStatement.TextMatrix(1, 0) & "'"
    Rec.Open mSQL, mCnn
    While Not (Rec.EOF Or Rec.BOF)
        
    Wend
End Sub


    Public Property Let VoucherID(mVID As Variant)
        mvarVoucherID = mVID
    End Property
    
    Public Property Get VoucherID() As Variant
        VoucherID = mvarVoucherID
    End Property
    

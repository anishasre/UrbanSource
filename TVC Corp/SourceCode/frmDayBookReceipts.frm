VERSION 5.00
Begin VB.Form frmDayBookReceipts 
   BackColor       =   &H00EBF7F7&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Day Book"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2655
      TabIndex        =   9
      Top             =   1620
      Width           =   3120
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EBF7F7&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2505
      Width           =   1530
   End
   Begin VB.ComboBox cmbCounters 
      Height          =   315
      Left            =   2670
      TabIndex        =   6
      Top             =   1230
      Width           =   3120
   End
   Begin VB.TextBox txtToDate 
      Height          =   285
      Left            =   4410
      TabIndex        =   4
      Top             =   870
      Width           =   1365
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00EBF7F7&
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2505
      Width           =   1530
   End
   Begin VB.TextBox txtFromDate 
      Height          =   285
      Left            =   2670
      TabIndex        =   1
      Top             =   870
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   1650
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Counter"
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
      Left            =   2025
      TabIndex        =   7
      Top             =   1275
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   135
      Picture         =   "frmDayBookReceipts.frx":0000
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Counter"
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
      Left            =   -1500
      TabIndex        =   5
      Top             =   -1260
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   4170
      TabIndex        =   3
      Top             =   930
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   2205
      TabIndex        =   0
      Top             =   900
      Width           =   360
   End
End
Attribute VB_Name = "frmDayBookReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillCombo()
    Dim mSQL As String
    mSQL = " Select Distinct faCounters.vchDescription, faCounters.intCounterID From faVouchers Inner Join"
    mSQL = mSQL + " faCOunters On faCounters.intCounterID = faVouchers.intCounterID"
    mSQL = mSQL + " Where dtDate Between '" & txtFromDate.Text & "'  AND  '" & txtToDate.Text & "'"
    Call PopulateList(cmbCounters, mSQL, , True, True, True)
End Sub

Private Sub cmdShow_Click()
    Dim mSQL As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDb As New clsDB
    Dim mCounterID As Long
    Dim mDailyTotal As Double
    Dim mDt As Date
    Dim mName As String
    Dim mChequeNo As String
    Dim mChequeDate As Date
    Dim mRef As String
    Dim mTransactionType As String
    Dim mCancelFlag As Integer
    
    
    objDb.SetConnection mCnn
    mSQL = "Select * From faVouchers Where dtDate Between '" & txtToDate.Text & "' And '" & txtFromDate & "'"
    
    mSQL = " Select * From faVouchers Left Join "
    mSQL = mSQL + " faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID Inner Join"
    mSQL = mSQL + " faCounters On faCounters.intCounterID = faVouchers.intCounterID Left Join"
    mSQL = mSQL + " faTransactionType On faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID"
    mSQL = mSQL + " Where dtDate Between '" & txtFromDate & "'  AND  '" & txtToDate & "'"
    If cmbCounters.ListIndex > 0 Then
        mSQL = mSQL + " And faVouchers.intCounterID = " & cmbCounters.itemData(cmbCounters.ListIndex)
    End If
    mSQL = mSQL + " Order By faVouchers.intCounterID, faVouchers.dtDate, faVouchers.intVoucherID "
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
    If Not (Rec.BOF And Rec.EOF) Then
        FileInitialize
        mDt = Rec!dtDate
        While Not Rec.EOF
            'On Error Resume Next
            If mCounterID <> Rec!intCounterID Then
                mCounterID = Rec!intCounterID
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO, "Counter :" & Rec!intCounterID; " - "; Rec!vchDescription
                Print #gbFileNO, "===================================================================================================== "
                Print #gbFileNO, "Date        Voucher No.        Amount  Name                  ChequeNo\Date        Ref.     Tran.Type "
                Print #gbFileNO, "-----------------------------------------------------------------------------------------------------"
            End If
            Print #gbFileNO, DdMmmYy(Rec!dtDate); PadL(Rec!intVoucherNo, 12); "  ";
            If Not IsNull(Rec!tnyCancelFlag) Then
                mCancelFlag = Rec!tnyCancelFlag
            Else
                mCancelFlag = 0
            End If
            If mCancelFlag Then
                Print #gbFileNO, "*** Canceled Voucher ***"
            Else
                Print #gbFileNO, PadL(Format(Rec!fltAmount, "0.00"), 12); "  ";
                mDailyTotal = mDailyTotal + Format(Rec!fltAmount, "0.00")
                
                mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                If Not IsNull(Rec!vchInstrumentNo) Then
                    mChequeNo = Rec!vchInstrumentNo
                    If IsDate(Rec!dtInstrumentDate) Then
                        mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                    End If
                End If
                
                mRef = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                mTransactionType = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                mSQL = ""
                mSQL = IIf(IsNull(Rec.Fields(12).Value), "", Rec.Fields(12).Value)
                Print #gbFileNO, PadR(mName, 20); "  ";
                Print #gbFileNO, PadR(mChequeNo, 20); " ";
                Print #gbFileNO, PadR(mRef, 10); " ";
                Print #gbFileNO, PadR(mTransactionType, 10);
                Print #gbFileNO, PadR(mSQL, 50); " ";
                Print #gbFileNO, PadR(Rec!intVoucherNo, 15)
                mChequeNo = ""
            End If
            Rec.MoveNext
            If Not Rec.EOF Then
                If mDt <> Rec!dtDate Then
                    mDt = Rec!dtDate
PrintTotals:
                    Print #gbFileNO, "                       --------------"
                    Print #gbFileNO, PadL(Format(mDailyTotal, "0.00"), 37)
                    Print #gbFileNO, "                       =============="
                End If
            Else
                GoTo PrintTotals:
            End If
        Wend
        Close #gbFileNO
        ShellPad
    End If 'If Not (Rec.BOF And Rec.EOF) Then
    
End Sub
    
    Private Sub Form_Activate()
        Me.Left = (frmMenu.Width - Me.Width) / 2
        Me.Top = 2500
    End Sub
    Private Sub Form_Load()
        txtFromDate.Text = DdMmmYy(gbTransactionDate)
        txtToDate.Text = DdMmmYy(gbTransactionDate)
        Call FillCombo
    End Sub
    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub
    
    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate)
        Call FillCombo
    End Sub
    Private Sub txtTodate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate)
        Call FillCombo
    End Sub

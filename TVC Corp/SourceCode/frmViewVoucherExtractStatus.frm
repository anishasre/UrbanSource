VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmViewVoucherExtractStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher Extract Status"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   Icon            =   "frmViewVoucherExtractStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   2025
      Width           =   8925
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2850
         Left            =   900
         TabIndex        =   7
         Top             =   225
         Width           =   6675
         _cx             =   11774
         _cy             =   5027
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
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
         Rows            =   7
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmViewVoucherExtractStatus.frx":1CCA
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
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
      TabIndex        =   3
      Top             =   675
      Width           =   8925
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<<"
         Height          =   285
         Left            =   4095
         TabIndex        =   12
         Top             =   855
         Width           =   375
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   285
         Left            =   6255
         TabIndex        =   11
         Top             =   855
         Width           =   375
      End
      Begin VB.TextBox txtCurrentMnth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   9
         Top             =   855
         Width           =   1725
      End
      Begin VB.TextBox txtApril 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         TabIndex        =   8
         Text            =   "APRIL"
         Top             =   855
         Width           =   1275
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8730
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ComboBox cmbYear 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "MONTH"
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
         Left            =   8505
         TabIndex        =   5
         Top             =   855
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "YEAR"
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
         Left            =   8460
         TabIndex        =   4
         Top             =   90
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   8925
      TabIndex        =   0
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "frmViewVoucherExtractStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub cmbYear_Click()
        txtCurrentMnth.Text = MonthName(Month(gbTransactionDate))
        txtCurrentMnth.Tag = Month(gbTransactionDate)
        Call FillGrid
    End Sub

    Private Sub cmdNext_Click()
        If txtCurrentMnth.Tag <> 3 Then
            If txtCurrentMnth.Tag = 12 Then
                txtCurrentMnth.Tag = 1
            Else
                txtCurrentMnth.Tag = txtCurrentMnth.Tag + 1
            End If
            txtCurrentMnth.Text = MonthName(txtCurrentMnth.Tag)
         End If
         Call FillGrid
    End Sub
    
    Private Sub cmdPrevious_Click()
        If txtCurrentMnth.Tag <> 4 Then
            If txtCurrentMnth.Tag = 1 Then
                txtCurrentMnth.Tag = 12
            Else
                txtCurrentMnth.Tag = txtCurrentMnth.Tag - 1
            End If
            txtCurrentMnth.Text = MonthName(txtCurrentMnth.Tag)
         End If
         Call FillGrid
    End Sub

    Private Sub Form_Load()
        Call FillYear
        cmbYear.Text = "2013"
        txtCurrentMnth.Text = MonthName(Month(gbTransactionDate))
        txtCurrentMnth.Tag = Month(gbTransactionDate)
        Call FillGrid
    End Sub
'''    Private Sub FillMonth()
'''
'''        cmbMonth.AddItem "April", 0
'''        cmbMonth.ItemData(0) = 4
'''
'''        cmbMonth.AddItem "May", 1
'''        cmbMonth.ItemData(1) = 5
'''
'''        cmbMonth.AddItem "June", 2
'''        cmbMonth.ItemData(2) = 6
'''
'''        cmbMonth.AddItem "July", 3
'''        cmbMonth.ItemData(3) = 7
'''
'''        cmbMonth.AddItem "August", 4
'''        cmbMonth.ItemData(4) = 8
'''
'''        cmbMonth.AddItem "September", 5
'''        cmbMonth.ItemData(5) = 9
'''
'''        cmbMonth.AddItem "October", 6
'''        cmbMonth.ItemData(6) = 10
'''
'''        cmbMonth.AddItem "November", 7
'''        cmbMonth.ItemData(7) = 11
'''
'''        cmbMonth.AddItem "December", 8
'''        cmbMonth.ItemData(8) = 12
'''
'''        cmbMonth.AddItem "January", 9
'''        cmbMonth.ItemData(9) = 1
'''
'''        cmbMonth.AddItem "February", 10
'''        cmbMonth.ItemData(10) = 2
'''
'''        cmbMonth.AddItem "March", 11
'''        cmbMonth.ItemData(11) = 3
'''    End Sub
    Private Sub FillYear()
        
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mDate As Date
        Dim mYearID As Variant
        Dim mCount, i As Integer
    
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQL = " SELECT   dtDate, intFinancialYearID mYear FROM faVouchers  WHERE intTransactionTypeID=3000"
        
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mYearID = Rec!mYear
        End If
        Rec.Close
        
        mCount = 1

        mYearID = mYearID + 1
        For i = mYearID To gbFinancialYearID
            cmbYear.AddItem i
            cmbYear.ItemData(cmbYear.NewIndex) = i
            mCount = mCount + 1
        Next i
        
        mCnn.Close
    End Sub
    
Private Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
   
    
    On Error GoTo err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " SELECT  COUNT(R)-COUNT(CR) 'R',COUNT(P) 'P',COUNT(C) 'C',COUNT(J) 'J',COUNT(CR) 'CancelledReceipts'"
    mSQL = mSQL + " From"
    mSQL = mSQL + " ("
    mSQL = mSQL + " SELECT  CASE WHEN tnyVoucherTypeID=10 THEN COUNT(faVouchers.intVoucherID)END AS 'R',"
    mSQL = mSQL + "     CASE WHEN tnyVoucherTypeID=20 THEN COUNT(faVouchers.intVoucherID)END AS 'P',"
    mSQL = mSQL + "     CASE WHEN tnyVoucherTypeID=30 THEN COUNT(faVouchers.intVoucherID)END AS 'C',"
    mSQL = mSQL + "     CASE WHEN tnyVoucherTypeID=40 THEN COUNT(faVouchers.intVoucherID)END AS 'J',"
    mSQL = mSQL + "     CASE WHEN tnyVoucherTypeID=10 AND tnyCancelFlag=1 THEN COUNT(faVouchers.intVoucherID)END AS 'CR'"
    mSQL = mSQL + " FROM  faVouchers"
    mSQL = mSQL + " WHERE intFinancialYearID =  " & cmbYear.Text
    mSQL = mSQL + " AND MONTH(dtDate)< = " & val(txtCurrentMnth.Tag)
    mSQL = mSQL + " GROUP BY  tnyVoucherTypeID,faVouchers.intVoucherID,tnyCancelFlag"
    mSQL = mSQL + " )A"
 
    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(0, 1) = IIf(IsNull(Rec!R), "0", Rec!R)
        vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!CancelledReceipts), "0", Rec!CancelledReceipts)
        vsGrid.TextMatrix(2, 1) = IIf(IsNull(Rec!P), "0", Rec!P)
        vsGrid.TextMatrix(3, 1) = IIf(IsNull(Rec!c), "0", Rec!c)
        vsGrid.TextMatrix(4, 1) = IIf(IsNull(Rec!J), "0", Rec!J)
    End If
    Rec.Close
    
    mSQL = ""
    mSQL = " SELECT SUM(DEBIT) DebitTot,SUM(CREDIT) CreditTot"
    mSQL = mSQL + " From"
    mSQL = mSQL + " ("
    mSQL = mSQL + " SELECT  CASE WHEN  tinDebitOrCreditFlag=0 THEN SUM(faTransactionChild.fltAmount) END AS 'DEBIT',"
    mSQL = mSQL + "     CASE WHEN  tinDebitOrCreditFlag=1 THEN SUM(faTransactionChild.fltAmount) END AS 'CREDIT'"
    mSQL = mSQL + " From faTransactions"
    mSQL = mSQL + " INNER JOIN faTransactionChild ON faTransactions.intTransactionID=faTransactionChild.intTransactionID"
    mSQL = mSQL + " Where intFinancialYearID = " & cmbYear.Text
    mSQL = mSQL + " AND MONTH(dtTransactionDate)< =" & val(txtCurrentMnth.Tag)
    mSQL = mSQL + " GROUP BY  intVoucherID,tinDebitOrCreditFlag"
    mSQL = mSQL + " )A"

    Rec.Open mSQL, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        vsGrid.TextMatrix(5, 1) = IIf(IsNull(Rec!DebitTot), "0", Rec!DebitTot)
        vsGrid.TextMatrix(6, 1) = IIf(IsNull(Rec!CreditTot), "0", Rec!CreditTot)
    End If
    Rec.Close
    Exit Sub
err:
    MsgBox err.Description
End Sub

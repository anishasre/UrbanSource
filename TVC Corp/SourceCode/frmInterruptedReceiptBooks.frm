VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmInterruptedReceiptBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Issue of Interrupted Receipt Books "
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterruptedReceiptBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbYear 
      Height          =   390
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1395
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3817
      TabIndex        =   6
      Top             =   5085
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5152
      TabIndex        =   8
      Top             =   5070
      Width           =   1275
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2482
      TabIndex        =   7
      Top             =   5070
      Width           =   1275
   End
   Begin VB.TextBox txtRecTo 
      Alignment       =   1  'Right Justify
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
      Left            =   5827
      MaxLength       =   10
      TabIndex        =   3
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox txtRecFrom 
      Alignment       =   1  'Right Justify
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
      Left            =   4470
      MaxLength       =   10
      TabIndex        =   2
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox txtReason 
      Enabled         =   0   'False
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
      Left            =   2685
      TabIndex        =   10
      Top             =   1920
      Width           =   4350
   End
   Begin VB.CheckBox chkClosed 
      Caption         =   "Closed"
      Enabled         =   0   'False
      Height          =   270
      Left            =   90
      TabIndex        =   9
      Top             =   1980
      Width           =   975
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2595
      Left            =   90
      TabIndex        =   5
      Top             =   2400
      Width           =   8760
      _cx             =   15452
      _cy             =   4577
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
      BackColorBkg    =   -2147483633
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
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInterruptedReceiptBooks.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
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
   Begin VB.TextBox txtNoOfReceipts 
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
      Left            =   6247
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   780
   End
   Begin VB.TextBox txtBookNo 
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
      Left            =   2685
      MaxLength       =   5
      TabIndex        =   1
      Top             =   900
      Width           =   1170
   End
   Begin VB.ComboBox cmbCounter 
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
      Left            =   2685
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   4365
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year"
      Height          =   300
      Left            =   1170
      TabIndex        =   18
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   1635
      Left            =   75
      Top             =   180
      Width           =   8820
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Count"
      Height          =   315
      Left            =   4920
      TabIndex        =   16
      Top             =   1485
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   5865
      TabIndex        =   15
      Top             =   630
      Width           =   315
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      Enabled         =   0   'False
      Height          =   270
      Left            =   1920
      TabIndex        =   14
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblNoOfReceipts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Receipts From"
      Height          =   270
      Left            =   4290
      TabIndex        =   13
      Top             =   630
      Width           =   1320
   End
   Begin VB.Label lblBookNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book No"
      Height          =   270
      Left            =   1905
      TabIndex        =   12
      Top             =   900
      Width           =   705
   End
   Begin VB.Label lblCounter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Counter"
      Height          =   270
      Left            =   1905
      TabIndex        =   11
      Top             =   315
      Width           =   690
   End
End
Attribute VB_Name = "frmInterruptedReceiptBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mBookIDs As String      '' To Get All Book IDs in the Vouchers According to a Book Number.
    Dim mLastPostingYearID As Integer
    '*********************************************************************************************'
    '               Form to issue Interrupt Receipt Book to a particular counter                  '
    '*********************************************************************************************'
    Private Sub FormInitialize()
        cmbCounter.ListIndex = 0
        '----------ADDED ON 08/05/2013 BY MINU--------------
        cmbYear.ListIndex = -1
        '---------------------------------------------------
        txtBookNo.Text = ""
        txtBookNo.Tag = ""
        txtRecFrom.Text = ""
        txtRecTo.Text = ""
        txtNoOfReceipts.Text = ""
        chkClosed.Value = 0
        chkClosed.Enabled = False
        txtReason.Text = ""
        txtReason.Enabled = False
        cmbCounter.Enabled = True
        txtBookNo.Enabled = True
        txtRecFrom.Enabled = True
        txtRecTo.Enabled = True
        cmbYear.Enabled = True
        mBookIDs = ""
        txtNoOfReceipts.ForeColor = vbBlack
    End Sub
    
    Private Sub FillvsGrid(Rec As ADODB.Recordset)
        Dim mRowCount       As Double
        Dim mStatus         As Variant
        
        vsGrid.Clear 1, 1
        vsGrid.Rows = 2
        mRowCount = 2
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
            vsGrid.TextMatrix(mRowCount, 2) = CStr(IIf(IsNull(Rec!numReceiptNoFrom), "", Rec!numReceiptNoFrom)) + " - " + CStr(IIf(IsNull(Rec!numReceiptNoTo), "", Rec!numReceiptNoTo))
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intCount), "", Rec!intCount)
            mStatus = IIf(IsNull(Rec!tnyClosed), "", Rec!tnyClosed)
            If mStatus <> "" Then
                If mStatus = 0 Then
                    vsGrid.TextMatrix(mRowCount, 5) = "Open"
                End If
                If mStatus = 1 Then
                    vsGrid.TextMatrix(mRowCount, 5) = "Closed"
                End If
            End If
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!intCounterID), "", Rec!intCounterID)
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intBookID), "", Rec!intBookID)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!tnyClosed), "", Rec!tnyClosed)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
            vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!numReceiptNoFrom), "", Rec!numReceiptNoFrom)
            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!intFinancialYearID), "", Rec!intFinancialYearID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
    End Sub
    
    Private Sub chkClosed_Click()
        If chkClosed.Value = 1 Then
            txtReason.Enabled = True
            txtReason.SetFocus
        Else
            txtReason.Enabled = False
        End If
    End Sub

    Private Sub cmbYear_LostFocus()
        '-----------------LAST POSTING VALIDATION------------------
        If cmbYear.Text <= mLastPostingYearID Then
            MsgBox "Transactions Locked !!!No More Transactions Is Possible for Current Date And less", vbInformation
            Exit Sub
        End If
        '-------------------------------------------------------------
        
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdSave.Enabled = True
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mArray  As Variant
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        Dim mReceiptNo      As String
        Dim mMaxReceipts    As Double
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '*********************************************************************************************'
        '        Procedure to issue Interrupt Receipt Book to a particular counter                    '
        '*********************************************************************************************'
        If cmbCounter.ListIndex = 0 Then
            MsgBox "Please select a Counter", vbInformation
            Exit Sub
        End If
         
        '-------------ADDED ON 08/05/2013 BY MINU------------------
        If cmbYear.ListIndex = -1 Then
            MsgBox "Please select the Financial Year", vbInformation
            Exit Sub
        End If
        '-----------------------------------------------------------
 
        If txtBookNo.Text = "" Then
            MsgBox "Please enter the Book No", vbInformation
            Exit Sub
        End If
        If txtNoOfReceipts.Text = "" Then
            MsgBox "Please enter the No.of Receipts", vbInformation
            Exit Sub
        End If
        If val(Trim(txtRecFrom.Text)) > val(Trim(txtRecTo.Text)) Then
            MsgBox "Please Check the Receipt's No"
            Exit Sub
        End If
        '''''' Checking whether there is multiple entries for the Same Counter '''''''
        mSql = "Select isNull(count(*),0)[Count] From faInterruptedReceiptBooks Where intBookID <> " & val(Trim(txtBookNo.Tag)) & " And tnyClosed <> 1 And intCounterID = " & cmbCounter.ItemData(cmbCounter.ListIndex)
        Rec.Open mSql, mCnn
        If Rec!count > 0 Then
            MsgBox "Already a book issued to this Counter", vbInformation
            Exit Sub
        End If
        Rec.Close
        '''''' Coinciding Receipt Number Validation ''''''
        If gbManualReceiptNewBool = False Then
            mReceiptNo = "9" + Right("00000" + CStr(txtBookNo.Text), 4) + "1" + Right("00000" + Trim(txtRecFrom.Text), 5)
        Else
            ''' New Validation
            mReceiptNo = "9" + Right("00" + CStr(gbFinancialYearID), 2) + Right("00" + CStr(gbCounterNo), 2) + Right("000000" + Trim(txtRecFrom.Text), 6)
        End If
        
        '''' NOTE: THIS CODE IS BLOCKED BY AIBY ON 31-10-2011
        ''''mSql = "Select  Case When isNull(max(intVoucherNo),0)<=isNull(Max(intReceiptNo),0) Then " & _
        ''''        "isNull(Max(intReceiptNo),0)Else isNull(max(intVoucherNo),0)End [ReceiptsNo] " & _
        ''''        "From faVouchers V Left Join faInterruptedCancelledReceipts FICR On V.intVoucherNo = FICR.intReceiptNo " & _
        ''''        "Where tnyVoucherGroupID = 4 And V.intFinancialYearID = " & gbFinancialYearID ''Maximum of Receipt Number Generated
        ''''Rec.Open mSql, mCnn
        ''''mMaxReceipts = Rec!ReceiptsNo
        ''''Rec.Close
        ''''mSql = "Select * From faInterruptedReceiptBooks Where intBookID <> " & val(txtBookNo.Tag) & _
        ''''    " And(" & txtRecFrom.Text & " Between isNull(numReceiptNofrom,0) and isNull(numReceiptNoTo,0)" & _
        ''''    "Or " & txtRecTo.Text & " Between isNull(numReceiptNofrom,0) and isNull(numReceiptNoTo,0)) And intFinancialYearID = " & gbFinancialYearID  '''' Ovelapping Validation
        ''''Rec.Open mSql, mCnn
        ''''If Not (Rec.EOF And Rec.BOF) Then
        ''''    If Rec!tnyClosed = 1 Then
        ''''        If mMaxReceipts >= val(mReceiptNo) Then
        ''''            MsgBox "Please Check the Receipt Number", vbInformation
        ''''            Exit Sub
        ''''        End If
        ''''    Else
        ''''        MsgBox "Please Check the Receipt Number ..", vbInformation
        ''''        Exit Sub
        ''''    End If
        ''''End If
        ''''Rec.Close   '''''' Coinciding Receipt Number Validation Over ''''''
        ''''
        
        '''' NOTE: NEW CODE ADDED : AIBY ON 31-10-2011
        '''' VALIDATION is If the same Book tried to issue its Receipt Number wont need to over lap
        '    Checking with Last Issues Receipt Number with New books Receipt Starting Number:
        
        ''----Commented By Anisha On 25-8-12
'''        mSql = "Select numReceiptNoFrom, numReceiptNoTo From faInterruptedReceiptBooks WHERE intBOOKID IN ("
'''        mSql = mSql + " Select ISNULL(Max(intBookID), 0) From faInterruptedReceiptBooks WHERE intBookNO = " & Trim(txtBookNo.Text) & " )"
'''        Rec.Open mSql, mCnn
'''        If Not (Rec.EOF And Rec.BOF) Then
'''            If val(txtRecFrom) <= Rec!numReceiptNoTo Then
'''                mSql = "The Same Book Number with Starting No " & Rec!numReceiptNoFrom & " to " & vbCrLf
'''                mSql = mSql + str(Rec!numReceiptNoTo) & " is already issued!"
'''                MsgBox mSql, vbCritical
'''                txtRecFrom.SetFocus
'''                Rec.Close
'''                Exit Sub
'''            End If
'''        End If
'''        Rec.Close
            ''-----------
        If chkClosed.Value = 1 Then
            chkClosed.Tag = 1
        Else
            chkClosed.Tag = 0
            txtReason.Text = ""
        End If
    
        mArray = Array(cmbCounter.ItemData(cmbCounter.ListIndex), _
                    txtBookNo.Text, _
                    txtNoOfReceipts, _
                    gbUserID, _
                    txtReason, _
                    cmbYear.Text, _
                    Date, _
                    chkClosed.Tag, _
                    IIf((txtBookNo.Tag = ""), Null, txtBookNo.Tag), _
                    val(Trim(txtRecFrom.Text)), _
                    val(Trim(txtRecTo.Text)))
        objdb.ExecuteSP "spSaveInterruptedReceiptBook", mArray, , , mCnn, adCmdStoredProc  'gbFinancialYearID
        cmdSave.Enabled = False
        Call FormInitialize
        mSql = "Select * From faInterruptedReceiptBooks"
        mSql = mSql + " Inner Join faCounters On faInterruptedReceiptBooks.intCounterID = faCounters.intCounterID"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
        End If
        Rec.Close
        MsgBox "Successfully Saved", vbInformation
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
'        Me.Width = 7065
'        Me.Height = 5025
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
                
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select vchDescription,intCounterID From faCounters Where intSectionID = 99 Order By vchDescription"
        PopulateList cmbCounter, mSql, , True, True, True, enuSourceString.Saankhya
        cmbCounter.ListIndex = 0
        
        '---------------------ADDED ON 08/05/2013 BY MINU-------------------------------
        cmbYear.AddItem gbFinancialYearID - 1
        cmbYear.ItemData(cmbYear.NewIndex) = 0
        cmbYear.AddItem gbFinancialYearID
        cmbYear.ItemData(cmbYear.NewIndex) = 1
        '------------------------------------------------------------------------------
        
        vsGrid.MergeRow(0) = True
        vsGrid.MergeCol(0) = True
        vsGrid.MergeCol(1) = True
        vsGrid.MergeCol(2) = True
        vsGrid.MergeCol(4) = True
        vsGrid.MergeCol(5) = True
        vsGrid.MergeCol(11) = True
        '---------------------------- Sinoj '
        If gbManualReceiptNewBool Then
            txtRecFrom.Locked = True
            txtRecTo.Locked = True
        End If
        '-----------------------------------'
        mSql = "Select * From faInterruptedReceiptBooks"
        mSql = mSql + " Inner Join faCounters On faInterruptedReceiptBooks.intCounterID = faCounters.intCounterID"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
        End If
        Rec.Close
        
        Call GetlastpostingYear
        
    End Sub

    Private Sub txtBookNo_Change()
        If gbManualReceiptNewBool = True Then
            If val(txtBookNo.Text) > 0 Then
                txtRecTo.Text = val(txtBookNo.Text) * 100
                txtRecFrom.Text = val(txtRecTo.Text) - 99
            Else
                txtRecTo.Text = ""
                txtRecFrom.Text = ""
            End If
        End If
    End Sub

    Private Sub txtBookNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtBookNo_LostFocus()
        Call findBookIDs(txtBookNo.Text)
    End Sub

    Private Sub txtNoOfReceipts_KeyPress(KeyAscii As Integer)
         If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtRecFrom_Change()
        If val(Trim(txtRecTo.Text)) < val(Trim(txtRecFrom.Text)) Then
'            txtRecTo.Text = Val(txtRecFrom.Text)
            txtNoOfReceipts.ForeColor = vbRed
        Else
            txtNoOfReceipts.ForeColor = vbBlack
        End If
        txtNoOfReceipts.Text = val(Trim(txtRecTo.Text)) - val(Trim(txtRecFrom.Text)) + 1  '' to Find No of the No of Receipts
    End Sub

    Private Sub txtRecFrom_KeyPress(KeyAscii As Integer)
        If gbManualReceiptNewBool = False Then
            If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtRecTo_Change()
    If val(Trim(txtRecTo.Text)) < val(Trim(txtRecFrom.Text)) Then
            'txtRecTo.Text = val(txtRecFrom.Text)
            txtNoOfReceipts.ForeColor = vbRed
        Else
            txtNoOfReceipts.ForeColor = vbBlack
        End If
        txtNoOfReceipts.Text = val(Trim(txtRecTo.Text)) - val(Trim(txtRecFrom.Text)) + 1  '' to Find No of the No of Receipts
    End Sub

    Private Sub txtRecTo_KeyPress(KeyAscii As Integer)
        If gbManualReceiptNewBool = False Then
            If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtRecTo_LostFocus()
        Call DuplicationCheck
        If val(Trim(txtRecTo.Text)) < val(Trim(txtRecFrom.Text)) Then
            txtNoOfReceipts.ForeColor = vbRed
            MsgBox "Please check the Receipt Number", vbInformation
            txtRecTo.Text = ""
'            txtRecTo.SetFocus
'            txtRecTo.Text = Val(txtRecFrom.Text)
        Else
            txtNoOfReceipts.ForeColor = vbBlack
        End If
        txtNoOfReceipts.Text = val(Trim(txtRecTo.Text)) - val(Trim(txtRecFrom.Text)) + 1  '' to Find No of the No of Receipts
        
    
    End Sub

    Private Sub DuplicationCheck()
        Dim mBookNo As String
        Dim mCnt    As Integer
        Dim mSql As String
        
        Dim mFrom As Double
        Dim mTo As Double
        
        mBookNo = txtBookNo.Text
        mFrom = val(txtRecFrom.Text)
        mTo = val(txtRecTo.Text)
        
        Dim mGridFrom As Long
        Dim mGridTo As Long
        Dim mTmp As String
        
        Dim mLoop As Integer
        Dim mValidFlag As Boolean
        
        Dim mYearID As Integer
        
        '
        'BLOCK [1]
        'NOTE: Find the year in which the Book is issued
        If cmbYear.ListIndex > 0 Then
            If cmbYear.ListIndex > -1 Then
                mYearID = cmbYear.ItemData(cmbYear.ListIndex)
            End If
        End If
        If mYearID < 1900 Then
            mYearID = gbFinancialYearID
        End If
        'END OF BLOCK [1]
        '
        
        For mLoop = 2 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 1) = mBookNo Then
                mTmp = vsGrid.TextMatrix(mLoop, 2)
                mGridFrom = val(Token(mTmp, "-"))
                mGridTo = val(mTmp)
                
                If (mGridFrom <= mFrom And mGridTo >= mFrom) Or _
                 (mGridFrom <= mTo And mGridTo >= mTo) Then
                 
                    If vsGrid.TextMatrix(mLoop, 11) = mYearID Then ' NOTE: Check Year is same or NOT
                        mSql = "The Same Book Number between Starting No " & txtRecFrom.Text & vbCrLf
                        mSql = mSql + " And Ending No " & txtRecTo.Text & vbCrLf
                        mSql = mSql + " is already issued!"
                        MsgBox mSql, vbCritical
                        txtRecFrom.Text = ""
                        txtRecFrom.SetFocus
                        cmdSave.Enabled = False
                        Exit Sub
                        
                    End If
                End If
               
            End If
        Next
    End Sub

    Private Sub vsGrid_Click()
        On Error GoTo err
        If vsGrid.Row < 2 Or vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
        Dim mRecFromTo As String
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        Call findBookIDs(vsGrid.TextMatrix(vsGrid.Row, 6))
        mSql = "Select  Case When isNull(max(intVoucherNo),0)<=isNull(Max(intReceiptNo),0) Then " & _
                "isNull(Max(intReceiptNo),0)Else isNull(max(intVoucherNo),0)End [ReceiptsNo] " & _
                "From faVouchers V Left Join faInterruptedCancelledReceipts FICR On V.intVoucherNo = FICR.intReceiptNo " & _
                "Where tnyVoucherGroupID = 4 And V.intBookNo = " & vsGrid.TextMatrix(vsGrid.Row, 7) ''Maximum of Receipt Number Generated
        Rec.Open mSql, mCnn
        cmbCounter.Enabled = True
            If Rec!ReceiptsNo > 0 Or vsGrid.TextMatrix(vsGrid.Row, 5) = "Closed" Then
''                cmbCounter.Enabled = False
                txtBookNo.Enabled = False
                txtRecFrom.Enabled = False
                txtRecTo.Enabled = False
                cmbYear.Enabled = False
            Else
                cmbCounter.Enabled = True
                txtBookNo.Enabled = True
                txtRecFrom.Enabled = True
                txtRecTo.Enabled = True
            End If
        Rec.Close
        chkClosed.Enabled = True
        If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
            If vsGrid.TextMatrix(vsGrid.Row, 8) = 0 Then
                chkClosed.Enabled = True
'                lblReason.Enabled = True
'                txtReason.Enabled = True
            Else
'                chkClosed.Enabled = False
                lblReason.Enabled = False
                txtReason.Enabled = True
            End If
        Else
            chkClosed.Enabled = False
            lblReason.Enabled = False
            txtReason.Enabled = False
        End If

        If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
            cmbCounter.Text = vsGrid.TextMatrix(vsGrid.Row, 0)
            txtBookNo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            txtBookNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 7)
            mRecFromTo = vsGrid.TextMatrix(vsGrid.Row, 2)
            txtRecFrom.Text = Trim(Token(mRecFromTo, "-"))
            txtRecTo.Text = Trim(mRecFromTo)
            txtNoOfReceipts.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            txtReason.Text = vsGrid.TextMatrix(vsGrid.Row, 9)
            If vsGrid.TextMatrix(vsGrid.Row, 8) = 1 Then
                chkClosed.Value = 1
            ElseIf vsGrid.TextMatrix(vsGrid.Row, 8) = 0 Then
                chkClosed.Value = 0
            End If
            cmbYear.Text = vsGrid.TextMatrix(vsGrid.Row, 11)
            
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub findBookIDs(intBookNo As String)
        mBookIDs = "(-2"
        If Trim(txtBookNo.Text) = "" Then
            mBookIDs = mBookIDs + ")"
            Exit Sub
        End If
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select Distinct intBookID From faInterruptedReceiptBooks FIB Inner Join faVouchers V On V.intBookNo = FIB.intBookID " & _
                "Where V.tnyVoucherGroupID = 4 And FIB.intBookNo = " & val(Trim(intBookNo))
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                mBookIDs = mBookIDs + "," + CStr(Rec!intBookID)
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        mBookIDs = mBookIDs + ")"
    End Sub
    Private Sub GetlastpostingYear()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "SELECT MAX(isnull(intFinYearID,0)) intFinYearID FROM faPostingIndex WHERE tnyStage=2 AND intMonthID=3"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            mLastPostingYearID = Rec!intFinYearID
        End If
        Rec.Close
        mCnn.Close
    End Sub

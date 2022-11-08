VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterruptedCancellationRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interrupted Cancellation  Request"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   345
      Left            =   1545
      TabIndex        =   7
      Top             =   4050
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   2813
      TabIndex        =   6
      Top             =   4050
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4080
      TabIndex        =   8
      Top             =   4050
      Width           =   1320
   End
   Begin VB.ComboBox cmbBookNo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   345
      Left            =   2865
      TabIndex        =   3
      Top             =   765
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      _Version        =   393216
      Format          =   60751873
      CurrentDate     =   40016
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2580
      Left            =   195
      TabIndex        =   5
      Top             =   1335
      Width           =   6555
      _cx             =   11562
      _cy             =   4551
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
      SelectionMode   =   0
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
      FormatString    =   $"frmInterruptedCancellationRequest.frx":0000
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
   Begin VB.TextBox txtReason 
      Height          =   315
      Left            =   4290
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   2355
   End
   Begin VB.TextBox txtReceiptDate 
      Height          =   315
      Left            =   1455
      TabIndex        =   2
      Top             =   750
      Width           =   1365
   End
   Begin VB.TextBox txtReceiptNo 
      Height          =   315
      Left            =   4290
      MaxLength       =   5
      TabIndex        =   1
      Top             =   390
      Width           =   1365
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
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
      Left            =   3615
      TabIndex        =   12
      Top             =   735
      Width           =   630
   End
   Begin VB.Label lblReceiptDate 
      AutoSize        =   -1  'True
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
      Left            =   285
      TabIndex        =   11
      Top             =   750
      Width           =   1125
   End
   Begin VB.Label lblReceiptNo 
      AutoSize        =   -1  'True
      Caption         =   "Receipt No"
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
      Left            =   3300
      TabIndex        =   10
      Top             =   420
      Width           =   945
   End
   Begin VB.Label lblBookNo 
      AutoSize        =   -1  'True
      Caption         =   "Book No"
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
      Left            =   705
      TabIndex        =   9
      Top             =   420
      Width           =   705
   End
End
Attribute VB_Name = "frmInterruptedCancellationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mStatus     As Variant
        
    Private Sub FormInitialize()
        cmbBookNo.ListIndex = 0
        txtReceiptDate.Text = ""
        txtReceiptNo.Text = ""
        txtReason.Text = ""
        'vsGrid.TextMatrix(vsGrid.Row, 4) = ""
        txtReceiptNo.Tag = ""
    End Sub
    Private Sub FillvsGrid(Rec As ADODB.Recordset)
        Dim mRowCount   As Double
        
        mStatus = ""
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intSerialNo), "", Rec!intSerialNo)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!dtReceiptDate), "", Rec!dtReceiptDate)
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            If mStatus <> "" Then
                If mStatus = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCount, 3) = True
                End If
                If mStatus = 0 Then
                    vsGrid.Cell(flexcpChecked, mRowCount, 3) = False
                End If
            End If
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!intBookID), "", Rec!intBookID)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!Remarks), "", Rec!Remarks)
            Rec.MoveNext
'            If (IsNull(Rec!intBookNo)) Then
'                MsgBox "ok"
'            End If
'            vsGrid.Rows = vsGrid.Rows + 1
            mRowCount = mRowCount + 1
        Wend
    End Sub
    
    Private Sub cmbBookNo_LostFocus()
'        If cmbBookNo.ListIndex > 0 Then
'            If GetDateWithRecNo = False Then
'                MsgBox "Invalid Book No or Receipt No", vbInformation
'            End If
'        End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdSave.Enabled = True
        Call GetNextReceiptNo
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim mArray      As Variant
        Dim mReceiptNo  As Double
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCount      As Variant
        Dim mRowCount   As Double
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '*********************************************************************************************'
        '               Procedure to request for the Interrupt Receipt Cancellation                   '
        '*********************************************************************************************'
        If cmbBookNo.ListIndex < 1 Then
            MsgBox "Please select the Book", vbInformation
            Exit Sub
        End If
        If txtReceiptNo.Text = "" Then
            MsgBox "Plese enter the Receipt No", vbInformation
            Exit Sub
        Else
'            mSQL = "Select intCount From faInterruptedReceiptBooks Where intBookNo = " & cmbBookNo.Text
'            Rec.Open mSQL, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mCount = IIf(IsNull(Rec!intCount), "", Rec!intCount)
'            End If
'            Rec.Close
'            If mCount <> "" Then
'                If CDbl(txtReceiptNo.Text) > mCount Then
'                    MsgBox "Please enter a valid Receipt No", vbInformation
'                    Exit Sub
'                End If
'            End If
        End If
        If txtReceiptDate.Text = "" Then
            MsgBox "Please enter the Receipt Date", vbInformation
            Exit Sub
        End If
        If Trim(txtReason.Text) = "" Then
            MsgBox "Please specify the reason", vbInformation
            Exit Sub
        End If
'        If GetDateWithRecNo = False Then            '' Function Call
'                MsgBox "Invalid Book No or Receipt No", vbInformation
'                Exit Sub
'            End If
'        If cmbBookNo.ListIndex <> -1 And txtReceiptNo.Text <> "" Then
'            If Len(txtReceiptNo.Text) = 1 Then
'                mReceiptNo = "9" + Right(0 + cmbBookNo.Text, 4) + txtReceiptNo.Text
'            ElseIf Len(txtReceiptNo.Text) = 2 Then
'                mReceiptNo = 109 + cmbBookNo.Text + "0" + txtReceiptNo.Text
'            Else
'                mReceiptNo = 109 + cmbBookNo.Text + txtReceiptNo.Text
'            End If
'        End If
        mSql = "Select * From faInterruptedReceiptBooks Where tnyClosed <> 1 And intBookNo = " & cmbBookNo.Text & " And " & Trim(txtReceiptNo.Text) & " Between numReceiptNoFrom And numReceiptNoTo"
        Rec.Open mSql, mCnn
        If Rec.EOF And Rec.BOF Then
            MsgBox "Invalid Receipt Number & Book Number", vbInformation
            Exit Sub
        End If
        Rec.Close
        mReceiptNo = "9" + Right("0000" + cmbBookNo.Text, 4) + "1" + Right("00000" + txtReceiptNo.Text, 5)
        If txtReceiptNo.Tag = "" Then
        'If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 4)) Then
            mSql = "Select * From faInterruptedCancelledReceipts"
            mSql = mSql + " Where intBookID = " & cmbBookNo.ItemData(cmbBookNo.ListIndex)
            mSql = mSql + " And intSerialNo =" & txtReceiptNo.Text
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                MsgBox "This Receipt is already Cancelled", vbInformation
                Exit Sub
            End If
            Rec.Close
        Else
            For mRowCount = 1 To vsGrid.Rows - 1
                If vsGrid.Row <> mRowCount Then
                    If vsGrid.TextMatrix(mRowCount, 4) = txtReceiptNo.Tag Then
                        If vsGrid.TextMatrix(mRowCount, 1) = txtReceiptNo.Text Then
                            MsgBox "This Receipt is already Cancelled", vbInformation
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
        
        mArray = Array(cmbBookNo.ItemData(cmbBookNo.ListIndex), _
                     cmbBookNo.Text, _
                     txtReceiptNo.Text, _
                     gbUserID, _
                     txtReceiptDate, _
                     mReceiptNo, _
                     txtReason.Text, _
                     txtReceiptNo.Tag, _
                     IIf(IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 1)), vsGrid.TextMatrix(vsGrid.Row, 1), "") _
                     )
        objDb.ExecuteSP "spSaveInterruptedReceiptCancellation", mArray, , , mCnn, adCmdStoredProc
        MsgBox "Successfully Saved", vbInformation
        cmdSave.Enabled = False
        Call FormInitialize
        mSql = "Select *,faInterruptedCancelledReceipts.vchRemarks[Remarks] From faInterruptedCancelledReceipts"
        mSql = mSql + " Inner Join faInterruptedReceiptBooks On faInterruptedCancelledReceipts.intBookID = faInterruptedReceiptBooks.intBookID"
        mSql = mSql + " Where numUserID =" & gbUserID
        mSql = mSql + " And tnyClosed = 0"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
        End If
        Rec.Close
        Call GetNextReceiptNo
    End Sub
    
    Private Sub dtpDate_CloseUp()
        txtReceiptDate.Text = dtpDate.value
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.Width = 7065
        Me.Height = 5025
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim arOut   As Variant
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        dtpDate.value = Date
        txtReceiptDate.Text = Date
        
        mSql = "Select intBookNo,intBookID From faInterruptedReceiptBooks"
        mSql = mSql + " Where intCounterID = " & gbCounterID
        mSql = mSql + " And tnyClosed <> 1"
        PopulateList cmbBookNo, mSql, , True, True, True, enuSourceString.Saankhya
        cmbBookNo.ListIndex = 0
        
        mSql = "Select *,faInterruptedCancelledReceipts.vchRemarks[Remarks] From faInterruptedCancelledReceipts"
        mSql = mSql + " Inner Join faInterruptedReceiptBooks On faInterruptedCancelledReceipts.intBookID = faInterruptedReceiptBooks.intBookID"
        mSql = mSql + " Where numUserID =" & gbUserID
        mSql = mSql + " And tnyClosed = 0"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
        End If
        Rec.Close
        Call GetNextReceiptNo
    End Sub
    
    Private Sub txtReceiptDate_LostFocus()
        If txtReceiptDate.Text <> "" Then
            txtReceiptDate.Text = CheckDateInMMM(txtReceiptDate.Text)
        End If
    End Sub

    Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtReceiptNo_LostFocus()
        Dim mReceiptNo  As String
        
        mReceiptNo = ""
        If txtReceiptNo.Text <> "" Then
            mReceiptNo = "0000" + Trim(txtReceiptNo.Text)
            txtReceiptNo.Text = Right(mReceiptNo, 5)
            If GetDateWithRecNo = False Then
                MsgBox "Invalid Receipt No", vbInformation
            End If
        End If
        
    End Sub

    Private Sub vsGrid_Click()
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 5)) Then
            cmbBookNo.Text = vsGrid.TextMatrix(vsGrid.Row, 0)
            'cmbBookNo.itemData(cmbBookNo.ListIndex) = vsGrid.TextMatrix(vsGrid.Row, 4)
            txtReceiptNo.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            txtReceiptNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 4)
            txtReceiptDate.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            txtReason.Text = vsGrid.TextMatrix(vsGrid.Row, 6)
            If vsGrid.TextMatrix(vsGrid.Row, 5) = 0 Then
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
        End If
    End Sub
    Private Function GetDateWithRecNo() As Boolean
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select dtDate From faVouchers Where tnyVoucherGroupID = 4 And Right(intVoucherNo,5) =  " & Trim(txtReceiptNo.Text) & " And intBookNo = " & cmbBookNo.ItemData(cmbBookNo.ListIndex)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtReceiptDate.Text = Format(Rec!dtDate, "dd-MMM-YYYY")
            GetDateWithRecNo = True
        Else
            GetDateWithRecNo = False
        End If
        Rec.Close
    End Function
    Private Sub GetNextReceiptNo()
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim mAryIn  As Variant
        Dim arOut   As Variant
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mAryIn = Array(gbCounterID)
        objDb.ExecuteSP "spGetNextReceiptNoInterrupted", mAryIn, arOut
        If IsNull(arOut(0, 0)) = False Then
            txtReceiptNo.Text = Right(CStr(arOut(0, 0)), 5)
            mSql = "SELECT intBookNo, intBookID, intCounterID, tnyClosed FROM faInterruptedReceiptBooks WHERE ISNULL(tnyClosed,0) = 0 And intCounterID = " & gbCounterID
            Rec.Open mSql, mCnn
            If Not (Rec.BOF And Rec.EOF) Then
                'cmbBookNo.Text = val(mID(arOut(0, 0), 2, 4))
                cmbBookNo.Text = Rec!intBookNo
            End If
            Rec.Close
            
            mSql = "Select dtRequestDate From faInterruptedRequests Where tnyStatus = 2 And intCounterID = " & gbCounterID
            mSql = mSql + " And intTypeID = 1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtReceiptDate.Text = Format(Rec!dtRequestDate, "dd-MMM-yyyy")
            End If
            Rec.Close
            
        End If
    End Sub
    

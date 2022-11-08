VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmMOReturned 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Money Order Return "
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMOReturned.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CanceL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3697
      TabIndex        =   6
      Top             =   3645
      Width           =   1395
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to Receipt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1792
      TabIndex        =   5
      Top             =   3645
      Width           =   1875
   End
   Begin VB.ComboBox cmbPensionType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2085
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   495
      Width           =   3570
   End
   Begin VB.CommandButton cmdSearchBill 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   450
   End
   Begin VSFlex8LCtl.VSFlexGrid vsBill 
      Height          =   2220
      Left            =   75
      TabIndex        =   4
      Top             =   1350
      Width           =   6750
      _cx             =   11906
      _cy             =   3916
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMOReturned.frx":1CCA
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
   Begin VB.TextBox txtPrefix 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2085
      TabIndex        =   1
      Top             =   840
      Width           =   1320
   End
   Begin VB.TextBox txtPensionerID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   2
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label lblPensionType 
      AutoSize        =   -1  'True
      Caption         =   "Pensioner Type"
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
      Left            =   690
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblPensionerID 
      AutoSize        =   -1  'True
      Caption         =   "Pensioner ID"
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
      Left            =   915
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "    Money Order Return"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmMOReturned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
'    Dim mNumberOfSelections As Variant

    '*********************************************************************************************'
    '               Form to Integrate Saankhya with Sevana Pension                                '
    '*********************************************************************************************'
    Private Sub cmbPensionType_Click()
        On Error GoTo err
        If (cmbPensionType.ListIndex > 0) Then
            If cmbPensionType.ItemData(cmbPensionType.ListIndex) > 0 Then
                txtPrefix.Text = CStr("1") + CStr(gbLocalBodyID) + CStr(Right(CStr("0" + CStr(cmbPensionType.ItemData(cmbPensionType.ListIndex))), 2))
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdCancel_Click()
        On Error GoTo err
        Dim objTranType As New clsTransactionType
        Unload Me
        If Not frmReceiptsCounter.InterruptEditMode Then
            objTranType.SetTransactionType (9999)
            frmReceiptsCounter.txtTransactionType.Text = objTranType.TransactionType
            frmReceiptsCounter.txtTransactionType.Tag = objTranType.TransactionTypeID
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdCopy_Click()
        Dim mRowCount   As Integer
        Dim objAcc      As New clsAccounts
        
        On Error GoTo err
        For mRowCount = 1 To vsBill.Rows - 1
            With frmReceiptsCounter.vsGrid
                .Rows = .Rows + 1         ' One Row Added
                If vsBill.TextMatrix(mRowCount, 7) = 1 Then
                    objAcc.SetAccountCode (250600200)             '' Setting Accounts with Head Code
                End If
                If vsBill.TextMatrix(mRowCount, 7) = 2 Then
                    objAcc.SetAccountCode (250601100)             '' Setting Accounts with Head Code
                End If
                If vsBill.TextMatrix(mRowCount, 7) = 3 Or vsBill.TextMatrix(mRowCount, 7) = 4 Then
                    objAcc.SetAccountCode (250600700)             '' Setting Accounts with Head Code
                End If
                If vsBill.TextMatrix(mRowCount, 7) = 5 Then
                    objAcc.SetAccountCode (250600600)             '' Setting Accounts with Head Code
                End If
                If vsBill.TextMatrix(mRowCount, 7) = 6 Then
                    objAcc.SetAccountCode (250600500)             '' Setting Accounts with Head Code
                End If
                .Cell(flexcpText, mRowCount, 0) = objAcc.AccountCode            'AccountHead
                .Cell(flexcpText, mRowCount, 1) = objAcc.AccountHead      'AccountHead
                .Cell(flexcpText, mRowCount, 2) = gbFinancialYearID        'YearID
                .Cell(flexcpText, mRowCount, 3) = gbCurrentPeriodID       'Period ID
                .Cell(flexcpText, mRowCount, 5) = vsBill.TextMatrix(mRowCount, 3)  'Current Amount
                .Cell(flexcpText, mRowCount, 6) = objAcc.AccountHeadID
                .Cell(flexcpText, mRowCount, 7) = gbFinancialYearID
                .Cell(flexcpText, mRowCount, 8) = 1
                .Cell(flexcpText, mRowCount, 10) = ""                      'Demand ID
                .Cell(flexcpText, mRowCount, 11) = vsBill.TextMatrix(mRowCount, 3)  'Current Amount                       'Amount Paid
                .Cell(flexcpChecked, mRowCount, 12) = 1                   'Checked
                .Cell(flexcpText, mRowCount, 17) = vsBill.TextMatrix(mRowCount, 5)
                .Cell(flexcpText, mRowCount, 18) = vsBill.TextMatrix(mRowCount, 6)
                .Cell(flexcpText, mRowCount, 19) = vsBill.TextMatrix(mRowCount, 7)
            End With
        Next
        Unload Me
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchBill_Click()
        On Error GoTo err
        If cmbPensionType.ListIndex < 0 Then
            MsgBox "Please select the Pension Type", vbInformation
            cmbPensionType.SetFocus
            Exit Sub
        End If
        If Trim(txtPrefix.Text) = "" Or Trim(txtPensionerID.Text) = "" Then
            MsgBox "Please enter the Pensioner ID", vbInformation
            txtPensionerID.SetFocus
            Exit Sub
        End If
        
        frmSearchMODetails.PensionerID = CStr(txtPrefix.Text) + CStr(txtPensionerID.Text)
        frmSearchMODetails.txtPrefix.Text = txtPrefix.Text
        frmSearchMODetails.txtPensionerID.Text = txtPensionerID.Text
        frmSearchMODetails.cmbPensionType.Text = cmbPensionType.Text
        frmSearchMODetails.Show vbModal
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Call cmdCancel_Click
        End If
    End Sub

    Private Sub Form_Load()
        Dim mSQL As String
                
        On Error GoTo err
        vsBill.Rows = 1
        mSQL = "Select chvPensionNameEnglish,tnyPensionTypeID From GM_PensionType"
        PopulateList cmbPensionType, mSQL, , True, True, True, SevanaPension
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub txtPensionerID_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
   
    Private Sub vsBill_KeyDown(KeyCode As Integer, Shift As Integer)
        If vsBill.Row > 0 Then
            If KeyCode = vbKeyDelete Then
                If MsgBox("Do you want to remove this row", vbYesNo) = vbYes Then
                    vsBill.RemoveItem (vsBill.Row)
                End If
            End If
        End If
    End Sub
    
    Private Sub vsBill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        vsBill.ToolTipText = "Press Delete to remove a row from Grid"
    End Sub

    Private Sub vsBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'        Dim mLoop As Long
'
'        If Row > 0 Then
'            If vsBill.Cell(flexcpChecked, Row, Col) = 2 Then
'                If Row = 1 Or vsBill.Cell(flexcpChecked, Row - 1, Col) = vbChecked Then
'                    vsBill.Cell(flexcpChecked, Row, Col) = vbChecked
'                    mNumberOfSelections = mNumberOfSelections + 1 'IIf(Row Mod 2 = 0, 1, 0)
'                Else
'                    Cancel = True
'                End If
'            Else ' Already  Checked
'                If vsBill.Cell(flexcpChecked, Row - 1, Col) = 1 Then
'                    For mLoop = 1 To vsBill.Rows - 1
'                        If vsBill.TextMatrix(Row, 10) <> vsBill.TextMatrix(mLoop, 10) Then
'                            vsBill.Cell(flexcpChecked, mLoop, 12) = 2
'                            'If vsbill.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
'                            'Cancel = True
'                            'End If
'                            mNumberOfSelections = mNumberOfSelections - 1
'                            'Exit For
'                        End If
'                    Next mLoop
'                Else
'                    Cancel = True
'                End If
'            End If
'        End If
    End Sub

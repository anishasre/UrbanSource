VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfDailyCollection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Cash Collection Detials"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   -15
      ScaleHeight     =   345
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   0
      Width           =   5115
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   4950
      Top             =   2430
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtTotal 
      Height          =   330
      Left            =   2475
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2175
      Width           =   1830
   End
   Begin VB.CommandButton cmdRemittance 
      Caption         =   "Remittance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Top             =   2175
      Width           =   1500
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1710
      Left            =   60
      TabIndex        =   2
      Top             =   375
      Width           =   4935
      _cx             =   8705
      _cy             =   3016
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorBkg    =   -2147483644
      BackColorAlternate=   -2147483626
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmListOfDailyCollection.frx":0000
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
      Begin VB.CheckBox chkSelect 
         Height          =   270
         Left            =   4425
         TabIndex        =   3
         Top             =   90
         Width           =   180
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1905
      TabIndex        =   4
      Top             =   2250
      Width           =   450
   End
End
Attribute VB_Name = "frmListOfDailyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim mNumberOfSelections As Integer
    
    Private Sub chkSelect_Click()
        If chkSelect.value = vbChecked Then
            If vsGrid.Rows > 1 Then
                vsGrid.Cell(flexcpChecked, 1, 2, vsGrid.Rows - 1, 2) = True
                Call Calculate
            End If
        ElseIf chkSelect.value = vbUnchecked Then
            If vsGrid.Rows > 1 Then
                vsGrid.Cell(flexcpChecked, 1, 2, vsGrid.Rows - 1, 2) = False
                txtTotal.Text = ""
            End If
        End If
        
        Call Calculate
        If chkSelect.value = vbUnchecked Then
            txtTotal.Text = ""
        End If
    End Sub

    Private Sub cmdRemittance_Click()
        Dim objAcc As New clsAccounts
        Dim i As Integer
        Dim mLastRow As Integer
        mLastRow = 0
        If val(txtTotal.Text) > 0 Then
            With frmContraEntry
                For i = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpChecked, i, 2) = 1 Then
                        mLastRow = i
                    End If
                Next
                If mLastRow > 0 Then
                .LastRemittanceDate = Format(vsGrid.TextMatrix(mLastRow, 0), "dd/MMM/yyyy")
                End If
                .copiedAmount = val(txtTotal.Text)
                .cmdSearchVoucherNo.Tag = -1
                .txtVoucherNo.Tag = -1
                .txtVoucherNo.Text = ""
                .txtReference.Text = ""
                .txtReference.Tag = "" '
                If .cmbInstruments.ListCount > 0 Then
                    .cmbInstruments.Text = "Cash"
                End If
                ' Credit Account Filling
                objAcc.SetAccountCode (gbAcHeadCodeCash)
                .txtAccountHead.Tag = objAcc.AccountHeadID
                .txtAccountCode.Text = objAcc.AccountCode
                .txtAccountHead.Text = objAcc.AccountHead
                .txtRef.Text = ""
                .txtIssuedDate.Text = ""
                .txtInstDate.Text = ""
                
'                .txtNarration.Text = "Remittance of JSK Collection Upto " & gbTransactionDate - 1
                If mLastRow > 0 Then
                .txtNarration.Text = "Remittance of JSK Collection Upto " & Format(vsGrid.TextMatrix(mLastRow, 0), "dd/MMM/yyyy")
                End If
                'Grid Filling
                .vsGrid.Rows = 1
                .vsGrid.Rows = 10
                Call objAcc.SetAccountID(gbDefaultBankID)
                .vsGrid.TextMatrix(1, 1) = objAcc.AccountCode
                .vsGrid.TextMatrix(1, 2) = objAcc.AccountHead
                .vsGrid.TextMatrix(1, 4) = Format(txtTotal.Text, "0.00")
                .txtDr.Text = Format(txtTotal.Text, "0.00")
'                .cmbInstruments.Locked = True
'                .vsGrid.Editable = flexEDNone
                .mRemittanceModule = 50
            End With
        Else
            frmContraEntry.mRemittanceModule = 0
            MsgBox "Amount Not Present", vbInformation
        End If
       cmdRemittance.Enabled = False
       Unload Me
       frmContraEntry.Visible = True
    End Sub

    Private Sub Calculate()
        Dim mTotal       As Double
        Dim mCount       As Integer
        txtTotal.Text = ""
        For mCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mCount, 2) = 1 Then
                If val(vsGrid.TextMatrix(mCount, 1)) <> 0 Then
                    mTotal = mTotal + Format(val(vsGrid.TextMatrix(mCount, 1)), "0.00")
                    txtTotal.Text = Format(mTotal, "0.00")
                End If
            End If
        Next
    End Sub
    Private Sub FillGrid()
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mRowCnt As Integer
        Dim mRecCnt As Integer
        Dim ID As Variant
        Dim ArrayIn As Variant
        Dim mSQL As Variant
        Dim mLastRemittance As Variant
        Dim mLastRemitanceExists    As Variant ' to check wheather this fielld is Empty or not
        Dim mCount As Variant
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
             mSQL = "SELECT dtLastRemittance from  faConfig"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
'                If IIf(IsNull(Rec!dtLastRemittance), "", Rec!dtLastRemittance) = gbTransactionDate Then
'                    Exit Sub
'                End If
                mLastRemittance = DdMmmYy(IIf(IsNull(Rec!dtLastRemittance), gbTransactionDate - 1, Rec!dtLastRemittance))
                mLastRemitanceExists = IIf(IsNull(Rec!dtLastRemittance), "", Rec!dtLastRemittance)
            End If
            Rec.Close
            mCount = 10
''            ArrayIn = Array(CDate(mLastRemittance) + 1, _
''                    gbTransactionDate - 1)
LRemit:
            If mLastRemitanceExists = "" Then
                ArrayIn = Array(CDate(mLastRemittance), _
                    gbTransactionDate)
            Else
                ArrayIn = Array(CDate(mLastRemittance) + 1, _
                        gbTransactionDate)
            End If
            Set Rec = objDB.ExecuteSP("spRptDailyRemitanceJSK", ArrayIn, , , mCnn, adCmdStoredProc)
            If mLastRemitanceExists = "" Then
                If (Rec.EOF And Rec.BOF) Then
                    mLastRemittance = DateAdd("d", -1, mLastRemittance)
                    mCount = mCount - 1
                    If mCount > 1 Then
                        GoTo LRemit
                    End If
                End If
            End If
            mRowCnt = 1
            mRecCnt = 1
            vsGrid.Rows = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(DdMmmYy(Rec!dtDate)), "", DdMmmYy(Rec!dtDate))
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!Amount), "", Format(Rec!Amount, "0.00"))
                    vsGrid.Cell(flexcpChecked, mRowCnt, 2) = vbChecked
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    mRecCnt = mRecCnt + 1
                Wend
                chkSelect.value = 1
                Call Calculate
            End If
            Rec.Close
        End If
    End Sub
'    Private Sub Form_Activate()
'        Me.Top = 0
'        Me.Left = 0
'    End Sub

    Private Sub Form_Load()
        winXPC.InitIDESubClassing
        Call FillGrid
    End Sub
    
    Private Sub vsGrid_Click()
        If vsGrid.Col = 2 Then
            vsGrid.Editable = flexEDKbdMouse
        Else
            vsGrid.Editable = flexEDNone
        End If
    End Sub

    Private Sub vsGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Dim mLoop As Long
        If Row > 0 Then
            If vsGrid.Cell(flexcpChecked, Row, Col) = 2 Then
                If Row = 1 Or vsGrid.Cell(flexcpChecked, Row - 1, Col) = vbChecked Then
                    vsGrid.Cell(flexcpChecked, Row, Col) = vbChecked
                    'mNumberOfSelections = mNumberOfSelections + 1 'IIf(Row Mod 2 = 0, 1, 0)
                Else
                    Cancel = True
                End If
            Else ' Already  Checked
                If vsGrid.Cell(flexcpChecked, Row, Col) = 1 Then
                    For mLoop = Row To vsGrid.Rows - 1
'                        If vsGrid.TextMatrix(Row, 10) <> vsGrid.TextMatrix(mLoop, 10) Then
                            vsGrid.Cell(flexcpChecked, mLoop, 2) = 2
                            'If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
                            'Cancel = True
                            'End If
                     '       mNumberOfSelections = mNumberOfSelections - 1
                            'Exit For
'                        End If
                    Next mLoop
                Else
                    Cancel = True
                End If
            End If
        End If
        Call Calculate
    End Sub

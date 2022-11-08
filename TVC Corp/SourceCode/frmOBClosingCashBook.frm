VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmOBClosingCashBook 
   BorderStyle     =   0  'None
   Caption         =   "ClosingCashBook"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14325
   Icon            =   "frmOBClosingCashBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNect 
      Caption         =   "Next"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4950
      TabIndex        =   8
      Top             =   4950
      Width           =   825
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4950
      Width           =   825
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3660
      Left            =   45
      TabIndex        =   6
      Top             =   1080
      Width           =   9960
      _cx             =   17568
      _cy             =   6456
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOBClosingCashBook.frx":1CCA
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
   Begin VB.TextBox txtDate 
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
      Left            =   2025
      TabIndex        =   4
      Top             =   675
      Width           =   1635
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8145
      TabIndex        =   2
      Top             =   4815
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5805
      TabIndex        =   1
      Top             =   4950
      Width           =   825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4095
      TabIndex        =   0
      Top             =   4950
      Width           =   825
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10080
      TabIndex        =   15
      Top             =   2205
      Width           =   195
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10080
      TabIndex        =   14
      Top             =   3465
      Width           =   195
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10080
      TabIndex        =   13
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the closing balance of Cash/Bank/Treasury Accounts as per the Single Entry Cash Book."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   915
      Left            =   10440
      TabIndex        =   12
      Top             =   3465
      Width           =   3660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter '0' balance in the Opening Balance screen, in the case of a Bank/Treasury Account opened in the Middle of  the year."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1050
      Left            =   10440
      TabIndex        =   11
      Top             =   2205
      Width           =   3570
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The Bank/Treasury Accounts in the Opening Cash Book screen only are listed here."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   10440
      TabIndex        =   10
      Top             =   1260
      Width           =   3570
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3705
      Left            =   10035
      TabIndex        =   9
      Top             =   1035
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Balance As On"
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
      Left            =   135
      TabIndex        =   5
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   7380
      TabIndex        =   3
      Top             =   4860
      Width           =   735
   End
End
Attribute VB_Name = "frmOBClosingCashBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Sub cmdClose_Click()
        If MsgBox("You haven't finished the Wizard, are you sure you want to quit?   ", vbQuestion + vbYesNo, "Close Wizard") = vbYes Then
            Unload Me
            frmOpeningWizard.cmdCancel_Click
        End If
    End Sub

    Private Sub cmdNect_Click()
        'Me.Hide
        'frmOpeningWizard.FrameNo = 5
        Unload Me
        frmOpeningWizard.cmdNext_Click
    End Sub

    Private Sub cmdPre_Click()
        Me.Hide
'        Unload Me
        frmOBPaymentTransactions.Form_Load
        frmOpeningWizard.cmdPre_Click
    End Sub

    Private Sub cmdSave_Click()
        Dim mAccID          As Integer
        Dim mAccCode        As String
        Dim mAmtOp          As Double
        Dim mAmtCL          As Double
        Dim mCnt            As Integer
        Dim mOBCashBookID   As Integer
        Dim mSql            As String
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim objAcc          As New clsAccounts
        Dim arrIn           As Variant
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mSql = "Delete FROM faOBCashBook Where isNull(fltClosing,0)<>0 And isNull(fltOpening,0)<>0"
'        mCnn.Execute (mSql)
        For mCnt = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mCnt, 3) <> "" Then
                mAccID = vsGrid.TextMatrix(mCnt, 3)
                objAcc.SetAccountID (mAccID)
                mSql = "select * from faOBCashBook where intAccountHeadID = " & mAccID & " "
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF And Rec.BOF) Then
                'If rec.RecordCount > 0 Then
                    mOBCashBookID = IIf(IsNull(Rec!intOBCashbookID), 0, Rec!intOBCashbookID)
                    mAmtOp = IIf(IsNull(Rec!fltOpening), 0, Rec!fltOpening)
                Else
                    mOBCashBookID = -1
                    mAmtOp = 0
                End If
                Rec.Close
                mAccCode = objAcc.AccountCode
                mAmtCL = val(vsGrid.TextMatrix(mCnt, 2))
                arrIn = Array(mOBCashBookID, mAccID, mAccCode, mAmtOp, mAmtCL)
                objdb.ExecuteSP "spSaveOBCashBook", arrIn, , , mCnn, adCmdStoredProc
            End If
        Next
        MsgBox "Successfully Saved..", vbApplicationModal
        cmdSave.Enabled = False
        Call FillOBcashBook
        Me.Hide
        'frmOpeningWizard.FrameNo = 5
        Unload Me
        frmOpeningWizard.cmdNext_Click
    End Sub

    Private Sub Form_Load()
        'vsGrid.ColComboList(0) = "|..."
        Call FillOBcashBook
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
    End Sub
    Private Sub FillOBcashBook()
        Dim mSql             As String
        Dim objdb            As New clsDB
        Dim mCnn             As New ADODB.Connection
        Dim Rec              As New ADODB.Recordset
        Dim mCnt             As Integer
        Dim mRowCnt          As Integer
        mSql = "Select convert(varchar, DATEADD(day,DATEDIFF(day,0,dtRPOpeningDate),-1),105) dtRPOpeningDate From faConfig"
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            txtDate.Text = IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate)
            txtDate.Enabled = False
        End If
        Rec.Close
        mSql = ""
        mSql = "Select faOBCashBook.vchAccountHeadCode,vchAccountHead,fltClosing,faOBCashBook.intAccountHeadID,intOBCashbookID "
        mSql = mSql + " From faOBCashBook Inner Join faAccountHeads On faOBCashBook.intAccountHeadID=faAccountHeads.intAccountHeadID"
        'Where isNull(fltClosing,0)<>0"
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        vsGrid.Rows = 1
        If Rec.RecordCount > 0 Then
            If Not (Rec.EOF And Rec.BOF) Then
                vsGrid.Rows = Rec.RecordCount + 1
                vsGrid.Col = 0
                vsGrid.Row = 1
                vsGrid.ColSel = 4
                vsGrid.RowSel = vsGrid.Rows - 1
                mSql = Rec.GetString(, , vbTab, Chr(13))
                vsGrid.Clip = mSql
            End If
            Call Calculate
'        Else
'            While Not (Rec.EOF Or Rec.BOF)
'                vsGrid.Rows = vsGrid.Rows + 1
'                vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
'                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'                vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intOBCashbookID), "", Rec!intOBCashbookID)
'            Wend
        End If
        Rec.Close
    End Sub
    Private Sub txtDate_LostFocus()
        If Not IsDate(txtDate.Text) Then
            txtDate.Text = DdMmmYy(gbStartingDate - 1)
        Else
            txtDate.Text = CheckDateInMMM(txtDate.Text)
        End If
    End Sub
    
    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Call Calculate
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Col = 1 Then
            If vsGrid.TextMatrix(Row, 1) = "" Then
                MsgBox "Please Select Account Head..."
                Exit Sub
            End If
        End If
    End Sub
    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim objdb            As New clsDB
        Dim mCnn             As New ADODB.Connection
        
        If KeyCode = vbKeyDelete Then
            If MsgBox(" Do you want to Delete the Record?", vbYesNo, "Saankhya") = vbYes Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                mCnn.Execute "Update faOBCashBook set fltClosing=0 Where intAccountHeadID=" & val(vsGrid.TextMatrix(vsGrid.Row, 3))
                vsGrid.RemoveItem (vsGrid.Row)
            End If
        End If
        If KeyCode = 13 Then
            If vsGrid.TextMatrix(vsGrid.Row, 4) <> "" And vsGrid.Row = vsGrid.Rows - 1 Then
                vsGrid.Rows = vsGrid.Rows + 1
            End If
        End If
        Call Calculate
    End Sub
    Private Sub Calculate()
        Dim mCnt    As Integer
        Dim mTotal  As Double
        mTotal = 0
        For mCnt = 1 To vsGrid.Rows - 1
            mTotal = mTotal + val(vsGrid.TextMatrix(mCnt, 2))
        Next
        txtTotal.Text = Format(mTotal, "#.00")
    End Sub
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim mToken As String
        If vsGrid.TextMatrix(vsGrid.Row - 1, 2) <> "" Then
            frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where intGroupID in (1,2) AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
            If gbSearchID <> -1 Then
                If vsGrid.FindRow(gbSearchID, 1, 3) > 0 Then
                    MsgBox "Selected Account Head Already in the List...."
                    Exit Sub
                End If
                vsGrid.TextMatrix(vsGrid.Row, 0) = Token(gbSearchStr, " ")
                vsGrid.TextMatrix(vsGrid.Row, 1) = Trim(gbSearchStr)
                vsGrid.TextMatrix(vsGrid.Row, 3) = gbSearchID
                gbSearchID = -1
                gbSearchStr = ""
                If vsGrid.FindRow(" ", 1, 3) = -1 Then
                    vsGrid.Rows = vsGrid.Rows + 1
                End If
            End If
        Else
            MsgBox "Please Complete Previous Row..."
            Exit Sub
        End If
      
    End Sub
    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGrid.Col = 2 Then
            If Not (((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) And KeyAscii <> 47) Then KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    End Sub
    Private Sub Form_Activate()
'        Me.Top = 2000
'        Me.Width = 10185
'        Me.Left = (frmMenu.Width - Me.Width) / 2
     End Sub


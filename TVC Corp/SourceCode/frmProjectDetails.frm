VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProjectDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Register"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRevisedProjects 
      Caption         =   "Revised Projects"
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
      Left            =   11655
      TabIndex        =   2
      Top             =   6840
      Width           =   1635
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   585
      Width           =   13380
      _cx             =   23601
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProjectDetails.frx":0000
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
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   13320
      Top             =   7020
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13380
      TabIndex        =   0
      Top             =   0
      Width           =   13380
   End
End
Attribute VB_Name = "frmProjectDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRevisedProjects_Click()
    Call FillGrid
End Sub

Private Sub Form_Load()
    XPC.InitSubClassing
    Call FormInitialize
    Call FillGrid
    'Call cmdUpdateNewAllotments
End Sub
Private Function CheckRequisitionRegister(mID As Integer) As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String

    If objDB.SetConnection(mCnn) Then
        mSQL = " Select intCountOfVouchers from faAllotments Where intID=" & mID
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If IIf(IsNull(Rec!intCountOfVouchers), 0, Rec!intCountOfVouchers) = 0 Then
                CheckRequisitionRegister = False
            Else
                CheckRequisitionRegister = True
            End If
        End If
        Rec.Close
    End If
    mCnn.Close
End Function
Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
    Call FillGrid
End Sub
Private Sub FormInitialize()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
            ctrl.Tag = ""
        ElseIf TypeOf ctrl Is OptionButton Then
            ctrl.value = False
        ElseIf TypeOf ctrl Is ComboBox Then
            If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
            ctrl.Tag = ""
        End If
    Next
End Sub
Public Sub FillGrid()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mRowCnt As Integer
    Dim mArrayIn As Variant
    Dim mLoop As Integer
    Dim mRevise As Integer
    
    'On Error GoTo Err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    If chkRevisedProjects.value = 1 Then
        mRevise = 1
    Else
        mRevise = 0
    End If
    
    'mArrayIn = Array(gbFinancialYearID, mRevise)
    mArrayIn = Array(2012, mRevise)
    Set Rec = objDB.ExecuteSP("spGetProjectRegisterDetails", mArrayIn, , , mCnn)
    mRowCnt = 1
    vsGrid.Clear 1, 1
    vsGrid.Rows = 1
    While Not (Rec.EOF Or Rec.BOF)
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
        vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate))
        'vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtAllotmentDate), Rec!dtRequisitionDate, Rec!dtAllotmentDate))
        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
        vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
        vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
        vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
        vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
        vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
        vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
        
        If Rec!tnyProjectStatus = 2 Then
            If vsGrid.TextMatrix(mRowCnt, 6) <> "" Then
                vsGrid.TextMatrix(mRowCnt, 11) = vbChecked
            ElseIf ChecksuExpenditure(IIf(IsNull(Rec!intID), 0, Rec!intID)) = True Then
                vsGrid.TextMatrix(mRowCnt, 11) = vbChecked
            End If
        ElseIf Rec!tnyProjectStatus = 1 Then
            vsGrid.TextMatrix(mRowCnt, 11) = vbUnchecked
            For mLoop = 0 To vsGrid.Cols - 1
                vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, mLoop) = &HE0E0E0
            Next mLoop
        Else
            vsGrid.TextMatrix(mRowCnt, 11) = vbUnchecked
        End If
        vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!intID), "", Rec!intID)
        vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!tnyProjectStatus), 0, Rec!tnyProjectStatus)
    Rec.MoveNext
    mRowCnt = mRowCnt + 1
    Wend
    Rec.Close
    Exit Sub
Err:
    MsgBox Err.Description
End Sub


        
    Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mTrAccHeadId As Integer
        
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                GetStatusFlag = -1
            End If
            Rec.Close
        End If
    End Function

Private Sub vsGrid_DblClick()
    Dim mCnn  As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSQL  As String
    Dim mLoop As Integer
    
    
    'BLOCK [1]
    'NOTE:- CHECKING Source of Fund Extraction is done or Not
    '       If done, no changes will be permitted in Requistion Register
        Dim mExtractedStatus As Integer
        Dim mMsg As String
        
        mMsg = ""
        mMsg = mMsg + "Previous year's Source wise transactions are all closed by Secretary" & vbCrLf
        mMsg = mMsg + "by brought down Source wise balances to new financial year by declaring the Source wise balances are correct." & vbCrLf
        mMsg = mMsg + "" & vbCrLf
        mMsg = mMsg + "Further changes in previous year's source wise transaction will" & vbCrLf
        mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted in Project Register" & vbCrLf
        
        mExtractedStatus = GetStatusFlag
        If mExtractedStatus = 2 Then
           MsgBox mMsg, vbInformation
          Exit Sub
        End If
    'END OF BLOCK[1]
    
    
    If CheckRequisitionRegister(vsGrid.TextMatrix(vsGrid.Row, 12)) = False Then
        MsgBox "Please Verify Requisition Register", vbInformation, "Saankhya"
        Exit Sub
    Else
        If vsGrid.Row > 0 Then
            For mLoop = 1 To vsGrid.Row - 1
               If vsGrid.TextMatrix(mLoop, 13) = 0 And vsGrid.Row <> 1 Then 'vsGrid.Cell(flexcpChecked, mLoop, 13)
                   MsgBox "Verify the Previous Requisition"
                   Exit Sub
               End If
            Next mLoop
            frmProjectRegisterDetails.ReqID = vsGrid.TextMatrix(vsGrid.Row, 12)
            frmProjectRegisterDetails.Show
        End If
    End If
End Sub
Private Function ChecksuExpenditure(mID As Integer) As Boolean
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objDB     As New clsDB
    Dim mSQL    As String
    
  
    If objDB.SetConnection(mCnn) Then
        mSQL = " SELECT *  FROM suExpenditures Where intAllotmentID=" & mID
        Rec.Open mSQL, mCnn, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
            ChecksuExpenditure = True
        Else
            ChecksuExpenditure = True
        End If
        Rec.Close
    End If
    'mCnn.Close
End Function


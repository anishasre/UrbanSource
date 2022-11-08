VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmRequisitionRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requisition Register"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   11475
      Top             =   6615
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6045
      Left            =   0
      TabIndex        =   1
      Top             =   495
      Width           =   11715
      _cx             =   20664
      _cy             =   10663
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRequisitionRegister.frx":0000
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   -45
      ScaleHeight     =   735
      ScaleWidth      =   11760
      TabIndex        =   0
      Top             =   0
      Width           =   11760
   End
End
Attribute VB_Name = "frmRequisitionRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Sub Form_Load()
     XPC.InitSubClassing
     Call FormInitialize
     Call FillGrid
End Sub
Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
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
    
    On Error GoTo Err
    objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
    
    mSQL = " Select * from faAllotments "
    mSQL = mSQL + " Left Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID "
    mSQL = mSQL + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID "
    mSQL = mSQL + " Where intFinancialYearID=2012 And Isnull(tnyStatus,0) = 1 And tnyStage=2 AND NOT ISNULL(tnyTypeID,0) IN (1,2) Order by dtRequisitionDate " 'vchRequisitionNo
    Rec.CursorLocation = adUseClient
        Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        mRowCnt = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        While Not (Rec.EOF Or Rec.BOF)
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
            vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
            vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
            vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
            vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!fltRequestedAmt), "0.00", Rec!fltRequestedAmt)
            vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!fltAuthorizedAmt), "0.00", Rec!fltAuthorizedAmt)
            If Rec!intCountOfVouchers = 1 Then
                vsGrid.TextMatrix(mRowCnt, 7) = vbChecked
            Else
                vsGrid.TextMatrix(mRowCnt, 7) = vbUnchecked
            End If
            vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intID), "", Rec!intID)
            vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
            vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
            vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
            vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!intSchemeID), "", Rec!intSchemeID)
            vsGrid.TextMatrix(mRowCnt, 13) = IIf(IsNull(Rec!intImplementingOfficersID), "", Rec!intImplementingOfficersID)
            If Not IsNull(Rec!dtAuthorizationDate) Then
                vsGrid.TextMatrix(mRowCnt, 14) = DdMmmYy(IIf(IsNull(Rec!dtAuthorizationDate), "", Rec!dtAuthorizationDate))
            End If
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
        mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted in Requisition Register" & vbCrLf
        
        mExtractedStatus = GetStatusFlag
        If mExtractedStatus = 2 Then
           MsgBox mMsg, vbInformation
          Exit Sub
        End If
    'END OF BLOCK[1]
    '
    
    If vsGrid.Row > 0 Then
        If vsGrid.Cell(flexcpChecked, vsGrid.Row - 1, 7) = 2 And vsGrid.Row <> 1 Then
            MsgBox "Verify the Previous Requisition"
            Exit Sub
        Else
            'frmRequisitionRegisterDetails.RequisitionID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
            frmRequisitionRegisterDetails.txtReqNo.Text = vsGrid.TextMatrix(vsGrid.Row, 0)
            frmRequisitionRegisterDetails.txtReqNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 8)
            frmRequisitionRegisterDetails.txtReqDate.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
            frmRequisitionRegisterDetails.txtReqDate.Tag = vsGrid.TextMatrix(vsGrid.Row, 14)
            frmRequisitionRegisterDetails.txtImpo.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
            
            frmRequisitionRegisterDetails.txtImpo.Tag = vsGrid.TextMatrix(vsGrid.Row, 13)
            frmRequisitionRegisterDetails.txtIMPODesig.Text = vsGrid.TextMatrix(vsGrid.Row, 11)
            frmRequisitionRegisterDetails.txtSource.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
            frmRequisitionRegisterDetails.txtSource.Tag = vsGrid.TextMatrix(vsGrid.Row, 9)
            
            frmRequisitionRegisterDetails.txtCategory.Text = vsGrid.TextMatrix(vsGrid.Row, 4)
            frmRequisitionRegisterDetails.txtCategory.Tag = vsGrid.TextMatrix(vsGrid.Row, 10)
            frmRequisitionRegisterDetails.vsGrid.TextMatrix(0, 1) = vsGrid.TextMatrix(vsGrid.Row, 5)
            frmRequisitionRegisterDetails.vsGrid.TextMatrix(1, 1) = vsGrid.TextMatrix(vsGrid.Row, 6)
            
            If vsGrid.TextMatrix(vsGrid.Row, 7) = vbChecked Then
                frmRequisitionRegisterDetails.lblmsg.Visible = True
                frmRequisitionRegisterDetails.lblmsg.Caption = "Already Approved"
                frmRequisitionRegisterDetails.cmdSave.Enabled = False
            End If
            
              If objDB.SetConnection(mCnn) Then
                    If vsGrid.TextMatrix(vsGrid.Row, 12) <> "" Then
                        mSQL = " Select intID,vchDescription from faDepSchPro Where intID= " & vsGrid.TextMatrix(vsGrid.Row, 12) & " "
                        Rec.Open mSQL, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                        frmRequisitionRegisterDetails.txtCategory.Tag = Rec!intID
                        frmRequisitionRegisterDetails.txtCategory.Text = Rec!vchDescription
                        frmRequisitionRegisterDetails.lblCategory.Caption = "Scheme"
                        End If
                        Rec.Close
                    End If
                End If
            frmRequisitionRegisterDetails.FillGrid
            frmRequisitionRegisterDetails.Show
    End If
    End If
End Sub

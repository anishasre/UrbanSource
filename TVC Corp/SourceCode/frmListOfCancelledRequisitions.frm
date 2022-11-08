VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfCancelledRequisitions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of Cancelled Requisitions"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   Icon            =   "frmListOfCancelledRequisitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   12795
      TabIndex        =   2
      Top             =   6000
      Width           =   12855
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   12120
         Top             =   720
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   12735
      _cx             =   22463
      _cy             =   8916
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
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListOfCancelledRequisitions.frx":1CCA
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12855
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmListOfCancelledRequisitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    
    Private Sub cmdNew_Click()
        frmCancelRequisitions.lblcaption = "This form records the details of Requisitions to be Cancelled"
        frmCancelRequisitions.Show vbModal
        Call FillGrid
    End Sub

     Private Sub Form_Load()
         XPC.InitSubClassing
         Call FormInitialize
         Call FillGrid
    End Sub
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
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdNew.Enabled = False
        End If
    End Sub
     Private Sub FillGrid()
        
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        Dim mRowCnt As Integer
        On Error GoTo err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        '  Changed by Aiby: 10-Nov-2011
        '        mSql = " SELECT faAllotments.vchRequisitionNo,faAllotments.dtRequisitionDate,faAllotments.vchNameofIMPO,faAllotments.vchDesignation,faAllotments.fltRequestedAmt,faReasons.vchReason,"
        '        mSql = mSql + " suSourceOfFund.vchSourceFundName,faTransactionCategory.vchTransactionCategory,"
        '        mSql = mSql + " faAllotments.tnyStatus,faAllotments.intID , faAllotments.intAuthorizedByUserID, faAllotments.intSourceID, faAllotments.intFundCategoryID,"
        '        mSql = mSql + " fltAuthorizedAmt From faAllotments"
        '        mSql = mSql + " left join faRequisitionRequest on faRequisitionRequest.vchRequisitionNo=faAllotments.vchRequisitionNo "
        '        mSql = mSql + " left join faReasons on faReasons.intReasonID=faRequisitionRequest.intReasonID "
        '        mSql = mSql + " Left Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
        '        mSql = mSql + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID"
        '        mSql = mSql + " where faAllotments.tnyStatus in (0,2)"
        '        mSql = mSql + " Order by dtRequisitionDate"
        '
        
        mSql = " SELECT faAllotments.vchRequisitionNo,faAllotments.dtRequisitionDate,faAllotments.vchNameofIMPO,faAllotments.vchDesignation,faAllotments.fltRequestedAmt,faReasons.vchReason,"
        mSql = mSql + " suSourceOfFund.vchSourceFundName,faTransactionCategory.vchTransactionCategory,"
        mSql = mSql + " faRequisitionRequest.tnyStatus,faAllotments.intID , faAllotments.intAuthorizedByUserID, faAllotments.intSourceID, faAllotments.intFundCategoryID,"
        mSql = mSql + " fltAuthorizedAmt From faAllotments"
        mSql = mSql + " INNER join faRequisitionRequest on faRequisitionRequest.intRequisitionID=faAllotments.intID "
        mSql = mSql + " left join faReasons on faReasons.intReasonID=faRequisitionRequest.intReasonID "
        mSql = mSql + " Left Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
        mSql = mSql + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID"
        mSql = mSql + " Order by dtRequisitionDate"
        Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mRowCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                'vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
                vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
                vsGrid.TextMatrix(mRowCnt, 1) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
                vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
                vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
                vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                
                ' Changed by Aiby :10-Nov-2011
                '    If Rec!tnyStatus = 1 Then
                '         vsGrid.TextMatrix(mRowCnt, 8) = "Waiting for Approval"
                '    ElseIf Rec!tnyStatus = 2 Then
                '         vsGrid.TextMatrix(mRowCnt, 8) = "Cancelled"
                '    End If
                
                If Rec!tnyStatus = 0 Then
                     vsGrid.TextMatrix(mRowCnt, 8) = "Waiting for Approval"
                ElseIf Rec!tnyStatus = 1 Then
                     vsGrid.TextMatrix(mRowCnt, 8) = "Approved"
                End If
                               
                vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!intID), "", Rec!intID)
                vsGrid.TextMatrix(mRowCnt, 10) = IIf(IsNull(Rec!intAuthorizedByUserID), "", Rec!intAuthorizedByUserID)
                vsGrid.TextMatrix(mRowCnt, 11) = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
                vsGrid.TextMatrix(mRowCnt, 12) = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub vsGrid_DblClick()
    
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSql As String
        Dim mArr As Variant
        
        If vsGrid.Row > 0 Then
            If vsGrid.TextMatrix(vsGrid.Row, 8) = "Cancelled" Then
                MsgBox "Requisition Cancelled", vbInformation
                Exit Sub
            End If
            frmCancelRequisitions.txtRequisitionNo.Text = vsGrid.TextMatrix(vsGrid.Row, 0)
            frmCancelRequisitions.txtRequisitionNo.Tag = vsGrid.TextMatrix(vsGrid.Row, 9)
            frmCancelRequisitions.txtReason.Text = vsGrid.TextMatrix(vsGrid.Row, 5)
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSql = " select * from faRequisitionRequest "
            mSql = mSql + " where intRequisitionId= " & vsGrid.TextMatrix(vsGrid.Row, 9) & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmCancelRequisitions.txtRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
            End If
            If vsGrid.TextMatrix(vsGrid.Row, 8) = "Approved" Then
                frmCancelRequisitions.cmdCancel.Tag = 1
                frmCancelRequisitions.cmdCancel.Visible = False
                frmCancelRequisitions.cmdRequest.Enabled = False
            Else
                frmCancelRequisitions.cmdCancel.Tag = 0
                frmCancelRequisitions.cmdCancel.Visible = True
            End If
            
            frmCancelRequisitions.Show vbModal
            Call FillGrid
        End If
        
    End Sub
     Private Sub Form_Paint()
        Me.Top = 0
        Me.Left = (Screen.Width - Me.Width) / 2
        Call FillGrid
     End Sub


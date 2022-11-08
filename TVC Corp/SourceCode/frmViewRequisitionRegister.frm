VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmViewRequisitionRegister 
   BackColor       =   &H00FAFAFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REQUISITION   REGISTER"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   375
      TabIndex        =   8
      Top             =   6090
      Width           =   6420
      Begin VB.TextBox txtProject 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1515
         TabIndex        =   13
         Top             =   795
         Width           =   4275
      End
      Begin VB.CommandButton cmdProject 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdSourceOfFund 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   345
         Width           =   375
      End
      Begin VB.TextBox txtSourceOfFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1515
         TabIndex        =   10
         Top             =   360
         Width           =   4275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROJECT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   750
         TabIndex        =   14
         Top             =   825
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOURCE OF FUND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   405
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   10080
      TabIndex        =   4
      Top             =   6075
      Width           =   1680
      Begin VB.CheckBox chkWaiting 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "WAITING..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   270
         TabIndex        =   15
         Top             =   705
         Width           =   1290
      End
      Begin VB.CheckBox chkCancelled 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CANCELLED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   270
         TabIndex        =   7
         Top             =   930
         Width           =   1290
      End
      Begin VB.CheckBox chkApproved 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "APPROVED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   480
         Width           =   1290
      End
      Begin VB.CheckBox chkAuthorized 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AUTHORISED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   255
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   6840
      TabIndex        =   1
      Top             =   6075
      Width           =   3165
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   765
         Width           =   1320
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1635
         TabIndex        =   3
         Top             =   405
         Width           =   1290
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   1290
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4875
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   12375
      _cx             =   21828
      _cy             =   8599
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16053492
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   16448
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmViewRequisitionRegister.frx":0000
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
   Begin VB.Label lblAmtCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4200
      TabIndex        =   18
      Top             =   5805
      Width           =   450
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000A0A89&
      Height          =   285
      Left            =   4830
      TabIndex        =   17
      Top             =   5745
      Width           =   1380
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   12750
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escape to Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   705
      TabIndex        =   16
      Top             =   7530
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      FillColor       =   &H00000080&
      Height          =   150
      Left            =   390
      Shape           =   3  'Circle
      Top             =   7560
      Width           =   240
   End
End
Attribute VB_Name = "frmViewRequisitionRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Option Explicit

    Private Sub chkApproved_Click()
        If chkApproved.value = vbChecked Then
            chkAuthorized.value = vbUnchecked
            chkWaiting.value = vbUnchecked
            chkCancelled.value = vbUnchecked
        End If
        Call FillGrid
    End Sub
    
    Private Sub chkAuthorized_Click()
        If chkAuthorized.value = vbChecked Then
            chkApproved.value = vbUnchecked
            chkWaiting.value = vbUnchecked
            chkCancelled.value = vbUnchecked
        End If
        Call FillGrid
    End Sub
    
    Private Sub chkCancelled_Click()
        If chkCancelled.value = vbChecked Then
            chkAuthorized.value = vbUnchecked
            chkApproved.value = vbUnchecked
            chkWaiting.value = vbUnchecked
        End If
        Call FillGrid
    End Sub
    
    Private Sub chkWaiting_Click()
        If chkWaiting.value = vbChecked Then
            chkAuthorized.value = vbUnchecked
            chkApproved.value = vbUnchecked
            chkCancelled.value = vbUnchecked
        End If
        Call FillGrid
    End Sub
    
    Private Sub cmbYear_Click()
        Call FillGrid
    End Sub
    
    Private Sub cmdProject_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        If cmbYear.ListIndex > 0 Then
            frmSearchMasters.SQLQry = "SELECT decProjectID, chvProjectSlNo + '" & " - " & "' + chvProjectnameEnglish From suProjectDetails WHERE tnyStatus = 9 AND intYearID = " & cmbYear.ItemData(cmbYear.ListIndex) & " Order by decProjectID "
        Else
            frmSearchMasters.SQLQry = "SELECT decProjectID, chvProjectSlNo + '" & " - " & "' + chvProjectnameEnglish From suProjectDetails WHERE tnyStatus = 9  Order by decProjectID "
        End If
       
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtProject.Text = gbSearchStr
            txtProject.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            Call FillGrid
        End If
        
    End Sub

    Private Sub cmdSourceOfFund_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund  "
        frmSearchMasters.Show vbModal
        'txtSourceOfFund.SetFocus
        If gbSearchID <> -1 Then
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            Call FillGrid
        End If
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Call FormInitializ
            Call FillGrid
        End If
    End Sub

    Private Sub Form_Load()
        Dim mSql As String
        
        Call FormInitializ
        
        
        On Error Resume Next
        mSql = "Select LTRIM(Str(intFinancialYear)) + '-' + LTRIM(Str(intFinancialYear+1)), intFinancialYearID  From faFinancialYear"
        PopulateList cmbYear, mSql, True, False, True, True
        cmbYear.Text = Trim(str(gbFinancialYearID)) + "-" + Trim(str((gbFinancialYearID + 1)))
        
        Call FillGrid
    End Sub
    
    Private Sub FormInitializ()
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            End If
        Next
    End Sub
    
    Private Sub FillGrid()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objDB   As New clsDB
        Dim mLoop   As Integer
        Dim mDate   As String
        Dim mAllotmentDate As String
        Dim mAmt As Double
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "SELECT  intID, dtRequisitionDate,vchRequisitionNo ,vchAllotmentNo,dtAllotmentDate, fltRequestedAmt fltAuthorizedAmt,"
        mSql = mSql + " suSourceOfFund.vchSourceFundName , vchProjectNo, vchPayOrderNo, faVouchers.intVoucherNo, faVouchers.dtDate, faAllotments.tnyStatus,faAllotments.tnyStage From faAllotments "
        mSql = mSql + " LEFT JOIN  faPayOrder ON faAllotments.intID = faPayOrder.intAllotmentID AND ISNULL(faPayOrder.tnyCancelled,0) <>1 "
        mSql = mSql + " LEFT JOIN  faVouchers ON faPayOrder.intVoucherID=faVouchers.intVoucherID "
        mSql = mSql + " INNER JOIN suSourceOfFund ON faAllotments.intSourceID = suSourceOfFund.intSourceFundID "
        mSql = mSql + " WHERE Not faAllotments.tnyStatus IS Null "
       
    
        If cmbYear.ListIndex > 0 Then
            mSql = mSql + " AND faAllotments.intFinancialYearID =  " & cmbYear.ItemData(cmbYear.ListIndex) & ""
            'mSql = mSql + " AND faAllotments.intFinancialYearID =  " & gbFinancialYearID - 1
        End If
        
        If val(txtProject.Tag) > 0 Then
            mSql = mSql + " AND numProjectID = " & txtProject.Tag
        End If
        
        If val(txtSourceOfFund.Tag) > 0 Then
            mSql = mSql + " AND intSourceFundID = " & val(txtSourceOfFund.Tag)
        End If
        
        If IsDate(txtFromDate.Text) And IsDate(txtToDate.Text) Then
            mSql = mSql + " AND dtRequisitionDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & " '"
        End If
        
        If chkAuthorized.value = vbChecked Then
            mSql = mSql + " AND ( faAllotments.tnyStage = 2 AND faAllotments.tnyStatus = 1 )"
        End If
        
        If chkApproved.value = vbChecked Then
            mSql = mSql + " AND ( faAllotments.tnyStage = 2 AND faAllotments.tnyStatus = 0 ) "
        End If
        
        If chkWaiting.value = vbChecked Then
            mSql = mSql + " AND ( faAllotments.tnyStage = 1 AND faAllotments.tnyStatus = 0 )"
        End If
        
        If chkCancelled.value = vbChecked Then
            mSql = mSql + " AND faAllotments.tnyStatus = 2 "
        End If
        
        mAmt = 0
        lblAmount.Caption = ""
        
        Rec.Open mSql, mCnn
        vsGrid.Rows = 1
        If Not (Rec.BOF And Rec.EOF) Then
            While Not (Rec.EOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!intID
                vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = DdMmmYy(IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate))
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                If Rec!dtAllotmentDate <> "" Then
                      mAllotmentDate = DdMmmYy(Rec!dtAllotmentDate)
                Else
                      mAllotmentDate = ""
                End If
                vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = mAllotmentDate
               
                vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!fltAuthorizedAmt), "", Rec!fltAuthorizedAmt)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                vsGrid.TextMatrix(vsGrid.Rows - 1, 9) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                If Rec!dtDate <> "" Then
                      mDate = DdMmmYy(Rec!dtDate)
                Else
                        mDate = ""
                End If
                vsGrid.TextMatrix(vsGrid.Rows - 1, 10) = mDate
                
                If Rec!tnyStatus = 1 And Rec!tnyStage = 2 Then
                     mAmt = mAmt + IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = "Authorized"
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, 11) = &HC3FFC3
                ElseIf Rec!tnyStatus = 0 And Rec!tnyStage = 2 Then
                     mAmt = mAmt + IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = "Approved"
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, 11) = &HD4FFFF
                ElseIf Rec!tnyStatus = 0 And Rec!tnyStage = 1 Then
                    mAmt = mAmt + IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = "Waiting..."
                    'vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, 11) = &HE3F8E3
                ElseIf Rec!tnyStatus = 2 Then
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 11) = "Cancelled"
                    vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, 11) = &HDCDCDC
                End If
                Rec.MoveNext
            Wend
            
            lblAmount.Caption = Format(mAmt, "0.00  ")
        End If
    End Sub
    
    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub
    
    Private Sub txtFromDate_LostFocus()
        If Not IsDate(txtFromDate.Text) Then
            txtFromDate.Text = DdMmmYy(gbStartingDate)
        Else
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    
    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub
    
    Private Sub txtToDate_LostFocus()
        If Not IsDate(txtToDate.Text) Then
             txtToDate.Text = DdMmmYy(gbTransactionDate)
        Else
             txtToDate.Text = CheckDateInMMM(Trim(txtToDate))
        End If
             
        If txtFromDate.Text <> "" Then
           Call FillGrid
        End If
    End Sub
        
    Private Sub vsGrid_DblClick()
        If vsGrid.Row > 0 Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 0)) > 0 Then
'                frmRequisition.RequisitionID = val(vsGrid.TextMatrix(vsGrid.Row, 0))
'                frmRequisition.Show vbModal
            End If
        End If
    End Sub

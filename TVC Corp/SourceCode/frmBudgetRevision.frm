VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmBudgetRevision 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Budget Revision"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraApprove 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1125
      Left            =   90
      TabIndex        =   23
      Top             =   5250
      Width           =   9705
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1155
         TabIndex        =   29
         ToolTipText     =   "Reference No. of Budget Revision Committee"
         Top             =   540
         Width           =   2295
      End
      Begin VB.Frame fraApproval 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Approval"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3870
         TabIndex        =   24
         Top             =   -120
         Width           =   5715
         Begin VB.OptionButton optRejected 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rejected"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   450
            TabIndex        =   27
            Top             =   660
            Width           =   975
         End
         Begin VB.OptionButton optApproved 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Approved"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   450
            TabIndex        =   26
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox txtStatusRemarks 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   2460
            TabIndex        =   25
            ToolTipText     =   "Reference No. of Budget Revision Committee"
            Top             =   330
            Width           =   3105
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1650
            TabIndex        =   28
            Top             =   360
            Width           =   765
         End
      End
      Begin MSComCtl2.DTPicker dtRevised 
         Height          =   315
         Left            =   1140
         TabIndex        =   30
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   17760259
         CurrentDate     =   39415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revised On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   32
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   31
         Top             =   510
         Width           =   660
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   9690
      Top             =   6330
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.ListBox lstBudgetCentre 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8220
      TabIndex        =   21
      Top             =   330
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   3630
      TabIndex        =   17
      Top             =   6420
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4890
      TabIndex        =   18
      Top             =   6420
      Width           =   1215
   End
   Begin VB.Frame FraBudgetCntr 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Budget Centre Details"
      Height          =   2385
      Left            =   90
      TabIndex        =   19
      Top             =   30
      Width           =   9705
      Begin VB.TextBox txtBudgetRevisionIDHide 
         Height          =   315
         Left            =   7920
         TabIndex        =   22
         Top             =   1740
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtFinancialYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   13
         Top             =   1980
         Width           =   2295
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   11
         Top             =   1680
         Width           =   4515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   9
         Top             =   1380
         Width           =   4515
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   7
         Top             =   1080
         Width           =   4515
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   5
         Top             =   780
         Width           =   4515
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Left            =   7650
         TabIndex        =   3
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtBudgetCentre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3930
         TabIndex        =   2
         Top             =   300
         Width           =   3705
      End
      Begin VB.TextBox txtBudgetCentreCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financial year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1365
         TabIndex        =   12
         Top             =   2010
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2085
         TabIndex        =   10
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   4
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   6
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Field"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2100
         TabIndex        =   8
         Top             =   1410
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Centre Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   825
         TabIndex        =   0
         Top             =   330
         Width           =   1680
      End
   End
   Begin VB.Frame fraAccntDetails 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Budget Allocation Details"
      Height          =   2805
      Left            =   90
      TabIndex        =   20
      Top             =   2430
      Width           =   9705
      Begin VB.TextBox txtRevisedAmt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6450
         TabIndex        =   16
         Top             =   2490
         Width           =   1395
      End
      Begin VB.TextBox txtAllotedAmt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Top             =   2490
         Width           =   1515
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2205
         Left            =   30
         TabIndex        =   14
         Top             =   270
         Width           =   9645
         _cx             =   17013
         _cy             =   3889
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBudgetRevision.frx":0000
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
   End
End
Attribute VB_Name = "frmBudgetRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
        Dim objDB                       As New clsDB
        Dim mCon                        As New ADODB.Connection
        Dim mSQL                        As String
        Dim Rec                         As New ADODB.Recordset
        Dim objBc                       As New clsBudgetCentre
        Dim mGridRows                   As Long
        Dim mEditFlag                   As Boolean
    
    Private Sub FormInitialize()
        txtBudgetCentreCode.Text = ""
        txtBudgetCentre.Text = ""
        txtFunctionary.Text = ""
        txtField.Text = ""
        txtFunction.Text = ""
        txtFund.Text = ""
        fraAccntDetails.Enabled = True
        fraAccntDetails.Enabled = True
        fraApproval.Enabled = True
        mEditFlag = False
        txtAllotedAmt.Text = ""
        txtRevisedAmt.Text = ""
        txtBudgetRevisionIDHide.Tag = ""
        vsGrid.Rows = 1
        vsGrid.Rows = 100
        txtRefNo.Text = ""
        dtRevised.value = Date
        optApproved.value = False
        optRejected.value = False
        txtStatusRemarks.Text = ""
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdSave_Click()
        Dim objDB As New clsDB
        Dim mintApprovalStatus      As Integer
        '----------------------------------------------------'
        ' Validations
        '----------------------------------------------------'
        ' Budget Centre
        objBc.SetBudgetCentre (Trim(txtBudgetCentreCode.Text))
        
        If objBc.BudgetCentreID < 1 Then
            MsgBox "Select a Budget Centre!", vbInformation
            txtBudgetCentreCode.SetFocus
            Exit Sub
        End If
        
        '----------------------------------------------------'
        '  UPDATING DATABASE
        '----------------------------------------------------'
        Dim arrInput                As Variant
        Dim arrOutput               As Variant
        Dim arrIn4Approve           As Variant
        Dim arrOut4Approve          As Variant
        Dim Rec                     As New ADODB.Recordset
        Dim Rec4NewAcHead           As New ADODB.Recordset
        Dim mCnn                    As ADODB.Connection
        Dim mintBudgetRevisionID    As Long
        Dim mLoopCnt                As Long
        Dim mintBudgetCentreID      As Long
        Dim objAcc                  As New clsAccounts
        Dim mTotalRevisedAmt        As Currency
        Dim mMajorAccountHeadID As Long
        
        
        mintBudgetCentreID = objBc.BudgetCentreID
        'gbUserTypeID = 4
      'gbUserTypeID = 2
        'gbUserID = 2
        If gbUserTypeID = 2 Then ' Approver
            If optApproved.value Then
                mintApprovalStatus = 1
            ElseIf optRejected.value Then
                mintApprovalStatus = 2
                mEditFlag = False
            End If
            arrInput = Array(val(txtBudgetRevisionIDHide.Tag), _
                        mintApprovalStatus, _
                        txtStatusRemarks.Text _
                        )
            objDB.SetConnection mCnn
            Set Rec = objDB.ExecuteSP("spUpdateBudgetRevisionHistory", arrInput, , , mCnn)
            'Set Rec4NewAcHead = GetRecordSet("spSaveBudgetRevisionHistory", mintBudgetCentreID)
             
'            If Not (Rec.BOF And Rec.EOF) Then
'                mMajorAccountHeadID = Rec!intMajorAccountHeadID
'                mMajorAccountHeadCode = Rec!vchMajorAccountHeadCode
'                mMajorAccountHead = Rec!vchMajorAccountHead
'                mMajorAccountTypeID = Rec!tinType
'            End If
            
            
        ElseIf gbUserTypeID > 2 Then ' Accounts officer or Operator
            '----------------------------------
            ' spSaveBudgetRevisionHistory
            '----------------------------------
            'intBudgetRevisionID , --1
            'dtBudgetRevision , --2
            'intBudgetCentreID , --3
            'vchRefNo , --4
            'intApproverUserID , --5
            'tinApprovalStatus , --6
            '/////dtStatusUpdatedOn ,   --7
            '/////vchStatusRemarks      --8
            '----------------------------------
            arrInput = Array(IIf(mEditFlag, val(txtBudgetRevisionIDHide.Tag), -1), _
                       Format(dtRevised.value, "DD/MmM/YYYY"), _
                       val(txtBudgetCentre.Tag), _
                       txtRefNo.Text, _
                       gbUserID, _
                       0 _
                       )
            objDB.SetConnection mCnn
            
            'mCnn.BeginTrans
            'On Error GoTo ErrRollBack
            
            Set Rec = objDB.ExecuteSP("spSaveBudgetRevisionHistory", arrInput, arrOutput, , mCnn)
            If IsNumeric(arrOutput(0, 0)) Then
                mintBudgetRevisionID = arrOutput(0, 0)
            End If
            
            
            mCnn.Execute "DELETE FROM faBudgetRevisionDetails WHERE intBudgetRevisionID = " & mintBudgetRevisionID
            For mLoopCnt = 1 To vsGrid.Rows - 1
                If Trim(vsGrid.TextMatrix(mLoopCnt, 1)) = "" Then
                    Exit For
                End If
                If val(vsGrid.TextMatrix(mLoopCnt, 4)) > 0 Then
                    objAcc.SetAccountCode (Trim(vsGrid.TextMatrix(mLoopCnt, 1)))
                    If objAcc.AccountHeadID < 1 Then
                        GoTo ErrRollBack
                    End If
                    mTotalRevisedAmt = mTotalRevisedAmt + Format(val(vsGrid.TextMatrix(mLoopCnt, 4)), "0.00")
                    arrInput = Array(mintBudgetRevisionID, _
                                objAcc.AccountHeadID, _
                                Format(val(vsGrid.TextMatrix(mLoopCnt, 4)), "0.00"), _
                                vsGrid.TextMatrix(mLoopCnt, 5) _
                                )
                    objDB.ExecuteSP "spSaveBudgetRevisionDetails", arrInput, , , mCnn
                    arrInput = Array(1, mintBudgetCentreID, objAcc.AccountHeadID, Format(val(vsGrid.TextMatrix(mLoopCnt, 4)), "0.00"))
                    objDB.ExecuteSP "spSaveBudgetAccountHead", arrInput, , , mCnn
                End If
            Next mLoopCnt
            If mTotalRevisedAmt = 0 Then
                GoTo ErrRollBack
            End If
            'mCnn.CommitTrans
            Set mCnn = Nothing
        End If
        Call FormInitialize
        gbUserTypeID = 0
        gbUserID = 0
        Exit Sub
        
ErrRollBack:
        'mCnn.RollbackTrans
        Set mCnn = Nothing
    End Sub
    
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        mGridRows = 0
        If gbUserTypeID = 2 Then        ' Approver
            
            fraAccntDetails.Enabled = False
            fraApprove.Enabled = True
        Else
            fraAccntDetails.Enabled = True
            fraApprove.Enabled = False
        End If
        vsGrid.ColComboList(1) = "|..."
    End Sub
    
    Private Sub Form_Activate()
            Me.Top = 0
            frmBudgetRevision.Left = (frmMenu.Width - Me.Width) / 2
            Call PopulateList(lstBudgetCentre, "Select vchBudgetCentre, intBudgetCentreID From faBudgetCentres Order By vchBudgetCentre", , , , True)
    End Sub

    Private Sub cmdSearch_Click()
            Call txtBudgetCentre_KeyDown(vbKeyF4, 0)
    End Sub


    Private Sub txtBudgetCentre_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            lstBudgetCentre.Width = 3500
            lstBudgetCentre.Left = 4000
            lstBudgetCentre.Height = 4000
            lstBudgetCentre.Visible = True
            lstBudgetCentre.SetFocus
        End If
    End Sub
    
    Private Sub lstBudgetCentre_DblClick()
            txtBudgetCentre.Text = lstBudgetCentre.Text
            If lstBudgetCentre.ListIndex > -1 Then
                objBc.SetBudgetCentreByID (lstBudgetCentre.ItemData(lstBudgetCentre.ListIndex))
                If objBc.BudgetCentreCode <> "" Then
                    txtBudgetCentreCode.Text = objBc.BudgetCentreCode
                End If
            End If
            lstBudgetCentre.Visible = False
            Call txtBudgetCentreCode_LostFocus
    End Sub
    
    Private Sub lstBudgetCentre_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                Call lstBudgetCentre_DblClick
            End If
    End Sub

    Private Sub txtBudgetCentre_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                PressTabKey
            End If
    End Sub
      
    Private Sub txtBudgetCentreCode_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
                Call txtBudgetCentre_KeyDown(vbKeyF4, 0)
            End If
    End Sub
    
    Private Sub txtBudgetCentreCode_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                PressTabKey
            End If
    End Sub
    
    Private Sub txtBudgetCentreCode_LostFocus()
            Dim mBudgetCenreID As Long
            Dim arrOut As Variant
            Dim RecNewHeads As New ADODB.Recordset
            txtBudgetCentreCode.Text = Trim(txtBudgetCentreCode)
            If Len(txtBudgetCentreCode) Then
                objBc.SetBudgetCentre (txtBudgetCentreCode.Text)
                mBudgetCenreID = objBc.BudgetCentreID
                If objBc.BudgetCentreID > -1 Then
                    'mEditFlag = True
                    txtBudgetCentreCode.Text = objBc.BudgetCentreCode
                    txtBudgetCentre.Text = objBc.BudgetCentre
                    txtBudgetCentre.Tag = objBc.BudgetCentreID
                    txtFunction.Text = objBc.FunctionName
                    txtFunction.Tag = objBc.FunctionID
                    txtFunctionary.Text = objBc.FunctionaryName
                    txtFunctionary.Tag = objBc.FunctionaryID
                    txtField.Text = objBc.FieldName
                    txtFund.Text = objBc.FundName
                    txtField.Tag = objBc.FieldID
                    txtFinancialYear.Tag = objBc.FinancialYearID
                    If txtFinancialYear.Tag <> 0 Then
                        Call DispFinancialYear(val(txtFinancialYear.Tag))
                    End If
                    objDB.SetConnection mCon
                        Set Rec = objDB.ExecuteSP("spGetLatestBudgetRevisionID", Array(mBudgetCenreID), arrOut, , mCon)
                        If IsNumeric(arrOut(0, 0)) Then
                            txtBudgetRevisionIDHide.Tag = arrOut(0, 0)
                        End If

                    Call FillAccountHeadDetails(objBc.BudgetCentreID)
                    Call CalculateAllotedAmt
                    Call CalculateRevisedAmt
                Else
                    mEditFlag = False
                    txtBudgetCentre.Text = ""
                    txtBudgetCentreCode.Text = ""
                End If
            End If
    End Sub
    
    Private Sub DispFinancialYear(mFinancialYearID)
            mSQL = "SELECT dtStartingDate, dtEndingDate, intFinancialYearID,tinCurrentFinancialYearFlag From faFinancialYear Where faFinancialYear.intFinancialYearID=" & mFinancialYearID
            Set Rec = GetRecordSet(mSQL)
            
            If Not (Rec.BOF And Rec.EOF) Then
                  txtFinancialYear.Text = Format(Rec!dtStartingDate, "Dd-Mmm-yyyy") & " -- " & Format(Rec!dtEndingDate, "Dd-Mmm-yyyy")
                  txtFinancialYear.Tag = Rec!intFinancialYearID
            End If
    End Sub
    
    Private Sub FillAccountHeadDetails(intBgtID As Long)
            vsGrid.Visible = False
            vsGrid.Rows = 1
            vsGrid.Rows = 100
            vsGrid.Visible = True
            mGridRows = 0
            Set Rec = objBc.GetAccountHeads(intBgtID)
            If Not (Rec.BOF And Rec.EOF) Then
            If Not IsNull(Rec!intBudgetRevisionID) Then
                mEditFlag = True
            End If
            dtRevised.value = Rec!dtBudgetRevision
            If IsNull(Rec!vchRefNo) Then
                txtRefNo.Text = ""
            Else
            
                txtRefNo.Text = Rec!vchRefNo
            End If
            While Not Rec.EOF
                mGridRows = mGridRows + 1
                vsGrid.AddItem mGridRows & vbTab & Rec!vchAccountHeadCode & vbTab & Rec!vchAccountHead & vbTab & Rec!fltEstimatedAmount & vbTab & Rec!fltRevisedAmount & vbTab & Rec!vchRemarks, mGridRows
                Rec.MoveNext
            Wend
            End If
            Rec.Close
            vsGrid.Visible = True
    End Sub
    
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Select Case Col
            Case 0, 2, 3
                Cancel = True
            Case 1, 4, 5
                Cancel = False
       End Select
       If Len(gbSearchStr) Then
            vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
            vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
            vsGrid.Col = vsGrid.Col + 2
            vsGrid.Redraw = flexRDDirect
            gbSearchStr = ""
        End If
    End Sub
    
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        If Col = 4 Then
            'If vsGrid.TextMatrix(Row, 0) <> "" Then
                vsGrid.TextMatrix(Row, 4) = Format(val(vsGrid.TextMatrix(Row, Col)), "0.00")
                Call CalculateRevisedAmt
            'End If
        End If
    End Sub

    Private Sub CalculateAllotedAmt()
        Dim mAmt As Double
        Dim mLoop As Long
                For mLoop = 1 To vsGrid.Rows - 1
                    mAmt = mAmt + val(vsGrid.TextMatrix(mLoop, 3))
                Next mLoop
                txtAllotedAmt.Text = Format(mAmt, "0.00")
    End Sub

    Private Sub CalculateRevisedAmt()
        Dim mAmt As Double
        Dim mLoop As Long
                For mLoop = 1 To vsGrid.Rows - 1
                    mAmt = mAmt + val(vsGrid.TextMatrix(mLoop, 4))
                Next mLoop
                txtRevisedAmt.Text = Format(mAmt, "0.00")
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinSecondaryAccountFlag=0"
        frmSearchAccountHeads.Show vbModal
    End Sub

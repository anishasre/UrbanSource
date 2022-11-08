VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmListOfRegisterOfBills 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register of Bills"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   12285
      TabIndex        =   5
      Top             =   7080
      Width           =   12345
      Begin VB.CommandButton cmdBillDetails 
         Caption         =   "&Billl Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6015
         TabIndex        =   9
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaymentOrder 
         Caption         =   "&PaymentOrder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7530
         TabIndex        =   8
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton cmdViewDemands 
         Caption         =   "&View Demand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4515
         TabIndex        =   7
         Top             =   60
         Width           =   1455
      End
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   10740
         Top             =   330
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3000
         TabIndex        =   6
         Top             =   60
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   12345
      TabIndex        =   0
      Top             =   0
      Width           =   12345
   End
   Begin VB.Frame Frame1 
      Height          =   6300
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   12345
      Begin VB.ComboBox cmbRegisterType 
         Height          =   315
         Left            =   8205
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   3225
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   5595
         Left            =   120
         TabIndex        =   3
         Top             =   675
         Width           =   11985
         _cx             =   21140
         _cy             =   9869
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmListOfRegisterOfBills.frx":0000
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Register Type"
         Height          =   195
         Left            =   7170
         TabIndex        =   4
         Top             =   300
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmListOfRegisterOfBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim mCheckDemandID As Variant
 Dim mTypeID As Variant
    Private Sub cmbRegisterType_Click()
        Call fillGrid
    End Sub

    Private Sub cmdBillDetails_Click()
    If vsGrid.Row > 0 Then
         mTypeID = vsGrid.TextMatrix(vsGrid.Row, 4)
         frmListofBillRegister.chkPaid.Visible = False
         frmListofBillRegister.cmbRegisterType.Enabled = False
         frmListofBillRegister.cmbRegisterType.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
         frmListofBillRegister.CheckBillID = 1
         frmListofBillRegister.CheckRegID = mTypeID
         frmListofBillRegister.Show vbModal
    End If
    End Sub
   Private Sub cmdNew_Click()
        frmRegisterOfBills.CheckDemand = 0
        frmRegisterOfBills.cmdSave.Caption = "Save"
        frmRegisterOfBills.Show vbModal
        Call fillGrid
   End Sub
    Private Sub cmdPaymentOrder_Click()
    If vsGrid.Row > 0 Then
         mTypeID = vsGrid.TextMatrix(vsGrid.Row, 4)
         frmListofBillRegister.chkPaid.Visible = False
         frmListofBillRegister.cmbRegisterType.Enabled = False
         frmListofBillRegister.cmbRegisterType.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
         frmListofBillRegister.CheckPaymentID = 1
         frmListofBillRegister.CheckRegID = mTypeID
         frmListofBillRegister.Show vbModal
    End If
    End Sub
   Private Sub cmdViewDemands_Click()
   If vsGrid.Row > 0 Then
        mTypeID = vsGrid.TextMatrix(vsGrid.Row, 4)
        frmListofBillRegister.chkPaid.Visible = False
        frmListofBillRegister.cmbRegisterType.Enabled = False
        frmListofBillRegister.cmbRegisterType.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
     
        frmListofBillRegister.CheckDemandID = 1
        frmListofBillRegister.CheckRegID = mTypeID
        frmListofBillRegister.Show vbModal
   Else
        frmListofBillRegister.Show vbModal
   End If
    End Sub
   Private Sub Form_Load()
        vsGrid.Cell(flexcpFontName, 0) = "Verdana"
        XPC.InitSubClassing
        Call PopulateList(cmbRegisterType, "Select vchRegType, intRegTypeID From faRegisterTypes Order By vchRegType", , , True, True)
        'cmbRegisterType.ListIndex = 1
   End Sub
   Private Sub Form_Activate()
'        Me.Top = 500
'        Me.Left = (frmMenu.Width - Me.Width) / 2
        Me.Left = 0
        Me.Top = 0
        Call fillGrid
    End Sub
    Private Sub fillGrid()
    Dim mSQL As String
    Dim objDb   As New clsDB
    Dim mCnn    As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim mRowCnt As Integer
    Dim mRecCnt As Integer
    Dim mTypeID As Integer
    
   
    
        If objDb.SetConnection(mCnn) Then
            mSQL = "SELECT faRegisterOfBills.intRegID,faRegisterOfBills.vchRegName, faRegisterOfBills.vchRefNo, faRegisterTypes.vchRegType"
            mSQL = mSQL + " FROM faRegisterOfBills LEFT JOIN faRegisterTypes ON faRegisterOfBills.intRegTypeID = faRegisterTypes.intRegTypeID"
            If cmbRegisterType.ListIndex > -1 Then
                mTypeID = cmbRegisterType.ItemData(cmbRegisterType.ListIndex)
                mSQL = mSQL + " WHERE faRegisterOfBills.intRegTypeID = " & mTypeID & " "
            End If
            Rec.CursorLocation = adUseClient
            Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            vsGrid.Rows = 1
            mRowCnt = 1
            mRecCnt = 1
            While Not (Rec.EOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = mRecCnt
                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
                vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRegType), "", Rec!vchRegType)
                vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
                mRecCnt = mRecCnt + 1
            Wend
            Rec.Close
        End If
        
        
        
''''''            If cmbRegisterType.ListIndex < 1 Then
''''''                msQl = " SELECT     faRegisterOfBills.intRegID,faRegisterOfBills.vchRegName, faRegisterOfBills.vchRefNo, faRegisterTypes.vchRegType"
''''''                msQl = msQl + " FROM faRegisterOfBills INNER JOIN faRegisterTypes ON faRegisterOfBills.intRegTypeID = faRegisterTypes.intRegTypeID order by faRegisterOfBills.intRegID "
''''''                Rec.CursorLocation = adUseClient
''''''                Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''''''                VSGrid.Rows = 2
''''''                mRowCnt = 1
''''''                mRecCnt = 1
''''''                VSGrid.Clear 1, 1
''''''                VSGrid.Rows = 1
''''''                While Not (Rec.EOF Or Rec.BOF)
''''''                    VSGrid.Rows = VSGrid.Rows + 1
''''''                    VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
''''''                    VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
''''''                    VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
''''''                    VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRegType), "", Rec!vchRegType)
''''''                    VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
''''''                    Rec.MoveNext
''''''
''''''                    mRowCnt = mRowCnt + 1
''''''                    mRecCnt = mRecCnt + 1
''''''
''''''                Wend
''''''                Rec.Close
''''''             Else
''''''
''''''                msQl = "SELECT faRegisterOfBills.intRegID,faRegisterOfBills.vchRegName, faRegisterOfBills.vchRefNo, faRegisterTypes.vchRegType"
''''''                msQl = msQl + " FROM faRegisterOfBills INNER JOIN faRegisterTypes ON faRegisterOfBills.intRegTypeID = faRegisterTypes.intRegTypeID WHERE faRegisterOfBills.intRegTypeID = " & mTypeID & " "
''''''                Rec.CursorLocation = adUseClient
''''''                Rec.Open msQl, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
''''''                VSGrid.Rows = 2
''''''                mRowCnt = 1
''''''                mRecCnt = 1
''''''                VSGrid.Clear 1, 1
''''''                While Not (Rec.EOF Or Rec.BOF)
''''''                    VSGrid.Rows = VSGrid.Rows + 1
''''''                    VSGrid.TextMatrix(mRowCnt, 0) = mRecCnt
''''''                    VSGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchRegName), "", Rec!vchRegName)
''''''                    VSGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
''''''                    VSGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchRegType), "", Rec!vchRegType)
''''''                    VSGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!intRegID), "", Rec!intRegID)
''''''                    Rec.MoveNext
''''''                    'vsGrid.Rows = vsGrid.Rows + 1
''''''                    mRowCnt = mRowCnt + 1
''''''                    mRecCnt = mRecCnt + 1
''''''                Wend
''''''                Rec.Close
''''''             End If
''''''         End If
    End Sub
    Private Sub VSGrid_DblClick()
    Dim mCnn As New ADODB.Connection
    Dim Rec  As New ADODB.Recordset
    Dim objDb As New clsDB
    Dim mSQL As String
        
        If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then Exit Sub
        
        If vsGrid.Row > 0 Then
        
            '----------IF DEMAND ALREADY GENERATED----------------------------------------
            
            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSQL = " SELECT intRegID,tnyStatus From faBillRegisters Where  (tnyStatus = 1 | 2 | 3)  And intRegID = " & vsGrid.TextMatrix(vsGrid.Row, 4) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                MsgBox "Demand Already Generated", vbInformation, "Saankhya"
                frmRegisterOfBills.CheckDemand = 1
                frmRegisterOfBills.cmdSave.Enabled = True
                frmRegisterOfBills.cmdDemand.Enabled = False
                frmRegisterOfBills.frm1.Enabled = True
                frmRegisterOfBills.txtRegName.Tag = Rec!intRegID
            Else
                frmRegisterOfBills.CheckDemand = 0
            End If
            Rec.Close
            
            '-----------------------------------------------------------------------------
            
            frmRegisterOfBills.cmbRegType = vsGrid.TextMatrix(vsGrid.Row, 3)
            frmRegisterOfBills.txtRefNo = vsGrid.TextMatrix(vsGrid.Row, 2)
            frmRegisterOfBills.txtRegName = vsGrid.TextMatrix(vsGrid.Row, 1)
            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSQL = " SELECT faRegisterOfBills.intRegID,faRegisterOfBills.vchRefTitle,faRegisterOfBills.tnyDay,faFunctions.intFunctionID,faFunctions.vchFunction,faFunctionaries.intFunctionaryID, faFunctionaries.vchFunctionary,faAccountHeads.intAccountHeadID, faAccountHeads.vchAccountHead, faPeriodicity.vchPeriodicity, "
            mSQL = mSQL + " faRegisterTypes.vchRegType FROM faRegisterOfBills INNER JOIN "
            mSQL = mSQL + " faFunctions ON faRegisterOfBills.intFunctionID = faFunctions.intFunctionID INNER JOIN "
            mSQL = mSQL + " faFunctionaries ON faRegisterOfBills.intFunctionaryID = faFunctionaries.intFunctionaryID INNER JOIN "
            mSQL = mSQL + " faAccountHeads ON faRegisterOfBills.intExpenditureHeadID = faAccountHeads.intAccountHeadID INNER JOIN "
            mSQL = mSQL + " faPeriodicity ON faRegisterOfBills.intPeriodicityID = faPeriodicity.intPeriodicityID INNER JOIN "
            mSQL = mSQL + " faRegisterTypes ON faRegisterOfBills.intRegTypeID = faRegisterTypes.intRegTypeID "
            mSQL = mSQL + " Where faRegisterOfBills.intRegID = " & vsGrid.TextMatrix(vsGrid.Row, 4) & " "
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                 frmRegisterOfBills.txtRegName.Tag = Rec!intRegID
                frmRegisterOfBills.txtFunction = Rec!vchFunction
                frmRegisterOfBills.txtFunctionary = Rec!vchFunctionary
                frmRegisterOfBills.txtAccountHead = Rec!vchAccountHead
                frmRegisterOfBills.txtRefTitle = Rec!vchRefTitle
                frmRegisterOfBills.cmbPeriod = Rec!vchPeriodicity
                frmRegisterOfBills.txtDueDate = Rec!tnyDay
               
                frmRegisterOfBills.txtFunctionary.Tag = Rec!intFunctionaryID
                frmRegisterOfBills.txtFunction.Tag = Rec!intFunctionID
                frmRegisterOfBills.txtAccountHead.Tag = Rec!intAccountHeadID
            End If
            Rec.Close
            frmRegisterOfBills.Show vbModal
            Call fillGrid
       End If
        
    End Sub


VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBudgetCentre 
   Caption         =   "Budget Centre"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11820
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   3525
      TabIndex        =   22
      Top             =   5925
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5122
      TabIndex        =   9
      Top             =   5925
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   5925
      Width           =   1575
   End
   Begin VB.ListBox lstRecords 
      Height          =   1680
      Left            =   11670
      TabIndex        =   21
      Top             =   660
      Width           =   5910
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3195
      Left            =   150
      TabIndex        =   8
      Top             =   2505
      Width           =   11550
      _cx             =   20373
      _cy             =   5636
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
      FormatString    =   $"frmBudgetCentre.frx":0000
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
   Begin VB.ComboBox cmbFinancialYear 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   2115
   End
   Begin VB.CommandButton cmdFunctionary 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10414
      TabIndex        =   6
      Top             =   1650
      Width           =   330
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10414
      TabIndex        =   7
      Top             =   1965
      Width           =   330
   End
   Begin VB.CommandButton cmdFund 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10414
      TabIndex        =   2
      Top             =   1035
      Width           =   330
   End
   Begin VB.CommandButton cmdBudgetCentre 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10414
      TabIndex        =   5
      Top             =   1335
      Width           =   330
   End
   Begin VB.TextBox txtFunctionary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4467
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1635
      Width           =   5925
   End
   Begin VB.TextBox txtFunctionaryCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2367
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1650
      Width           =   2085
   End
   Begin VB.TextBox txtFunction 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4467
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1950
      Width           =   5925
   End
   Begin VB.TextBox txtFunctionCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2367
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1950
      Width           =   2085
   End
   Begin VB.TextBox txtFund 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4467
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1020
      Width           =   5925
   End
   Begin VB.TextBox txtFundCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2367
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1020
      Width           =   2085
   End
   Begin VB.TextBox txtBudgetCentre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4467
      TabIndex        =   4
      Top             =   1335
      Width           =   5925
   End
   Begin VB.TextBox txtBudgetCentreCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2367
      TabIndex        =   3
      Top             =   1335
      Width           =   2085
   End
   Begin VB.Label lblFinancialYear 
      AutoSize        =   -1  'True
      Caption         =   "Financial Year"
      Height          =   270
      Left            =   1110
      TabIndex        =   20
      Top             =   675
      Width           =   1230
   End
   Begin VB.Label lblFunctionary 
      Caption         =   "Functionary"
      Height          =   285
      Left            =   1290
      TabIndex        =   17
      Top             =   1650
      Width           =   1020
   End
   Begin VB.Label lblFunction 
      Caption         =   "Function"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   1965
      Width           =   735
   End
   Begin VB.Label lblFund 
      Caption         =   "Fund"
      Height          =   240
      Left            =   1875
      TabIndex        =   13
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label lblBudgetCentre 
      Caption         =   "Budget Centre"
      Height          =   270
      Left            =   1065
      TabIndex        =   0
      Top             =   1320
      Width           =   1245
   End
End
Attribute VB_Name = "frmBudgetCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Sub FormInitialize()
        'cmbFinancialYear.ListIndex = -1
        txtBudgetCentre.Text = ""
        txtBudgetCentreCode.Text = ""
        txtBudgetCentreCode.Tag = ""
        'txtFundCode.Text = ""
        'txtFund.Text = ""
        txtFunctionCode.Text = ""
        txtFunction.Text = ""
        txtFunctionaryCode.Text = ""
        txtFunctionary.Text = ""
        vsGrid.Clear 1, 1
    End Sub

    Private Sub cmbFinancialYear_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdBudgetCentre_Click()
        Dim mSql As String
        If cmbFinancialYear.ListIndex <> -1 Then
            mSql = "Select vchBudgetCentre,intBudgetCentreID From faBudgetCentres Where intFinancialYearID = " & cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)
        Else
            mSql = "Select vchBudgetCentre,intBudgetCentreID From faBudgetCentres"
        End If
        mSql = mSql + " Order By vchBudgetCentre"
        Call PopulateList(lstRecords, mSql, , , , True)
        lstRecords.Tag = "2"
        lstRecords.Visible = True
        lstRecords.Left = 4470
        lstRecords.Top = 1020
        lstRecords.SetFocus
    End Sub



    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdFunction_Click()
        Dim mSql   As String
        frmSearchFunction.Show vbModal     'Modified 24/11/2009
        'mSplit = Split(gbSearchStr, " ")
        txtFunctionCode.Text = Token(gbSearchStr, " ")
        txtFunction.Text = Trim(gbSearchStr)
        txtFunction.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
       
        'mSQL = "Select (vchFunction + , , intFunctionID) From faFunctions Order By vchFunction"
        'Call PopulateList(lstRecords, mSQL, , , , True)
        'lstRecords.Tag = "3"
        'lstRecords.Visible = True
        'lstRecords.Left = 4470
        'lstRecords.Top = 1020
        'lstRecords.SetFocus
    End Sub

    Private Sub cmdFunctionary_Click()
        Dim mSql As String
        frmSearchFunctionary.Show vbModal 'Modified 24/11/2009
        txtFunctionaryCode.Text = Token(gbSearchStr, " ")
        txtFunctionary.Text = Trim(gbSearchStr)
        txtFunctionary.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        
        'mSQL = "Select vchFunctionary, intFunctionaryID From faFunctionaries Order By vchFunctionary"
        'Call PopulateList(lstRecords, mSQL, , , , True)
        'lstRecords.Tag = "4"
        'lstRecords.Visible = True
        'lstRecords.Left = 4470
        'lstRecords.Top = 1020
        'lstRecords.SetFocus
    End Sub

    Private Sub cmdFund_Click()
        Dim mSql As String
        
        mSql = "Select vchFund, intFundID From faFunds Where tnyActiveFlag = 1 Order By vchFund"
        Call PopulateList(lstRecords, mSql, , , , True)
        lstRecords.Tag = "1"
        lstRecords.Visible = True
        lstRecords.Left = 4470
        lstRecords.Top = 1020
        lstRecords.SetFocus
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdSave.Enabled = True
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDb
        Dim mArrIn              As Variant
        Dim mArrOut             As Variant
        Dim mArrInAccHead       As Variant
        Dim mRowCount           As Variant
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        
        If cmbFinancialYear.ListIndex = -1 Then
            MsgBox "Please select the Financial Year"
            cmbFinancialYear.SetFocus
            Exit Sub
        End If
        If txtBudgetCentreCode.Text = "" Then
            MsgBox "Please enter the BudgetCentreCode"
            txtBudgetCentreCode.SetFocus
            Exit Sub
        End If
        If txtBudgetCentre.Text = "" Then
            MsgBox "Please enter the BudgetCentreName"
            txtBudgetCentre.SetFocus
            Exit Sub
        End If
        If txtFund.Text = "" Then
            MsgBox "Please Select a Fund", vbCritical
            cmdFund.SetFocus
            Exit Sub
        End If
        If txtFunction.Text = "" Then
            MsgBox "Please Select a Function", vbCritical
            cmdFunction.SetFocus
            Exit Sub
        End If
        If txtFunctionary.Text = "" Then
            MsgBox "Please Select a Functionary", vbCritical
            cmdFunctionary.SetFocus
            Exit Sub
        End If
        
        mCnn.BeginTrans
        
        mArrIn = Array((IIf(txtBudgetCentreCode.Tag = "", -1, val(txtBudgetCentreCode.Tag))), _
                            txtBudgetCentreCode.Text, _
                            txtBudgetCentre.Text, _
                            txtFunctionary.Tag, _
                            txtFunction.Tag, _
                            Null, _
                            txtFundCode.Tag, _
                            cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex), _
                            gbLocalBodyID, _
                            0 _
                        )
        objDb.ExecuteSP "spSaveBudgetCentre", mArrIn, mArrOut, , mCnn, adCmdStoredProc
        
        If mArrOut(0, 0) <> "" Then
            mCnn.Execute "Delete from faBudgetAccountHeads Where intBudgetCentreId=" & mArrOut(0, 0)
        End If
        
        For mRowCount = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mRowCount, 0) <> "" Then
                If vsGrid.TextMatrix(mRowCount, 3) <> "" Then
                    If IsNumeric(vsGrid.TextMatrix(mRowCount, 3)) And val(vsGrid.TextMatrix(mRowCount, 3)) > 0 Then
                        mArrInAccHead = Array(0, mArrOut(0, 0), vsGrid.TextMatrix(mRowCount, 5), val(vsGrid.TextMatrix(mRowCount, 4)), val(vsGrid.TextMatrix(mRowCount, 3)), val(vsGrid.TextMatrix(mRowCount, 6)))
                       objDb.ExecuteSP "spSaveBudgetAccountHead", mArrInAccHead, , , mCnn
                    End If
                Else
                    MsgBox "Please enter the Amount", vbInformation
                    vsGrid.SetFocus
                    mCnn.RollbackTrans
                    Exit Sub
                End If
            End If
        Next
       cmdSave.Enabled = False
       mCnn.CommitTrans
       Call FormInitialize
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim objDb       As New clsDb
        Dim mSql        As String
        Dim mLoop       As Integer
        Dim mItem       As String
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        Me.Height = 7150
        Me.Width = 11940
        lstRecords.Visible = False
        
        For mLoop = gbFinancialYearID - 1 To gbFinancialYearID + 1
            mItem = CStr(mLoop) & "-" & CStr(mLoop + 1)
            cmbFinancialYear.AddItem (mItem)
            cmbFinancialYear.ItemData(cmbFinancialYear.NewIndex) = mLoop
        Next
        vsGrid.ColComboList(0) = "|..."
'
'        mItem = "#0; "
'        mItem = mItem & "|#" & 1 & "; Income"
'        mItem = mItem & "|#" & 2 & "; Expenditure"
'        vsGrid.ColComboList(2) = mItem
        
    End Sub
    
    Private Sub lstRecords_DblClick()
         Dim mCnn As New ADODB.Connection
         Dim Rec As New ADODB.Recordset
         Dim objDb As New clsDb
         Dim mSearchID As Long
         
         mSearchID = lstRecords.ItemData(lstRecords.ListIndex)
         objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
         
         lstRecords.Visible = False
         Select Case val(lstRecords.Tag)
                Case 1: txtFund.SetFocus
                        Rec.Open "Select vchFundCode as Fundcode from fafunds where  intfundID =" & mSearchID, mCnn
                        txtFund.Text = lstRecords.Text
                        txtFundCode.Text = IIf(IsNull(Rec!FundCode), "", Rec!FundCode)
                        txtFundCode.Tag = lstRecords.ItemData(lstRecords.ListIndex)
                        Rec.Close
                Case 2: txtBudgetCentre.SetFocus
                        Rec.Open "Select tnyStatus,vchBudgetCentreCode From faBudgetCentres Where intBudgetCentreID=" & mSearchID, mCnn
'                        If IsNull(Rec!tnyStatus) = False Then
'                            txtBudgetCentre.Tag = Rec!tnyStatus
'                        End If
                        txtBudgetCentre.Text = lstRecords.Text
                        txtBudgetCentreCode.Text = IIf(IsNull(Rec!vchBudgetCentreCode), "", Rec!vchBudgetCentreCode)
                        txtBudgetCentreCode.Tag = lstRecords.ItemData(lstRecords.ListIndex)
                        Rec.Close
                        Call txtBudgetCentreCode_LostFocus
                Case 3: txtFunction.SetFocus
                        Rec.Open "Select vchFunctionCode as Functioncode from fafunctions where  intfunctionID =" & mSearchID, mCnn
                        txtFunction.Text = IIf(IsNull(lstRecords.Text), "", lstRecords.Text)
                        txtFunctionCode.Text = IIf(IsNull(Rec!FunctionCode), "", Rec!FunctionCode)
                        txtFunction.Tag = lstRecords.ItemData(lstRecords.ListIndex)
                        Rec.Close
                Case 4: txtFunctionary.SetFocus
                        Rec.Open "Select vchFunctionaryCode as Functionarycode from fafunctionaries where  intfunctionaryID =" & mSearchID, mCnn
                        txtFunctionary.Text = lstRecords.Text
                        txtFunctionaryCode.Text = IIf(IsNull(Rec!FunctionaryCode), "", Rec!FunctionaryCode)
                        txtFunctionary.Tag = lstRecords.ItemData(lstRecords.ListIndex)
                        Rec.Close
         End Select
    
    End Sub

    Private Sub lstRecords_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            lstRecords_DblClick
        End If
    End Sub

    Private Sub lstRecords_LostFocus()
        lstRecords.Visible = False
    End Sub
    
    Private Sub txtBudgetCentreCode_LostFocus()
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim mSql                As String
        Dim objDb               As New clsDb
        Dim mSQLAccountHeads    As String
        Dim mRowCount           As Long
        Dim mFinancialYearID    As Variant
        
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If txtBudgetCentreCode.Text <> "" Then
            mSql = "Select * From faBudgetCentres"
            mSql = mSql + " Left Join faBudgetAccountHeads On faBudgetCentres.intBudgetCentreID = faBudgetAccountHeads.intBudgetCentreID"
            mSql = mSql + " Left Join faAccountHeads On faBudgetAccountHeads.intAccountHeadID = faAccountHeads.intAccountHeadID"
            mSql = mSql + " Left Join faFunds On faBudgetCentres.intFundID = faFunds.intFundID"
            mSql = mSql + " Left Join faFunctions On faBudgetCentres.intFunctionID = faFunctions.intFunctionID"
            mSql = mSql + " Left Join faFunctionaries On faBudgetCentres.intFunctionaryID = faFunctionaries.intFunctionaryID"
            mSql = mSql + " Where vchBudgetCentreCode = '" & txtBudgetCentreCode.Text & "'"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFinancialYearID = IIf(IsNull(Rec.Fields(6)), "", Rec.Fields(6)) 'intFinancialYearID
                If mFinancialYearID <> "" Then
                    cmbFinancialYear.Text = CStr(mFinancialYearID) & "-" & CStr(mFinancialYearID + 1)
                End If
                txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
                txtFundCode.Text = IIf(IsNull(Rec!vchFundCode), "", Rec!vchFundCode)
                txtFundCode.Tag = IIf(IsNull(Rec!intFundID), "", Rec!intFundID)
                txtBudgetCentre.Text = IIf(IsNull(Rec!vchBudgetCentre), "", Rec!vchBudgetCentre)
                txtBudgetCentreCode.Text = IIf(IsNull(Rec!vchBudgetCentreCode), "", Rec!vchBudgetCentreCode)
                txtBudgetCentreCode.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intBudgetCentreID
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                txtFunctionCode.Text = IIf(IsNull(Rec!vchFunctionCode), "", Rec!vchFunctionCode)
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionaryCode.Text = IIf(IsNull(Rec!vchFunctionaryCode), "", Rec!vchFunctionaryCode)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                   
                mRowCount = 1
                While Not Rec.EOF
                    vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltEstimatedAmount), "", Rec!fltEstimatedAmount)
                    vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltBudgetActualofPrevYr), "", Rec!fltBudgetActualofPrevYr)
                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!tnyType), "", Rec!tnyType)
                    If vsGrid.TextMatrix(mRowCount, 6) <> "" Then
                        If vsGrid.TextMatrix(mRowCount, 6) = 1 Then
                            vsGrid.Cell(flexcpText, mRowCount, 2) = "Income"
                        End If
                        If vsGrid.TextMatrix(mRowCount, 6) = 2 Then
                            vsGrid.Cell(flexcpText, mRowCount, 2) = "Expenditure"
                        End If
                    End If
                    mRowCount = mRowCount + 1
                    vsGrid.Rows = vsGrid.Rows + 1
                    Rec.MoveNext
                Wend
            End If
            Rec.Close
        End If
    End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If vsGrid.TextMatrix(vsGrid.Row, 3) <> "" Then
            If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) = False Then
                MsgBox "Must Enter Numeric values", vbCritical
                vsGrid.TextMatrix(vsGrid.Row, 3) = 0
                vsGrid.TextMatrix(vsGrid.Row, 3) = Format(vsGrid.TextMatrix(vsGrid.Row, 3), "0.00")
            Else
                vsGrid.TextMatrix(vsGrid.Row, 3) = Format(vsGrid.TextMatrix(vsGrid.Row, 3), "0.00")
            End If
        End If
        If vsGrid.TextMatrix(vsGrid.Row, 4) <> "" Then
            If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 4)) = False Then
                MsgBox "Must Enter Numeric values", vbCritical
                vsGrid.TextMatrix(vsGrid.Row, 4) = 0
                vsGrid.TextMatrix(vsGrid.Row, 4) = Format(vsGrid.TextMatrix(vsGrid.Row, 4), "0.00")
            Else
                vsGrid.TextMatrix(vsGrid.Row, 4) = Format(vsGrid.TextMatrix(vsGrid.Row, 4), "0.00")
            End If
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGrid.Row > 1 Then
            If vsGrid.TextMatrix(vsGrid.Row - 1, 0) = "" Or _
               (val(vsGrid.TextMatrix(vsGrid.Row - 1, 3)) <= 0) Then
               'Val(vsGrid.TextMatrix(vsGrid.Row - 1, 5)) <= 0)
               Cancel = True
               Exit Sub
            End If
        End If
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim mSql As String
       
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
        
        vsGrid.Rows = vsGrid.Rows + 1
        vsGrid.TextMatrix(vsGrid.Row, 0) = Token(gbSearchStr, " ")
        vsGrid.TextMatrix(vsGrid.Row, 1) = gbSearchStr
        vsGrid.TextMatrix(vsGrid.Row, 5) = gbSearchID
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        If vsGrid.Col = 0 Then
            'mItem = "#0; "
            If Left(vsGrid.TextMatrix(vsGrid.Row, 0), 1) = 1 Then
                vsGrid.TextMatrix(vsGrid.Row, 2) = "Income"
                vsGrid.TextMatrix(vsGrid.Row, 6) = 1
                'mItem = mItem & "|#" & 1 & "; Income"
            ElseIf Left(vsGrid.TextMatrix(vsGrid.Row, 0), 1) = 2 Then
                vsGrid.TextMatrix(vsGrid.Row, 2) = "Expenditure"
                vsGrid.TextMatrix(vsGrid.Row, 6) = 2
                'mItem = mItem & "|#" & 2 & "; Expenditure"
            End If
            'vsGrid.ColComboList(2) = mItem
        End If
    End Sub

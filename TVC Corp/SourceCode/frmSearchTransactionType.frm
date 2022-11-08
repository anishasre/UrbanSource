VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchTransactionType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Transaction Types"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkListAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "List All"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8370
      TabIndex        =   9
      Top             =   30
      Width           =   855
   End
   Begin VB.ListBox lstTransactionTypes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF7&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4590
      Left            =   90
      TabIndex        =   8
      Top             =   1095
      Width           =   8925
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -1530
      Top             =   6480
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fmeMode 
      Caption         =   "Mode of Transaction Types"
      Enabled         =   0   'False
      Height          =   615
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   8925
      Begin VB.OptionButton optPayment 
         Caption         =   "Payment Transaction Types"
         Height          =   270
         Left            =   5700
         TabIndex        =   7
         Top             =   210
         Width           =   2655
      End
      Begin VB.OptionButton optReceipt 
         Caption         =   "Receipt Transaction Types"
         Height          =   270
         Left            =   2490
         TabIndex        =   6
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   8610
      TabIndex        =   4
      Top             =   5925
      Width           =   375
   End
   Begin VB.TextBox txtSearchKey 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   5940
      Width           =   7035
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4815
      Left            =   8970
      TabIndex        =   3
      Top             =   6330
      Visible         =   0   'False
      Width           =   8985
      _cx             =   15849
      _cy             =   8493
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
      Rows            =   15
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchTransactionType.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Press F8 to search Transaction Type by Account Heads"
      Height          =   270
      Left            =   3915
      TabIndex        =   10
      Top             =   6315
      Width           =   4665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   5955
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Transaction Types"
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
      TabIndex        =   2
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "frmSearchTransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private intModeOfTransaction As Integer '   1= Receipt ; 2 = payment    '
    Private mStrQry As String        ' For call the form from Demand Interface
    Private mSQLQty As String       ' Can sent complete SQL Query
 
    Private Sub chkListAll_Click()
        If chkListAll.value = 1 Then
            If intModeOfTransaction = 1 Then 'R e c e i p t   M o d e
                FillAllTransaction ("")
            End If
        End If
    End Sub
   
    Private Sub cmdsearch_Click()
        If CheckModeOfTransaction = False Then
            Call FillAllTransaction(Trim(txtSearchKey.Text))
        End If
    End Sub
    
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        'intModeOfTransaction = 2
        If CheckModeOfTransaction = False Then
            Call FillReciptTransaction(Trim(txtSearchKey.Text))
        End If
    End Sub
    '=========================================================================='
    '                           Filling Transaction Types                      '
    '=========================================================================='
    Private Sub FillAllTransaction(ByVal strData As String)
        On Error GoTo err:
            Dim mSql As String
            If IsNull(strData) Then
                mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Order By vchTransactionType"
            Else
                mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where vchTransactionType Like '%" & strData & "%' Order By vchTransactionType"
            End If
            PopulateList lstTransactionTypes, mSql, , , , True
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub FillReciptTransaction(ByVal strData As String)
         On Error GoTo err:
            chkListAll.Visible = True
            Dim mSql As String
            If IsNull(strData) Then
                mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where intGroupID = 10 And isNull(tnyHidden,0)=0 Order By vchTransactionType"
            Else
                mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where vchTransactionType Like '%" & strData & "%' and intGroupID = 10 And isNull(tnyHidden,0)=0  Order By vchTransactionType"
            End If
            If mStrQry <> "" Then
                mSql = mStrQry
            End If
            PopulateList lstTransactionTypes, mSql, , , , True
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub FillQuery()
        On Error GoTo err:
        If mSQLQty <> "" Then
            PopulateList lstTransactionTypes, mSQLQty, , , , True
            chkListAll.Visible = False
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub FillPaymentTransaction(ByVal strData As String)
        On Error GoTo err:
        chkListAll.Visible = True
        Dim mSql As String
        If IsNull(strData) Then
            PopulateList lstTransactionTypes, "Select vchTransactionType, intTransactionTypeID from faTransactionType Where intGroupID = 20 And isNull(tnyHidden,0)=0 Order By vchTransactionType", , , , True
            'mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where intGroupID = 20 And isNull(tnyHidden,0)=0 Order By vchTransactionType"
        Else
            mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where vchTransactionType Like '%" & strData & "%' and intGroupID = 20 And isNull(tnyHidden,0)=0 Order By vchTransactionType"
            PopulateList lstTransactionTypes, mSql, , , , True
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub FillAccountHeadWiseTransaction(ByVal strAccountHeadCode As String)
        'On Error GoTo Err:
        Dim mSql As String
        If Not IsNull(strAccountHeadCode) Then
            mSql = "Select vchTransactionType, faTransactionType.intTransactionTypeID From faTransactionType INNER JOIN "
            mSql = mSql + " faTransactionTypeChild ON faTransactionTypeChild.intTransactionTypeID = faTransactionType.intTransactionTypeID "
            mSql = mSql + " WHERE faTransactionType.intGroupID = 10 And isNull(tnyHidden,0)= 0 AND vchAccountHeadCode = '" & strAccountHeadCode & "' Order By vchTransactionType"
            PopulateList lstTransactionTypes, mSql, , , , True
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
        ElseIf KeyCode = 13 Then
            If lstTransactionTypes.ListIndex > -1 Then
                gbSearchStr = lstTransactionTypes.Text
                gbSearchID = lstTransactionTypes.ItemData(lstTransactionTypes.ListIndex)
                Unload Me
            End If
        ElseIf KeyCode = vbKeyF8 Then
            '---------------------------------------------------------------'
            ' NOTE: Funciton Key F8 Loads the AccountHead Search Screen     '
            '     : Selection of Account head filters the Transaction Type  '
            '---------------------------------------------------------------'
            frmSearchAccountHeads.Show vbModal
            If gbSearchID > 0 Then
                
                Dim mCode As String
                mCode = ""
                If Len(gbSearchStr) > 9 Then
                    mCode = Left(gbSearchStr, 9)
                End If
                chkListAll.value = 0
                FillAccountHeadWiseTransaction (mCode)
                
                gbSearchID = -1
                gbSearchCode = ""
                gbSearchStr = ""
            End If
        End If
    End Sub
    Private Sub lstTransactionTypes_DblClick()
        Call Form_KeyDown(13, 0)
    End Sub
    
    Private Sub vsGrid_Click()
        On Error GoTo err:
        vsGrid.Cell(flexcpBackColor, 1, 0, vsGrid.Rows - 1, 1) = vbWhite
        If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
            vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 1) = &HC0C0FF
        Else
            vsGrid.Cell(flexcpBackColor, vsGrid.Row, 0, vsGrid.Row, 1) = vbWhite
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo err:
        If KeyCode = vbKeyEscape Then
            Unload Me
        ElseIf KeyCode = 13 Then
            If vsGrid.TextMatrix(vsGrid.Row, 1) <> "" Then
                gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 1)
                gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 0)
                Unload Me
            End If
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub vsGrid_DblClick()
         Call vsGrid_KeyDown(13, 0)
    End Sub
    
    Private Sub txtSearchKey_Change()
        On Error GoTo err:
        Dim mIndex As Long
        Dim mStr As String
        
        mStr = txtSearchKey.Text
        mIndex = SendMyMessage(lstTransactionTypes.hwnd, LB_FINDSTRING, -1, ByVal mStr)
        If mIndex > -1 Then
            lstTransactionTypes.ListIndex = mIndex
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
        On Error GoTo err:
        '38 Up Arrow
        If (KeyCode = 38 Or KeyCode = 40) Then
            Dim objAcc As New clsAccounts
            If KeyCode = 38 And lstTransactionTypes.ListIndex > 0 Then
                lstTransactionTypes.ListIndex = lstTransactionTypes.ListIndex - 1
            End If
            '40 = Down Arrow
            If KeyCode = 40 And lstTransactionTypes.ListIndex < (lstTransactionTypes.ListCount - 1) Then
                lstTransactionTypes.ListIndex = lstTransactionTypes.ListIndex + 1
                'Debug.Print lstTransactionTypes.ListCount - 1, lstTransactionTypes.ListIndex
            End If
            If lstTransactionTypes.ListIndex > -1 Then
                objAcc.SetAccounts (lstTransactionTypes.ItemData(lstTransactionTypes.ListIndex))
                txtSearchKey.Text = objAcc.AccountCode
            End If
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    Private Function CheckModeOfTransaction() As Boolean
        On Error GoTo err:
        optReceipt.value = False
        optPayment.value = False
        If intModeOfTransaction = 1 Then 'Receipt Mode
            optReceipt.value = True
            optPayment.value = False
            Call FillReciptTransaction(Trim(txtSearchKey.Text))
            CheckModeOfTransaction = True
            Exit Function
        ElseIf intModeOfTransaction = 2 Then
            optPayment = True
            optReceipt = False
            If mSQLQty <> "" Then
                Call FillQuery
            Else '
                Call FillPaymentTransaction(Trim(txtSearchKey.Text))
            End If
            CheckModeOfTransaction = True
            Exit Function
        ElseIf intModeOfTransaction = 3 Then
            optPayment = True
            optReceipt = False
            If mSQLQty <> "" Then
                Call FillQuery
            Else '
                Call FillPaymentUnAuthorizedTransaction(Trim(txtSearchKey.Text))
            End If
            CheckModeOfTransaction = True
            Exit Function
        End If
        
        CheckModeOfTransaction = False
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Sub FillPaymentUnAuthorizedTransaction(ByVal strData As String) 'FOR UNAUTHORIZED DRAWAL
        On Error GoTo err:
        chkListAll.Visible = True
        Dim mSql As String
        If IsNull(strData) Then
            PopulateList lstTransactionTypes, "Select vchTransactionType, intTransactionTypeID from faTransactionType Where  intTransactionTypeID IN (1201,1391) Order By vchTransactionType", , , , True
          
        Else
            mSql = "Select vchTransactionType, intTransactionTypeID from faTransactionType Where vchTransactionType Like '%" & strData & "%' and  intTransactionTypeID IN (1201,1391) Order By vchTransactionType"
            PopulateList lstTransactionTypes, mSql, , , , True
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Public Property Let ModeOfTransaction(mData As Integer)
        intModeOfTransaction = mData
    End Property
    Public Property Let StrQuery(mData As String)
        mStrQry = mData
    End Property
    Public Property Let SQLQry(mData As String)
        mSQLQty = mData
    End Property
    

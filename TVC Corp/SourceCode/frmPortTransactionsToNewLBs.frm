VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPortTransactionsToNewLBs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PORT TRANSACTIONS TO NEW LBs"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   7200
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "TRANSFER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   6720
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar PgrBar 
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "VERIFY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "EXTRACT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   13
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdMonthDown 
      Caption         =   "<<"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   525
   End
   Begin VB.CommandButton cmdMonthUp 
      Caption         =   ">>"
      Height          =   345
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   525
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox cmbDB 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   292
         Width           =   2415
      End
      Begin VB.ComboBox cmbServer 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   292
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "DATABASE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "SERVER"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   11235
      TabIndex        =   3
      Top             =   0
      Width           =   11295
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.ComboBox cmbLBName 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "LOCAL BODY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4215
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   11055
      _cx             =   19500
      _cy             =   7435
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
      Rows            =   20
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPortTransactionsToNewLBs.frx":0000
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
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "OCTOBER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   11
      Top             =   1365
      Width           =   795
   End
End
Attribute VB_Name = "frmPortTransactionsToNewLBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim dtMonthLastDate  As Variant
    Dim mTimer  As Integer
    Dim mControlVariableForProgressBar As Integer
    Private Sub Command1_Click()
        Timer1.Enabled = True
    End Sub
    Private Sub cmdExtract_Click()
        If SaveValidation Then
            Call MonthLastDay
            Call ExtractToTmp
'''            Call FillGridExtract
        End If
    End Sub
    
    Private Sub CheckProgressBar()
        PgrBar.Max = 10000 + 1
        While PgrBar.value < PgrBar.Max
            PgrBar.value = PgrBar.value + 1
        Wend
    End Sub
    Private Sub ExtractToTmp()
        Dim mCnnServer  As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        Dim mRowCnt As Integer
    
        Call ServerConnection(mCnnServer)
        mCnnServer.CommandTimeout = 1000000000
        objdb.ExecuteSP "spExtractTotmpPO", , , , mCnnServer 'EXTRACT PO(PO,Child,Address)
        objdb.ExecuteSP "spExtractTotmpTransactions", , , , mCnnServer 'EXTRACT transactions(trans,Child)
        objdb.ExecuteSP "spExtractTotmpVouchers", , , , mCnnServer 'EXTRACT Vouchers(voucher,Child,address,sub)
        
        If PgrBar.value < PgrBar.Max - 2 Then
            PgrBar.value = PgrBar.Max - 1
        End If
        
        mCnnServer.Close
        
    End Sub
    
    Private Function SaveValidation() As Boolean
            If cmbLBName.ListIndex = -1 Then
                MsgBox "Please Select the LocalBody", vbInformation
                SaveValidation = False
                Exit Function
            End If
            If cmbServer.ListIndex = -1 Then
                MsgBox "Please Select the Server Name", vbInformation
                SaveValidation = False
                Exit Function
            End If
            If cmbDB.ListIndex = -1 Then
                MsgBox "Please Select the DataBase", vbInformation
                SaveValidation = False
                Exit Function
            End If
            SaveValidation = True
    End Function
    
    Private Sub ServerConnection(mCnnServer As ADODB.Connection)
        'Dim mCnnServer  As New ADODB.Connection
        mCnnServer.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=" + Trim(cmbServer.Text) + ";"
        mCnnServer.Open
    End Sub
    
    Private Sub DatabaseConnection(mCnnClient As ADODB.Connection)
        mCnnClient.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnFa;uid=FAUser;pwd=FAUser;database=" + Trim(cmbDB.Text) + ";"
        mCnnClient.Open
    End Sub
    Private Sub FillGridExtract()
'''    Dim mCnnServer  As New ADODB.Connection
'''    Dim objDB As New clsDB
'''    Dim Rec   As New ADODB.Recordset
'''    Dim mSQL  As String
'''    Dim mRowCnt As Integer
'''
'''
'''    On Error GoTo err
'''    'objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
'''    Call ServerConnection(mCnnServer)
'''
'''    mSQL = " SELECT tmpTransactionChild.intAccountHeadID,vchAccountHeadCode,vchAccountHead,isnull(sum(fltAmount*((tmpTransactionChild.tinDebitOrCreditFlag*2)-1)),0) fltAmount "
'''    mSQL = mSQL + " From tmpTransactionChild"
'''    mSQL = mSQL + " INNER  JOIN tmpTransactions ON tmpTransactions.intTransactionID=tmpTransactionChild.intTransactionID"
'''    mSQL = mSQL + " INNER  JOIN  faAccountHeads ON tmpTransactionChild.intAccountheadID=faAccountHeads.intAccountheadID"
'''    'mSQL = mSQL + " --INNER JOIN faBAnks ON faBanks.intAccountHeadID=faAccountHeads.intAccountheadID"
'''    mSQL = mSQL + " WHERE   dtTransactionDate < = ' & dtMonthLastDate &  ' AND ( vchAccountHeadCode LIKE '450250%' OR vchAccountHeadCode LIKE '450650%')"
'''    mSQL = mSQL + " AND faAccountHeads.intGroupID=2 AND (tmpTransactions.tnyStatus <> 4 OR tmpTransactions.tnyStatus IS NULL)"
'''    mSQL = mSQL + " Group by vchAccountHead,vchAccountHeadCode,tmpTransactionChild.intAccountHeadID"
'''
'''    Rec.CursorLocation = adUseClient
'''
'''    Rec.Open mSQL, mCnnServer, adOpenDynamic, adLockOptimistic, adLockReadOnly
'''    mRowCnt = 1
'''    vsGrid.Clear 1, 1
'''    vsGrid.Rows = 1
'''    While Not (Rec.EOF Or Rec.BOF)
'''        vsGrid.Rows = vsGrid.Rows + 1
'''        vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'''        vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
'''        vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'''        vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'''        Rec.MoveNext
'''        mRowCnt = mRowCnt + 1
'''    Wend
'''
'''    Rec.Close
'''    mCnnServer.Close
'''    Exit Sub
'''err:
'''    MsgBox err.Description
End Sub

    Private Sub cmdMonthDown_Click()
        If lblMonth.Caption = "FEBRUARY" Then
            lblMonth.Caption = "JANUARY"
        ElseIf lblMonth.Caption = "JANUARY" Then
            lblMonth.Caption = "DECEMBER"
        ElseIf lblMonth.Caption = "DECEMBER" Then
            lblMonth.Caption = "NOVEMBER"
        ElseIf lblMonth.Caption = "NOVEMBER" Then
            lblMonth.Caption = "OCTOBER"
        End If
        Call MonthLastDay
        Call FillGridExtract
    End Sub

    Private Sub cmdMonthUp_Click()
        If lblMonth.Caption = "OCTOBER" Then
            lblMonth.Caption = "NOVEMBER"
        ElseIf lblMonth.Caption = "NOVEMBER" Then
            lblMonth.Caption = "DECEMBER"
        ElseIf lblMonth.Caption = "DECEMBER" Then
            lblMonth.Caption = "JANUARY"
        ElseIf lblMonth.Caption = "JANUARY" Then
            lblMonth.Caption = "FEBRUARY"
        End If
        Call MonthLastDay
        Call FillGridExtract
    End Sub


    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 1200
    End Sub

    Private Sub Form_Load()
        lblMonth.Caption = "OCTOBER"
        Call FillCombo
        Call MonthLastDay
        'Call FillGridExtract
    End Sub
    
    Private Sub FillCombo()
        Dim mSql As String
                    
        mSql = "SELECT vchLBName,intID FROM tmpMergedLBs"
        PopulateList cmbLBName, mSql, , , True, True, enuSourceString.Saankhya
        
        mSql = "SELECT name,dbid FROM master..sysdatabases"
        PopulateList cmbServer, mSql, , , True, True, enuSourceString.Saankhya
        
        PopulateList cmbDB, mSql, , , True, True, enuSourceString.Saankhya
    End Sub
                
    Private Function MonthLastDay()
        If lblMonth.Caption = "OCTOBER" Then
            dtMonthLastDate = "31/OCT/2015"
        ElseIf lblMonth.Caption = "NOVEMBER" Then
             dtMonthLastDate = "30/NOV/2015"
        ElseIf lblMonth.Caption = "DECEMBER" Then
            dtMonthLastDate = "31/DEC/2015"
        ElseIf lblMonth.Caption = "JANUARY" Then
            dtMonthLastDate = "31/JAN/2016"
        ElseIf lblMonth.Caption = "FEBRUARY" Then
            dtMonthLastDate = "29/FEB/2016"
        End If
    End Function

    Private Sub Timer1_Timer()
''''        If mStartFlag = False And mControlVariableForProgressBar > 5 Then
''''            mStartFlag = True
''''            Call ExtractToTmp
''''        End If
'''
'''        If mTimer = 0 Then
'''            mTimer = 1
'''            'Exit Sub
'''        ElseIf mTimer = 1 Then
'''            mTimer = 0
'''            'Exit Sub
'''        End If
'''
'''        If mControlVariableForProgressBar < 20 Then
'''            mControlVariableForProgressBar = mControlVariableForProgressBar + 1
'''        Else
'''            mControlVariableForProgressBar = 0
'''            If PgrBar.value < PgrBar.Max Then
'''                PgrBar.value = PgrBar.value + 1
'''            Else
'''                PgrBar.value = 1
'''            End If
'''        End If
'''        Call FillGridExtract
    End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSourceWiseOpening 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSourceWiseOpening.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOpeningBalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   9450
      TabIndex        =   11
      Top             =   1530
      Width           =   1905
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   345
      Left            =   11370
      TabIndex        =   10
      Top             =   1140
      Width           =   315
   End
   Begin VB.TextBox txtAccountHead 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   6780
      TabIndex        =   9
      Top             =   1170
      Width           =   4575
   End
   Begin VB.TextBox txtAccountHeadCode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   6780
      TabIndex        =   8
      Top             =   840
      Width           =   1275
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3510
      Top             =   6390
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   7890
      TabIndex        =   6
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Clos&E"
      Height          =   405
      Left            =   10470
      TabIndex        =   5
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   9180
      TabIndex        =   4
      Top             =   5760
      Width           =   1245
   End
   Begin VB.TextBox txtNetAmount 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5130
      TabIndex        =   3
      Top             =   6330
      Width           =   1620
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6405
      Left            =   60
      TabIndex        =   1
      Top             =   -60
      Width           =   6735
      _cx             =   11880
      _cy             =   11298
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12632256
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
      Rows            =   24
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSourceWiseOpening.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      Editable        =   1
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      Height          =   270
      Left            =   8010
      TabIndex        =   12
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head"
      Height          =   270
      Left            =   6780
      TabIndex        =   7
      Top             =   570
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   4080
      TabIndex        =   2
      Top             =   6360
      Width           =   990
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Source Wise Opening Entry Form"
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
      Left            =   6810
      TabIndex        =   0
      Top             =   90
      Width           =   5055
   End
End
Attribute VB_Name = "frmSourceWiseOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdSave_Click()
        If SaveValidation And SaveSourceWiseOpening Then
            MsgBox "Source Wise Opening Entered Sucessfully", vbInformation
            Call FormInitialize
        End If
    End Sub

    Private Sub FormInitialize()
        On Error GoTo Err:
            Dim mLoop As Integer
            For mLoop = 1 To vsGrid.Rows - 1
                vsGrid.TextMatrix(mLoop, 2) = ""
            Next
            txtAccountHead.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHeadCode.Tag = ""
            txtNetAmount.Text = ""
            txtOpeningBalance = ""
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmdSearch_Click()
        On Error GoTo Err:
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = 2"
            frmSearchAccountHeads.Show 1
            If gbSearchID <> -1 Then
                txtAccountHeadCode.Text = Left(gbSearchStr, 9)
                txtAccountHeadCode.Tag = gbSearchID
                txtAccountHead.Text = gbSearchStr
                Call GetOBForLedger
                vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = Val(txtOpeningBalance.Text)
            End If
            gbSearchStr = ""
            gbSearchID = -1
            
            Call FillGrid
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function GetOBForLedger() As Boolean
        On Error GoTo Err:
            Dim objDb As New clsDB
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            
            If objDb.SetConnection(mCnn) Then
                mSql = "Select fltOpeningBalance as OB From faAccountHeads Where intAccountHeadID = " & Val(txtAccountHeadCode.Tag)
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtOpeningBalance.Text = IIf(IsNull(Rec!OB), 0, Val(Rec!OB))
                End If
                GetOBForLedger = True
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function FillGrid() As Boolean
        On Error GoTo Err:
            Dim objDb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim mLoop As Integer
            
            If objDb.SetConnection(mCnn) Then
                mSql = "Select intSourceFundID, vchSourceFundName,'','-1' From suSourceOfFund"
                Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                If Not (Rec.EOF Or Rec.BOF) Then
                    vsGrid.Rows = Rec.RecordCount + 1
                    vsGrid.Col = 0
                    vsGrid.Row = 1
                    vsGrid.ColHidden(0) = True
                    'vsGrid.ColHidden(3) = True
                    vsGrid.ColSel = 3
                    vsGrid.RowSel = vsGrid.Rows - 1
                    mSql = Rec.GetString(, , vbTab, Chr(13))
                    vsGrid.Clip = mSql
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = Val(txtOpeningBalance.Text)
                End If
                If Rec.State = 1 Then Rec.Close
                mSql = "Select intSourceWiseOpeningID,intSourceFundID,fltOpeningAmt From faSourceWiseOpening Where intAccountHeadID = " & Val(txtAccountHeadCode.Tag)
                Rec.Open mSql, mCnn
                While Not (Rec.EOF Or Rec.BOF)
                    For mLoop = 1 To vsGrid.Rows - 2
                        If vsGrid.TextMatrix(mLoop, 0) = Rec!intSourceFundID Then
                            vsGrid.TextMatrix(mLoop, 2) = Rec!fltOpeningAmt
                            vsGrid.TextMatrix(mLoop, 3) = Rec!intSourceWiseOpeningID
                        End If
                    Next
                    Rec.MoveNext
                Wend
                
                FillGrid = True
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function SaveValidation() As Boolean
        On Error GoTo Err:
            Dim blnFlag As Boolean
            Dim mRowCnt As Integer
            
            For mRowCnt = 1 To vsGrid.Rows - 2
                If vsGrid.TextMatrix(mRowCnt, 2) <> "" Then
                    blnFlag = True
                End If
            Next
            If blnFlag = False Then
                MsgBox "Please Enter Atleast One Opening Across the Source in the Grid", vbInformation
                vsGrid.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If Val(txtAccountHeadCode.Tag) = 0 Then
                MsgBox "Please Select the Account Head Across the Source to be Entered", vbInformation
                cmdSearch.SetFocus
                SaveValidation = False
                Exit Function
            End If
            SaveValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Function SaveSourceWiseOpening() As Boolean
        On Error GoTo Err:
            Dim objDb As New clsDB
            Dim mSql As String
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim mRowCnt As Integer
            
            If objDb.SetConnection(mCnn) Then
                For mRowCnt = 1 To vsGrid.Rows - 2
                    If vsGrid.TextMatrix(mRowCnt, 2) <> "" Then
                        aryIn = Array(vsGrid.TextMatrix(mRowCnt, 3), _
                                        Val(txtAccountHeadCode.Tag), _
                                        Val(txtAccountHeadCode.Text), _
                                        vsGrid.TextMatrix(mRowCnt, 0), _
                                        vsGrid.TextMatrix(mRowCnt, 2), _
                                        gbFinancialYearID, _
                                        gbLocalBodyID)
                        objDb.ExecuteSP "spSaveSourceWiceOpening", aryIn, , , mCnn, adCmdStoredProc
                        mSql = " Update suSourceOfFund Set fltOpeningAmt =  "
                        mSql = mSql + " ( Select Sum(fltOpeningAmt) TotalOpening from faSourceWiseOpening Where intSourceFundID = " & vsGrid.TextMatrix(mRowCnt, 0) & ") "
                        mSql = mSql + " Where intSourceFundID = " & vsGrid.TextMatrix(mRowCnt, 0)
                        mCnn.Execute mSql
                    End If
                Next
                SaveSourceWiseOpening = True
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function

    Private Sub Form_Activate()
        WindowsXPC1.InitIDESubClassing
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Call FillGrid
    End Sub
    
    Private Function GetTotalAmount() As Double
        On Error GoTo Err:
            Dim mLoop As Integer
            Dim mAmt As Double
            mAmt = 0
            For mLoop = 1 To vsGrid.Rows - 1
                mAmt = Val(vsGrid.TextMatrix(mLoop, 2)) + mAmt
            Next
            GetTotalAmount = mAmt
        Exit Function
Err:
        MsgBox (Error$)
    End Function
   
    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim mRowCnt As Integer
        Dim mTotal As Double
        mTotal = 0
        For mRowCnt = 1 To vsGrid.Rows - 2
            mTotal = mTotal + Val(vsGrid.TextMatrix(mRowCnt, 2))
        Next
        vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = Val(txtOpeningBalance.Text) - mTotal
        
        txtNetAmount.Text = GetTotalAmount
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Col = 1 Then
            Cancel = True
        End If
        If Row = vsGrid.Rows - 1 Then
            Cancel = True
        End If
        If txtOpeningBalance.Text = "" Then
            cmdSearch.SetFocus
            Cancel = True
        End If
    End Sub

    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            vsGrid.Row = vsGrid.Row + 1
        End If
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Col = 2 Then
            If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
                    KeyAscii = 0
            End If
        End If
    End Sub

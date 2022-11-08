VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmTransactionTypes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransactionTypes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   9090
      TabIndex        =   14
      Top             =   6120
      Width           =   1245
   End
   Begin VB.ListBox lstTransactionType 
      Height          =   2220
      Left            =   2250
      TabIndex        =   2
      Top             =   6450
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Clos&E"
      Height          =   405
      Left            =   10380
      TabIndex        =   15
      Top             =   6120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   7800
      TabIndex        =   13
      Top             =   6120
      Width           =   1245
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   -3480
      Top             =   6390
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   0
      TabIndex        =   27
      Top             =   3360
      Width           =   11625
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove From Grid"
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   1980
         Width           =   2085
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add To Grid"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   1560
         Width           =   2085
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         TabIndex        =   16
         Top             =   750
         Width           =   1575
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2295
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   7965
         _cx             =   14049
         _cy             =   4048
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
         Rows            =   7
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTransactionTypes.frx":1CCA
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   270
         Left            =   210
         TabIndex        =   38
         Top             =   1110
         Width           =   210
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   270
         Left            =   180
         TabIndex        =   37
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "    Account Head Code"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   210
         TabIndex        =   36
         Top             =   330
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      TabIndex        =   20
      Top             =   300
      Width           =   11625
      Begin VB.TextBox txtSourceOfFund 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   34
         Top             =   1410
         Width           =   4815
      End
      Begin VB.CommandButton cmdSearchSource 
         Caption         =   "..."
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
         Left            =   6450
         TabIndex        =   33
         Top             =   1410
         Width           =   315
      End
      Begin VB.ComboBox cmbType 
         Height          =   390
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1770
         Width           =   2955
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   390
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1350
         Width           =   2955
      End
      Begin VB.CommandButton cmdSearchBank 
         Caption         =   "..."
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
         Left            =   6450
         TabIndex        =   6
         Top             =   1800
         Width           =   315
      End
      Begin VB.TextBox txtBankHeadCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1800
         Width           =   4815
      End
      Begin VB.ComboBox cmbSection 
         Height          =   390
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2955
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   390
         ItemData        =   "frmTransactionTypes.frx":1DE7
         Left            =   1590
         List            =   "frmTransactionTypes.frx":1DF6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   2955
      End
      Begin VB.ComboBox cmbFund 
         Height          =   390
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   2955
      End
      Begin VB.ComboBox cmbFunction 
         Height          =   390
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   2955
      End
      Begin VB.ComboBox cmbFunctionary 
         Height          =   390
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   570
         Width           =   2955
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
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
         Left            =   6450
         TabIndex        =   1
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   0
         Top             =   210
         Width           =   4815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source of Fund"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   270
         Left            =   7020
         TabIndex        =   32
         Top             =   1860
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   270
         Left            =   7020
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Head Code"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   270
         Left            =   7020
         TabIndex        =   24
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         Height          =   270
         Left            =   7020
         TabIndex        =   23
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         Height          =   270
         Left            =   7020
         TabIndex        =   22
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Transaction Type Definition Form"
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
      Left            =   30
      TabIndex        =   29
      Top             =   30
      Width           =   11625
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Accont Head Details"
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
      Height          =   300
      Left            =   30
      TabIndex        =   28
      Top             =   3030
      Width           =   11595
   End
End
Attribute VB_Name = "frmTransactionTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Sub Clear_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdAdd_Click()
        On Error GoTo Err:
            Dim mRowCnt As Integer
            Dim mSql As String
            Dim Rec As New ADODB.RecordSet
            Dim objDb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim mCnt As Integer
            
            If objDb.SetConnection(mCnn) Then
                mSql = "Select * from faAccountHeads Where vchAccountHeadCode between '" & Trim(txtFrom.Text) & "' and '" & Trim(txtTo.Text) & "'"
                Rec.Open mSql, mCnn
                mCnt = 1
                For mRowCnt = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mRowCnt, 1) = "" Then
                        Exit For
                    End If
                    mCnt = mCnt + 1
                Next
                
                mRowCnt = mCnt
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
                    vsGrid.TextMatrix(mRowCnt, 1) = Rec!intAccountHeadID
                    vsGrid.TextMatrix(mRowCnt, 2) = Rec!vchAccountHeadCode
                    vsGrid.TextMatrix(mRowCnt, 3) = Rec!vchAccountHead
                    Rec.MoveNext
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCnt = mRowCnt + 1
                Wend
            Else
                MsgBox "Connection to Finance does not exist, Please Contact your System Administrator"
            End If
            
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub FormInitialize()
        On Error GoTo Err:
            txtBankHeadCode.Text = ""
            txtBankHeadCode.Tag = ""
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
            cmbCategory.ListIndex = -1
            cmbFunction.ListIndex = -1
            cmbFunctionary.ListIndex = -1
            cmbFund.ListIndex = -1
            cmbGroup.ListIndex = -1
            cmbSection.ListIndex = -1
            cmbType.ListIndex = -1
            'vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            vsGrid.Rows = 7
            txtSourceOfFund.Text = ""
            txtSourceOfFund.Tag = ""
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmdRemove_Click()
'        On Error GoTo Err:
'            Dim mCnn As New ADODB.Connection
'            Dim objDb As New clsDB
'            Dim mSql As String
'
'            If Val(txtTransactionType.Tag) = 0 Then
'                MsgBox "Select the Tr.Type", vbInformation
'                txtTransactionType.SetFocus
'                Exit Sub
'            End If
'
'            If objDb.SetConnection(mCnn) Then
'                mSql = "Update faTransactionType Set tnyHidden = 1 Where intTransactionTypeID = " & Val(txtTransactionType.Tag)
'                mCnn.Execute mSql
'                MsgBox "Removed"
'                Call FormInitialize
'            Else
'                MsgBox "Invalid Connection"
'            End If
'        Exit Sub
'Err:
'        MsgBox (Error$)
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDB
            Dim mSql As String
            Dim mRowCnt As Integer

            If vsGrid.TextMatrix(vsGrid.Row, 2) = "" Then
                MsgBox "Please Select Head From the Grid", vbInformation
                vsGrid.SetFocus
                Exit Sub
            End If
            
            vsGrid.RemoveItem (vsGrid.Row)
            
            For mRowCnt = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mRowCnt, 1) = "" Then Exit For
                vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
            Next
            
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim mVarrIn As Variant
        Dim mVarOut As Variant
        Dim mCount As Integer
        
        Dim mCategoryID As Variant
        Dim mTypeID As Variant
        Dim mFundID As Variant
        Dim mFunctionID As Variant
        Dim mFunctionaryID As Variant
        
        '''' Validations ''''
        If Trim(txtTransactionType.Text) = "" Then
            MsgBox "Please Enter Transaction Type", vbCritical, "Saankhya"
            txtTransactionType.SetFocus
            Exit Sub
        End If
        
'        If cmbGroup.ListIndex = 1 Then
'            If cmbFunctionary.ListIndex < 1 Then
'                MsgBox "functionary Must be Selected", vbCritical, "Saankhya"
'                cmbFunctionary.SetFocus
'                Exit Sub
'            End If
'            If cmbFunction.ListIndex < 1 Then
'                MsgBox "Function must be Selected", vbCritical, "Saankhya"
'                cmbFunction.SetFocus
'                Exit Sub
'            End If
'        End If
        
        If cmbFund.ListIndex < 1 Then
            MsgBox "Fund must be Selected", vbCritical, "Saankhya"
            cmbFund.SetFocus
            Exit Sub
        End If
        If cmbSection.ListIndex < 1 Then
            MsgBox "Please Select the Section", vbCritical
            cmbSection.SetFocus
            Exit Sub
        End If
        If cmbGroup.ListIndex < 1 Then
            MsgBox "Please Select the proper group", vbCritical
            cmbGroup.SetFocus
            Exit Sub
        End If
        For mCount = 1 To vsGrid.Rows
            If vsGrid.TextMatrix(mCount, 1) = "" Then Exit For
            If Trim(vsGrid.TextMatrix(mCount, 2)) = "" And Trim(vsGrid.TextMatrix(mCount, 3)) = "" Then
                MsgBox "Please Select AccountHead ", vbCritical
                vsGrid.SetFocus
                Exit Sub
            End If
        Next
        '''' Validations ''''
        
        If cmbType.ListIndex = -1 Then
            mTypeID = Null
        Else
            mTypeID = cmbType.ItemData(cmbType.ListIndex)
        End If
        
        If cmbCategory.ListIndex = -1 Then
            mCategoryID = Null
        Else
            mCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
        End If
        
        If cmbFund.ListIndex = -1 Then
            mFundID = Null
        Else
            mFundID = cmbFund.ItemData(cmbFund.ListIndex)
        End If
        
        If cmbFunction.ListIndex = -1 Then
            mFunctionID = Null
        Else
            mFunctionID = cmbFunction.ItemData(cmbFunction.ListIndex)
        End If
        
        If cmbFunctionary.ListIndex = -1 Then
            mFunctionaryID = Null
        Else
            mFunctionaryID = cmbFunctionary.ItemData(cmbFunctionary.ListIndex)
        End If
        
        mVarrIn = Array(IIf(txtTransactionType.Tag = "", -1, Val(txtTransactionType.Tag)), _
                        Trim(txtTransactionType.Text), _
                        0, mFundID, 115, 1, _
                        cmbGroup.ItemData(cmbGroup.ListIndex), _
                        mID(cmbGroup.List(cmbGroup.ListIndex), 1, 1), _
                        mFunctionID, _
                        mFunctionaryID, Null, Null, _
                        cmbSection.ItemData(cmbSection.ListIndex), _
                        txtBankHeadCode.Tag, _
                        mCategoryID, _
                        mTypeID, _
                        Val(txtSourceOfFund.Tag) _
                        )
        If objDb.SetConnection(mCnn) Then
            objDb.ExecuteSP "spSaveTransactionType", mVarrIn, mVarOut, , mCnn, adCmdStoredProc
            mCnn.Execute "DELETE FROM faTransactionTypeChild where intTransactionTypeID=" & mVarOut(0, 0)
            For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mCount, 1) = "" Then Exit For
                mVarrIn = Array(Val(mVarOut(0, 0)), _
                                mCount, _
                                Val(vsGrid.TextMatrix(mCount, 1)), _
                                IIf(vsGrid.TextMatrix(mCount, 7) = "", Null, Val(vsGrid.TextMatrix(mCount, 7))), _
                                Null, _
                                115, _
                                0, _
                                vsGrid.TextMatrix(mCount, 2), _
                                Null, _
                                Null, _
                                Null, _
                                0, _
                                Null _
                            )
                objDb.ExecuteSP "spSaveTransactionTypeChild", mVarrIn, , , mCnn, adCmdStoredProc
            Next
            MsgBox "Transaction Type Saved", vbInformation, "Saankhya"
            Call FormInitialize
        End If
    End Sub

    Private Sub cmdSearchBank_Click()
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = 2"
        frmSearchAccountHeads.Show 1
        If gbSearchID <> -1 Then
            txtBankHeadCode.Text = CStr(gbSearchStr)
            txtBankHeadCode.Tag = Left(gbSearchStr, 9)
        End If
        gbSearchStr = ""
        gbSearchID = -1
    End Sub

    Private Sub cmdSearchSource_Click()
        On Error GoTo Err:
            Dim mSql As String
            Dim objDb As New clsDB
            Dim mCnn As New ADODB.Connection
        
            mSql = "Select intSourceFundID, vchSourceFundName From suSourceOfFund Where intSourceFundID <> 24 Order By vchSourceFundName"
            frmSearchMasters.SQLQry = mSql
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearchTransactionType_Click()
'        Dim objDb As New clsDB
'        Dim mCnn As New ADODB.Connection
'        Dim mSql As String
'            objDb.SetConnection mCnn
'        mSql = "Select vchTransactionType,intTransactionTypeID from faTransactionType Order By vchTransactionType"
'        lstTransactionType.Visible = True
'        PopulateList lstTransactionType, mSql, , True, , True, enuSourceString.Saankhya

        On Error GoTo Err:
            Dim mSql As String
            Dim objDb As New clsDB
            Dim mCnn As New ADODB.Connection
        
            mSql = "Select intTransactionTypeID, vchTransactionType From faTransactionType Where isnull(tnyHidden,0) <> 1 Order By vchTransactionType"
            frmSearchMasters.SQLQry = mSql
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            txtTransactionType.Text = gbSearchStr
            txtTransactionType.Tag = gbSearchID
            
            gbSearchStr = ""
            gbSearchID = -1
            
            txtTransactionType_LostFocus
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mCount As Integer
        
        If KeyCode = vbKeyDelete And IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 1)) Then
            If MsgBox("Are Sure, You want to Delete the Current row : " & vsGrid.Row, vbYesNo, "Saankhya") = vbYes Then
                objDb.SetConnection mCnn
                mCnn.Execute "Delete from faTransactionTypeChild where intTransactionTypeID= " & Val(txtTransactionType.Tag) & " And intOrder= " & vsGrid.TextMatrix(vsGrid.Row, 0)
                vsGrid.RemoveItem (vsGrid.Row)
                For mCount = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mCount, 1) = "" Then Exit For
                    mCnn.Execute "Update faTransactionTypeChild Set intOrder= " & mCount & " where intTransactionTypeID= " & Val(txtTransactionType.Tag) & " And intOrder= " & vsGrid.TextMatrix(mCount, 0)
                Next
                Call FillFunction
            End If
        End If
    End Sub

    Private Sub Form_Load()
        winXPC.InitIDESubClassing
        vsGrid.ColComboList(2) = "..."
        
        ''' Fill Combos '''
        PopulateList cmbFunctionary, "Select vchFunctionary,intFunctionaryID From faFunctionaries Order By vchFunctionary", , True, , True
        PopulateList cmbFunction, "Select vchFunction,intFunctionID From faFunctions Order By vchFunction", , True, , True
        PopulateList cmbFund, "Select vchFund,intFundID From faFunds Order By vchFund", , True, , True
        PopulateList cmbSection, "Select vchSectionName,intSectionID from faSection Order By vchSectionName", , True, , True
        PopulateList cmbCategory, "Select vchTransactionCategory,intCategoryID from faTransactionCategory Order By intCategoryID", , True, , True
        PopulateList cmbType, "Select vchNatureOfTransaction,intTypeID from faNatureOfTransaction Order By intTypeID", , True, , True
        
        vsGrid.ColComboList(4) = "Debit|Credit"
    End Sub

Private Sub Frame1_Click()
    lstTransactionType.Visible = False
End Sub

    Private Sub lstTransactionType_DblClick()
        If lstTransactionType.ListIndex > 0 Then
            txtTransactionType.Text = lstTransactionType.Text
            txtTransactionType.Tag = lstTransactionType.ItemData(lstTransactionType.ListIndex)
            lstTransactionType.Visible = False
            txtTransactionType_LostFocus
        End If
    End Sub

    Private Sub lstTransactionType_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call lstTransactionType_DblClick
        End If
    End Sub

    Private Sub lstTransactionType_LostFocus()
        lstTransactionType.Visible = False
    End Sub

    Private Sub txtTransactionType_LostFocus()
        Call FillFunction
    End Sub
    
    Private Sub FillFunction()
        Dim mSql As String
        Dim mCount As Integer
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.RecordSet
            objDb.SetConnection mCnn
        mSql = "Select  faTransactionType.intGroupID[GroupID],*, a.intAccountHeadID as HeadID, a.vchAccountHeadCode as Code, a.vchAccountHead as Head, b.vchAccountHeadCode as BankCode, b.vchAccountHead as BankName  From faTransactionType"
        mSql = mSql + " Left Join faTransactionTypeChild On faTransactionType.intTransactionTypeID=faTransactionTypeChild.intTransactionTypeID"
        mSql = mSql + " Left Join faFunctionaries On faFunctionaries.intFunctionaryID=faTransactionType.intFunctionaryID"
        mSql = mSql + " Left Join faFunctions On faFunctions.intFunctionID=faTransactionType.intFunctionID"
        mSql = mSql + " Left Join faSection On faSection.intSectionID=faTransactionType.intSectionID"
        mSql = mSql + " Left Join faFunds On faFunds.intFundID=faTransactionType.intFundID"
        mSql = mSql + " Left Join faAccountHeads a On faTransactionTypeChild.intAccountHeadID=a.intAccountHeadID"
        mSql = mSql + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faTransactionType.intCategoryID "
        mSql = mSql + " Left Join faNatureOfTransaction On faNatureOfTransaction.intTypeID = faTransactionType.intTypeID "
        mSql = mSql + " Left Join faAccountHeads b on b.vchAccountHeadCode = faTransactionType.vchBankHeadCode "
        mSql = mSql + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID = faTransactionType.intSourceFundID "
        mSql = mSql + " Where faTransactionType.vchTransactionType = '" & txtTransactionType.Text & "'"
        mSql = mSql + " Order By intOrder"
        Rec.Open mSql, mCnn
        vsGrid.Rows = 1
        vsGrid.Rows = 50
        
        cmbFunction.ListIndex = -1
        cmbFund.ListIndex = -1
        cmbFunctionary.ListIndex = -1
        cmbGroup.ListIndex = -1
        cmbSection.ListIndex = -1
        
        
        If Not Rec.EOF And Not Rec.BOF Then
            On Error Resume Next
            txtTransactionType.Tag = Rec(1)
            cmbFunctionary.Text = Rec!vchFunctionary
            cmbFunction.Text = Rec!vchFunction
            If Rec!GroupID = 10 Then cmbGroup.Text = "Receipt"
            If Rec!GroupID = 20 Then cmbGroup.Text = "Payment"
            If IsNull(Rec!vchSectionName) Then
                cmbSection.ListIndex = -1
            Else
                cmbSection.Text = Rec!vchSectionName
            End If
            If IsNull(Rec!vchFund) Then
                cmbFund.ListIndex = -1
            Else
                cmbFund.Text = Rec!vchFund
            End If
            If IsNull(Rec!vchNatureOfTransaction) Then
                cmbType.ListIndex = -1
            Else
                cmbType.Text = Rec!vchNatureOfTransaction
            End If
            If IsNull(Rec!vchTransactionCategory) Then
                cmbCategory.ListIndex = -1
            Else
                cmbCategory.Text = Rec!vchTransactionCategory
            End If
            If IsNull(Rec!intSourceFundID) Then
                txtSourceOfFund.Text = ""
                txtSourceOfFund.Tag = ""
            Else
                txtSourceOfFund.Text = Rec!vchSourceFundName
                txtSourceOfFund.Tag = Rec!intSourceFundID
            End If
            
            txtBankHeadCode.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!BankName)
            txtBankHeadCode.Tag = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!BankCode)
            mCount = 1
            While Not Rec.EOF
                vsGrid.TextMatrix(mCount, 0) = mCount
                vsGrid.TextMatrix(mCount, 1) = Rec!HeadID
                vsGrid.TextMatrix(mCount, 2) = Rec!Code
                vsGrid.TextMatrix(mCount, 3) = Rec!Head
                If (Rec!tinDebitOrCredit = 1) Then
                    vsGrid.TextMatrix(mCount, 4) = "Debit"
                ElseIf (Rec!tinDebitOrCredit = 0) Then
                    vsGrid.TextMatrix(mCount, 4) = "Credit"
                End If
                vsGrid.TextMatrix(mCount, 7) = Rec!tinDebitOrCredit
                vsGrid.TextMatrix(mCount, 5) = Rec!intGroupID
                vsGrid.TextMatrix(mCount, 6) = Rec!tnyNetPayFlag
                vsGrid.Rows = vsGrid.Rows + 1
                Rec.MoveNext
                mCount = mCount + 1
            Wend
            On Error GoTo 0
        Else
            txtTransactionType.Tag = -1
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Col = 3 Then
            Cancel = True
        End If
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If Trim(vsGrid.TextMatrix(Row - 1, Col)) <> "" Then
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Order By faAccountHeads.vchAccountHeadCode"
            frmSearchAccountHeads.Show 1
            If gbSearchID <> -1 Then
                vsGrid.TextMatrix(Row, 0) = Row
                vsGrid.TextMatrix(Row, 2) = Token(gbSearchStr, " ")
                vsGrid.TextMatrix(Row, 3) = Trim(gbSearchStr)
                vsGrid.TextMatrix(Row, 1) = gbSearchID
            End If
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub vsGrid_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
        If vsGrid.Col = 4 Then
            If vsGrid.TextMatrix(Row, 4) = "Debit" Then vsGrid.TextMatrix(Row, 7) = 1
            If vsGrid.TextMatrix(Row, 4) = "Credit" Then vsGrid.TextMatrix(Row, 7) = 0
        End If
    End Sub

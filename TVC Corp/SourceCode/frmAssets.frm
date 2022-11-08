VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAssets 
   Caption         =   "Assets"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAssetLink 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   5880
      Width           =   255
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   7170
      Top             =   6135
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid VSGrid 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   7335
      _cx             =   12938
      _cy             =   5530
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAssets.frx":0000
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7335
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtAssetCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   210
         Width           =   2535
      End
      Begin VB.TextBox txtAssetName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   7
         Top             =   570
         Width           =   2535
      End
      Begin VB.ComboBox cmbCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   975
         Width           =   2535
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtPeriod 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtTotalCost 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Head "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Asset Code "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Asset Name "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Asset Type "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Period :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cost :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Link With Assets Register"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   5880
      Width           =   2010
   End
End
Attribute VB_Name = "frmAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub chkAssetLink_Click()
        Call FillGrid
    End Sub

    Private Sub cmbCategory_Click()
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mSql As String
             
        If cmbCategory.ListIndex > -1 Then
            txtAccountHead.Enabled = True
            cmdSearchAccountHead.Enabled = True
            mSql = "Update faAssetsType set vchMinorAccountHeadCode=Replace(vchMinorAccountHeadCode,' ','')"
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            Call FillGrid
        End If
    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdSave_Click()
         Dim mCnn    As New ADODB.Connection
         Dim objdb   As New clsDB
         Dim mintID  As Variant
         Dim mStatus As Variant
         Dim mArrIn  As Variant
         Dim mArrOut As Variant

         If objdb.SetConnection(mCnn) Then
            If Trim(txtAssetName.Text) = "" Then
                MsgBox "Enter the Assets Name", vbInformation, "Saankhya"
                Exit Sub
            End If
            If cmbCategory.ListIndex = -1 Then
                MsgBox "Select a Category", vbInformation, "Saankhya"
                Exit Sub
            End If
            If Trim(txtYear.Text) = "" Then
                MsgBox "Enter the Year", vbInformation, "Saankhya"
                Exit Sub
            End If
            
            mintID = IIf(txtAssetName.Tag = "", -1, val(txtAssetName.Tag))
            mArrIn = Array(mintID, IIf(IsNull(txtAssetCode.Text), Null, txtAssetCode.Text), _
                            cmbCategory.ItemData(cmbCategory.ListIndex), _
                            txtAssetName.Text, _
                            txtAccountHead.Tag, _
                            txtYear.Text, _
                            Null, _
                            Null, _
                            0 _
                            )
            objdb.ExecuteSP "spSaveAssets", mArrIn, mArrOut, , mCnn, adCmdStoredProc
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
            txtAssetCode.Text = mArrOut(0, 0)
  
         Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
         End If
         cmdSave.Enabled = False
    End Sub

    Private Sub cmdSearchAccountHead_Click()
        Dim mToken As String
        
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID FROM faAssetsType INNER JOIN faMinorAccountHeads ON faAssetsType.vchMinorAccountHeadCode = faMinorAccountHeads.vchMinorAccountHeadCode INNER JOIN faAccountHeads ON faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID where faAccountHeads.intMinorAccountHeadID = " & cmbCategory.ItemData(cmbCategory.ListIndex) & "  "
        frmSearchAccountHeads.Show vbModal
        mToken = Token(gbSearchStr, " ")
           If gbSearchID <> -1 Then
               txtAccountHead.Text = Trim(gbSearchStr)
               txtAccountHead.Tag = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        'Call PopulateList(cmbCategory, "select vchMinorAccountHead,intID from faAssetsTYpe", , , True, True)
        Call PopulateList(cmbCategory, "select faAssetsType.vchMinorAccountHead,faMinorAccountHeads.intMinorAccountHeadID from faAssetsType Inner join faMinorAccountHeads on RTrim(faAssetsType.vchMinorAccountHeadCode) = faMinorAccountHeads.vchMinorAccountHeadCode", , , True, True)
    End Sub
    Private Sub FillGrid()
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mRowCnt As Integer

        If objdb.SetConnection(mCnn) Then
            mSql = " select * from faAssets where intAssetType = " & cmbCategory.ItemData(cmbCategory.ListIndex) & " "
            If chkAssetLink.value = vbChecked Then
                mSql = mSql + " and tnyLinkWithAssetMaster=1 "
            End If
            mSql = mSql + " order by intAssetID desc "
            '''mSQL = mSQL + " from faAccountHeads where intMinorAccountHeadID = " & cmbCategory.ItemData(cmbCategory.ListIndex) & " "
            Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            mRowCnt = 1
            vsGrid.Clear 1, 1
            vsGrid.Rows = 1
            While Not (Rec.EOF Or Rec.BOF)
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(mRowCnt, 0) = mRowCnt
                vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!vchAssetCode), "", Rec!vchAssetCode)
                vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!vchAssetName), "", Rec!vchAssetName)
                vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!intYear), "", Rec!intYear)
                vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!intAssetID), "", Rec!intAssetID)
                vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!intAssetType), "", Rec!intAssetType)
                Rec.MoveNext
                mRowCnt = mRowCnt + 1
            Wend
            Rec.Close
        End If
    End Sub
   Private Sub vsGrid_Click()
        Dim mCnn As New ADODB.Connection
        Dim Rec  As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
            If vsGrid.Row > 0 Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                mSql = "select vchAccountHead from faAccountHeads where intAccountHeadID= " & vsGrid.TextMatrix(vsGrid.Row, 5) & " "
                Rec.Open mSql, mCnn
                While Not (Rec.EOF Or Rec.BOF)
                    txtAccountHead.Text = Rec!vchAccountHead
                Rec.MoveNext
                Wend
                Rec.Close
                txtAssetCode.Text = vsGrid.TextMatrix(vsGrid.Row, 1)
                txtAssetName.Text = vsGrid.TextMatrix(vsGrid.Row, 2)
                txtAssetName.Tag = vsGrid.TextMatrix(vsGrid.Row, 6)
                txtYear.Text = vsGrid.TextMatrix(vsGrid.Row, 3)
                txtAccountHead.Tag = vsGrid.TextMatrix(vsGrid.Row, 5)
                cmdSave.Caption = "Update"
                If vsGrid.Cell(flexcpChecked, vsGrid.Row, 4) = flexcpChecked Then
                    vsGrid.Editable = flexEDNone
                End If
            End If
    End Sub
    Private Sub vsGrid_DblClick()
        If vsGrid.TextMatrix(vsGrid.Row, 0) <> "" Then
            gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 2)
            gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 6)
            gbSearchCode = vsGrid.TextMatrix(vsGrid.Row, 6)
            Unload Me
        End If
    End Sub
     Private Sub FormInitialize()
        txtAssetCode.Text = ""
        txtAssetName.Text = ""
        txtAssetName.Tag = ""
        txtYear.Text = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
     End Sub

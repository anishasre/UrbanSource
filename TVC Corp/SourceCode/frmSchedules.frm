VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReportSchedules 
   Caption         =   "Define Schedule"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9840
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   7980
      TabIndex        =   13
      Top             =   5550
      Width           =   1245
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   750
      TabIndex        =   11
      Top             =   4770
      Width           =   3825
   End
   Begin VB.ListBox lstSelected 
      Height          =   645
      Left            =   3480
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4830
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   150
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   150
      Width           =   2085
   End
   Begin VB.ListBox lstAccountHeads 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   60
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1170
      Width           =   4485
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4650
      TabIndex        =   1
      Top             =   2505
      Width           =   450
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4650
      TabIndex        =   0
      Top             =   3105
      Width           =   450
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   -3570
      Top             =   6360
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   3885
      Left            =   5175
      TabIndex        =   2
      Top             =   1170
      Width           =   4575
      _cx             =   8070
      _cy             =   6853
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchedules.frx":0000
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Search"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   4815
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Schedule Title"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   3690
      TabIndex        =   6
      Top             =   195
      Width           =   795
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "  Account Heads"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   870
      Width           =   4890
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "       Selected Heads"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4905
      TabIndex        =   4
      Top             =   870
      Width           =   4845
   End
End
Attribute VB_Name = "frmReportSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mEditFlag As Boolean
    
    Private Sub cmdAdd_Click()
        AddToGrid
    End Sub

    Private Sub cmdRemove_Click()
        RemoveFromGrid
    End Sub

    Private Sub cmdSave_Click()
        
    End Sub

    Private Sub Form_Load()
        WindowsXPC.InitIDESubClassing
        FormInitialize
        PopulateList lstAccountHeads, "Select vchAccountHead,intAccountHeadID from faAccountHeads order by vchAccountHead", , False, True, True, enuSourceString.Saankhya
    End Sub
    
    Private Sub FormInitialize()
        mEditFlag = False
        vsGrid.Rows = 1
        Me.Height = 6915
        Me.Width = 9960
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            End If
        Next
    End Sub
    
    Private Sub AddToGrid()
        Dim mCount As Integer
        If lstAccountHeads.ListCount > 0 Then
            For mCount = 0 To lstAccountHeads.ListCount - 1
                If lstAccountHeads.Selected(mCount) Then
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 0) = lstAccountHeads.List(mCount)
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1) = lstAccountHeads.ItemData(mCount)
                    vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 2) = mCount
                    lstSelected.AddItem vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1)
                    lstSelected.ItemData(lstSelected.NewIndex) = vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1)
                End If
            Next mCount
            Call FillAccountHeads
        End If
    End Sub

    Private Sub lstAccountHeads_Click()
        If lstAccountHeads.ListIndex > 0 Then
            Dim objAccounts As New clsAccounts
            objAccounts.SetAccounts (lstAccountHeads.ItemData(lstAccountHeads.ListIndex))
            lstAccountHeads.ToolTipText = objAccounts.AccountCode
        End If
    End Sub

    Private Sub lstAccountHeads_DblClick()
        AddToGrid
    End Sub
    
    Private Sub RemoveFromGrid()
        Dim mLoop As Long
        Dim mChildLoop As Long
        Dim mIndex As Long
        Dim mTempString As String
        Dim mCount As Integer
        If vsGrid.Rows > 1 Then
            For mCount = 1 To vsGrid.Rows - 1
                If mCount < vsGrid.Rows Then
                    If vsGrid.IsSelected(mCount) = True Then
                        mTempString = vsGrid.Cell(flexcpText, mCount, 1)
                        mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTempString)
                        If mIndex <> -1 Then
                            lstSelected.RemoveItem (mIndex)
                        End If
                        vsGrid.RemoveItem (mCount)
                    End If
                End If
            Next mCount
        End If
        For mCount = 1 To vsGrid.Rows - 1
            vsGrid.Cell(flexcpText, mCount, 2) = mCount
        Next mCount
        Call FillAccountHeads
    End Sub
    
    Private Sub FillAccountHeads()
        lstAccountHeads.Clear
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mIndex As Integer
        Dim mTempString As String
        If objDB.SetConnection(mCnn) Then
            Set Rec = objDB.ExecuteSP("Select vchAccountHead,intAccountHeadID from faAccountHeads Order By vchAccountHEad", , , , mCnn, adCmdText)
        End If
        While Not Rec.EOF
            mTempString = Rec!intAccountHeadID
            mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTempString)
            If mIndex = -1 Then
                lstAccountHeads.AddItem Rec!vchAccountHead
                lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = Rec!intAccountHeadID
            End If
            Rec.MoveNext
        Wend
    End Sub

    Private Sub txtSearch_Change()
        Dim mLoop As Long
        If lstAccountHeads.ListCount > 0 Then
            lstAccountHeads.Selected(lstAccountHeads.ListIndex) = False
            For mLoop = 0 To lstAccountHeads.ListCount - 1
                If IsNumeric(Trim(txtSearch.Text)) Then
                    If lstAccountHeads.ItemData(mLoop) = Trim(txtSearch.Text) Then
                        lstAccountHeads.ListIndex = mLoop
                        lstAccountHeads.Selected(mLoop) = True
                        Exit For
                    End If
                End If
            Next mLoop
        End If
    End Sub

    Private Sub txtSearch_KeyPress(KeyAscii As Integer)
       ' MsgBox KeyAscii
    End Sub

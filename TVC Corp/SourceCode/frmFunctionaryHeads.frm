VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmFunctionaryHeads 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F u n c t i o n a r y   H e a d s"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   5130
      TabIndex        =   9
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3780
      TabIndex        =   8
      Top             =   6060
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9870
      TabIndex        =   10
      Top             =   0
      Width           =   9870
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   9870
      Begin VB.ComboBox cmbAccountGroups 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4830
         Width           =   3000
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   16
         Top             =   600
         Width           =   3870
      End
      Begin VB.ListBox lstMasters 
         BackColor       =   &H80000018&
         Height          =   4545
         ItemData        =   "frmFunctionaryHeads.frx":0000
         Left            =   8430
         List            =   "frmFunctionaryHeads.frx":0002
         TabIndex        =   15
         Top             =   570
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdFunctionaries 
         Caption         =   "..."
         Height          =   285
         Left            =   5220
         TabIndex        =   13
         Top             =   570
         Width           =   315
      End
      Begin VB.ListBox lstSelected 
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   345
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
         Left            =   4710
         TabIndex        =   7
         Top             =   3285
         Width           =   450
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
         Left            =   4710
         TabIndex        =   4
         Top             =   2685
         Width           =   450
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   3645
         Left            =   5235
         TabIndex        =   6
         Top             =   1350
         Width           =   4575
         _cx             =   8070
         _cy             =   6429
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
         FormatString    =   $"frmFunctionaryHeads.frx":0004
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
         Top             =   1350
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Groups"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   18
         Top             =   4875
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   30
         Left            =   60
         TabIndex        =   11
         Top             =   1020
         Width           =   9750
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
         Left            =   4965
         TabIndex        =   5
         Top             =   1050
         Width           =   4845
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
         Left            =   60
         TabIndex        =   2
         Top             =   1050
         Width           =   4890
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   9810
      End
   End
End
Attribute VB_Name = "frmFunctionaryHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        
        Option Explicit
        Dim mGridRows As Long
        
        Private Sub FillAccountHeads(intFunctionaryID As Double)
                Dim objDB As New clsDB
                Dim Rec As New ADODB.Recordset
                Dim mSQL As String
                
                vsGrid.Visible = False
                vsGrid.Rows = 1
                vsGrid.Rows = 50
                mGridRows = 0
                mSQL = "Select faAccountHeads.vchAccountHead, faAccountHeads.intAccountHeadID, faAccountHeads.tinType"
                mSQL = mSQL + " From faFunctionaryHeads"
                mSQL = mSQL + " Inner Join faAccountHeads"
                mSQL = mSQL + " On faAccountHeads.intAccountHeadID = faFunctionaryHeads.intAccountHeadID"
                mSQL = mSQL + " Where faFunctionaryHeads.intFunctionaryID = " & intFunctionaryID
                
                Set Rec = GetRecordSet(mSQL)
                If Not (Rec.BOF And Rec.EOF) Then
                    While Not Rec.EOF
                        mGridRows = mGridRows + 1
                        vsGrid.AddItem Rec!vchAccountHead & vbTab & Rec!intAccountHeadId & vbTab & Rec!tinType, mGridRows
                        lstSelected.AddItem Rec!intAccountHeadId
                        Rec.MoveNext
                    Wend
                End If
                Rec.Close
                vsGrid.Visible = True
        End Sub
        
        Private Sub FillList()
                Dim objDB As New clsDB
                Dim arrInput As Variant
                Dim Rec As New ADODB.Recordset
                Dim mTemp As String
                Dim mIndex As Long
                
                Select Case cmbAccountGroups.Text
                    Case Is = "Income":         arrInput = Array(1)
                    Case Is = "Expenditures":   arrInput = Array(2)
                    Case Is = "Liabilities":    arrInput = Array(3)
                    Case Is = "Assets":         arrInput = Array(4)
                    Case Else:                  arrInput = Array(0)
                End Select
                
                Rec.CursorLocation = adUseClient
                Set Rec = objDB.ExecuteSP("spGetAccountHeadsByType", arrInput)
                If Not (Rec.EOF And Rec.BOF) Then
                    lstAccountHeads.Clear
                    While Not Rec.EOF
                        mTemp = Rec!intAccountHeadId
                        mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTemp)
                        If mIndex = -1 Then
                            lstAccountHeads.AddItem Rec!vchAccountHead
                            lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = Rec!intAccountHeadId
                        End If
                        Rec.MoveNext
                    Wend
                End If
                Rec.Close
        End Sub
       
        Private Sub FormInitialize()
                txtFunctionary.Text = ""
                vsGrid.Clear 1, 1
                lstSelected.Clear
                cmbAccountGroups.ListIndex = -1
        End Sub
        
        Private Sub cmbAccountGroups_Click()
                Call FillList
        End Sub

        Private Sub cmbAccountGroups_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                    PressTabKey
                End If
        End Sub

        Private Sub cmdAdd_Click()
                Dim mCount As Long
                Dim mGroupID As Integer
                Dim mIndex As Long
                
                For mCount = 0 To lstAccountHeads.ListCount - 1
                    If lstAccountHeads.Selected(mCount) Then
                    Select Case cmbAccountGroups.Text
                        Case Is = "Income": mGroupID = 1
                        Case Is = "Expenditures": mGroupID = 2
                        Case Is = "Liabilities": mGroupID = 3
                        Case Is = "Assets": mGroupID = 4
                        Case Else: mGroupID = 0
                    End Select
                    vsGrid.AddItem lstAccountHeads.List(mCount) & vbTab & lstAccountHeads.ItemData(mCount) & vbTab & CStr(mGroupID), mGridRows + 1
                    lstSelected.AddItem lstAccountHeads.ItemData(mCount)
                    mGridRows = mGridRows + 1
                    End If
                Next mCount
                '--------------------------------------------------------------------------'
                ' Refilling Account head on Left side                                      '
                '--------------------------------------------------------------------------'
                  mIndex = lstAccountHeads.ListIndex
                  Call FillList
                  If lstAccountHeads.ListIndex > 1 Then lstAccountHeads.ListIndex = mIndex
                '--------------------------------------------------------------------------'
        End Sub

        Private Sub cmdCancel_Click()
                Unload Me
        End Sub
        
        Private Sub cmdRemove_Click()
                Dim mLoop As Long
                Dim mChildLoop As Long
                Dim mTemp As String
                Dim mIndex As Long
Skip:
                vsGrid.Refresh
                For mLoop = 0 To vsGrid.SelectedRows - 1
                    If vsGrid.TextMatrix(vsGrid.SelectedRow(mLoop), 0) <> "" Then
                        mTemp = Val(vsGrid.TextMatrix(vsGrid.SelectedRow(mLoop), 1))
                        mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTemp)
                        If mIndex > -1 Then
                            lstSelected.RemoveItem (mIndex)
                            vsGrid.RemoveItem vsGrid.SelectedRow(mLoop)
                            If mGridRows > 0 Then mGridRows = mGridRows - 1
                            GoTo Skip
                        End If
                    End If
                Next mLoop
                FillList
        End Sub

        Private Sub cmdSave_Click()
                Dim mLoop As Long
                Dim objFunctionary As New clsFunctionary
                Dim objDB As New clsDB
                Dim arrInput As Variant
                Dim mCnn As New ADODB.Connection
    
                objDB.SetConnection mCnn
                'mCnn.BeginTrans
                'On Error GoTo ErrHandler
                objFunctionary.SetFunctionaryByID Val(txtFunctionary.Tag)
                mCnn.Execute "Delete From faFunctionaryHeads Where intFunctionaryID = " & objFunctionary.FunctionaryID
                For mLoop = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mLoop, 0) = "" And Val(vsGrid.TextMatrix(mLoop, 1)) <= 0 Then
                        Exit For
                    ElseIf Val(vsGrid.TextMatrix(mLoop, 1)) > 0 Then
                        arrInput = Array((objFunctionary.FunctionaryID), Val(vsGrid.TextMatrix(mLoop, 1)))
                        objDB.ExecuteSP "spSaveFunctionaryHead", arrInput, , , mCnn
                    End If
                Next mLoop
                'mCnn.RollbackTrans
                Set mCnn = Nothing
                FormInitialize
                Exit Sub
ErrHandler:
                'mCnn.RollbackTrans
                Set mCnn = Nothing
        End Sub

        Private Sub Form_Activate()
                Me.Top = 0
                frmFunctionaryHeads.Left = (frmMenu.Width - Me.Width) / 2
        End Sub
            
        Private Sub Form_Load()
                cmbAccountGroups.AddItem "Income"
                cmbAccountGroups.AddItem "Expenditures"
                cmbAccountGroups.AddItem "Liabilities"
                cmbAccountGroups.AddItem "Assets"
                mGridRows = 0
                Call FillList
        End Sub
        
        Private Sub lstAccountHeads_DblClick()
                Call cmdAdd_Click
        End Sub

        Private Sub lstAccountHeads_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then
                PressTabKey
                End If
        End Sub
        
        Private Sub lstMasters_DblClick()
                gbSearchStr = lstMasters.Text
                gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
                txtFunctionary.SetFocus
                lstMasters.Visible = False
        End Sub
        
        Private Sub lstMasters_GotFocus()
                lstMasters.Width = 4000
                lstMasters.Left = 5760
        End Sub
    
        
          
     
    

        Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
                If Col = 0 Then
                Cancel = True
                End If
        End Sub

        Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
                If Col = 1 Then
                    If vsGrid.TextMatrix(Row, 0) <> "" Then
                    vsGrid.TextMatrix(Row, 1) = Format(Val(vsGrid.TextMatrix(Row, Col)), "0.00")
                    End If
                End If
        End Sub

        Private Sub vsGrid_Click()
                If vsGrid.Row = 0 Then
                vsGrid.Sort = flexSortStringAscending
                End If
        End Sub
        
        

        Private Sub cmdFunctionaries_Click()
                Dim mSQL As String
                mSQL = "Select vchFunctionary, intFunctionaryID From faFunctionaries Order By vchFunctionary"
                Call PopulateList(lstMasters, mSQL, , True, , True)
                lstMasters.Tag = "2"
                lstMasters.Visible = True
                lstMasters.SetFocus
        End Sub

        Private Sub txtFunctionary_GotFocus()
                If gbSearchStr <> "" Then
                    txtFunctionary.Text = gbSearchStr
                    txtFunctionary.Tag = gbSearchID
                    Call FillAccountHeads(gbSearchID)
                    cmbAccountGroups.ListIndex = -1
                    Call FillList
                    gbSearchStr = ""
                    gbSearchID = -1
                End If
        End Sub
        
        Private Sub txtFunctionary_KeyDown(KeyCode As Integer, Shift As Integer)
                If KeyCode = vbKeyF4 Then
                    Call cmdFunctionaries_Click
                ElseIf KeyCode = vbKeyDelete Then
                    txtFunctionary.Text = ""
                    txtFunctionary.Tag = ""
                End If
        End Sub
        
        Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
                If KeyAscii = 13 Then PressTabKey
        End Sub
        
        Private Sub DispFinancialYear(mFinancialYearID)
                Dim mSQL As String
                Dim Rec As New ADODB.Recordset
                
                mSQL = "SELECT dtStartingDate, dtEndingDate, intFinancialYearID,tinCurrentFinancialYearFlag From faFinancialYear Where faFinancialYear.intFinancialYearID=" & mFinancialYearID
                Set Rec = GetRecordSet(mSQL)
                If Not (Rec.BOF And Rec.EOF) Then
                    'txtFinancialYear.Text = Format(Rec!dtStartingDate, "Dd-Mmm-yyyy") & " -- " & Format(Rec!dtEndingDate, "Dd-Mmm-yyyy")
                    'txtFinancialYear.Tag = Rec!intFinancialYearID
                End If
        End Sub

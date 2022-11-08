VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBudgetAllocation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " B u d g e t   A l l o c a t i o n "
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
      TabIndex        =   16
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3780
      TabIndex        =   15
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
      TabIndex        =   17
      Top             =   0
      Width           =   9870
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   -30
      TabIndex        =   0
      Top             =   660
      Width           =   9870
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
         Left            =   7830
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Top             =   840
         Width           =   3360
      End
      Begin VB.ListBox lstMasters 
         BackColor       =   &H80000018&
         Height          =   4545
         ItemData        =   "frmBudgetAllocation.frx":0000
         Left            =   8970
         List            =   "frmBudgetAllocation.frx":0002
         TabIndex        =   33
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtFinancialYear 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7740
         TabIndex        =   35
         Top             =   420
         Width           =   2040
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   1170
         Width           =   3360
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5430
         TabIndex        =   27
         Top             =   1170
         Width           =   3360
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5430
         TabIndex        =   26
         Top             =   840
         Width           =   3360
      End
      Begin VB.CommandButton cmdFunctionaries 
         Caption         =   "..."
         Height          =   285
         Left            =   4590
         TabIndex        =   25
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdFields 
         Caption         =   "..."
         Height          =   285
         Left            =   8820
         TabIndex        =   24
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdFunctions 
         Caption         =   "..."
         Height          =   285
         Left            =   4590
         TabIndex        =   23
         Top             =   1170
         Width           =   315
      End
      Begin VB.CommandButton cmdFunds 
         Caption         =   "..."
         Height          =   285
         Left            =   8820
         TabIndex        =   22
         Top             =   1170
         Width           =   315
      End
      Begin VB.ListBox lstSelected 
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   1710
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmbSearch 
         Caption         =   "..."
         Height          =   285
         Left            =   5970
         TabIndex        =   6
         Top             =   420
         Width           =   315
      End
      Begin VB.TextBox txtBudgetCentre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2820
         TabIndex        =   4
         Top             =   420
         Width           =   3105
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   2355
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
         TabIndex        =   14
         Top             =   3825
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
         TabIndex        =   11
         Top             =   3225
         Width           =   450
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2805
         Left            =   5220
         TabIndex        =   13
         Top             =   2400
         Width           =   4515
         _cx             =   7964
         _cy             =   4948
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBudgetAllocation.frx":0004
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
         Height          =   2760
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   2430
         Width           =   4515
      End
      Begin VB.ComboBox cmbAccountGroups 
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2130
      End
      Begin VB.TextBox txtBudgetCentreCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Financial year"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6450
         TabIndex        =   34
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fund"
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
         Left            =   4950
         TabIndex        =   32
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Function"
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
         Left            =   390
         TabIndex        =   31
         Top             =   1170
         Width           =   780
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
         Left            =   90
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Field"
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
         Left            =   4980
         TabIndex        =   29
         Top             =   870
         Width           =   420
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
         Left            =   90
         TabIndex        =   7
         Top             =   1725
         Width           =   1380
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
         TabIndex        =   20
         Top             =   1560
         Width           =   9750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount Allocated :"
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
         Left            =   5115
         TabIndex        =   18
         Top             =   1725
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "       Allocated Account Heads"
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
         Left            =   4875
         TabIndex        =   12
         Top             =   2115
         Width           =   4815
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
         Left            =   30
         TabIndex        =   9
         Top             =   2115
         Width           =   4800
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "  Budget Centre"
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
         Top             =   90
         Width           =   9810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Left            =   750
         TabIndex        =   2
         Top             =   450
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmBudgetAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        
        Option Explicit
        
        Dim objBudgetCentre As New clsBudgetCentre
        Dim mGridRows As Long
        Private mSelectedHeads As New Collection
                
        Private Sub FillAccountHeads(intBgtID As Long)
            Dim objBgt As New clsBudgetCentre
            Dim Rec As New ADODB.Recordset
            
            vsGrid.Visible = False
            vsGrid.Rows = 1
            vsGrid.Rows = 50
            vsGrid.Visible = True
            mGridRows = 0
            
            Set Rec = objBgt.GetAccountHeads(intBgtID)
            If Not (Rec.BOF And Rec.EOF) Then
                While Not Rec.EOF
                    mGridRows = mGridRows + 1
                    vsGrid.AddItem Rec!vchAccountHeadCode + "  " + Rec!vchAccountHead & vbTab & Rec!fltEstimatedAmount & vbTab & Rec!intAccountHeadID & vbTab & Rec!tinType, mGridRows
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
            Dim mTemp2 As String
            Dim mIndex As Long
            
            Select Case cmbAccountGroups.Text
                Case Is = "Income"
                    arrInput = Array(1)
                Case Is = "Expenditures"
                    arrInput = Array(2)
                Case Is = "Liabilities"
                    arrInput = Array(3)
                Case Is = "Assets"
                    arrInput = Array(4)
                Case Else
                    arrInput = Array(100)
            End Select
            Rec.CursorLocation = adUseClient
            Set Rec = objDB.ExecuteSP("spGetAccountHeadsByType", arrInput)
            If Not (Rec.EOF And Rec.BOF) Then
                lstAccountHeads.Clear
                While Not Rec.EOF
                    mTemp = Rec!vchAccountHead
                    mTemp2 = Rec!vchAccountHeadCode
                    'If Rec!vchAccountHeadCode = "110150200" Then Stop
                    mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTemp)
                    If mIndex = -1 Then
                        lstAccountHeads.AddItem mTemp2 + "     " + mTemp
                        lstAccountHeads.ItemData(lstAccountHeads.NewIndex) = Rec!intAccountHeadID
                    End If
                    Rec.MoveNext
                Wend
            End If
        End Sub
       
        Private Sub Calculate()
            Dim mAmt As Double
            Dim mLoop As Long
            For mLoop = 1 To vsGrid.Rows - 1
                mAmt = mAmt + Val(vsGrid.TextMatrix(mLoop, 1))
            Next mLoop
            txtAmount.Text = Format(mAmt, "0.00")
        End Sub
        
        Private Sub FormInitialize()
            txtBudgetCentreCode.Text = ""
            txtBudgetCentre.Text = ""
            txtAmount.Text = ""
            lstAccountHeads.Clear
            vsGrid.Clear 1, 1
            lstSelected.Clear
            
        End Sub
        
        Private Sub cmbAccountGroups_Click()
            Call FillList
        End Sub

        Private Sub cmbAccountGroups_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                PressTabKey
            End If
        End Sub

        Private Sub cmbSearch_Click()
            Call txtBudgetCentre_KeyDown(vbKeyF4, 0)
        End Sub

        Private Sub cmdAdd_Click()
            Dim mCount As Long
            Dim mGroupID As Integer
            For mCount = 0 To lstAccountHeads.ListCount - 1
                If lstAccountHeads.Selected(mCount) Then
                    Select Case cmbAccountGroups.Text
                        Case Is = "Income"
                            mGroupID = 1
                        Case Is = "Expenditures"
                            mGroupID = 2
                        Case Is = "Liabilities"
                            mGroupID = 3
                        Case Is = "Assets"
                            mGroupID = 4
                        Case Else
                            mGroupID = 100
                    End Select
                    vsGrid.AddItem lstAccountHeads.List(mCount) & vbTab & "0.00" & vbTab & lstAccountHeads.ItemData(mCount) & vbTab & CStr(mGroupID), mGridRows + 1
                    lstSelected.AddItem lstAccountHeads.List(mCount)
                    lstSelected.ItemData(lstSelected.NewIndex) = lstAccountHeads.ItemData(mCount)
                    mGridRows = mGridRows + 1
                End If
            Next mCount
            '------------------------------------------'
            ' Refilling Account head on Left side
            '------------------------------------------'
            Call FillList
            Call Calculate
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
            For mLoop = 0 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mLoop, 0) <> "" Then
                For mChildLoop = 0 To lstSelected.ListCount - 1
                    If lstSelected.List(mChildLoop) = vsGrid.TextMatrix(mLoop, 0) Then
                        Exit For
                    End If
                Next mChildLoop
                If mChildLoop = lstSelected.ListCount Then
                    vsGrid.RemoveItem (mLoop)
                    GoTo Skip:
                End If
                End If
            Next mLoop
            
            mTemp = vsGrid.TextMatrix(vsGrid.Row, 0)
            mIndex = SendMyMessage(lstSelected.hwnd, LB_FINDSTRING, -1, ByVal mTemp)
            If mIndex > -1 Then
                lstSelected.RemoveItem (mIndex)
            End If
            FillList
            
            
        End Sub
        
        Private Sub cmdSave_Click()
            Dim mLoop As Long
            Dim objBgt As New clsBudgetCentre
            Dim objDB As New clsDB
            Dim arrInput As Variant
            Dim mCnn As New ADODB.Connection
            
            objBgt.SetBudgetCentre CStr(Trim(txtBudgetCentreCode.Text))
            If objBgt.BudgetCentreID > -1 Then
                objDB.SetConnection mCnn
                'mCnn.BeginTrans
                'On Error GoTo ErrHandler
                mCnn.Execute "Delete From faBudgetAccountHeads Where intBudgetCentreID = " & objBgt.BudgetCentreID
                For mLoop = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mLoop, 0) = "" And Val(vsGrid.TextMatrix(mLoop, 1)) <= 0 Then
                        Exit For
                    ElseIf Val(vsGrid.TextMatrix(mLoop, 2)) > 0 And Val(vsGrid.TextMatrix(mLoop, 1)) > 0 Then
                        arrInput = Array(0, _
                                    (objBgt.BudgetCentreID), _
                                    Val(vsGrid.TextMatrix(mLoop, 2)), _
                                    Format(Val(vsGrid.TextMatrix(mLoop, 1)), "0.00"))
                        objDB.ExecuteSP "spSaveBudgetAccountHead", arrInput, , , mCnn
                    End If
                Next mLoop
                'mCnn.RollbackTrans
                Set mCnn = Nothing
                FormInitialize
                Exit Sub
ErrHandler:
                'mCnn.RollbackTrans
                Set mCnn = Nothing
            End If
        End Sub

        Private Sub Form_Activate()
            Me.Top = 0
            frmBudgetAllocation.Left = (frmMenu.Width - Me.Width) / 2
            Call PopulateList(lstBudgetCentre, "Select vchBudgetCentre, intBudgetCentreID From faBudgetCentres Order By vchBudgetCentre", , , , True)
        End Sub
            
        Private Sub Form_Load()
            cmbAccountGroups.AddItem "Income"
            cmbAccountGroups.AddItem "Expenditures"
            cmbAccountGroups.AddItem "Liabilities"
            cmbAccountGroups.AddItem "Assets"
            mGridRows = 0
        End Sub
        
        Private Sub lstAccountHeads_DblClick()
            Call cmdAdd_Click
        End Sub

        Private Sub lstAccountHeads_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                PressTabKey
            End If
        End Sub
        
        Private Sub lstBudgetCentre_DblClick()
            
            txtBudgetCentre.Text = lstBudgetCentre.Text
            If lstBudgetCentre.ListIndex > -1 Then
                objBudgetCentre.SetBudgetCentreByID (lstBudgetCentre.ItemData(lstBudgetCentre.ListIndex))
                If objBudgetCentre.BudgetCentreCode <> "" Then
                    txtBudgetCentreCode.Text = objBudgetCentre.BudgetCentreCode
                End If
            End If
            Call txtBudgetCentreCode_GotFocus
            '-------------------------------'
            ' Sorry to do like this  : Aiby '
            '-------------------------------'
                Call PressTabKey
                Call PressTabKey
                Call PressTabKey
            '-------------------------------'
            lstBudgetCentre.Visible = False
        End Sub
        
        Private Sub lstBudgetCentre_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                Call lstBudgetCentre_DblClick
            End If
        End Sub

    Private Sub lstMasters_DblClick()
        gbSearchStr = lstMasters.Text
        gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
        Select Case Val(lstMasters.Tag)
            Case 1: txtFunction.SetFocus
            Case 2: txtFunctionary.SetFocus
            Case 3: txtField.SetFocus
            Case 4: txtFund.SetFocus
            Case 5: txtBudgetCentre.SetFocus
        End Select
        lstMasters.Visible = False
        
    End Sub
    Private Sub lstMasters_GotFocus()
        lstMasters.Width = 4000
        lstMasters.Left = 5760
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
        
        Private Sub txtBudgetCentre_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                PressTabKey
            End If
        End Sub
      
        Private Sub txtBudgetCentreCode_GotFocus()
            txtBudgetCentreCode.Text = Trim(txtBudgetCentreCode)
            If Len(txtBudgetCentreCode) Then
                objBudgetCentre.SetBudgetCentre (txtBudgetCentreCode.Text)
                If objBudgetCentre.BudgetCentreID > -1 Then
                    txtBudgetCentreCode.Text = objBudgetCentre.BudgetCentreCode
                    txtBudgetCentre.Text = objBudgetCentre.BudgetCentre
                    txtFunction.Text = objBudgetCentre.FunctionName
                    txtFunction.Tag = objBudgetCentre.FunctionID
                    txtFunctionary.Text = objBudgetCentre.FunctionaryName
                    txtFunctionary.Tag = objBudgetCentre.FunctionaryID
                    txtField.Text = objBudgetCentre.FieldName
                    txtFund.Text = objBudgetCentre.FundName
                    txtField.Tag = objBudgetCentre.FieldID
                    txtFinancialYear.Tag = objBudgetCentre.FinancialYearID
                    If txtFinancialYear.Tag <> 0 Then
                        Call DispFinancialYear(Val(txtFinancialYear.Tag))
                    End If
                    
                    Call FillAccountHeads(objBudgetCentre.BudgetCentreID)
                Else
                   txtBudgetCentre.Text = ""
                   txtBudgetCentreCode.Text = ""
                End If
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
            txtBudgetCentreCode.Text = Trim(txtBudgetCentreCode)
            If Len(txtBudgetCentreCode) Then
                objBudgetCentre.SetBudgetCentre (txtBudgetCentreCode.Text)
                If objBudgetCentre.BudgetCentreID > -1 Then
                    txtBudgetCentreCode.Text = objBudgetCentre.BudgetCentreCode
                    txtBudgetCentre.Text = objBudgetCentre.BudgetCentre
                    txtFunction.Text = objBudgetCentre.FunctionName
                    txtFunction.Tag = objBudgetCentre.FunctionID
                    txtFunctionary.Text = objBudgetCentre.FunctionaryName
                    txtFunctionary.Tag = objBudgetCentre.FunctionaryID
                    txtField.Text = objBudgetCentre.FieldName
                    txtFund.Text = objBudgetCentre.FundName
                    txtField.Tag = objBudgetCentre.FieldID
                    txtFinancialYear.Tag = objBudgetCentre.FinancialYearID
                    If txtFinancialYear.Tag <> 0 Then
                        Call DispFinancialYear(Val(txtFinancialYear.Tag))
                    End If
                    
                    Call FillAccountHeads(objBudgetCentre.BudgetCentreID)
                Else
                   txtBudgetCentre.Text = ""
                   txtBudgetCentreCode.Text = ""
                End If
            End If
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
                    Call Calculate
                End If
            End If
        End Sub

        Private Sub vsGrid_Click()
             If vsGrid.Row = 0 Then
                vsGrid.Sort = flexSortStringAscending
             End If
        End Sub
        Private Sub cmdFields_Click()
                Dim mSQL As String
                mSQL = "Select vchField, intFieldID From faFields Order By vchField"
                Call PopulateList(lstMasters, mSQL, , True, , True)
                lstMasters.Tag = "3"
                lstMasters.Visible = True
                lstMasters.SetFocus
        End Sub

        Private Sub cmdFunctionaries_Click()
                Dim mSQL As String
                mSQL = "Select vchFunctionary, intFunctionaryID From faFunctionaries Order By vchFunctionary"
                Call PopulateList(lstMasters, mSQL, , True, , True)
                lstMasters.Tag = "2"
                lstMasters.Visible = True
                lstMasters.SetFocus
        End Sub

        Private Sub cmdFunctions_Click()
                Dim mSQL As String
                mSQL = "Select vchFunction, intFunctionID From faFunctions Order By vchFunction"
                Call PopulateList(lstMasters, mSQL, , True, , True)
                lstMasters.Tag = "1"
                lstMasters.Visible = True
                lstMasters.SetFocus
        End Sub

        Private Sub cmdFunds_Click()
                Dim mSQL As String
                mSQL = "Select vchFund, intFundID From faFunds Where tnyActiveFlag = 1 Order By vchFund"
                Call PopulateList(lstMasters, mSQL, , True, , True)
                lstMasters.Tag = "4"
                lstMasters.Visible = True
                lstMasters.SetFocus
        End Sub
        Private Sub txtField_GotFocus()
            If gbSearchStr <> "" Then
                txtField.Text = gbSearchStr
                txtField.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchID = -1
            End If
        End Sub
        
        Private Sub txtField_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
                Call cmdFields_Click
            End If
        End Sub
        
        Private Sub txtField_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then PressTabKey
        End Sub
        
        Private Sub txtFunction_GotFocus()
            If gbSearchStr <> "" Then
                txtFunction.Text = gbSearchStr
                txtFunction.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchID = -1
            End If
        End Sub
        
        Private Sub txtFunction_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
                Call cmdFunctions_Click
            End If
        End Sub
        
        Private Sub txtFunction_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then PressTabKey
        End Sub
        
        Private Sub txtFunctionary_GotFocus()
            If gbSearchStr <> "" Then
                txtFunctionary.Text = gbSearchStr
                txtFunctionary.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchID = -1
            End If
        End Sub
        
        Private Sub txtFunctionary_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
                Call cmdFunctionaries_Click
            End If
        End Sub
        
        Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then PressTabKey
        End Sub
        
        Private Sub txtFund_GotFocus()
            If gbSearchStr <> "" Then
                txtFund.Text = gbSearchStr
                txtFund.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchID = -1
            End If
        End Sub
        
        Private Sub txtFund_KeyDown(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
                Call cmdFunds_Click
            End If
        End Sub
        
        Private Sub txtFund_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then PressTabKey
        End Sub
        
        Private Sub DispFinancialYear(mFinancialYearID)
                Dim mSQL As String
                Dim Rec As New ADODB.Recordset
                
                mSQL = "SELECT dtStartingDate, dtEndingDate, intFinancialYearID,tinCurrentFinancialYearFlag From faFinancialYear Where faFinancialYear.intFinancialYearID=" & mFinancialYearID
                Set Rec = GetRecordSet(mSQL)
                
                If Not (Rec.BOF And Rec.EOF) Then
                      txtFinancialYear.Text = Format(Rec!dtStartingDate, "Dd-Mmm-yyyy") & " -- " & Format(Rec!dtEndingDate, "Dd-Mmm-yyyy")
                        txtFinancialYear.Tag = Rec!intFinancialYearID
                End If
            End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmOpeningBalace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Balance"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemoveLeft 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   690
      TabIndex        =   14
      Top             =   5700
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cance&L"
      Height          =   375
      Left            =   5730
      TabIndex        =   9
      Top             =   6135
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4710
      TabIndex        =   8
      Top             =   6135
      Width           =   915
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
      ScaleWidth      =   11640
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---Date---"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9495
         TabIndex        =   16
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5385
      Left            =   0
      TabIndex        =   1
      Top             =   690
      Width           =   11295
      Begin VB.CommandButton cmdByHead 
         Caption         =   "..."
         Height          =   315
         Left            =   10740
         TabIndex        =   18
         Top             =   4560
         Width           =   345
      End
      Begin VB.TextBox txtByHead 
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   4590
         Width           =   10515
      End
      Begin VB.CommandButton cmdRemoveRight 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5700
         TabIndex        =   15
         Top             =   4965
         Width           =   1305
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "...."
         Height          =   345
         Left            =   150
         TabIndex        =   11
         Top             =   5040
         Width           =   435
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
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
         Left            =   8610
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   4995
         Width           =   2235
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   4995
         Width           =   2235
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridL 
         Height          =   3645
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   5445
         _cx             =   9604
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
         Rows            =   18
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOpeningBalace.frx":0000
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
      Begin VSFlex8LCtl.VSFlexGrid vsGridR 
         Height          =   3615
         Left            =   5670
         TabIndex        =   6
         Top             =   930
         Width           =   5445
         _cx             =   9604
         _cy             =   6376
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
         Rows            =   18
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOpeningBalace.frx":00A2
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
      Begin VB.Label lblDifferenceLiabilities 
         AutoSize        =   -1  'True
         Caption         =   "-"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1665
         TabIndex        =   13
         Top             =   4635
         Width           =   45
      End
      Begin VB.Label lblDifferenceAsset 
         AutoSize        =   -1  'True
         Caption         =   "-"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7185
         TabIndex        =   12
         Top             =   4635
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "OPENING BALANCE SHEET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   30
         TabIndex        =   10
         Top             =   120
         Width           =   11220
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "  Liabilities"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   630
         Width           =   5430
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Assets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   5670
         TabIndex        =   3
         Top             =   630
         Width           =   5445
      End
   End
End
Attribute VB_Name = "frmOpeningBalace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private mvsLRow As Long
    Private mvsRRow As Long
    Public mFlag As Variant

    Private Sub DeleteRows(fg As VSFlexGrid)
        Dim mLoop As Long
        mLoop = 1
        Do While (mLoop < fg.Rows)
            If fg.IsSelected(mLoop) Then
                fg.RemoveItem (mLoop)
            Else
                mLoop = mLoop + 1
            End If
        Loop
        
    End Sub
    
    Private Sub Display()
'        Dim objDB As New clsDB
'        Dim Rec As New ADODB.Recordset
'        Dim mCnn As New ADODB.Connection
'        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
'
'        Set Rec = GetRecordSet("Select * From faAccountHeads Where fltOpeningBalance > 0 AND tinType IN (3,4) Order By vchAccountHead")
'        If Not (Rec.BOF And Rec.EOF) Then
'            While Not Rec.EOF
'                If Rec!tinType = 3 Then
'                mvsLRow = mvsLRow + 1
'                vsGridL.AddItem Rec!vchAccountHead & vbTab & Format(Rec!fltOpeningBalance, "0.00") & vbTab & Rec!intAccountHeadID, mvsLRow
'                ElseIf Rec!tinType = 4 Then
'                mvsRRow = mvsRRow + 1
'                vsGridR.AddItem Rec!vchAccountHead & vbTab & Format(Rec!fltOpeningBalance, "0.00") & vbTab & Rec!intAccountHeadID, mvsRRow
'                End If
'            Rec.MoveNext
'            Wend
'        End If
'        Rec.Close
        
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSQL As String
        
        vsGridL.Rows = 1
        vsGridR.Rows = 1
        vsGridL.Rows = 2
        vsGridR.Rows = 2
        mvsLRow = 0
        mvsRRow = 0
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "Select faTransactions.intTransactionID,faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,fltAmount,tinType From faTransactionChild Inner Join faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                "Inner Join faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID " & _
                "Where intTransactionTypeID in (5000,6000) And intSerialNo <> 1 Order By faTransactions.intTransactionID,intSerialNo"
        Rec.Open mSQL, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                If Rec!tinType = 3 Then
                    mvsLRow = mvsLRow + 1
                    vsGridL.AddItem Rec!vchAccountHead & vbTab & Format(Rec!fltAmount, "0.00") & vbTab & Rec!intAccountHeadID, mvsLRow
                    vsGridL.Tag = Rec!intTransactionID
                ElseIf Rec!tinType = 4 Then
                    mvsRRow = mvsRRow + 1
                    vsGridR.AddItem Rec!vchAccountHead & vbTab & Format(Rec!fltAmount, "0.00") & vbTab & Rec!intAccountHeadID, mvsRRow
                    vsGridR.Tag = Rec!intTransactionID
                End If
            Rec.MoveNext
            Wend
        End If
        Rec.Close
    End Sub
    
    Private Sub Calculate()
        Dim mLoop As Long
        Dim mAmtLib As Double
        Dim mAmtAsst As Double
        Dim mAmtAsset As Double
        Dim mDifference As Double
        lblDifferenceAsset.Caption = ""
        lblDifferenceLiabilities.Caption = ""
        For mLoop = 1 To vsGridL.Rows - 1
            If vsGridL.TextMatrix(mLoop, 0) <> "" And Val(vsGridL.TextMatrix(mLoop, 1)) <> 0 Then
                mAmtLib = mAmtLib + Format(Val(vsGridL.TextMatrix(mLoop, 1)), "0.00")
            End If
        Next
        For mLoop = 1 To vsGridR.Rows - 1
            If vsGridR.TextMatrix(mLoop, 0) <> "" And Val(vsGridR.TextMatrix(mLoop, 1)) <> 0 Then
                mAmtAsst = mAmtAsst + Format(Val(vsGridR.TextMatrix(mLoop, 1)), "0.00")
            End If
        Next
        txtCredit.Text = Format(mAmtLib, "0.00")
        txtDebit.Text = Format(mAmtAsst, "0.00")
        mDifference = Val(txtCredit.Text) - Val(txtDebit.Text)
        If mDifference > 0 Then
            lblDifferenceLiabilities.Caption = Format(mDifference, "0.00")
        ElseIf mDifference < 0 Then
             lblDifferenceAsset.Caption = Format(mDifference, "0.00")
        End If
        
    End Sub
    

    Private Sub cmdByHead_Click()
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where intAccountHeadid=887"
        frmSearchAccountHeads.Show vbModal
           If gbSearchID <> -1 Then
               txtByHead.Text = gbSearchStr
               txtByHead.Tag = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub

    Private Sub cmdCancel_Click()
        On Error Resume Next
        vsGridL.Select 1, 0, mvsLRow, 0
        vsGridL.Sort = flexSortStringAscending
    End Sub
    
    Private Sub cmdRemoveLeft_Click()
     If vsGridL.TextMatrix(vsGridL.Row, 0) <> "" Then
        vsGridL.RemoveItem (vsGridL.Row)
     End If
    End Sub

    Private Sub cmdRemoveRight_Click()
    If vsGridR.TextMatrix(vsGridR.Row, 0) <> "" Then
        vsGridR.RemoveItem (vsGridR.Row)
    End If
    
    End Sub

    Private Sub cmdSave_Click()
        Dim mLoop       As Long
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim arrInput    As Variant
        Dim arrOut      As Variant
        Dim mSQL        As String
        Dim mRow        As Integer
        Dim Rec         As New ADODB.Recordset
        Dim mAmt        As Double
        Dim recAccountHeadID        As New ADODB.Recordset
        
        objDB.SetConnection mCnn
        For mLoop = 1 To vsGridL.Rows - 1
            If Val(vsGridL.TextMatrix(mLoop, 1)) <> 0 And Val(vsGridL.TextMatrix(mLoop, 2)) <> 0 Then
                arrInput = Array(Val(vsGridL.TextMatrix(mLoop, 2)), Val(vsGridL.TextMatrix(mLoop, 1)))
                objDB.ExecuteSP "spSaveAccountHeadOpening", arrInput, , , mCnn
            End If
        Next mLoop
        
        For mLoop = 1 To vsGridR.Rows - 1
            If Val(vsGridR.TextMatrix(mLoop, 1)) <> 0 And Val(vsGridR.TextMatrix(mLoop, 2)) <> 0 Then
               arrInput = Array(Val(vsGridR.TextMatrix(mLoop, 2)), Val(vsGridR.TextMatrix(mLoop, 1)))
                objDB.ExecuteSP "spSaveAccountHeadOpening", arrInput, , , mCnn
            End If
        Next mLoop
'        mSQL = "Select Count(*) From faTransactions Where intTransactionID = 0"
'                Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'                If Rec.Fields(0).Value = 0 Then
'                    Dim intTransactionID_1   As Double
'                    Dim mintLocalBodyID_2  As Long
'                    Dim mintFinancialYearID_3  As Long
'                    Dim mdtTransactionDate_4   As Date
'                    Dim mintExternalApplicationID_5    As Long
'                    Dim mintExternalApplicationModuleID_6  As Long
'                    Dim mintFunctionID_7   As Variant
'                    Dim mintFunctionaryID_8   As Variant
'                    Dim mintFieldID_9 As Variant
'                    Dim mintFundID_10 As Variant
'                    Dim mintBudgetCentreID_11  As Variant
'                    Dim mvchNarration_12   As String
'                    Dim mintTransactionTypeID_13   As Variant
'                    Dim mintVoucherNo_14   As Variant
'                    Dim mintProcessID_15    As Variant
'                    Dim mintGroupID_17    As Variant
'                    Dim mvchGroup_16   As String
'                    Dim mintKeyID_18   As Variant
'                    Dim mnumSubLedgerID_19    As Variant
'                    Dim mintUserID_20  As Variant
'
'                    intTransactionID_1 = 0
'                    mintLocalBodyID_2 = gbLocalBodyID
'                    mintFinancialYearID_3 = gbFinancialYearID
'                    mdtTransactionDate_4 = gbStartingDate
'                    mintExternalApplicationID_5 = AppID.Saankhya
'                    mintExternalApplicationModuleID_6 = 0
'                    mintFunctionID_7 = Null
'                    mintFunctionaryID_8 = Null
'                    mintFieldID_9 = Null
'                    mintFundID_10 = Null
'                    mintBudgetCentreID_11 = Null
'                    mvchNarration_12 = "Opening Balance"
'                    mintTransactionTypeID_13 = Null
'                    mintVoucherNo_14 = Null
'                    mintProcessID_15 = Null
'                    mvchGroup_16 = "JV"
'                    mintGroupID_17 = 40
'                    mintKeyID_18 = Null
'                    mnumSubLedgerID_19 = Null
'                    mintUserID_20 = 0
'
'                    arrInput = Array( _
'                                    intTransactionID_1, _
'                                    mintLocalBodyID_2, _
'                                    mintFinancialYearID_3, _
'                                    mdtTransactionDate_4, _
'                                    mintExternalApplicationID_5, _
'                                    mintExternalApplicationModuleID_6, _
'                                    mintFunctionID_7, _
'                                    mintFunctionaryID_8, _
'                                    mintFieldID_9, _
'                                    mintFundID_10, _
'                                    mintBudgetCentreID_11, _
'                                    mvchNarration_12, _
'                                    mintTransactionTypeID_13, _
'                                    mintProcessID_15, _
'                                    mvchGroup_16, _
'                                    mintGroupID_17, _
'                                    mintKeyID_18, _
'                                    mnumSubLedgerID_19, _
'                                    gbUserID, _
'                                    mintVoucherNo_14)
'
'                    objDB.ExecuteSP "spSaveTransactions", arrInput, arrOut, , mCnn
''                End If
'                If mAmt <> 0 Then
'                    If Not recAccountHeadID.EOF Then
'                            arrInput = Array(0, _
'                            2, _
'                            recAccountHeadID.Fields(0), _
'                            Format(mAmt, "0.00"), _
'                             1, _
'                            Null, _
'                            "Opening Balance", _
'                            Null _
'                            )
'                        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
'                    End If
'                Else
'                    mCnn.Execute "Delete From fatransactionchild Where intaccountheadid='" & IIf(IsNull(mFlag), 0, mFlag) & "'And inttransactionid=0"
'                End If
'                MsgBox "Successfully Saved"
'
'
'    End Sub

            Dim intTransactionID_1   As Double
            Dim mintLocalBodyID_2  As Long
            Dim mintFinancialYearID_3  As Long
            Dim mdtTransactionDate_4   As Date
            Dim mintExternalApplicationID_5    As Long
            Dim mintExternalApplicationModuleID_6  As Long
            Dim mintFunctionID_7   As Variant
            Dim mintFunctionaryID_8   As Variant
            Dim mintFieldID_9 As Variant
            Dim mintFundID_10 As Variant
            Dim mintBudgetCentreID_11  As Variant
            Dim mvchNarration_12   As String
            Dim mintTransactionTypeID_13   As Variant
            Dim mintVoucherNo_14   As Variant
            Dim mintProcessID_15    As Variant
            Dim mintGroupID_17    As Variant
            Dim mvchGroup_16   As String
            Dim mintKeyID_18   As Variant
            Dim mnumSubLedgerID_19    As Variant
            Dim mintUserID_20  As Variant
            
            If Val(txtCredit.Text) <> Val(txtDebit.Text) Then
                MsgBox "Incorrect information" & vbNewLine & "Assets and Liabilities must be Equal", vbInformation
                Exit Sub
            End If
            If txtByHead.Text = "" Then
                MsgBox "Please Select AccountHead", vbInformation
                cmdByHead.SetFocus
                Exit Sub
            End If
        '--------------------------------------------------------------------------------------------------'
        '                                           Liability JV                                           '
        '--------------------------------------------------------------------------------------------------'
        intTransactionID_1 = IIf(Val(vsGridL.Tag) = 0, -1, vsGridL.Tag)
        mintLocalBodyID_2 = gbLocalBodyID
        mintFinancialYearID_3 = gbFinancialYearID
        mdtTransactionDate_4 = lblDate.Caption
        mintExternalApplicationID_5 = AppID.Saankhya
        mintExternalApplicationModuleID_6 = 0
        mintFunctionID_7 = Null
        mintFunctionaryID_8 = Null
        mintFieldID_9 = Null
        mintFundID_10 = Null
        mintBudgetCentreID_11 = Null
        mvchNarration_12 = "Opening Balance"
        mintTransactionTypeID_13 = 5000                     ' For the Liability JV TransactionTypeID = 5000
        mintVoucherNo_14 = Null
        mintProcessID_15 = Null
        mvchGroup_16 = "JV"
        mintGroupID_17 = 40
        mintKeyID_18 = Null
        mnumSubLedgerID_19 = Null
        mintUserID_20 = 0
        
        arrInput = Array( _
        intTransactionID_1, _
        mintLocalBodyID_2, _
        mintFinancialYearID_3, _
        mdtTransactionDate_4, _
        mintExternalApplicationID_5, _
        mintExternalApplicationModuleID_6, _
        mintFunctionID_7, _
        mintFunctionaryID_8, _
        mintFieldID_9, _
        mintFundID_10, _
        mintBudgetCentreID_11, _
        mvchNarration_12, _
        mintTransactionTypeID_13, _
        mintProcessID_15, _
        mvchGroup_16, _
        mintGroupID_17, _
        mintKeyID_18, _
        mnumSubLedgerID_19, _
        gbUserID, _
        mintVoucherNo_14)
        
        objDB.ExecuteSP "spSaveTransactions", arrInput, arrOut, , mCnn
        arrInput = Array(arrOut(0, 0), _
                            1, _
                            Val(txtByHead.Tag), _
                            Val(txtCredit.Text), _
                            1, _
                            Null, _
                            "Opening Balance", _
                            Null _
                            )
        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
        For mRow = 1 To vsGridL.Rows - 1
            If vsGridL.TextMatrix(mRow, 1) <> "" And vsGridL.TextMatrix(mRow, 4) <> "" Then
                arrInput = Array(arrOut(0, 0), _
                                mRow + 1, _
                                Val(vsGridL.TextMatrix(mRow, 4)), _
                                Val(vsGridL.TextMatrix(mRow, 1)), _
                                0, _
                                Val(txtByHead.Tag), _
                                "Opening Balance", _
                                Null _
                                )
                objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End If
        Next mRow
        '-----------------------------------------------------------------------------------
        ' Credit Head in Asset
        For mRow = 1 To vsGridR.Rows - 1
            If vsGridR.TextMatrix(mRow, 4) <> "" And vsGridR.TextMatrix(mRow, 1) < 0 Then
                arrInput = Array(arrOut(0, 0), _
                                mRow + 1, _
                                Val(vsGridR.TextMatrix(mRow, 4)), _
                                Val(vsGridR.TextMatrix(mRow, 1)), _
                                1, _
                                Val(txtByHead.Tag), _
                                "Opening Balance", _
                                Null _
                                )
                objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End If
        Next mRow
        '-----------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------'
        '                                           Asset JV                                               '
        '--------------------------------------------------------------------------------------------------'
        intTransactionID_1 = IIf(Val(vsGridR.Tag) = 0, -1, vsGridR.Tag)
        mintLocalBodyID_2 = gbLocalBodyID
        mintFinancialYearID_3 = gbFinancialYearID
        mdtTransactionDate_4 = lblDate.Caption
        mintExternalApplicationID_5 = AppID.Saankhya
        mintExternalApplicationModuleID_6 = 0
        mintFunctionID_7 = Null
        mintFunctionaryID_8 = Null
        mintFieldID_9 = Null
        mintFundID_10 = Null
        mintBudgetCentreID_11 = Null
        mvchNarration_12 = "Opening Balance"
        mintTransactionTypeID_13 = 6000                     ' For the Asset JV TransactionTypeID = 6000
        mintVoucherNo_14 = Null
        mintProcessID_15 = Null
        mvchGroup_16 = "JV"
        mintGroupID_17 = 40
        mintKeyID_18 = Null
        mnumSubLedgerID_19 = Null
        mintUserID_20 = 0
        
        arrInput = Array( _
        intTransactionID_1, _
        mintLocalBodyID_2, _
        mintFinancialYearID_3, _
        mdtTransactionDate_4, _
        mintExternalApplicationID_5, _
        mintExternalApplicationModuleID_6, _
        mintFunctionID_7, _
        mintFunctionaryID_8, _
        mintFieldID_9, _
        mintFundID_10, _
        mintBudgetCentreID_11, _
        mvchNarration_12, _
        mintTransactionTypeID_13, _
        mintProcessID_15, _
        mvchGroup_16, _
        mintGroupID_17, _
        mintKeyID_18, _
        mnumSubLedgerID_19, _
        gbUserID, _
        mintVoucherNo_14)
        
        objDB.ExecuteSP "spSaveTransactions", arrInput, arrOut, , mCnn
        arrInput = Array(arrOut(0, 0), _
                            1, _
                            Val(txtByHead.Tag), _
                            Val(txtCredit.Text), _
                            0, _
                            Null, _
                            "Opening Balance", _
                            Null _
                            )
        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
        For mRow = 1 To vsGridR.Rows - 1
            If vsGridR.TextMatrix(mRow, 4) <> "" And vsGridR.TextMatrix(mRow, 1) > 0 Then
                arrInput = Array(arrOut(0, 0), _
                                mRow + 1, _
                                Val(vsGridR.TextMatrix(mRow, 4)), _
                                Val(vsGridR.TextMatrix(mRow, 1)), _
                                1, _
                                Val(txtByHead.Tag), _
                                "Opening Balance", _
                                Null _
                                )
                objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            End If
        Next mRow
    End Sub
    
    Private Sub cmdSearch_Click()
        frmSelectAccountHeads.Show vbModal
        frmSelectAccountHeads.ZOrder (0)
    End Sub
        
    Private Sub Form_Load()
        Dim objDB As New clsDB
        Dim mSQL As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSQL = "Select Min(dtTransactionDate)-1 [dtDate] From faTransactions"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            lblDate.Caption = Format(Rec!dtDate, "dd-MMM-yyyy")
        Else
            lblDate.Caption = gbStartingDate - 1
        End If
        
        Me.Top = 150
        Me.Left = (frmMenu.Width - Me.Width) / 2
        mvsLRow = 0
        mvsRRow = 0
        Call Display
        Call Calculate
        vsGridL.ColComboList(0) = "|..."
        vsGridR.ColComboList(0) = "|..."
        cmdSearch.Visible = False
        
    End Sub

    Private Sub vsGridL_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If vsGridL.TextMatrix(Row, 0) <> "" And vsGridL.TextMatrix(Row, 1) <> "" Then
            vsGridL.Rows = vsGridL.Rows + 1
        End If
        Call Calculate
    End Sub
            


    Private Sub vsGridL_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And tinType = 3"
        frmSearchAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            vsGridL.TextMatrix(vsGridL.Row, 0) = gbSearchStr
            vsGridL.TextMatrix(vsGridL.Row, 4) = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
            'vsGridL.Rows = vsGridL.Rows + 1
        End If
    End Sub

    Private Sub vsGridL_CellChanged(ByVal Row As Long, ByVal Col As Long)
        If Col = 1 Then
            vsGridL.TextMatrix(Row, Col) = Format(Val(vsGridL.TextMatrix(Row, 1)), "0.00")
        End If
        
    End Sub
    
    Private Sub vsGridL_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            Call DeleteRows(vsGridL)
        End If
    End Sub

    Private Sub vsGridR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If vsGridR.TextMatrix(Row, 0) <> "" And vsGridR.TextMatrix(Row, 1) <> "" Then
            vsGridR.Rows = vsGridR.Rows + 1
        End If
        Call Calculate
    End Sub
    
    Private Sub vsGridR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select ( vchAccountHeadCode + '  ' + vchAccountHead) as vchAccountHeadCode, intAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And tinType = 4"
        frmSearchAccountHeads.Show vbModal
           If gbSearchID <> -1 Then
               vsGridR.TextMatrix(vsGridR.Row, 0) = gbSearchStr
               vsGridR.TextMatrix(vsGridR.Row, 4) = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub

    Private Sub vsGridR_CellChanged(ByVal Row As Long, ByVal Col As Long)
        If Col = 1 Then
            vsGridR.TextMatrix(Row, 1) = Format(Val(vsGridR.TextMatrix(Row, 1)), "0.00")
        End If
    End Sub
    
    Private Sub vsGridR_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyDelete Then
            Call DeleteRows(vsGridR)
        End If
    End Sub
    
    Public Property Let LRows(mVal As Long)
        mvsLRow = mVal
    End Property
    
    Public Property Get LRows() As Long
        LRows = mvsLRow
    End Property
    
    Public Property Let RRows(mVal As Long)
        mvsRRow = mVal
    End Property
    
    Public Property Get RRows() As Long
        RRows = mvsRRow
    End Property


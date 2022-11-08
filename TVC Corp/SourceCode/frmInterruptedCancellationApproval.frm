VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmInterruptedCancellationApproval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interrupted Receipt Cancellation List For Approval"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2550
      Left            =   30
      TabIndex        =   0
      Top             =   255
      Width           =   7980
      _cx             =   14076
      _cy             =   4498
      Appearance      =   2
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInterruptedCancellationApproval.frx":0000
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
End
Attribute VB_Name = "frmInterruptedCancellationApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Function AutoWordWrap(vs As VSFlexGrid)
        With vs
            .AutoSizeMode = flexAutoSizeRowHeight
            .WordWrap = True
            .AutoSize 0, .Cols - 1
            .Cell(flexcpAlignment, 1, 5, .Rows - 1, .Cols - 1) = 0
        End With
    End Function


    Private Sub FillvsGrid(Rec As ADODB.Recordset)
        Dim mRowCount   As Double
        Dim mStatus     As Variant
        
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        While Not Rec.EOF
            vsGrid.AddItem ""
            vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intSerialNo), "", Rec!intSerialNo)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!dtReceiptDate), "", Rec!dtReceiptDate)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!Remarks), "", Rec!Remarks)
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            If mStatus <> "" Then
                If mStatus = 1 Then
                    vsGrid.Cell(flexcpChecked, mRowCount, 6) = True
                End If
                If mStatus = 0 Then
                    vsGrid.Cell(flexcpChecked, mRowCount, 6) = False
                End If
            End If
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intCounterID), "", Rec!intCounterID)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!intBookID), "", Rec!intBookID)
            Rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
'''    Private Sub cmdReject_Click()
'''        Dim mRowCount   As Integer
'''        For mRowCount = 1 To vsGrid.Rows - 1
'''            If vsGrid.Cell(flexcpChecked, mRowCount, 6) = 1 Then
'''                frmReject.Mode = 4
'''                frmReject.RequestTypeID = vsGrid.TextMatrix(mRowCount, 3)
'''                frmReject.Show vbModal
'''                cmdReject.Enabled = False
'''                cmdSave.Enabled = False
'''            End If
'''        Next
'''    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mArray      As Variant
        Dim mRowCount   As Double
        Dim Rec         As New ADODB.Recordset
        Dim mSQL        As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '*********************************************************************************************'
        '               Procedure to approve the Interrupt Receipt Cancellation                       '
        '*********************************************************************************************'
        
        For mRowCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mRowCount, 6) = 1 Then
                mArray = Array(vsGrid.TextMatrix(mRowCount, 9), _
                           vsGrid.TextMatrix(mRowCount, 3), _
                           gbUserID _
                          )
                objdb.ExecuteSP "spApproveInterruptedReceiptCancellation", mArray, , , mCnn, adCmdStoredProc
            End If
        Next
        MsgBox "Successfully Saved", vbInformation
        mSQL = "Select *,faInterruptedCancelledReceipts.vchRemarks[Remarks] From faInterruptedCancelledReceipts"
        mSQL = mSQL + " Inner Join faUser On faInterruptedCancelledReceipts.numUserID = faUser.numUserID"
        mSQL = mSQL + " Inner Join faInterruptedReceiptBooks On faInterruptedCancelledReceipts.intBookID = faInterruptedReceiptBooks.intBookID"
        mSQL = mSQL + " Inner Join faCounters On faInterruptedReceiptBooks.intCounterID = faCounters.intCounterID"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
            AutoWordWrap vsGrid
        End If
        Rec.Close
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.Width = 8175
        Me.Height = 4020
    End Sub

    Private Sub Form_Load()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQL = "Select *,faInterruptedCancelledReceipts.vchRemarks[Remarks] From faInterruptedCancelledReceipts"
        mSQL = mSQL + " Inner Join faUser On faInterruptedCancelledReceipts.numUserID = faUser.numUserID"
        mSQL = mSQL + " Inner Join faInterruptedReceiptBooks On faInterruptedCancelledReceipts.intBookID = faInterruptedReceiptBooks.intBookID"
        mSQL = mSQL + " Inner Join faCounters On faInterruptedReceiptBooks.intCounterID = faCounters.intCounterID"
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Call FillvsGrid(Rec)
            AutoWordWrap vsGrid
        End If
        Rec.Close
    End Sub


    Private Sub VSGrid_Click()
        If vsGrid.col = 6 Then
            vsGrid.Editable = flexEDKbdMouse
        Else
            vsGrid.Editable = flexEDNone
        End If
    End Sub

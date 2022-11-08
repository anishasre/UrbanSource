VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDeductionExcludePayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Excludes Payment"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _cx             =   15478
      _cy             =   3731
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      BackColorFixed  =   13559526
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   3
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDeductionExcludePayment.frx":0000
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
End
Attribute VB_Name = "frmDeductionExcludePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mRow As Variant
Private Sub cmdSave_Click()
    Dim mLoop As Integer
    Dim mPLoop As Integer
     'For mPLoop = 1 To frmPaymentOrder.vsGrid.Rows - 1
        For mLoop = 1 To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, mLoop, 6) = 1 Then
            frmPaymentOrder.vsGrid.TextMatrix(mLoop, 6) = 1
            frmPaymentOrder.vsGrid.Cell(flexcpForeColor, mLoop, 1, , 6) = &H8000000C
            Else
            frmPaymentOrder.vsGrid.TextMatrix(mLoop, 6) = 0
            frmPaymentOrder.vsGrid.Cell(flexcpForeColor, mLoop, 1, , 6) = &H0&
        End If
        Next mLoop
        'frmPaymentOrder.vsGrid.TextMatrix(mPLoop, 6) = 1
    'Next mPLoop
    Unload Me
End Sub
Private Sub Form_Load()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mLoop As Integer
                mRow = 0
                vsGrid.Rows = 1
                While (frmPaymentOrder.vsGrid.TextMatrix(mRow + 1, 1) <> "")
                   vsGrid.Rows = vsGrid.Rows + 1
                   mRow = vsGrid.Rows - 1

                   vsGrid.TextMatrix(mRow, 1) = frmPaymentOrder.vsGrid.TextMatrix(mRow, 1)
                   vsGrid.TextMatrix(mRow, 2) = frmPaymentOrder.vsGrid.TextMatrix(mRow, 2)
                   vsGrid.TextMatrix(mRow, 3) = frmPaymentOrder.vsGrid.TextMatrix(mRow, 3)
                   vsGrid.TextMatrix(mRow, 4) = frmPaymentOrder.vsGrid.TextMatrix(mRow, 4)
                   vsGrid.TextMatrix(mRow, 6) = frmPaymentOrder.vsGrid.TextMatrix(mRow, 6)

                If val(frmPaymentOrder.vsGrid.TextMatrix(mRow, 1)) <> gbAcHeadDeductionExcludeProfTax Then  'Profession Tax
                If val(frmPaymentOrder.vsGrid.TextMatrix(mRow, 1)) > gbAcHeadDeductionStart And val(frmPaymentOrder.vsGrid.TextMatrix(mRow, 1)) < gbAcHeadDeductionEnd Then
                   vsGrid.Cell(flexcpChecked, mRow, 6) = False
                   vsGrid.TextMatrix(mRow, 7) = 1
                   vsGrid.Cell(flexcpForeColor, mRow, 1, , 6) = &H8000000C
                Else
                   vsGrid.TextMatrix(mRow, 7) = 0
                End If
                End If
              Wend
              'vsGrid.ColSort(2) = flexSortGenericAscending
              'Call Display
  End Sub
Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If val(vsGrid.TextMatrix(Row, 7)) = 1 Then
        Cancel = True
    End If
End Sub
Private Sub Display()
Dim i As Variant
Dim j As Variant
For i = 0 To frmPaymentOrder.vsGrid.Rows - 1
For j = 1 To frmPaymentOrder.vsGrid.Cols
vsGrid.TextMatrix(i, j - 1) = frmPaymentOrder.vsGrid.TextMatrix(i, j - 1)
       If val(frmPaymentOrder.vsGrid.TextMatrix(i, 1)) <> gbAcHeadDeductionExcludeProfTax Then  'Profession Tax
        If val(frmPaymentOrder.vsGrid.TextMatrix(i, 1)) > gbAcHeadDeductionStart And val(frmPaymentOrder.vsGrid.TextMatrix(i, 1)) < gbAcHeadDeductionEnd Then
          vsGrid.Cell(flexcpChecked, i, 6) = False
          vsGrid.TextMatrix(i, 7) = 1
          vsGrid.Cell(flexcpForeColor, i, 1, , 6) = &H8000000C
          Else
          vsGrid.TextMatrix(i, 7) = 0
          End If
          End If
Next j
Next i
End Sub



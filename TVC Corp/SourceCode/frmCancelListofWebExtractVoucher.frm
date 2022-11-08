VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmCancelListofWebExtractVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmCancelListofWebExtractVoucher"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14595
      _cx             =   25744
      _cy             =   11033
      Appearance      =   1
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCancelListofWebExtractVoucher.frx":0000
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
Attribute VB_Name = "frmCancelListofWebExtractVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
 Call FillGrid

End Sub
Private Sub FillGrid()
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim RecSub      As New ADODB.Recordset
        Dim mRowCnt     As Integer
        
        If objdb.SetConnection(mCnn) Then
            
                mSql = " SELECT faVouchers.intVoucherNo,faVouchers.dtDate ,isnull(faReverseEntry.tnystatus,0) stat,* from faReverseEntry Inner Join faWebExtracts "
                mSql = mSql + " On  faWebExtracts.intwebExtractID=faReverseEntry.numDemandID Inner Join faVouchers  On  "
                mSql = mSql + " faVouchers.intVoucherID=faWebExtracts.numKeyID Where faReverseEntry.intCategoryID = 80"
               
                Rec.CursorLocation = adUseClient
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                mRowCnt = 1
            
                vsGrid.Clear 1, 1
                vsGrid.Rows = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!dtRequestDate), "", CheckDateInMMM(Rec!dtRequestDate))
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!numbillcontrolcode), "", Rec!numbillcontrolcode)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
                    If (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 10 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "R"
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 20 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "P"
                    ElseIf (IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)) = 40 Then
                        vsGrid.TextMatrix(mRowCnt, 4) = "JV"
                    End If
                    vsGrid.TextMatrix(mRowCnt, 5) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsGrid.TextMatrix(mRowCnt, 6) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    
'                    vsGrid.TextMatrix(mRowCnt, 7) = IIf(IsNull(Rec!VrDate), "", Format(Rec!VrDate, "DD-MMM-YYYY"))
'                    vsGrid.TextMatrix(mRowCnt, 8) = IIf(IsNull(Rec!intWebExtractID), "", Rec!intWebExtractID)
'                    vsGrid.TextMatrix(mRowCnt, 9) = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)

                    If (IIf(IsNull(Rec!Stat), "", Rec!Stat)) = 0 Then
                        vsGrid.TextMatrix(mRowCnt, 10) = "Requested for Reverse"
                    ElseIf (IIf(IsNull(Rec!Stat), "", Rec!Stat)) = 1 Then
                        vsGrid.TextMatrix(mRowCnt, 10) = "Request Verified"
                    ElseIf (IIf(IsNull(Rec!Stat), "", Rec!Stat)) = 2 Then
                        vsGrid.TextMatrix(mRowCnt, 10) = "Request Approved"
                    End If
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                Wend
                Rec.Close
                
       End If
    End Sub

Private Sub vsGrid_DblClick()

     If vsGrid.TextMatrix(vsGrid.Row, 1) > 1 Then
        frmCancelWebExtractVoucher.DispayWebExtractVoucherCancelLisstdetails (vsGrid.TextMatrix(vsGrid.Row, 1))
        frmCancelWebExtractVoucher.Show vbModal
       
    Else
        MsgBox "Please Select Voucher Generated E - bill", vbInformation
    End If

End Sub

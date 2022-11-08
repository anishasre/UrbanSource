VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmViewHeadWiseConsolidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Collection - Head Wise Consolidation - "
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInit 
      Caption         =   "Initialize"
      Height          =   285
      Left            =   2445
      TabIndex        =   3
      Top             =   5790
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   4245
      TabIndex        =   4
      Top             =   5655
      Width           =   1380
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5895
      Top             =   5610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   885
      TabIndex        =   0
      Top             =   255
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483633
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   19988483
      CurrentDate     =   40397
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4740
      Left            =   330
      TabIndex        =   2
      Top             =   795
      Width           =   9930
      _cx             =   17515
      _cy             =   8361
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
      FormatString    =   ""
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   255
      TabIndex        =   1
      Top             =   390
      Width           =   345
   End
End
Attribute VB_Name = "frmViewHeadWiseConsolidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
    CommonDialog.FileName = ""
    Call CommonDialog.ShowSave
    'CommonDialog.ShowOpen
'    CommonDialog.Action = 1
    If Len(CommonDialog.FileName) = 0 Then Exit Sub
    
    MousePointer = MousePointerConstants.vbHourglass
    vsGrid.SaveGrid CommonDialog.FileName, flexFileExcel
    MousePointer = MousePointerConstants.vbDefault
End Sub



'''Private Sub cmdInit_Click()
'''    With vsGrid
'''        If .Rows = 0 Then .Rows = 1
'''        .Cols = 5
'''        .TextMatrix(0, 0) = "Head Code"
'''        .TextMatrix(0, 1) = "Head"
'''        .TextMatrix(0, 2) = "Cash"
'''        .TextMatrix(0, 3) = "Bank"
'''        .TextMatrix(0, 4) = "Total"
'''
'''        .ColWidth(0) = .Width * 10 / 100
'''        .ColWidth(1) = .Width * 37 / 100
'''        .ColWidth(2) = .Width * 15 / 100
'''        .ColWidth(3) = .Width * 15 / 100
'''        .ColWidth(4) = .Width * 20 / 100
'''
'''        Dim objDB As New clsDb
'''        Dim mCnn As New ADODB.Connection
'''        Dim Rec As New ADODB.Recordset
'''        Dim arrInput As Variant
'''        objDB.SetConnection mCnn
'''
'''        arrInput = Array(DdMmmYy(dtpDate.Value))
'''        Rec.CursorLocation = adUseClient
'''
'''        'Set Rec = objDB.ExecuteSP("spDailyHeadWiseConsolidatedCollection", ArrInput, , , , adCmdStoredProc)
'''        Rec.Open "spDailyHeadWiseConsolidatedCollection '" & DdMmmYy(dtpDate.Value) & "'", mCnn, adOpenDynamic
'''        If Not (Rec.BOF And Rec.EOF) Then
'''            Dim mSQL As String
'''            vsGrid.Rows = Rec.RecordCount + 1
'''            vsGrid.Col = 0
'''            vsGrid.Row = 1
'''            vsGrid.ColSel = Rec.Fields.count - 1
'''            vsGrid.RowSel = vsGrid.Rows - 1
'''            mSQL = Rec.GetString(, , vbTab, Chr(13))
'''            vsGrid.Clip = mSQL
'''        End If
'''        Rec.Close
'''
'''    End With
'''End Sub

Private Sub dtpDate_Click()
    Call initFn
End Sub

'    Private Sub dtpDate_CloseUp()
'        dtpDate.Value = CheckDateInMMM(dtpDate.Value)
'    End Sub
Private Sub Form_Load()
    CommonDialog.Filter = "Microsoft Excel Workbooks(*.xls)|*.xls"
    CommonDialog.DefaultExt = "xls"
    dtpDate.Value = gbTransactionDate
End Sub

Private Sub initFn()
    With vsGrid
        If .Rows = 0 Then .Rows = 1
        .Cols = 5
        .TextMatrix(0, 0) = "Head Code"
        .TextMatrix(0, 1) = "Head"
        .TextMatrix(0, 2) = "Cash"
        .TextMatrix(0, 3) = "Bank"
        .TextMatrix(0, 4) = "Total"
        
        .ColWidth(0) = .Width * 10 / 100
        .ColWidth(1) = .Width * 37 / 100
        .ColWidth(2) = .Width * 15 / 100
        .ColWidth(3) = .Width * 15 / 100
        .ColWidth(4) = .Width * 20 / 100
        
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        objDB.SetConnection mCnn
        
        arrInput = Array(DdMmmYy(dtpDate.Value))
        Rec.CursorLocation = adUseClient
        
        'Set Rec = objDB.ExecuteSP("spDailyHeadWiseConsolidatedCollection", ArrInput, , , , adCmdStoredProc)
        Rec.Open "spDailyHeadWiseConsolidatedCollection '" & DdMmmYy(dtpDate.Value) & "'", mCnn, adOpenDynamic
       
        vsGrid.Clear 1, 1
        If Not (Rec.BOF And Rec.EOF) Then
            Dim mSQL As String
            vsGrid.Rows = Rec.RecordCount + 1
            vsGrid.Col = 0
            vsGrid.Row = 1
            vsGrid.ColSel = Rec.Fields.count - 1
            vsGrid.RowSel = vsGrid.Rows - 1
            mSQL = Rec.GetString(, , vbTab, Chr(13))
            vsGrid.Clip = mSQL
        End If
        Rec.Close
        
    End With
End Sub

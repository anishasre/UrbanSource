VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmAFSClosingBalanceSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Closing Balance Sheet"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   Icon            =   "frmAFSClosingBalanceSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4905
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   12150
      _cx             =   21431
      _cy             =   8652
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
      BackColorBkg    =   16777215
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
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAFSClosingBalanceSheet.frx":1CCA
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
Attribute VB_Name = "frmAFSClosingBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub Form_Load()
        vsGrid.MergeCells = flexMergeFree
        vsGrid.MergeRow(0) = True
        Call FillGrid
    End Sub
    
    Private Sub FillGrid()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb  As New clsDB
        Dim arrIn   As Variant
        Dim mACnt    As Integer
        Dim mLCnt    As Integer
        vsGrid.Clear 2, 1
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            arrIn = Array(gbTransactionDate, 1)
            Set Rec = objDb.ExecuteSP("SpRptBalanceSheet", arrIn, , , mCnn, adCmdStoredProc)
            mACnt = 2
            mLCnt = 2
            vsGrid.Rows = 3
            While Not (Rec.EOF)
                If (Rec!AccountHeadCode = "ASSETS") Then
                    vsGrid.TextMatrix(mACnt, 4) = Rec!vchMajorAccountHeadCode
                    vsGrid.TextMatrix(mACnt, 5) = Rec!Accounts
                    vsGrid.TextMatrix(mACnt, 6) = Rec!vchScheduleTitle
                    vsGrid.TextMatrix(mACnt, 7) = Rec(7)
                    mACnt = mACnt + 1
                Else
                   vsGrid.TextMatrix(mLCnt, 0) = Rec!vchMajorAccountHeadCode
                   vsGrid.TextMatrix(mLCnt, 1) = Rec!Accounts
                   vsGrid.TextMatrix(mLCnt, 2) = Rec!vchScheduleTitle
                   vsGrid.TextMatrix(mLCnt, 3) = Rec(7)
                   mLCnt = mLCnt + 1
                End If
                vsGrid.Rows = vsGrid.Rows + 1
                Rec.MoveNext
                
            Wend
        End If
    End Sub

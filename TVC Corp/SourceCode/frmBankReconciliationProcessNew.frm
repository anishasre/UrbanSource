VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmBankReconciliationProcessNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmBankReconciliationProcessNew"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   16485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReconcile 
      Caption         =   "Reconcile"
      Height          =   7275
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Width           =   16335
      Begin VB.TextBox txtBankDate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9990
         TabIndex        =   26
         Top             =   6540
         Width           =   1485
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   1140
         TabIndex        =   21
         Top             =   6810
         Width           =   2115
      End
      Begin VB.TextBox txtInstrument 
         Height          =   285
         Left            =   1140
         TabIndex        =   19
         Top             =   6450
         Width           =   2115
      End
      Begin VSFlex8LCtl.VSFlexGrid vsVoucher 
         Height          =   3135
         Left            =   60
         TabIndex        =   14
         Top             =   120
         Width           =   13710
         _cx             =   24183
         _cy             =   5530
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBankReconciliationProcessNew.frx":0000
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
      Begin VSFlex8LCtl.VSFlexGrid vsBank 
         Height          =   3225
         Left            =   60
         TabIndex        =   15
         Top             =   3270
         Width           =   13710
         _cx             =   24183
         _cy             =   5689
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBankReconciliationProcessNew.frx":0101
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
      Begin VB.Label lblBank 
         Caption         =   "#"
         Height          =   315
         Left            =   13860
         TabIndex        =   29
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lblVr 
         Caption         =   "#"
         Height          =   315
         Left            =   13860
         TabIndex        =   28
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Bank Entry date"
         Height          =   225
         Left            =   8700
         TabIndex        =   27
         Top             =   6600
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   6810
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Instrument"
         Height          =   315
         Left            =   210
         TabIndex        =   20
         Top             =   6480
         Width           =   915
      End
   End
   Begin VB.Frame fraMutualBank 
      Caption         =   "MutualBank"
      Height          =   6825
      Left            =   210
      TabIndex        =   32
      Top             =   1410
      Width           =   15945
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   1020
         TabIndex        =   35
         Top             =   5310
         Width           =   2115
      End
      Begin VB.TextBox txtBankAmount 
         Height          =   285
         Left            =   1050
         TabIndex        =   34
         Top             =   5700
         Width           =   2115
      End
      Begin VSFlex8LCtl.VSFlexGrid vsbankMutual 
         Height          =   4815
         Left            =   330
         TabIndex        =   33
         Top             =   300
         Width           =   14520
         _cx             =   25612
         _cy             =   8493
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBankReconciliationProcessNew.frx":022E
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
      Begin VB.Label lblDeposit 
         Caption         =   "#"
         Height          =   315
         Left            =   5610
         TabIndex        =   39
         Top             =   5700
         Width           =   885
      End
      Begin VB.Label lblWithdraw 
         Caption         =   "#"
         Height          =   315
         Left            =   5610
         TabIndex        =   38
         Top             =   5340
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque No"
         Height          =   315
         Left            =   90
         TabIndex        =   37
         Top             =   5340
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   5670
         Width           =   855
      End
   End
   Begin VB.CheckBox chkMutualBank 
      Caption         =   "Mutual Reconcile Bank"
      Height          =   315
      Left            =   4440
      TabIndex        =   31
      Top             =   780
      Width           =   1965
   End
   Begin VB.Frame fraCheqAmt 
      Caption         =   "cham"
      Height          =   6855
      Left            =   570
      TabIndex        =   10
      Top             =   1560
      Width           =   15915
      Begin VB.CheckBox chkAll 
         Caption         =   "Check All"
         Height          =   315
         Left            =   7350
         TabIndex        =   11
         Top             =   210
         Width           =   315
      End
      Begin VSFlex8LCtl.VSFlexGrid fgcheAmtsame 
         Height          =   6465
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   15750
         _cx             =   27781
         _cy             =   11404
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBankReconciliationProcessNew.frx":035B
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   30
      Width           =   16455
      Begin VB.Label lblStage 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         Height          =   375
         Left            =   5790
         TabIndex        =   30
         Top             =   0
         Width           =   10575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Display only Unreconciled items"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   150
         TabIndex        =   25
         Top             =   60
         Width           =   5445
      End
   End
   Begin VB.Frame fraMutual 
      Caption         =   "Mutual"
      Height          =   6855
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   15915
      Begin VSFlex8LCtl.VSFlexGrid vsMutual 
         Height          =   6465
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Width           =   15750
         _cx             =   27781
         _cy             =   11404
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBankReconciliationProcessNew.frx":0530
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
      Begin VB.CheckBox Check1 
         Caption         =   "Check All"
         Height          =   315
         Left            =   6870
         TabIndex        =   17
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CheckBox chkReconcile 
      Caption         =   "Reconcile"
      Height          =   315
      Left            =   6540
      TabIndex        =   23
      Top             =   780
      Width           =   1155
   End
   Begin VB.CommandButton cmdReconcileVoucher 
      Caption         =   "&Reconcile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7740
      TabIndex        =   9
      Top             =   8820
      Width           =   1005
   End
   Begin VB.CheckBox chkMutual 
      Caption         =   "Mutual Reconcile"
      Height          =   315
      Left            =   2850
      TabIndex        =   7
      Top             =   780
      Width           =   1575
   End
   Begin VB.CheckBox chkCheqAmt 
      Caption         =   "Cheque No and Amount Same"
      Height          =   315
      Left            =   270
      TabIndex        =   6
      Top             =   780
      Width           =   2505
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Search"
      Height          =   345
      Left            =   8010
      TabIndex        =   5
      Top             =   780
      Width           =   1035
   End
   Begin VB.TextBox txtD1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   9570
      TabIndex        =   4
      Top             =   450
      Width           =   1485
   End
   Begin VB.CommandButton cmdBankSearch 
      Caption         =   "---"
      Height          =   345
      Left            =   8010
      TabIndex        =   2
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox txtBankName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   2340
      TabIndex        =   1
      Top             =   405
      Width           =   5610
   End
   Begin VB.TextBox txtBankCode 
      Enabled         =   0   'False
      Height          =   345
      Left            =   825
      TabIndex        =   0
      Top             =   405
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Voucher date"
      Height          =   225
      Left            =   8490
      TabIndex        =   8
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   225
      Left            =   330
      TabIndex        =   3
      Top             =   435
      Width           =   420
   End
End
Attribute VB_Name = "frmBankReconciliationProcessNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim mSearchID As Variant
    Dim mvarManuallyReconciled As Variant
    Dim mvarRemarks As Variant
    Dim mSelectedAmt As Double
    Dim mSelectedScroll As Double
    Dim mBankEntryDate As Date
    Private Sub chkAll_Click()
          If chkAll.Value = vbChecked Then
              If fgcheAmtsame.Rows > 1 Then
                  fgcheAmtsame.Cell(flexcpChecked, 1, 5, fgcheAmtsame.Rows - 1, 5) = True
                  fgcheAmtsame.Cell(flexcpChecked, 1, 11, fgcheAmtsame.Rows - 1, 11) = True
              End If
          ElseIf chkAll.Value = vbUnchecked Then
              If fgcheAmtsame.Rows > 1 Then
                  fgcheAmtsame.Cell(flexcpChecked, 1, 5, fgcheAmtsame.Rows - 1, 5) = False
                  fgcheAmtsame.Cell(flexcpChecked, 1, 11, fgcheAmtsame.Rows - 1, 1) = False
        
              End If
          End If
    End Sub



Private Sub chkCheqAmt_Click()
    If chkCheqAmt.Value = vbChecked Then
        fraCheqAmt.Visible = True
'        chkCheqAmt.Value = vbUnchecked
'        'chkMutual.Value = vbchecked
'    Else
'        chkCheqAmt.Value = vbChecked
'        chkMutual.Value = vbUnchecked
    End If
End Sub

    Private Sub chkCheqAmt_Validate(Cancel As Boolean)
    
        If chkCheqAmt.Value = vbChecked Then
               
            chkCheqAmt.Value = vbChecked
            chkMutual.Value = vbUnchecked
            chkReconcile.Value = vbUnchecked
        Else
            chkCheqAmt.Value = vbUnchecked
        End If
    
    End Sub

    Private Sub chkMutual_Click()
        If chkMutual.Value = vbChecked Then
        fraMutual.Visible = True
'            chkMutual.Value = vbUnchecked
'            'chkMutual.Value = vbchecked
'        Else
'            chkMutual.Value = vbChecked
'            chkCheqAmt.Value = vbUnchecked
        End If
    End Sub

    Private Sub chkMutual_Validate(Cancel As Boolean)
        If chkMutual.Value = vbChecked Then
           
            chkMutual.Value = vbChecked
            
        chkCheqAmt.Value = vbUnchecked
        chkReconcile.Value = vbUnchecked
    Else
        chkMutual.Value = vbUnchecked
    End If
    End Sub

    Private Sub chkMutualBank_Click()
        If chkMutualBank.Value = vbChecked Then
            fraMutualBank.Visible = True

        End If
    End Sub

    Private Sub chkMutualBank_Validate(Cancel As Boolean)
        If chkMutualBank.Value = vbChecked Then
           
            chkMutualBank.Value = vbChecked
            
            chkCheqAmt.Value = vbUnchecked
            chkMutual.Value = vbUnchecked
            chkReconcile.Value = vbUnchecked
        Else
            chkMutualBank.Value = vbUnchecked
        End If
    End Sub

    Private Sub chkReconcile_Click()
        If chkMutual.Value = vbChecked Then
            fraReconcile.Visible = True

        End If
    End Sub

Private Sub chkReconcile_Validate(Cancel As Boolean)
    If chkReconcile.Value = vbChecked Then
           
            chkReconcile.Value = vbChecked
            
        chkCheqAmt.Value = vbUnchecked
        chkMutual.Value = vbUnchecked
    Else
        chkReconcile.Value = vbUnchecked
    End If
End Sub

    Private Sub cmdBankSearch_Click()
        Dim mSQL As String
        Dim mCount As Integer
        mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
        frmSearchAccountHeads.SQLString = mSQL
        frmSearchAccountHeads.Show vbModal
        mCount = InStr(1, gbSearchStr, " ")
        mSearchID = gbSearchID
        txtBankCode.Text = IIf(IsNull(Left(gbSearchStr, mCount)), "", Left(gbSearchStr, mCount))
        txtBankCode.Tag = mSearchID
        If mCount <> 0 Then
            txtBankName.Text = IIf(IsNull(mID(gbSearchStr, mCount)), "", mID(gbSearchStr, mCount))
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

    Private Sub cmdGet_Click()
        Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDB As New clsDB
            Dim mSQL As String
            Dim mRowCount As Long
            Dim mAccID As Integer
            Dim mD1 As Date
            Dim mD2 As Date
            Dim mVrID As Double
        If val(txtBankCode.Tag) < 1 Then
            MsgBox "Please Select Bank", vbApplicationModal
            Exit Sub
        Else
            mAccID = val(txtBankCode.Tag)
            fraCheqAmt.Tag = mAccID
        End If
        If txtD1.Text <> "" Then
            
        Else
        
'            If txtD2.Text <> "" Then
'            Else
        If chkReconcile.Value = vbUnchecked Then
'               MsgBox ("Please Enter Date"), vbApplicationModal
'               Exit Sub
        End If
'            End If
        End If
        If chkCheqAmt.Value = vbUnchecked Then
            If chkMutual.Value = vbUnchecked Then
                If chkReconcile.Value = vbUnchecked Then
                    If chkMutualBank.Value = vbUnchecked Then
                        MsgBox "Please select Any one option", vbApplicationModal
                        Exit Sub
                    End If
                End If
            End If
        Else
            
        End If
        
        objDB.SetConnection mCnn
        If chkCheqAmt.Value = vbChecked Then
            fraCheqAmt.Visible = True
            fraMutual.Visible = False
            fraReconcile.Visible = False
            fraMutualBank.Visible = False
            lblStage.Caption = " Chequeno and Amount are same in both side"
            mSQL = " Select A.* From"
            mSQL = mSQL + " (Select fltAmount,faVouchers.intVoucherNo,vchInstrumentNo,vchCheQueNo,fltDrAmount,fltCrAmount,faVouchers.intVoucherID VrID,intReconciliationID ReconID,dtDate,dtBankEntryDate"
            mSQL = mSQL + " ,vchDescription,vchParticulars"
            mSQL = mSQL + " From faVouchers"
            mSQL = mSQL + " Inner Join fabankReconciliationEntries On faVouchers.intKeyID1=fabankReconciliationEntries.intBankAccountHeadID"
            mSQL = mSQL + " and faVouchers.vchInstrumentNO=fabankReconciliationEntries.vchChequeNo and fltAmount=fltDrAmount"
            mSQL = mSQL + " Where tnyVouchertypeID in (20,30) and faVouchers.intKeyID1=" & mAccID & " and faVouchers.tnyReconciled is Null"
            mSQL = mSQL + " and fabankReconciliationEntries.tnyReconciled is Null"
            mSQL = mSQL + " Union All"
            mSQL = mSQL + " Select fltAmount,faVouchers.intVoucherNo,vchInstrumentNo,vchCheQueNo,fltDrAmount,fltCrAmount,faVouchers.intVoucherID VrID,intReconciliationID ReconID,dtDate,dtBankEntryDate"
            mSQL = mSQL + " ,vchDescription,vchParticulars"
            mSQL = mSQL + " From faVouchers"
            mSQL = mSQL + " Inner Join fabankReconciliationEntries On faVouchers.intKeyID1=fabankReconciliationEntries.intBankAccountHeadID"
            mSQL = mSQL + " and faVouchers.vchInstrumentNO=fabankReconciliationEntries.vchChequeNo and fltAmount=fltCrAmount"
            mSQL = mSQL + " Where tnyVouchertypeID in (10,30) and faVouchers.intKeyID1=" & mAccID & " and faVouchers.tnyReconciled is Null"
            mSQL = mSQL + " and fabankReconciliationEntries.tnyReconciled is Null )A"
            mSQL = mSQL + " Where A.dtDate between '" & DdMmmYy(txtD1.Text) & "' and '" & DdMmmYy(txtD1.Text) & "'"
            mSQL = mSQL + " "
            mSQL = mSQL + " Order By A.dtDate"
            
            Rec.Open mSQL, mCnn
            fgcheAmtsame.Rows = 1
            mRowCount = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    fgcheAmtsame.Rows = fgcheAmtsame.Rows + 1
                    fgcheAmtsame.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    fgcheAmtsame.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    fgcheAmtsame.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    fgcheAmtsame.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    fgcheAmtsame.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    fgcheAmtsame.TextMatrix(mRowCount, 5) = flexChecked
                    fgcheAmtsame.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!ReconID), "", Rec!ReconID)
                    fgcheAmtsame.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                    fgcheAmtsame.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                    fgcheAmtsame.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                    If (IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount) > 0) Then
                        fgcheAmtsame.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
                    Else
                        fgcheAmtsame.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
                    End If
                    fgcheAmtsame.TextMatrix(mRowCount, 11) = flexChecked
                    fgcheAmtsame.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!VrID), "", Rec!VrID)
                    fgcheAmtsame.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!ReconID), "", Rec!ReconID)
                    
                    mRowCount = mRowCount + 1
                    
                    Rec.MoveNext
                Wend
            Else
                MsgBox "Data not Found", vbApplicationModal
                Exit Sub
            End If
         ElseIf (chkMutual.Value = vbChecked) Then
         
            ''List only Reversed Vouchers
            lblStage.Caption = " Reversed Vouchers Only"
            fraCheqAmt.Visible = False
            fraMutual.Visible = True
            fraReconcile.Visible = False
            fraMutualBank.Visible = False
            mSQL = " Select * From faVouchers"
            mSQL = mSQL + " Where faVouchers.tnyReversed=1 And numLinkKeyId is null and tnyReconciled is null and faVouchers.dtDate='" & DdMmmYy(txtD1.Text) & "' and faVouchers.intKeyID1=" & mAccID
            Rec.Open mSQL, mCnn
            vsMutual.Rows = 1
            mRowCount = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    vsMutual.Rows = vsMutual.Rows + 1
                    vsMutual.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsMutual.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    vsMutual.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    vsMutual.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    vsMutual.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsMutual.TextMatrix(mRowCount, 5) = flexChecked
    
                    vsMutual.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    mVrID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    Call GetVoucherDetails(mVrID, mRowCount)
                    mRowCount = mRowCount + 1
                    Rec.MoveNext
                Wend
            Else
                MsgBox "Data not Found", vbApplicationModal
                Exit Sub
            End If
         ElseIf (chkReconcile.Value = vbChecked) Then
            lblStage.Caption = "List all Vouchers Except Reversed and cancelled"
            fraCheqAmt.Visible = False
            fraMutual.Visible = False
            fraReconcile.Visible = True
            fraMutualBank.Visible = False
            mSQL = " Select * From faVouchers"
            mSQL = mSQL + " Where faVouchers.tnyReversed is null And isnull(tnyStatus,0)=0 and tnyReconciled is null"
            mSQL = mSQL + " and faVouchers.intKeyID1=" & mAccID
            If txtD1 <> "" Then
                mSQL = mSQL + "  and faVouchers.dtDate='" & DdMmmYy(txtD1.Text) & "' "
            End If
            If txtInstrument.Text <> "" Then
                mSQL = mSQL + "  and vchInstrumentNo like '%" & txtInstrument.Text & "%'"
            End If
            If val(txtAmount.Text) > 0 Then
                mSQL = mSQL + "  and fltAmount =" & txtAmount.Text & ""
            End If
            Rec.Open mSQL, mCnn
            vsVoucher.Rows = 1
            mRowCount = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    vsVoucher.Rows = vsVoucher.Rows + 1
                    vsVoucher.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                    vsVoucher.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                    vsVoucher.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    vsVoucher.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    vsVoucher.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    vsVoucher.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                    mRowCount = mRowCount + 1
                Rec.MoveNext
                Wend
            End If
            Rec.Close
            mSQL = " Select * From fabankReconciliationEntries"
            mSQL = mSQL + " Where tnyReconciled is null "
            If txtBankDate.Text <> "" Then
                mSQL = mSQL + "  and dtBankEntryDate='" & DdMmmYy(txtBankDate.Text) & "' "
            End If
            mSQL = mSQL + "  and intBankAccountHeadID=" & mAccID
            If txtInstrument.Text <> "" Then
                mSQL = mSQL + "  and vchChequeNo like '%" & txtInstrument.Text & "%'"
            End If
            If val(txtAmount.Text) > 0 Then
                mSQL = mSQL + "  and (fltCrAmount =" & txtAmount.Text & " Or fltDrAmount=" & txtAmount.Text & ")"
            End If
            Rec.Open mSQL, mCnn
            vsBank.Rows = 1
            mRowCount = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    vsBank.Rows = vsBank.Rows + 1
                    vsBank.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                    vsBank.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                    vsBank.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                    vsBank.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                    vsBank.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
                    vsBank.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
                    'vsBank.TextMatrix(mRowCount, 6) = flexChecked
                    vsBank.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
'                    vsVoucher.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
'                    mVrID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)

                    mRowCount = mRowCount + 1
                Rec.MoveNext
                Wend
                
            End If
        ElseIf (chkMutualBank.Value = vbChecked) Then
            lblStage.Caption = "List all UnReconciled Bank Passbook details"
            fraCheqAmt.Visible = False
            fraMutual.Visible = False
            fraReconcile.Visible = False
            fraMutualBank.Visible = True
            
            mSQL = " Select * From fabankReconciliationEntries"
            mSQL = mSQL + " Where tnyReconciled is null "
            If txtD1.Text <> "" Then
                mSQL = mSQL + "  and dtBankEntryDate='" & DdMmmYy(txtD1.Text) & "' "
            End If
            mSQL = mSQL + "  and intBankAccountHeadID=" & mAccID
            If txtChequeNo.Text <> "" Then
                mSQL = mSQL + "  and vchChequeNo like '%" & txtChequeNo.Text & "%'"
            End If
            If val(txtBankAmount.Text) > 0 Then
                mSQL = mSQL + "  and (fltCrAmount =" & txtBankAmount.Text & " Or fltDrAmount=" & txtBankAmount.Text & ")"
            End If
            Rec.Open mSQL, mCnn
            vsbankMutual.Rows = 1
            mRowCount = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF Or Rec.BOF)
                    vsbankMutual.Rows = vsbankMutual.Rows + 1
                    vsbankMutual.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
                    vsbankMutual.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!dtBankEntryDate), "", Rec!dtBankEntryDate)
                    vsbankMutual.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchChequeNo), "", Rec!vchChequeNo)
                    vsbankMutual.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
                    vsbankMutual.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
                    vsbankMutual.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
                    'vsBank.TextMatrix(mRowCount, 6) = flexChecked
                    vsbankMutual.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
'                    vsVoucher.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
'                    mVrID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)

                    mRowCount = mRowCount + 1
                Rec.MoveNext
                Wend
                
            End If
        End If
    End Sub
    Private Sub GetVoucherDetails(ByVal mVrID As Double, mRow As Long)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mSQL As String
        
        objDB.SetConnection mCnn
        mSQL = " Select * From faVouchers"
        mSQL = mSQL + " Where faVouchers.tnyReversed=1 And numLinkKeyId=" & mVrID
        Rec.Open mSQL, mCnn
            While Not (Rec.EOF Or Rec.BOF)
                vsMutual.TextMatrix(mRow, 6) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                vsMutual.TextMatrix(mRow, 7) = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                vsMutual.TextMatrix(mRow, 8) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                vsMutual.TextMatrix(mRow, 9) = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                vsMutual.TextMatrix(mRow, 10) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                vsMutual.TextMatrix(mRow, 11) = flexChecked
'                vsMutual.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!ReconID), "", Rec!ReconID)
'                vsMutual.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!dtbankEntrydate), "", Rec!dtbankEntrydate)
'                vsMutual.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!vchCheQueNo), "", Rec!vchCheQueNo)
'                vsMutual.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
'                If (IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount) > 0) Then
'                    vsMutual.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
'                Else
'                    vsMutual.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!fltCrAmount), "", Rec!fltCrAmount)
'                End If
'                vsMutual.TextMatrix(mRowCount, 11) = flexChecked
                vsMutual.TextMatrix(mRow, 13) = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
'                vsMutual.TextMatrix(mRowCount, 13) = IIf(IsNull(Rec!ReconID), "", Rec!ReconID)
                
            Rec.MoveNext
            Wend
    End Sub
    
    Private Sub cmdReconcileVoucher_Click()
        If chkCheqAmt.Value = vbChecked Then
            Call AutoReconcile
        ElseIf chkMutual.Value = vbChecked Then
            Call MutualReconcile
        ElseIf chkReconcile.Value = vbChecked Then
            Call Reconcile
        ElseIf chkMutualBank.Value = vbChecked Then
            Call BankMutual
        End If
        Call cmdGet_Click
    End Sub
    Private Sub BankMutual()
        Dim mSQL As String
        Dim mLoop As Integer
        Dim mVrID As Double
        Dim mVrNo As Double
        Dim mVrReNo As Double
        Dim mReconId   As Double
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mRevDate  As Date
        objDB.SetConnection mCnn
        Dim mReconWith As Double
        Dim mReconDepo As Double
        Dim mCount As Integer
        If lblWithdraw.Caption <> "#" Then
            If lblWithdraw.Caption = lblDeposit.Caption Then
'                mReconWith = lblWithdraw.Tag
'                mReconDepo = lblDeposit.Tag
                For mCount = 1 To vsbankMutual.Rows - 1
                    If vsbankMutual.Cell(flexcpChecked, mCount, 6) = vbChecked Then
                         mSQL = "Update faBankReconciliationEntries Set vchRemarks = vchRemarks+'Bank Mutual' "
                        ' mSQL = mSQL + " intVoucherNo =  " & mVrID
                         mSQL = mSQL + ", tnyReconciled = 5"
                         mSQL = mSQL + ", fltDifference = 0"
                         mSQL = mSQL + " Where intReconciliationID = " & vsbankMutual.TextMatrix(mCount, 7)
                         mCnn.Execute mSQL

                    End If
                Next
                MsgBox "Reconciled Successfully", vbApplicationModal
            Else
                MsgBox "Withdrawal Amount and Deposit Amount are not Matching", vbApplicationModal
                Exit Sub
            End If
        End If
    End Sub
    Private Sub Reconcile()
        Dim mSQL As String
        Dim mLoop As Integer
        Dim mVrID As Double
        Dim mVrNo As Double
        Dim mVrReNo As Double
        Dim mReconId   As Double
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mRevDate  As Date
        objDB.SetConnection mCnn
        
        If lblBank.Caption = lblVr.Caption Then
            mVrID = val(lblVr.Tag)
            mReconId = val(lblBank.Tag)
            For mLoop = 1 To vsVoucher.Rows - 1
                mVrID = vsVoucher.TextMatrix(mLoop, 6)
                If vsVoucher.Cell(flexcpChecked, mLoop, 5) = vbChecked Then
                    mSQL = " Update faVouchers "
                    mSQL = mSQL + " Set tnyReconciled = 2"
                    mSQL = mSQL + ", numTockenID = " & mReconId
                    mSQL = mSQL + ", dtRealisationDate = '" & DdMmmYy(gbTransactionDate) & "'"
                    mSQL = mSQL + ", tnysync = Null"
                    mSQL = mSQL + ", vchRemarks = vchRemarks+'Auto-Reconciliation'"
                    mSQL = mSQL + " Where intVoucherID = " & mVrID
                    mCnn.Execute mSQL
                    
                    mSQL = " Update faTransactionChild "
                    mSQL = mSQL + " Set numTockenID = 2"
                    mSQL = mSQL + ", dtReconcileDate = '" & DdMmmYy(gbTransactionDate) & "'"
                    mSQL = mSQL + ", tnysync = Null"
                    mSQL = mSQL + " Where intTransactionID in (Select intTransactionID From faTransactions Where intVoucherID=" & mVrID & ")"
                    mSQL = mSQL + "And intAccountHeadID=" & val(fraCheqAmt.Tag)
                    mCnn.Execute mSQL
                    
                End If
            Next
            For mLoop = 1 To vsBank.Rows - 1
                mReconId = vsBank.TextMatrix(mLoop, 7)
                If vsBank.Cell(flexcpChecked, mLoop, 6) = vbChecked Then
                    mSQL = "Update faBankReconciliationEntries Set vchRemarks = vchRemarks+'Auto-Reconciliation' ,"
                    mSQL = mSQL + " intVoucherNo =  " & mVrID
                    mSQL = mSQL + ", tnyReconciled = 2"
                    mSQL = mSQL + ", fltDifference = 0"
                    mSQL = mSQL + " Where intReconciliationID = " & mReconId
                    mCnn.Execute mSQL
                End If
            Next
        Else
            MsgBox "Voucher Amount and bank stmt Amount are not Matching", vbApplicationModal
            Exit Sub
        End If
        
    End Sub
    
    Private Sub MutualReconcile()
        ''Reversed item
        Dim mSQL As String
        Dim mLoop As Integer
        Dim mVrID As Double
        Dim mVrNo As Double
        Dim mVrReNo As Double
        Dim mReconId   As Double
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mRevDate  As Date
        objDB.SetConnection mCnn
        If chkMutual.Value = vbChecked Then
            If vsMutual.Rows > 1 Then
                For mLoop = 1 To fgcheAmtsame.Rows - 1
                    mVrID = vsMutual.TextMatrix(mLoop, 12)
                    mVrReNo = vsMutual.TextMatrix(mLoop, 13)
                    mVrNo = vsMutual.TextMatrix(mLoop, 0)
                    mRevDate = vsMutual.TextMatrix(mLoop, 7)
                    If vsMutual.Cell(flexcpChecked, mLoop, 5) = vbChecked Then
                        mSQL = " Update faVouchers "
                        mSQL = mSQL + " Set tnyReconciled = 5"
                        mSQL = mSQL + ", numTockenID = " & mVrNo
                        mSQL = mSQL + ", dtRealisationDate = '" & DdMmmYy(mRevDate) & "'"
                        mSQL = mSQL + ", tnysync = Null"
                        mSQL = mSQL + ", vchRemarks = vchRemarks+'Auto-Reconciliation'"
                        mSQL = mSQL + " Where intVoucherID in ( " & mVrID & "," & mVrReNo & ")"
                        mCnn.Execute mSQL
                        
                        mSQL = " Update faTransactionChild "
                        mSQL = mSQL + " Set numTockenID = " & mVrNo
                        mSQL = mSQL + ", dtReconcileDate = '" & DdMmmYy(mRevDate) & "'"
                        mSQL = mSQL + ", tnysync = Null"
                        mSQL = mSQL + " Where intTransactionID in (Select intTransactionID From faTransactions Where intVoucherID in ( " & mVrID & "," & mVrReNo & "))"
                        mSQL = mSQL + "And intAccountHeadID=" & val(fraCheqAmt.Tag)
                        mCnn.Execute mSQL
                    End If
                    
                Next
            End If
        End If
    End Sub
    Private Sub AutoReconcile()
        '' with cheque no and amount same
        Dim mSQL As String
        Dim mLoop As Integer
        Dim mVrID As Double
        Dim mVrNo As Double
        Dim mReconId   As Double
        Dim mBankDate As Date
        Dim objDB As New clsDB
                Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            
        objDB.SetConnection mCnn
        
        If chkCheqAmt.Value = vbChecked Then
            If fgcheAmtsame.Rows > 1 Then
                For mLoop = 1 To fgcheAmtsame.Rows - 1
                    mVrID = fgcheAmtsame.TextMatrix(mLoop, 12)
                    mReconId = fgcheAmtsame.TextMatrix(mLoop, 13)
                    mVrNo = fgcheAmtsame.TextMatrix(mLoop, 0)
                    mBankDate = fgcheAmtsame.TextMatrix(mLoop, 7)
                    If fgcheAmtsame.Cell(flexcpChecked, mLoop, 5) = vbChecked Then
                        mSQL = " Update faVouchers "
                        mSQL = mSQL + " Set tnyReconciled = 2"
                        mSQL = mSQL + ", numTockenID = " & mReconId
                        mSQL = mSQL + ", dtRealisationDate = '" & DdMmmYy(mBankDate) & "'"
                        mSQL = mSQL + ", tnysync = Null"
                        mSQL = mSQL + ", vchRemarks = vchRemarks+'Auto-Reconciliation'"
                        mSQL = mSQL + " Where intVoucherID = " & mVrID
                        mCnn.Execute mSQL
                        
                        mSQL = " Update faTransactionChild "
                        mSQL = mSQL + " Set numTockenID = 2"
                        mSQL = mSQL + ", dtReconcileDate = '" & DdMmmYy(mBankDate) & "'"
                        mSQL = mSQL + ", tnysync = Null"
                        mSQL = mSQL + " Where intTransactionID in (Select intTransactionID From faTransactions Where intVoucherID=" & mVrID & ")"
                        mSQL = mSQL + "And intAccountHeadID=" & val(fraCheqAmt.Tag)
                        mCnn.Execute mSQL
                        
                        mSQL = "Update faBankReconciliationEntries Set vchRemarks = vchRemarks+'Auto-Reconciliation' ,"
                        mSQL = mSQL + " intVoucherNo =  " & mVrNo
                        mSQL = mSQL + ", tnyReconciled = 2"
                        mSQL = mSQL + ", fltDifference = 0"
                        mSQL = mSQL + " Where intReconciliationID = " & mReconId
                        mCnn.Execute mSQL
                    End If
                    
                Next
                MsgBox "Reconciled Successfully", vbApplicationModal
                Call cmdGet_Click
                
            End If
        End If
    End Sub

Private Sub fgcheAmtsame_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    If fgcheAmtsame.Col = 5 Then
'            If fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 5) = flexChecked Then
'                 fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 11) = flexChecked
'            Else
'                fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 11) = flexUnchecked
'            End If
'        Else
'          '  vsGrid.Editable = flexEDNone
'        End If
End Sub

    Private Sub fgcheAmtsame_Click()
        If fgcheAmtsame.Col = 5 Then
            If fgcheAmtsame.Cell(flexcpChecked, fgcheAmtsame.Row, 5) = flexChecked Then
                 'fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 11) = flexChecked
                 fgcheAmtsame.Cell(flexcpChecked, fgcheAmtsame.Row, 5) = flexUnchecked
                 fgcheAmtsame.Cell(flexcpChecked, fgcheAmtsame.Row, 11) = flexUnchecked
            Else
                fgcheAmtsame.Cell(flexcpChecked, fgcheAmtsame.Row, 5) = flexChecked
                fgcheAmtsame.Cell(flexcpChecked, fgcheAmtsame.Row, 11) = flexChecked
'                fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 5) = flexUnchecked
'                fgcheAmtsame.TextMatrix(fgcheAmtsame.Row, 11) = flexUnchecked
            End If
        Else
          '  vsGrid.Editable = flexEDNone
        End If
    End Sub

    Private Sub Form_Load()
        txtD1.Text = DdMmmYy(gbStartingDate)
        txtD1.Text = DdMmmYy(gbStartingDate)
    End Sub



    Private Sub txtBankDate_LostFocus()
        If (IsDate(txtBankDate)) Then
            If Trim(txtBankDate) <> "" Then
                    txtBankDate = DdMmmYy(txtBankDate)
            Else
                txtBankDate.Text = DdMmmYy(gbStartingDate)
            End If
        Else
            MsgBox "Wrong date format", vbApplicationModal
            Exit Sub
        End If
    End Sub

    Private Sub txtD1_LostFocus()
        If (IsDate(txtD1)) Then
            If Trim(txtD1) <> "" Then
                    txtD1 = DdMmmYy(txtD1)
            Else
                txtD1.Text = DdMmmYy(gbStartingDate)
            End If
        Else
            MsgBox "Wrong date format", vbApplicationModal
            Exit Sub
        End If
    End Sub

'    Private Sub txtD2_LostFocus()
'        If (IsDate(txtD2)) Then
'            If Trim(txtD2) <> "" Then
'                txtD2 = DdMmmYy(txtD2)
'            Else
'                txtD2.Text = DdMmmYy(gbStartingDate)
'            End If
'        Else
'            MsgBox "Wrong date format", vbApplicationModal
'            Exit Sub
'        End If
'    End Sub



    Private Sub vsBank_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 6 Then
            If vsBank.Cell(flexcpChecked, Row, 6) = vbChecked Then
                Call calculateVrAmt
                'lblVr.Tag = vsVoucher.Cell(flexcpChecked, Row, 7)
''                If vsVoucher.Cell(flexcpChecked, Row, 8) = vbChecked Then
'                    mSelectedScroll = mSelectedScroll + val(vsBank.Cell(flexcpText, Row, 4)) + val(vsBank.Cell(flexcpText, Row, 5)) * -1
''                Else
''                    mSelectedAmt = mSelectedAmt - val(vsVoucher.Cell(flexcpText, Row, 3)) - val(vsVoucher.Cell(flexcpText, Row, 4)) * -1
''                End If
Else
Call calculateVrAmt
            End If
'            lblVr.Caption = Format(Abs(mSelectedScroll), "0.00")
            
        End If
    End Sub
    Private Sub vsBank_Click()
        If vsBank.Col = 6 Then
            vsBank.Editable = flexEDKbdMouse
            Call calculateBankAmt
        Else
            vsBank.Editable = flexEDNone
        End If
    End Sub



    Private Sub vsbankMutual_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 6 Then
            If vsbankMutual.Cell(flexcpChecked, Row, 6) = vbChecked Then
                Call calculateVrAmt
                'lblVr.Tag = vsVoucher.Cell(flexcpChecked, Row, 7)
''                If vsVoucher.Cell(flexcpChecked, Row, 8) = vbChecked Then
'                    mSelectedScroll = mSelectedScroll + val(vsBank.Cell(flexcpText, Row, 4)) + val(vsBank.Cell(flexcpText, Row, 5)) * -1
''                Else
''                    mSelectedAmt = mSelectedAmt - val(vsVoucher.Cell(flexcpText, Row, 3)) - val(vsVoucher.Cell(flexcpText, Row, 4)) * -1
''                End If
            End If
'            lblVr.Caption = Format(Abs(mSelectedScroll), "0.00")
            
        End If
    End Sub
    Private Sub vsbankMutual_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If Col = 6 Then
            vsbankMutual.Editable = flexEDKbdMouse
        Else
            vsbankMutual.Editable = flexEDNone
        End If
    End Sub
    Private Sub vsbankMutual_Click()
        If vsbankMutual.Col = 6 Then
            vsbankMutual.Editable = flexEDKbdMouse
            Call calculateBankAmt
        Else
            vsbankMutual.Editable = flexEDNone
        End If
    End Sub
    Private Sub vsMutual_Click()
        If vsMutual.Col = 5 Then
            If vsMutual.Cell(flexcpChecked, vsMutual.Row, 5) = flexChecked Then
                 vsMutual.Cell(flexcpChecked, vsMutual.Row, 5) = flexUnchecked
                 vsMutual.Cell(flexcpChecked, vsMutual.Row, 11) = flexUnchecked
            Else
                vsMutual.Cell(flexcpChecked, vsMutual.Row, 5) = flexChecked
                vsMutual.Cell(flexcpChecked, vsMutual.Row, 11) = flexChecked

            End If
        Else
        End If
    End Sub
    Private Sub vsVoucher_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 5 Then
                If vsVoucher.Cell(flexcpChecked, Row, 5) <> vbChecked Then
                    Call calculateVrAmt
                    lblVr.Tag = vsVoucher.Cell(flexcpChecked, Row, 6)
    '                'If vsVoucher.Cell(flexcpChecked, Row, 8) = vbChecked Then
    '                    mSelectedAmt = mSelectedAmt + val(vsVoucher.Cell(flexcpText, Row, 4))
    ''                'Else
    ''                    mSelectedAmt = mSelectedAmt - val(vsVoucher.Cell(flexcpText, Row, 3)) - val(vsVoucher.Cell(flexcpText, Row, 4)) * -1
    ''                End If
                End If
    '            lblVr.Caption = Format(Abs(mSelectedAmt), "0.00")
            
        Else
            vsVoucher.Editable = flexEDNone
        End If
    End Sub
'    Private Sub vsVoucher_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'        If Col = 5 Then
'            vsVoucher.Editable = flexEDKbdMouse
'        Else
'        vsVoucher.Editable = flexEDNone
'        End If
'    End Sub
    Private Sub vsVoucher_Click()
'        If vsVoucher.Col = 5 Then
'
'            'vsVoucher.Editable = flexEDKbdMouse
'            Call calculateVrAmt
'        Else
'            'vsVoucher.Editable = flexEDNone
'        End If
        If vsVoucher.Col = 5 Then
            vsVoucher.Editable = flexEDKbdMouse
            Call calculateVrAmt
        Else
            vsVoucher.Editable = flexEDNone
        End If
    End Sub

    Private Sub calculateVrAmt()
        ' For Reconcile with bank and Voucher
        Dim mCount As Integer
        mSelectedAmt = 0
        mSelectedScroll = 0
            For mCount = 1 To vsVoucher.Rows - 1
                If vsVoucher.Cell(flexcpChecked, mCount, 5) = vbChecked Then
                   mSelectedAmt = mSelectedAmt + val(vsVoucher.Cell(flexcpText, mCount, 4))
                End If
            Next
            For mCount = 1 To vsBank.Rows - 1
                If vsBank.Cell(flexcpChecked, mCount, 6) = vbChecked Then
                   mSelectedScroll = mSelectedScroll + val(vsBank.Cell(flexcpText, mCount, 4)) + val(vsBank.Cell(flexcpText, mCount, 5)) * -1
                End If
            Next
         lblVr.Caption = mSelectedAmt
         lblBank.Caption = Abs(mSelectedScroll)
    End Sub
    Private Sub calculateBankAmt()
        Dim mCount As Integer
        Dim mwBankAmt As Double
        Dim mdBankAmt As Double
        mwBankAmt = 0
        mdBankAmt = 0
         lblWithdraw.Caption = mwBankAmt
         lblDeposit.Caption = mdBankAmt
            For mCount = 1 To vsbankMutual.Rows - 1
                If vsbankMutual.Cell(flexcpChecked, mCount, 6) = vbChecked Then
                   mwBankAmt = mwBankAmt + val(vsbankMutual.Cell(flexcpText, mCount, 4))
                   mdBankAmt = mdBankAmt + val(vsbankMutual.Cell(flexcpText, mCount, 5))
                   
'                   lblWithdraw.Tag = vsbankMutual.TextMatrix(mCount, 7)
'                    lblDeposit.Tag = vsbankMutual.TextMatrix(mCount, 7)
                End If
            Next
         lblWithdraw.Caption = mwBankAmt
         lblDeposit.Caption = mdBankAmt
    End Sub
    Private Sub vsVoucher_DblClick()
        Dim mLoop As Long
    
            If vsVoucher.Row = -1 Then Exit Sub
            
            If vsVoucher.Row > 0 Then
                If vsVoucher.Cell(flexcpChecked, vsVoucher.Row, 5) = flexChecked Then
                    lblVr.Caption = vsVoucher.TextMatrix(vsVoucher.Row, 4)
                    lblVr.Tag = vsVoucher.TextMatrix(vsVoucher.Row, 6)
                End If
            End If
    End Sub

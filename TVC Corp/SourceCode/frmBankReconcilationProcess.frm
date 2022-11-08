VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankReconcilationProcess 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Reconciliation Process"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   18705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   14730
      TabIndex        =   50
      Top             =   8160
      Width           =   1425
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto Reconcile"
      Height          =   345
      Left            =   9210
      TabIndex        =   49
      Top             =   300
      Width           =   1395
   End
   Begin VB.CommandButton cmdAutoRecon 
      Caption         =   "Auto2"
      Height          =   375
      Left            =   16365
      TabIndex        =   48
      Top             =   8610
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton cmdReconciliationReport 
      Caption         =   "Reconciliation Report"
      Height          =   420
      Left            =   16350
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton cmdAutoReconcile 
      Caption         =   "Auto Reconcile"
      Height          =   375
      Left            =   16305
      TabIndex        =   46
      Top             =   7515
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Repor&t"
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
      Left            =   8430
      TabIndex        =   34
      Top             =   8490
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtVInstNo 
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
      Height          =   330
      Left            =   11400
      TabIndex        =   28
      Top             =   7815
      Width           =   1035
   End
   Begin VB.TextBox txtVDrAmt 
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
      Height          =   330
      Left            =   12435
      TabIndex        =   30
      Top             =   7815
      Width           =   1035
   End
   Begin VB.TextBox txtVCrAmt 
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
      Height          =   330
      Left            =   13470
      TabIndex        =   32
      Top             =   7815
      Width           =   1035
   End
   Begin VB.TextBox txtBankCrAmt 
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
      Height          =   330
      Left            =   4215
      TabIndex        =   18
      Top             =   8085
      Width           =   1035
   End
   Begin VB.TextBox txtBankDrAmt 
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
      Height          =   330
      Left            =   3180
      TabIndex        =   16
      Top             =   8085
      Width           =   1035
   End
   Begin VB.TextBox txtBankInstNo 
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
      Height          =   330
      Left            =   2145
      TabIndex        =   14
      Top             =   8085
      Width           =   1035
   End
   Begin VB.CheckBox chkWithDrawalsOnly 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3225
      TabIndex        =   5
      Top             =   1215
      Width           =   255
   End
   Begin VB.CheckBox chkDepositOnly 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4275
      TabIndex        =   6
      Top             =   1215
      Width           =   255
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
      Left            =   8445
      TabIndex        =   33
      Top             =   7740
      Width           =   1005
   End
   Begin VB.CheckBox chkUnReconciledBank 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Un-Reconciled"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3690
      TabIndex        =   12
      Top             =   7590
      Width           =   1770
   End
   Begin VB.CheckBox chkUnReconciledVouchers 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Un-Reconciled"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11520
      TabIndex        =   26
      Top             =   7380
      Width           =   1770
   End
   Begin VB.TextBox txtD4 
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
      Left            =   10020
      TabIndex        =   25
      Text            =   "30-Sep-2008"
      Top             =   7350
      Width           =   1335
   End
   Begin VB.TextBox txtD3 
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
      Left            =   8430
      TabIndex        =   24
      Text            =   "01-Apr-2008"
      Top             =   7350
      Width           =   1335
   End
   Begin VB.TextBox txtD2 
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
      Left            =   1755
      TabIndex        =   11
      Top             =   7590
      Width           =   1335
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
      Left            =   165
      TabIndex        =   10
      Top             =   7590
      Width           =   1335
   End
   Begin VSFlex8LCtl.VSFlexGrid vsTitleGrid 
      Height          =   600
      Left            =   60
      TabIndex        =   42
      Top             =   1485
      Width           =   8265
      _cx             =   14579
      _cy             =   1058
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   13012223
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
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBankReconcilationProcess.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.ComboBox cmbVoucherType 
      Height          =   345
      Left            =   14985
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1095
      Width           =   1545
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   345
      ItemData        =   "frmBankReconcilationProcess.frx":0126
      Left            =   6255
      List            =   "frmBankReconcilationProcess.frx":0128
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1110
      Width           =   1875
   End
   Begin VB.CheckBox chkMonth 
      BackColor       =   &H00C0FFFF&
      Caption         =   "List This Month Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   225
      Left            =   16425
      TabIndex        =   20
      Top             =   120
      Width           =   2130
   End
   Begin VB.TextBox txtBankCode 
      Enabled         =   0   'False
      Height          =   345
      Left            =   615
      TabIndex        =   1
      Top             =   285
      Width           =   1515
   End
   Begin VB.TextBox txtBankName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   2130
      TabIndex        =   2
      Top             =   285
      Width           =   5610
   End
   Begin VB.CommandButton cmdBankSearch 
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7740
      TabIndex        =   3
      Top             =   270
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   210
      TabIndex        =   37
      Top             =   8460
      Width           =   3525
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   90
         TabIndex        =   41
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UnReconciled"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   435
         TabIndex        =   40
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   1890
         TabIndex        =   39
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reconciled"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   38
         Top             =   120
         Width           =   1050
      End
   End
   Begin MSComctlLib.ProgressBar pbBank 
      Height          =   195
      Left            =   45
      TabIndex        =   36
      Top             =   9180
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdReconcile 
      BackColor       =   &H8000000A&
      Caption         =   "Reconcile"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      MaskColor       =   &H00000040&
      TabIndex        =   35
      Top             =   7605
      Width           =   1380
   End
   Begin VSFlex8LCtl.VSFlexGrid fgVoucherStatement 
      Height          =   5805
      Left            =   8625
      TabIndex        =   23
      Top             =   1515
      Width           =   10080
      _cx             =   17780
      _cy             =   10239
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
      Rows            =   21
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBankReconcilationProcess.frx":012A
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
   Begin VSFlex8LCtl.VSFlexGrid fgBankStatement 
      Height          =   5475
      Left            =   45
      TabIndex        =   9
      Top             =   2070
      Width           =   8280
      _cx             =   14605
      _cy             =   9657
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
      Rows            =   20
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBankReconcilationProcess.frx":02B3
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
   Begin VB.Label lblSelectedScrollAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   5985
      TabIndex        =   19
      Top             =   8145
      Width           =   210
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      Height          =   225
      Left            =   13635
      TabIndex        =   31
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      Height          =   225
      Left            =   12465
      TabIndex        =   29
      Top             =   7560
      Width           =   435
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inst. No"
      Height          =   225
      Left            =   11430
      TabIndex        =   27
      Top             =   7560
      Width           =   630
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      Height          =   225
      Left            =   4230
      TabIndex        =   17
      Top             =   7875
      Width           =   645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withdrawal"
      Height          =   225
      Left            =   3195
      TabIndex        =   15
      Top             =   7875
      Width           =   915
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inst. No"
      Height          =   225
      Left            =   2205
      TabIndex        =   13
      Top             =   7875
      Width           =   630
   End
   Begin VB.Label lblSelectedAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   14220
      TabIndex        =   45
      Top             =   7350
      Width           =   210
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   9825
      TabIndex        =   44
      Top             =   7305
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1560
      TabIndex        =   43
      Top             =   7545
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   225
      Left            =   5655
      TabIndex        =   7
      Top             =   1170
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B a n k    S t a t e m e n t    D e t a i l s"
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   1215
      Width           =   2970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V o u c h e r   D e t a i l s"
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   8505
      TabIndex        =   21
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuLine 
         Caption         =   ""
      End
      Begin VB.Menu mnuManuallyReconcile 
         Caption         =   "Manually Reconcile "
      End
      Begin VB.Menu mnuUnReconcile 
         Caption         =   "Un-Reconcile"
      End
      Begin VB.Menu mnuVoucherMutual 
         Caption         =   "Voucher Mutual"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVoucherMutualUnReconcile 
         Caption         =   "Voucher Mutual Un-Reconcile"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExitReconciliation 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frmBankReconcilationProcess"
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
    Dim mAdjdiff As Double
    
    Private Sub chkDepositOnly_Click()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        If chkDepositOnly.Value = 1 Then
            mFlag = True
        Else
            mFlag = False
        End If
        For mLoopCount = 1 To fgBankStatement.Rows - 1
            If val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) > 0 Then
                fgBankStatement.RowHidden(mLoopCount) = mFlag
            End If
        Next mLoopCount
    End Sub

    Private Sub chkMonth_Click()
        fgBankStatement.Clear 1, 1
        fgVoucherStatement.Clear 1, 1
        Call FillVoucherStatement
        Call FillBankStatement
    End Sub
    
    Private Sub chkUnReconciledBank_Click()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        If chkUnReconciledBank.Value = 1 Then
            mFlag = True
        Else
            mFlag = False
        End If
        
        For mLoopCount = 1 To fgBankStatement.Rows - 1
            If fgBankStatement.Cell(flexcpChecked, mLoopCount, 5) = 1 Then
                fgBankStatement.RowHidden(mLoopCount) = mFlag
            End If
        Next mLoopCount
    End Sub
    
    Private Sub chkUnReconciledVouchers_Click()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        If chkUnReconciledVouchers.Value = 1 Then
            mFlag = True
        Else
            mFlag = False
        End If
        For mLoopCount = 1 To fgVoucherStatement.Rows - 1
            If fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 6) = 1 Then
                fgVoucherStatement.RowHidden(mLoopCount) = mFlag
            End If
        Next mLoopCount
    End Sub
    
    Private Sub chkWithDrawalsOnly_Click()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        
        If chkWithDrawalsOnly.Value = 1 Then
            mFlag = True
        Else
            mFlag = False
        End If
        For mLoopCount = 1 To fgBankStatement.Rows - 1
            If val(fgBankStatement.Cell(flexcpText, mLoopCount, 4)) > 0 Then
                fgBankStatement.RowHidden(mLoopCount) = mFlag
            End If
        Next mLoopCount
    End Sub

    Private Sub cmbMonth_Click()
        Dim mMonthIndex As Integer
        If txtBankCode.Text <> "" Then
            '---------------------------------------------------------------------------------'
            'Note:- Finding Range of Dates According the month selected
            '---------------------------------------------------------------------------------'
            mMonthIndex = cmbMonth.ItemData(cmbMonth.ListIndex)
            If gbLBPanchayat <> 1 Then
            If mMonthIndex > 3 Then
                txtD1.Text = CheckDateInMMM(DateSerial(gbFinancialYearID, mMonthIndex, 1))
            Else
                txtD1.Text = CheckDateInMMM(DateSerial(gbFinancialYearID + 1, mMonthIndex, 1))
            End If
            
            End If
            If Not IsDate(txtD1) Then
                txtD1.Text = DdMmmYy(gbStartingDate)
            End If
            
            If mMonthIndex > 3 Then
                txtD2.Text = CheckDateInMMM(DateSerial(gbFinancialYearID, mMonthIndex + 1, 1) - 1)
            Else
                txtD2.Text = CheckDateInMMM(DateSerial(gbFinancialYearID + 1, mMonthIndex + 1, 1) - 1)
            End If
            
            txtD3.Text = txtD1.Text
            txtD4.Text = txtD2.Text
            '---------------------------------------------------------------------------------'
            'Note:- Filling Grids
            '---------------------------------------------------------------------------------'
            Call FillVoucherStatement
            Call FillBankStatement
            
        Else
            MsgBox "Please Select The Bank before Selecting the Month", vbInformation
            Exit Sub
        End If
    End Sub
    
    Private Sub cmbVoucherType_Click()
        Call FillVoucherStatement
    End Sub
    
Private Sub cmdAuto_Click()
    frmBankReconciliationProcessNew.Visible = True
   frmBankReconciliationProcessNew.ZOrder (0)
End Sub

Private Sub cmdAutoRecon_Click()
    '-------------------------------------------------------------'
    ' To Analyse Bank Reconciliation By Comparing Three tables    '
    '-------------------------------------------------------------'
        If gbSeatGroupID = 100 Then
            If MsgBox("Do you want to Run Auto Reconciliation process?!", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Permission denied for this operation", vbInformation
            Exit Sub
        End If
    
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim mAmt As Double
        Dim mBankAmt As Double
        Dim mLoopControl As Long
        Dim RecBank As New ADODB.Recordset
        Dim Rec As New ADODB.Recordset
        Dim mStr As String
        Dim mDrOrCrFlag As Integer
        Dim mVDate As Date
        Dim mOVDate As Date
        Dim mOInst As String
        Dim mInst As String
        Dim mD1 As Date
        Dim mD2 As Date
        
        If mSearchID < 1 Then
            MsgBox "Please select the Bank..!", vbInformation
            Exit Sub
        End If
        If cmbMonth.ListIndex > 0 Then
            If cmbMonth.ItemData(cmbMonth.ListIndex) Then
                If cmbMonth.ItemData(cmbMonth.ListIndex) < 4 Then
                    mD1 = DateSerial(gbFinancialYearID + 1, cmbMonth.ItemData(cmbMonth.ListIndex), 1)
                    mD2 = DateAdd("m", 1, mD1) - 1
                Else
                    mD1 = DateSerial(gbFinancialYearID, cmbMonth.ItemData(cmbMonth.ListIndex), 1)
                    mD2 = DateAdd("m", 1, mD1) - 1
                End If
            Else
                MsgBox "@Select the month ,please", vbInformation
            End If
        Else
            MsgBox "Select the month, please!", vbInformation
            Exit Sub
        End If
        
        If Not (IsDate(mD1) And IsDate(mD2)) Then
            MsgBox " Please select the month!", vbInformation
            Exit Sub
        End If
        mSql = " Select * From faBankReconciliationEntries Where intBankAccountHeadID = " & mSearchID
        mSql = mSql + " And tnyReconciled IS NULL "
        mSql = mSql + " And dtBankEntryDate Between '" & DdMmmYy(mD1) & "' AND '" & DdMmmYy(mD2) & "'"
        mSql = mSql + " Order By numTockenID"
        objdb.SetConnection mCnn
        RecBank.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        FileInitialize
        While Not RecBank.EOF
            If Not IsNull(RecBank!intReconciliationID) Then
                If IsNumeric(RecBank!fltDrAmount) Then
                    mBankAmt = RecBank!fltDrAmount
                    mDrOrCrFlag = 1
                Else
                    mBankAmt = RecBank!fltCrAmount
                    mDrOrCrFlag = 0
                End If
                
                mSql = " Select fltAmount,dtInstrumentDate, vchInstrumentNo From faOpeningVouchers Where intAccountHeadID = 1506 "
                mSql = mSql + " And vchInstrumentNo = '" & RecBank!vchChequeNo & "'"
                'mSQL = mSQL + " And tinDebitOrCreditFlag = " & mDrOrCrFlag
                'mSQL = mSQL + " And dtInstrumentDate <= " & RecBank!dtBankEntryDate
                'Rec.Close
                
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                While Not Rec.EOF
                    mAmt = mAmt + Rec!fltAmount
                    mOVDate = Rec!dtInstrumentDate
                    mOInst = Rec!vchInstrumentNo
                    Rec.MoveNext
                    mStr = "O"
                Wend
                Rec.Close
                
                If mDrOrCrFlag Then
                    mDrOrCrFlag = 20
                Else
                    mDrOrCrFlag = 10
                End If
                mSql = " Select fltAmount, dtDate,vchInstrumentNo From faVouchers Where intKeyID1 = 1506 "
                mSql = mSql + " And vchInstrumentNo = '" & RecBank!vchChequeNo & "'"
                mSql = mSql + " And tnyVoucherTypeId  = " & mDrOrCrFlag
                Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                While Not Rec.EOF
                    mAmt = mAmt + Rec!fltAmount
                    
                    mVDate = Rec!dtDate
                    mInst = Rec!vchInstrumentNo
                    mStr = mStr + "C"
                    Rec.MoveNext
                Wend
                If mBankAmt = mAmt Then
                    Print #gbFileNO, RecBank!intReconciliationID, mBankAmt, mAmt, RecBank!vchChequeNo, RecBank!dtBankEntryDate, mInst, mVDate, mOInst, mOVDate
                    
                    
                    
                    mSql = " Update faOpeningVouchers"
                    mSql = mSql + " Set tnyReconciled = 2"
                    mSql = mSql + ", numTockenID = " & ""
                    mSql = mSql + ", dtRealisationDate = ''"
                    mSql = mSql + ", vchRemarks = 'Auto Reconciled'"
                    mSql = mSql + "  Where intAccountHeadID = 1506 "
                    mSql = mSql + " And vchInstrumentNo = '" & RecBank!vchChequeNo & "'"
                    
                    
                    
                    mSql = " Update faVouchers "
                    mSql = mSql + " Set tnyReconciled = 2"
                    mSql = mSql + ", numTockenID = " & ""
                    mSql = mSql + ", dtRealisationDate = ''"
                    mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                    mSql = mSql + " Where intVoucherID =  "
                    
                
                    
                    
                    mSql = " Update faOpeningVouchers Set tnyReconciled = 0 Where intAccountHeadID = 1506 AND numTockenID = " & RecBank!intReconciliationID
                    mCnn.Execute mSql
                    
                    mSql = " Update faVouchers Set tnysync=Null,tnyReconciled = 0  Where  intKeyID1 = 1506 AND numTockenID = " & RecBank!intReconciliationID
                    mCnn.Execute mSql
                    
                    mSql = "Update faBankReconciliationEntries Set tnyReconciled = 0 Where intBankAccountHeadID = 1506 And intReconciliationID = " & RecBank!intReconciliationID
                    mCnn.Execute mSql
                End If
                mAmt = 0
                mStr = ""
                Rec.Close
            End If
            RecBank.MoveNext
        Wend
        RecBank.Close
        Close #gbFileNO
        ShellPad
End Sub

    Private Sub cmdAutoReconcile_Click()
        If gbSeatGroupID = 100 Then
            If MsgBox("Do you want to Run Auto Reconciliation process?!", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                Call AutoReconcile
            End If
        Else
            MsgBox "Permission denied for this operation", vbInformation
        End If
    End Sub

    Private Sub cmdBankSearch_Click()
        Dim mSql As String
        Dim mCount As Integer
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        mCount = InStr(1, gbSearchStr, " ")
        mSearchID = gbSearchID
        txtBankCode.Text = IIf(IsNull(Left(gbSearchStr, mCount)), "", Left(gbSearchStr, mCount))
        If mCount <> 0 Then
            txtBankName.Text = IIf(IsNull(mID(gbSearchStr, mCount)), "", mID(gbSearchStr, mCount))
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub
    
    Private Sub cmdReconcile_Click()
'        If chkMonth.Value = 1 Then
'            MsgBox "Please Uncheck - List This Month Olny - Caption for Better Reconciliation", vbInformation
'            Exit Sub
'        End If
'         If txtBankCode.Text <> "" And cmbMonth.ListIndex <> -1 Then
'            Dim mVchrRowCount As Long
'            Dim mBankRowCount As Long
'            Dim mSQL As String
'            Dim Rec As New ADODB.Recordset
'            Dim objDB As New clsDB
'            Dim mCnn As New ADODB.Connection
'            Dim mTotCount As Long
'            Dim mSumAmount As Double
'                objDB.SetConnection mCnn
'            pbBank.Value = 0
'            mSumAmount = 0
'            mTotCount = fgBankStatement.Rows
'            pbBank.Max = mTotCount
'
'            For mBankRowCount = 1 To fgBankStatement.Rows - 2
'                For mVchrRowCount = 1 To fgVoucherStatement.Rows - 2
'                    If Trim(fgBankStatement.TextMatrix(mBankRowCount, 2)) = Trim(fgVoucherStatement.TextMatrix(mVchrRowCount, 3)) Then
'                        mSumAmount = mSumAmount + fgVoucherStatement.TextMatrix(mVchrRowCount, 4)
'                    Else
'                        On Error Resume Next
'                        If fgBankStatement.TextMatrix(mBankRowCount, 3) = mSumAmount Then
'                            mSQL = "Set DateFormat DMY Update faBankReconciliationEntries Set tnyReconciled = 1 ,intVoucherNo = " & fgVoucherStatement.TextMatrix(mVchrRowCount, 1) & " Where vchChequeNo = '" & fgBankStatement.TextMatrix(mBankRowCount, 2) & "'"
'                            mCnn.Execute mSQL
'                            mSQL = "Set DateFormat DMY Update faVouchers Set tnyReconciled = 1 , numTockenID = " & fgBankStatement.TextMatrix(mBankRowCount, 0) & " , dtRealisationDate = '" & fgBankStatement.TextMatrix(mBankRowCount, 1) & "' Where vchInstrumentNo = '" & fgBankStatement.TextMatrix(mBankRowCount, 2) & "'"
'                            mCnn.Execute mSQL
'                        End If
'                        mSumAmount = 0
'                    End If
'                Next
'                pbBank.Value = pbBank.Value + 1
'            Next
'            Call FillBankStatement
'            Call FillVoucherStatement
'        Else
'            MsgBox "Please Select the Month and Bank", vbQuestion
'            cmdBankSearch.SetFocus
'        End If
    End Sub
    
    Private Sub cmdReconcileVoucher_Click()
        If mSearchID < 1 Then
            MsgBox "Please select the Bank..!", vbInformation
            Exit Sub
        End If
        If val(vsTitleGrid.Tag) > 0 Then
            mvarManuallyReconciled = False
            frmManualReconcile.VoucherFlag = True
            frmManualReconcile.txtVoucherNo.Tag = vsTitleGrid.TextMatrix(1, 0)
            frmManualReconcile.VoucherFlag = True
            frmManualReconcile.Show vbModal
        End If
    End Sub

    Private Sub cmdReport_Click()
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mDate As Date
        Dim mD1 As Date
        Dim mD2 As Date
        
        Dim mDrAmt As Double
        Dim mCrAmt As Double
        
        Dim mDrRAmt As Double
        Dim mCrPAmt As Double
        
        Dim mDrBAmt As Double
        Dim mCrBAmt As Double
        
        Dim mBankBalance As Double
        Dim mBankPassBookBalance As Double
        
        Dim mInput As Variant
        
        If IsDate(txtD1) Then
            mD1 = txtD1.Text
        Else
            mD1 = gbStartingDate
        End If
        If IsDate(txtD2) Then
            mD2 = txtD2.Text
        Else
            mD2 = gbEndingDate
        End If
        
        objdb.SetConnection mCnn
        If IsDate(txtD2.Text) Then
            mSql = "Select isNull(Sum(fltCrAmount),0)-isNull(Sum( fltDrAmount),0) fltAmount From faBankReconciliationEntries "
            mSql = mSql + " Where tnyOpening = 0 AND intBankAccountHeadID = " & mSearchID & " And dtBankEntryDate <= '" & txtD2 & "'"
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
            If Not (Rec.BOF And Rec.EOF) Then
                mBankPassBookBalance = Format(Rec!fltAmount, "0.00")
                Debug.Print Format(Rec!fltAmount, "0.00")
            End If
            Rec.Close
            
            mSql = "Select fltOpening * ((tinDebitOrCreditFlag*2)-1 )fltOpening From faBanks Where intAccountHeadID = " & mSearchID
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
            If Not (Rec.BOF And Rec.EOF) Then
                mBankPassBookBalance = Format(Rec!fltOpening, "0.00") + mBankPassBookBalance
                Debug.Print Format(Rec!fltOpening, "0.00")
                Debug.Print mBankPassBookBalance
            End If
            Rec.Close
        End If
        
        mInput = Array(mSearchID, mD1, mD2)
        Set Rec = objdb.ExecuteSP("spGetClosingBalance", mInput, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            mBankBalance = Rec!NetBalance
        Else
            MsgBox "Bank Balance Didn't able to find as per the date!!", vbInformation
            'Exit Sub
        End If
        Rec.Close
        FileInitialize
        'mBankBalance = 1002900 '26141223.58
        'mBankPassBookBalance = 24558717.71
        
        mSql = ""
        mSql = mSql + "       SELECT      dbo.faOpeningVouchers.intTransactionID intTransactionID    ,"
        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherID      intVoucherID   ,"
        mSql = mSql + "           Null intBookNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherNo intVoucherNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.vchInstrumentNo vchInstrumentNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.tnyVoucherTypeID tnyVoucherTypeID,"
        mSql = mSql + "           dbo.faOpeningVouchers.tnyReconciled tnyReconciled,"
        mSql = mSql + "           dbo.faOpeningVouchers.numTockenID numTockenID,"
        mSql = mSql + "           dbo.faOpeningVouchers.dtDate dtDate,"
        mSql = mSql + "           Null vchGroup         ,"
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then"
        mSql = mSql + "               dbo.faOpeningVouchers.fltAmount * -1"
        mSql = mSql + "           End fltCrAmount,"
        mSql = mSql + "       Case when tinDebitOrCreditFlag = 1 then"
        mSql = mSql + "              dbo.faOpeningVouchers.fltAmount"
        mSql = mSql + "           End fltDrAmount,"
        mSql = mSql + "       dbo.faOpeningVouchers.tinDebitOrCreditFlag tinDebitOrCreditFlag,"
        mSql = mSql + "           dbo.faOpeningVouchers.vchNarration vchNarration,"
        mSql = mSql + "           1 tnyOpeningFlag, "
        mSql = mSql + "           0 fltAmount "
        mSql = mSql + "       From dbo.faOpeningVouchers"
        mSql = mSql + "       Where faOpeningVouchers.intAccountHeadID = " & mSearchID
'        mSql = mSql + "       And tnyReconciled is Null Or (tnyReconciled is Not Null And dtRealisationDate > '" & DdMmmYy(mD2) & "')"
        mSql = mSql + "       And tnyReconciled is Null Or (tnyReconciled is Not Null And dtRealisationDate > '" & DdMmmYy(mD2) & "')  Or (tnyReconciled = 8 And dtRealisationDate <='" & DdMmmYy(mD2) & "' And faOpeningVouchers.intAccountHeadID = " & mSearchID & ")"
        mSql = mSql + "       Union All "
        
        'mSQL = ""
        mSql = mSql + "       SELECT"
        mSql = mSql + "          dbo.faTransactionChild.intTransactionID     , "
        mSql = mSql + "           dbo.faTransactions.intVoucherID         , "
        mSql = mSql + "           dbo.faVouchers.intBookNo, "
        mSql = mSql + "           dbo.faVouchers.intVoucherNo, "
        mSql = mSql + "           dbo.faVouchers.vchInstrumentNo, "
        mSql = mSql + "           dbo.faVouchers.tnyVoucherTypeID, "
        mSql = mSql + "           dbo.faVouchers.tnyReconciled, "
        mSql = mSql + "           dbo.faVouchers.numTockenID, "
        
        mSql = mSql + "           dbo.faVouchers.dtDate, "
        mSql = mSql + "           dbo.faTransactions.vchGroup         , "
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then "
        mSql = mSql + "               dbo.faTransactionChild.fltAmount * -1 "
        mSql = mSql + "           End fltCrAmount, "
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 1 then "
        mSql = mSql + "              dbo.faTransactionChild.fltAmount "
        mSql = mSql + "           End fltDrAmount, "
        mSql = mSql + "           dbo.faTransactionChild.tinDebitOrCreditFlag , "
        mSql = mSql + "           dbo.faTransactions.vchNarration, "
        mSql = mSql + "           0 tnyOpeningFlag, "
        mSql = mSql + "           dbo.faVouchers.fltAmount "
        mSql = mSql + "        FROM        dbo.faTransactionChild      "
        mSql = mSql + "        INNER JOIN  dbo.faTransactions  ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID"
        mSql = mSql + "        INNER  JOIN  dbo.faVouchers          ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID"
        mSql = mSql + "        LEFT  JOIN  faBankReconciliationEntries ON faBankReconciliationEntries.intReconciliationID = faTransactionChild.numTockenID"
        mSql = mSql + "        WHERE   ( "
        mSql = mSql + "                 dbo.faTransactionChild.intAccountHeadID = " & mSearchID
        mSql = mSql + "                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL) "
        mSql = mSql + "                 AND faTransactions.intTransactionID > 0"
        mSql = mSql + "                 AND faTransactions.dtTransactionDate <= '" & DdMmmYy(mD2) & "'"
              mSql = mSql + "                 AND faTransactionChild.numTockenID is Null Or (faTransactionChild.numTockenID is Not Null And faBankReconciliationEntries.dtBankEntryDate > '" & DdMmmYy(mD2) & "' And faTransactions.dtTransactionDate <= '" & DdMmmYy(mD2) & "')"
        'mSql = mSql + "                 AND faTransactionChild.numTockenID is Null"
        mSql = mSql + "                 )"
        mSql = mSql + "     Order By   dtDate,tnyVoucherTypeID, faVouchers.intVoucherID"
        
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            
            Print #gbFileNO, "-----------------------------------------------------------------------------------"
            Print #gbFileNO, " Date           Voucher No   Type          Debit Amt      Credit Amt    Voucher Amt"
            Print #gbFileNO, "-----------------------------------------------------------------------------------"
            While Not Rec.EOF
                If IIf(IsNull(Rec!tnyReconciled), 0, Rec!tnyReconciled) = 8 Then
                    If Rec!tnyVoucherTypeID = 10 Then mDrRAmt = mDrRAmt - IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                    If Rec!tnyVoucherTypeID = 20 Then mCrPAmt = mCrPAmt - IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                Else
                    Select Case Rec!tnyVoucherTypeID
                        Case Is = 10: mDrRAmt = mDrRAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                        Case Is = 20: mCrPAmt = mCrPAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                        Case Else:
                        If Rec!intVoucherNo = 31000051 Or Rec!intVoucherNo = 31000052 Then
                            Debug.Print ""
                        End If
                            If IsNumeric(Rec!fltDrAmount) Then
                                mDrRAmt = mDrRAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                            ElseIf IsNumeric(Rec!fltCrAmount) Then
                                mCrPAmt = mCrPAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                            End If
                    End Select
                End If
                mDrAmt = mDrAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                mCrAmt = mCrAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                
                Print #gbFileNO, Rec!dtDate, Rec!intVoucherNo, Rec!tnyVoucherTypeID, PadL(Format(Rec!fltDrAmount, "0.00"), 11), PadL(Format(Abs(Rec!fltCrAmount), "0.00"), 11), PadL(Format(Rec!fltAmount, "0.00"), 12)
                
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        
        Print #gbFileNO, "==================================================================================="
        Print #gbFileNO, "   BANK SCROLL"
        Print #gbFileNO, "==================================================================================="
        
        mSql = "        Select intReconciliationID, "
        mSql = mSql + "       dtBankEntryDate,"
        mSql = mSql + "       intVoucherID,"
        mSql = mSql + "       intVoucherNo,"
        mSql = mSql + "       vchChequeNo,"
        mSql = mSql + "       dtChequeDate,"
        mSql = mSql + "       fltDrAmount,"
        mSql = mSql + "       fltCrAmount,"
        mSql = mSql + "       vchParticulars,"
        mSql = mSql + "       vchRemarks"
        mSql = mSql + " From faBankReconciliationEntries"
        mSql = mSql + " Where dtBankEntryDate <= '" & DdMmmYy(mD2) & "'"
        mSql = mSql + " And tnyReconciled is Null"
        mSql = mSql + " And intBankAccountHeadID = " & mSearchID
        
        mSql = mSql + " Order By dtBankEntryDate "
        
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                If Not (IsNull(Rec!fltDrAmount) Or Rec!fltDrAmount = 0) Then
                    mDrBAmt = mDrBAmt + Rec!fltDrAmount
                Else 'If Not (IsNull(Rec!fltCrAmount) Or Rec!fltCrAmount = 0) Then
                    mCrBAmt = mCrBAmt + Rec!fltCrAmount
                End If
                Print #gbFileNO, Rec!dtBankEntryDate, Rec!intReconciliationID, "  ", PadL(Format(Rec!fltDrAmount, "0.00"), 11), PadL(Format(Abs(Rec!fltCrAmount), "0.00"), 11)
                Rec.MoveNext
            Wend
        End If
        Print #gbFileNO, "-----------------------------------------------------------------------------------"
            
        Print #gbFileNO, "                Bank Balance As On : "; PadL(Format(mBankBalance, "0.00"), 12)
        Print #gbFileNO, "   Cheque Issued But Not Presented : "; PadL(Format(Abs(mCrPAmt), "0.00"), 12)
        Print #gbFileNO, "         Directly Credited by Bank : "; PadL(Format(mCrBAmt, "0.00"), 12)
        
        Print #gbFileNO, "Cheque Deposited But Not Collected : "; PadL(Format(mDrRAmt, "0.00"), 12)
        Print #gbFileNO, "          Directly Debited by Bank : "; PadL(Format(mDrBAmt, "0.00"), 12)
        Print #gbFileNO, "                 Pass Book Balance : "; PadL(Format(mBankPassBookBalance, "0.00"), 12)
        
        Print #gbFileNO, "                                   : "; PadL(Format(mBankBalance + Abs(mCrPAmt) + mCrBAmt - mDrRAmt - mDrBAmt, "0.00"), 12)
        Close #gbFileNO
        ShellPad
    End Sub
        
    Private Sub Command1_Click()
        Call AutoReconcile
    End Sub
        
    Private Sub Command2_Click()
        Call ReconciliationReport(mSearchID, txtD1, txtD2)
    End Sub

    Private Sub fgBankStatement_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 8 Then
            mAdjdiff = val(fgBankStatement.Cell(flexcpText, Row, 9))
            If fgBankStatement.Cell(flexcpChecked, Row, 5) <> vbChecked Then
                If fgBankStatement.Cell(flexcpChecked, Row, 8) = vbChecked Then
                    If val(fgBankStatement.Cell(flexcpText, Row, 4)) = 0 Then
                        'mSelectedScroll = mSelectedScroll + (val(fgBankStatement.Cell(flexcpText, Row, 3)) + mAdjdiff) + (val(fgBankStatement.Cell(flexcpText, Row, 4)) + mAdjdiff) * -1
                        mSelectedScroll = mSelectedScroll + (val(fgBankStatement.Cell(flexcpText, Row, 3)) + mAdjdiff)
                    Else
                       mSelectedScroll = mSelectedScroll + (val(fgBankStatement.Cell(flexcpText, Row, 4)) + mAdjdiff) * -1
                    End If
                    mBankEntryDate = Format(fgBankStatement.TextMatrix(fgBankStatement.Row, 1), "dd/mmm/yyyy")
                    
                Else
                    If val(fgBankStatement.Cell(flexcpText, Row, 4)) = 0 Then
                        mSelectedScroll = mSelectedScroll - (val(fgBankStatement.Cell(flexcpText, Row, 3)) + mAdjdiff)
                        'mSelectedScroll = mSelectedScroll - val(fgBankStatement.Cell(flexcpText, Row, 3)) - val(fgBankStatement.Cell(flexcpText, Row, 4)) * -1
                    Else
                        mSelectedScroll = mSelectedScroll - (val(fgBankStatement.Cell(flexcpText, Row, 4)) + mAdjdiff) * -1
                    End If
                End If
            End If
            lblSelectedScrollAmt.Caption = Format(Abs(mSelectedScroll), "0.00")
        End If
    End Sub

    Private Sub fgBankStatement_DblClick()
        Dim mLoop As Long
        If fgBankStatement.Row = -1 Then Exit Sub
        vsTitleGrid.Clear 1, 1
        If fgBankStatement.Row = -1 Then Exit Sub
        If fgBankStatement.Row > 0 Then
            vsTitleGrid.Tag = fgBankStatement.Row
            For mLoop = 0 To fgBankStatement.Cols - 1
                vsTitleGrid.TextMatrix(1, mLoop) = fgBankStatement.TextMatrix(fgBankStatement.Row, mLoop)
            Next mLoop
        End If
        
        If Trim(txtVInstNo) = "" Or mSelectedAmt = 0 Then
            mSelectedAmt = 0
            txtVInstNo.Text = fgBankStatement.TextMatrix(fgBankStatement.Row, 2)
            Call txtVInstNo_LostFocus
            fgBankStatement.Row = vsTitleGrid.Tag
        End If
        '----
        fgBankStatement.Select 1, 8, fgBankStatement.Rows - 1, 7
        fgBankStatement.Clear 2, 0
        fgBankStatement.Select 0, 0
        mSelectedScroll = 0
        lblSelectedScrollAmt = 0
    End Sub
    
    Private Sub fgBankStatement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 Then
            If fgBankStatement.Cell(flexcpChecked, fgBankStatement.Row, 5) = vbChecked Then
                mnuUnReconcile.Visible = True
                mnuManuallyReconcile.Visible = False
            Else
                mnuUnReconcile.Visible = False
                mnuManuallyReconcile.Visible = True
            End If
            mnuVoucherMutual.Visible = False
            mnuVoucherMutualUnReconcile.Visible = False
            Call PopupMenu(mnuPopup)
        End If
    End Sub
    
    Private Sub fgBankStatement_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If fgBankStatement.Col = 8 Then
            fgBankStatement.Editable = flexEDKbdMouse
            If fgBankStatement.Cell(flexcpChecked, fgBankStatement.Row, 5) = vbChecked Then
                fgBankStatement.Cell(flexcpChecked, fgBankStatement.Row, fgBankStatement.Col) = vbUnchecked
            End If
        Else
            fgBankStatement.Editable = flexEDNone
        End If
    End Sub

    Private Sub fgVoucherStatement_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 7 Then
            If fgVoucherStatement.Cell(flexcpChecked, Row, 7) = vbChecked Then
                mSelectedAmt = mSelectedAmt + val(fgVoucherStatement.Cell(flexcpText, Row, 4)) + val(fgVoucherStatement.Cell(flexcpText, Row, 5)) * -1
            Else
                mSelectedAmt = mSelectedAmt - val(fgVoucherStatement.Cell(flexcpText, Row, 4)) - val(fgVoucherStatement.Cell(flexcpText, Row, 5)) * -1
            End If
            lblSelectedAmt.Caption = Format(Abs(mSelectedAmt), "0.00")
        End If
    End Sub
    
    Private Sub fgVoucherStatement_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        'NOTE:- An attempt to Select a Reconciled Item again then it will Not Allowed to Select again
        If Col = 7 Then
            If fgVoucherStatement.Cell(flexcpChecked, Row, 6) = vbChecked Then
                Cancel = True
            End If
        End If
    End Sub
    
    Private Sub fgVoucherStatement_Click()
        If fgVoucherStatement.Row = -1 Then Exit Sub
        If fgVoucherStatement.Col = 7 Or fgVoucherStatement.Col = 13 Then
'            vsTitleGrid.Tag = fgVoucherStatement.Row
            fgVoucherStatement.Editable = flexEDKbdMouse
            If fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, 5) = vbChecked Then
                fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, fgVoucherStatement.Col) = vbUnchecked
            End If
        Else
            fgVoucherStatement.Editable = flexEDNone
        End If
        If fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, 6) = vbChecked Or fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, 7) = vbChecked Then
            fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, 13) = vbUnchecked
        End If
    ''''    If fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 3) <> "" Then
    ''''        On Error Resume Next
    ''''        frmSearchForBankReconciliation.mInstNO = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 3)
    ''''        frmSearchForBankReconciliation.Show vbModal
    '''''''        Call FillVoucherStatement
    '''''''        Call FillBankStatement
    ''''    Else
    ''''        frmSearchForBankReconciliation.mAmt = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 4)
    ''''        frmSearchForBankReconciliation.Show vbModal
    ''''    End If
    ''''    If frmSearchForBankReconciliation.Flag = 1 Then
    ''''        fgVoucherStatement.Cell(flexcpBackColor, fgVoucherStatement.Row, 0, fgVoucherStatement.Row, 6) = &HFF00&
    ''''    ElseIf frmSearchForBankReconciliation.Flag = 2 Then
    ''''        fgVoucherStatement.Cell(flexcpBackColor, fgVoucherStatement.Row, 0, fgVoucherStatement.Row, 6) = &H8080FF
    ''''    End If
    End Sub
    
    Private Sub fgVoucherStatement_DblClick()
        '-------------------------------------------------------------------------------------------------------------------'
        '    If fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 3) <> "" Then                                         '
        '        On Error Resume Next                                                                                       '
        '        frmSearchForBankReconciliation.mInstNO = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 3)          '
        '        frmSearchForBankReconciliation.VoucherID = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 7)        '
        '        frmSearchForBankReconciliation.Show vbModal                                                                '
        '''        Call FillVoucherStatement                                                                                '
        '''        Call FillBankStatement                                                                                   '
        '    Else                                                                                                           '
        '        frmSearchForBankReconciliation.mAmt = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 4)             '
        '        frmSearchForBankReconciliation.VoucherID = fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 7)        '
        '        frmSearchForBankReconciliation.Show vbModal                                                                '
        '    End If                                                                                                         '
        '    If frmSearchForBankReconciliation.Flag = 1 Then                                                                '
        '        fgVoucherStatement.Cell(flexcpBackColor, fgVoucherStatement.Row, 0, fgVoucherStatement.Row, 6) = &HFF00&   '
        '    ElseIf frmSearchForBankReconciliation.Flag = 2 Then                                                            '
        '        fgVoucherStatement.Cell(flexcpBackColor, fgVoucherStatement.Row, 0, fgVoucherStatement.Row, 6) = &H8080FF  '
        '    End If                                                                                                         '
        '-------------------------------------------------------------------------------------------------------------------'
    End Sub
    
    Private Sub fgVoucherStatement_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        KeyAscii = 0
    End Sub

    Private Sub fgVoucherStatement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call menuAllVisible(False)
        If Button = 2 Then
            mnuUnReconcile.Visible = False
            mnuManuallyReconcile.Visible = False
            If val(fgVoucherStatement.Cell(flexcpChecked, fgVoucherStatement.Row, 13)) = vbChecked Then
                mnuVoucherMutual.Visible = True
                mnuVoucherMutualUnReconcile.Visible = False
            ElseIf fgVoucherStatement.Cell(flexcpText, fgVoucherStatement.Row, 12) = 5 Then
                mnuVoucherMutual.Visible = False
                mnuVoucherMutualUnReconcile.Visible = True
            End If
            Call PopupMenu(mnuPopup)
        End If
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub
    
    Private Sub Form_Load()
        cmbVoucherType.AddItem ""
        cmbVoucherType.AddItem "Receipts"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 10
        cmbVoucherType.AddItem "Payments"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 20
        cmbVoucherType.AddItem "Contra"
        cmbVoucherType.ItemData(cmbVoucherType.NewIndex) = 30
        txtD1.Text = DdMmmYy(gbStartingDate)
        txtD2.Text = DdMmmYy(gbTransactionDate)
        Call fillMonthCombo
        Call FillRoundingAdjustment
        If gbLBID = 167 Then
        
            cmdAuto.Visible = True
        Else
            cmdAuto.Visible = False
        End If
    End Sub
    Private Sub FillRoundingAdjustment()
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSql    As String
        
'        IF NOT EXISTS(SELECT *
'          From INFORMATION_SCHEMA.Columns
'          WHERE  TABLE_NAME = 'fabankReconciliationEntries'
'                 AND COLUMN_NAME = 'fltRoundAdj')
'Alter Table fabankReconciliationEntries
'add fltRoundAdj  Numeric(18,2)
'GO
        objdb.SetConnection mCnn
        mSql = "Update fabankReconciliationEntries set fltRoundAdj=1-(ABS(fltCrAmount) - FLOOR(ABS(fltCrAmount)))"
        mSql = mSql + " Where (Abs(fltCrAmount) - FLOOR(Abs(fltCrAmount))) > 0 and  isnull(tnyReconciled,0)=0"
        mCnn.Execute mSql
        
        mSql = ""
        mSql = "Update fabankReconciliationEntries set fltRoundAdj=1-(ABS(fltDrAmount) - FLOOR(ABS(fltDrAmount)))"
        mSql = mSql + " Where (Abs(fltDrAmount) - FLOOR(Abs(fltDrAmount))) > 0 and  isnull(tnyReconciled,0)=0"
        mCnn.Execute mSql
        ''''Select fltCrAmount,fltDrAmount,* From fabankReconciliationEntries Where intBankAccountHeadID=1661
        ''''And (ABS(fltCrAmount) - FLOOR(ABS(fltCrAmount))) >0
        
        ''''Update fabankReconciliationEntries set fltRoundAdj=1-(ABS(fltCrAmount) - FLOOR(ABS(fltCrAmount)))
        ''''Where intBankAccountHeadID = 1661
        ''''And (ABS(fltCrAmount) - FLOOR(ABS(fltCrAmount))) >0
    End Sub
    
    Private Sub fillMonthCombo()
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
    End Sub
    
    Private Sub FillVoucherStatement()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mRowCount As Long
        Dim mD1 As Date
        Dim mD2 As Date
        Dim arrInput As Variant
        mSelectedAmt = 0
        lblSelectedAmt.Caption = Format(mSelectedAmt, "0.00")
        chkUnReconciledVouchers.Value = 0
        
        If cmbMonth.ListIndex = -1 Then
            MsgBox "Plaese Select Month", vbApplicationModal
            Exit Sub
        End If
        objdb.SetConnection mCnn
    
'''        mSql = ""
'''        mSql = mSql + " Select intVoucherID, tnyVoucherTypeID, intVoucherNo, vchInstrumentNo, dtDate, fltAmount, tnyReconciled, numTockenID from( "
'''        mSql = mSql + " Select intVoucherID, tnyVoucherTypeID, intVoucherNo, vchInstrumentNo, dtDate, fltAmount, tnyReconciled, numTockenID from faVouchers  Where "
'''            If chkMonth.Value = 0 Then
'''                mSql = mSql + " Month(dtDate) Between 4 And " & cmbMonth.ItemData(cmbMonth.ListIndex)
'''            Else
'''                mSql = mSql + " Month(dtDate) = " & cmbMonth.ItemData(cmbMonth.ListIndex)
'''            End If
'''        mSql = mSql + " And tnyReconciled is Null and tnyVoucherTypeID < 40  And (intKeyID1 = " & val(mSearchID) & " Or intKeyID1 is Null )"
'''        mSql = mSql + " Union All"
'''        mSql = mSql + " Select faVouchers.intVoucherID, tnyVoucherTypeID, intVoucherNo, vchInstrumentNo, dtDate, faVoucherChild.fltAmount, tnyReconciled, numTockenID From faVouchers"
'''        mSql = mSql + " Inner Join faVoucherChild On faVoucherChild.intVoucherID = faVouchers.intVoucherID Where"
'''            If chkMonth.Value = 0 Then
'''                mSql = mSql + " Month(dtDate) Between 4 And " & cmbMonth.ItemData(cmbMonth.ListIndex)
'''            Else
'''                mSql = mSql + " Month(dtDate) = " & cmbMonth.ItemData(cmbMonth.ListIndex)
'''            End If
'''        mSql = mSql + " And tnyReconciled is Null and tnyVoucherTypeID < 40  And (faVoucherChild.intAccountHeadID = " & val(mSearchID) & " ) "
'''        mSql = mSql + " ) A"
'''        If IsDate(txtD3) Then mD1 = txtD3 Else mD1 = gbStartingDate
'''        If IsDate(txtD4) Then mD2 = txtD4 Else mD2 = gbEndingDate
'''
'''        mSql = ""
'''        mSql = mSql + "       SELECT     Distinct dbo.faOpeningVouchers.intTransactionID intTransactionID    ,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherID      intVoucherID   ,"
'''        mSql = mSql + "           Null intBookNo,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherNo intVoucherNo,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.vchInstrumentNo vchInstrumentNo,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.tnyVoucherTypeID tnyVoucherTypeID,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.tnyReconciled tnyReconciled,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.numTockenID numTockenID,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.dtDate dtDate,"
'''        mSql = mSql + "           Null vchGroup         ,"
'''        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then"
'''        mSql = mSql + "               dbo.faOpeningVouchers.fltAmount * -1"
'''        mSql = mSql + "           End fltCrAmount,"
'''        mSql = mSql + "       Case when tinDebitOrCreditFlag = 1 then"
'''        mSql = mSql + "              dbo.faOpeningVouchers.fltAmount"
'''        mSql = mSql + "           End fltDrAmount,"
'''        mSql = mSql + "       dbo.faOpeningVouchers.tinDebitOrCreditFlag tinDebitOrCreditFlag,"
'''        mSql = mSql + "           dbo.faOpeningVouchers.vchNarration vchNarration,"
'''        mSql = mSql + "           1 tnyOpeningFlag"
'''        mSql = mSql + "       From dbo.faOpeningVouchers"
'''        mSql = mSql + "       Where  intAccountHeadID = " & mSearchID
'''        mSql = mSql + "       Union All "
'''
'''        mSql = mSql + "       SELECT Distinct "
'''        mSql = mSql + "          dbo.faTransactionChild.intTransactionID     , "
'''        mSql = mSql + "           dbo.faTransactions.intVoucherID         , "
'''        mSql = mSql + "           dbo.faVouchers.intBookNo, "
'''        mSql = mSql + "           dbo.faVouchers.intVoucherNo, "
'''        mSql = mSql + "           dbo.faVouchers.vchInstrumentNo, "
'''        mSql = mSql + "           dbo.faVouchers.tnyVoucherTypeID, "
'''        mSql = mSql + "           dbo.faVouchers.tnyReconciled, "
'''        mSql = mSql + "           dbo.faTransactionChild.numTockenID, "
'''
'''        mSql = mSql + "           dbo.faVouchers.dtDate, "
'''        mSql = mSql + "           dbo.faTransactions.vchGroup         , "
'''        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then "
'''        mSql = mSql + "               dbo.faTransactionChild.fltAmount * -1 "
'''        mSql = mSql + "           End fltCrAmount, "
'''        mSql = mSql + "           Case when tinDebitOrCreditFlag = 1 then "
'''        mSql = mSql + "              dbo.faTransactionChild.fltAmount "
'''        mSql = mSql + "           End fltDrAmount, "
'''        mSql = mSql + "           dbo.faTransactionChild.tinDebitOrCreditFlag , "
'''        mSql = mSql + "           dbo.faTransactions.vchNarration, "
'''        mSql = mSql + "           0 tnyOpeningFlag"
'''        mSql = mSql + "        FROM dbo.faTransactionChild      "
'''        mSql = mSql + "        INNER JOIN  dbo.faTransactions  ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID"
'''        mSql = mSql + "        INNER JOIN  dbo.faVouchers ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID"
'''        mSql = mSql + "        WHERE   ( "
'''        mSql = mSql + "                 dbo.faTransactionChild.intAccountHeadID = " & mSearchID
'''        mSql = mSql + "                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL) "
'''        mSql = mSql + "                 AND faTransactions.intTransactionID > 0"
'''        mSql = mSql + "             AND faTransactions.dtTransactionDate Between '" & DdMmmYy(mD1) & "' AND '" & DdMmmYy(mD2) & "'"
'''
'''        If cmbVoucherType.ListIndex > 0 Then
'''            mSql = mSql + "             AND faTransactions.intGroupID = " & cmbVoucherType.ItemData(cmbVoucherType.ListIndex)
'''        End If
'''        mSql = mSql + "                 )"
'''        mSql = mSql + "     Order By dtDate" ', faVouchers.intVoucherID"
        
        
        
        If IsDate(txtD3) Then mD1 = txtD3 Else mD1 = gbTransactionDate
        If IsDate(txtD4) Then mD2 = txtD4 Else mD2 = gbTransactionDate
        fgVoucherStatement.AutoSearch = flexSearchFromCursor
        arrInput = Array(mSearchID, CDate(mD1), CDate(mD2))
        Set Rec = objdb.ExecuteSP("spGetReconcileVouchers", arrInput)
        'Rec.Open mSql, mCnn
        
        fgVoucherStatement.Rows = 1
        fgVoucherStatement.Rows = 2
        mRowCount = 1
        
        While Not (Rec.EOF Or Rec.BOF)
            If Rec!tnyVoucherTypeID = 10 Then fgVoucherStatement.TextMatrix(mRowCount, 0) = "R"
            If Rec!tnyVoucherTypeID = 20 Then fgVoucherStatement.TextMatrix(mRowCount, 0) = "P"
            If Rec!tnyVoucherTypeID = 30 Then
                fgVoucherStatement.TextMatrix(mRowCount, 0) = "C"
            End If
            If Rec!tnyVoucherTypeID = 40 Then fgVoucherStatement.TextMatrix(mRowCount, 0) = "J"
            fgVoucherStatement.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            fgVoucherStatement.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!dtDate), "", Format(Rec!dtDate, "dd/mmm/yyyy"))
            fgVoucherStatement.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            
            If Rec!tinDebitOrCreditFlag Then
                fgVoucherStatement.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltDrAmount), "", Rec!fltDrAmount)
            Else
                fgVoucherStatement.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltCrAmount), "", Abs(Rec!fltCrAmount))
            End If
            
            If IsNull(Rec!numTockenID) Then
                fgVoucherStatement.Cell(flexcpChecked, mRowCount, 6) = vbUnchecked
                'fgVoucherStatement.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 6) = &H00D6FFD6&
                
            Else
                fgVoucherStatement.Cell(flexcpChecked, mRowCount, 6) = vbChecked
                fgVoucherStatement.Cell(flexcpBackColor, mRowCount, 0, mRowCount, 6) = &HD6FFD6
            End If
            
            fgVoucherStatement.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!numTockenID), "", Rec!numTockenID)
            fgVoucherStatement.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!vchNarration), "", "  " & Trim(Rec!vchNarration))
            fgVoucherStatement.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
            fgVoucherStatement.TextMatrix(mRowCount, 11) = Rec!tnyOpeningFlag
            fgVoucherStatement.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!tnyReconciled), -1, Rec!tnyReconciled) ''Mutual Reconcilliation = 5 (Voucher)
            mRowCount = mRowCount + 1
            fgVoucherStatement.Rows = fgVoucherStatement.Rows + 1
            Rec.MoveNext
        Wend
    End Sub
    
    Private Sub FillBankStatement()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mRowCount As Long
        Dim mD1 As Date
        Dim mD2 As Date
        
        mAdjdiff = 0
        mSelectedScroll = 0
        objdb.SetConnection mCnn
        
        Rec.CursorLocation = adUseClient
        mSql = "Select intReconciliationID,Convert(varchar(11),dtBankEntryDate,106) dtBankEntryDate ,vchChequeNo, fltDrAmount,fltCrAmount,tnyReconciled, intVoucherNo, fltDifference,0,fltRoundAdj FROM faBankReconciliationEntries"
                                                        
        If IsDate(txtD1) Then mD1 = txtD1 Else mD1 = gbStartingDate
        If IsDate(txtD2) Then mD2 = txtD2 Else mD2 = gbEndingDate
        mSql = mSql + " Where dtBankEntryDate Between '" & DdMmmYy(mD1) & "' and '" & DdMmmYy(mD2) & "'"
        mSql = mSql + " And intBankAccountHeadID = " & mSearchID
        mSql = mSql + " Order By convert(smallDatetime,dtBankEntryDate) "
        Rec.Open mSql, mCnn
        
        fgBankStatement.Rows = 1
        fgBankStatement.Rows = Rec.RecordCount + 1
        If fgBankStatement.Rows > 1 Then
            fgBankStatement.Col = 0
            fgBankStatement.Row = 1
            fgBankStatement.ColSel = 9
            fgBankStatement.RowSel = fgBankStatement.Rows - 1
        
            mSql = Rec.GetString(, , vbTab, Chr(13))
            fgBankStatement.Clip = mSql
        End If
    End Sub
    
    Private Sub FillAfterReconciliation()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mRowCount As Long
            objdb.SetConnection mCnn
        mSql = " Select * from faVouchers  "
        mSql = mSql + " Inner Join faBankReconciliationEntries On faVouchers.vchInstrumentNo = faBankReconciliationEntries.vchChequeNo "
        mSql = mSql + " Where faBankReconciliationEntries.tnyReconciled = 1 "
        mSql = mSql + " and DatePart(mm,dtDate) Between '4' and ' " & cmbMonth.ItemData(cmbMonth.ListIndex) & "'"
        Rec.Open mSql
    End Sub
    
    Public Sub FillAfterSearchReconcile(ByRef mInstrNo As Long)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mRowCount As Integer
        
        objdb.SetConnection mCnn
        mSql = "Select * from faVouchers"
        mSql = mSql + " Inner Join faBankReconciliationEntries "
        mSql = mSql + " On faVouchers.vchInstrumentNo = faBankReconciliationEntries.vchChequeNo"
        mSql = mSql + " Where faVouchers.vchInstrumentNo = '" & mInstrNo & "'"
        Rec.Open mSql, mCnn
        'fgBankStatement.TextMatrix(0, 0) = Rec!
    End Sub
    
    Private Sub mnuManuallyReconcile_Click()
        '-------------------------------------------------------------------------------'
        ' Right Click Popup Menu - For Manual Reconciliation                            '
        '-------------------------------------------------------------------------------'
'        If Val(vsTitleGrid.Tag) > 0 And Val(vsTitleGrid.TextMatrix(1, 0)) > 0 Then
'            '---------------------------------------------------------------------------'
'            'NOTE:-Grid Col(0) contain Selected Scroll Row's ReconciliationID(TokenID)
'            'NOTE:-Tag holds Grid Row ID
'            '    :-Reconciliation Will be done in frmManualReconcile Form
'            '---------------------------------------------------------------------------'
'            mvarManuallyReconciled = False
'            frmManualReconcile.VoucherFlag = False
'            frmManualReconcile.txtVoucherNo.Tag = vsTitleGrid.TextMatrix(1, 0)
'            frmManualReconcile.Show vbModal
'            If mvarManuallyReconciled Then
'                fgBankStatement.Cell(flexcpChecked, Val(vsTitleGrid.Tag), 5) = 1
'                fgBankStatement.Cell(flexcpText, Val(vsTitleGrid.Tag), 6) = ""
'            End If
'
'        End If

        Dim mLoop As Integer
        Dim mDeposit As Double
        Dim mWithDraw As Double
        mDeposit = 0
        mWithDraw = 0
        For mLoop = 1 To fgBankStatement.Rows - 1
            If fgBankStatement.RowHidden(mLoop) = False And _
                fgBankStatement.Cell(flexcpChecked, mLoop, 5) = 2 And _
                    fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked Then
                mWithDraw = mWithDraw + val(fgBankStatement.TextMatrix(mLoop, 3))       '''Debit Credit Calculations in Bank Scroll
                mDeposit = mDeposit + val(fgBankStatement.TextMatrix(mLoop, 4))
            End If
        Next mLoop
        If mWithDraw <> mDeposit Then
            MsgBox "Deposit and Withdrawal Amounts must be equal", vbInformation
            Exit Sub
        End If
        If val(vsTitleGrid.Tag) > 0 And val(vsTitleGrid.TextMatrix(1, 0)) > 0 Then
        '---------------------------------------------------------------------------'
        'NOTE:-Grid Col(0) contain Selected Scroll Row's ReconciliationID(TokenID)
        'NOTE:-Tag holds Grid Row ID
        '    :-Reconciliation Will be done in frmManualReconcile Form
        '---------------------------------------------------------------------------'
            mvarManuallyReconciled = False
            frmManualReconcile.VoucherFlag = False
            frmManualReconcile.txtVoucherNo.Tag = vsTitleGrid.TextMatrix(1, 0)
            frmManualReconcile.Show vbModal
            If mvarManuallyReconciled Then
                fgBankStatement.Cell(flexcpChecked, val(vsTitleGrid.Tag), 5) = 1
                fgBankStatement.Cell(flexcpText, val(vsTitleGrid.Tag), 6) = ""
            End If
        End If
    End Sub
    
    Private Sub mnuUnReconcile_Click()
        If val(vsTitleGrid.Tag) > 0 And val(vsTitleGrid.TextMatrix(1, 0)) > 0 Then
            If vsTitleGrid.Cell(flexcpChecked, 1, 5) = vbChecked Then
                MsgBox val(vsTitleGrid.Tag) & " --- " & vsTitleGrid.Cell(flexcpChecked, 1, 5)
                Dim mTokenID As Long
                Dim objdb As New clsDB
                Dim mCnn As New ADODB.Connection
                Dim Rec As New ADODB.Recordset
                Dim mSql As String
                
                objdb.SetConnection mCnn
                mTokenID = val(vsTitleGrid.TextMatrix(1, 0))
                If mTokenID > 0 Then
                    mSql = "Update faVouchers Set tnysync=Null,dtRealisationDate = Null, vchRemarks= Null, tnyReconciled= Null, numTockenID= Null Where numTockenID = " & mTokenID
                    mCnn.Execute mSql
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '                                   For Multiple Reconciliation's unReconcilie                             '
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Rec.Open "Select intMaxID From faBankReconciliationEntries Where intReconciliationID = " & mTokenID, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        If IsNull(Rec!intMaxID) = False Then
                            mSql = "Update faBankReconciliationEntries Set tnyReconciled = Null ,vchRemarks = Null ,intVoucherNo = Null, fltDifference = Null Where numTockenID = " & Rec!intMaxID
                            mCnn.Execute mSql
                        End If
                    End If
                    Rec.Close
                    
                    mSql = "Update faBankReconciliationEntries Set tnyReconciled = Null ,vchRemarks = Null ,intVoucherNo = Null, fltDifference = Null Where intReconciliationID = " & mTokenID
                    mCnn.Execute mSql
                    
                    mSql = "Delete From faOpeningVouchers Where tnyReconciled = 8 And numTockenID = " & mTokenID & ";" & vbNewLine
                    mSql = mSql + "Update faOpeningVouchers Set tnyReconciled = Null, numTockenID = Null, dtRealisationDate = Null, vchRemarks = Null Where numTockenID = " & mTokenID
                    mCnn.Execute mSql
                    
                    mSql = "Update faTransactionChild Set numTockenID = Null, dtReconcileDate = Null Where numTockenID = " & mTokenID
                    mCnn.Execute mSql
                    
                    Call FillBankStatement
                    Call FillVoucherStatement
                End If
            End If
        End If
    End Sub

    Private Sub mnuVoucherToBankReconcile_Click()
        
'        Dim mLoop As Integer
'        Dim mDeposit As Double
'        Dim mWithDraw As Double
'        mDeposit = 0
'        mWithDraw = 0
'        For mLoop = 1 To fgBankStatement.Rows - 1
'            If fgBankStatement.RowHidden(mLoop) = False And _
'                fgBankStatement.Cell(flexcpChecked, mLoop, 5) = 2 And _
'                    fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked Then
'                mWithDraw = mWithDraw + Val(fgBankStatement.TextMatrix(mLoop, 3))       '''Debit Credit Calculations in Bank Scroll
'                mDeposit = mDeposit + Val(fgBankStatement.TextMatrix(mLoop, 4))
'            End If
'        Next mLoop
'        If Not (Val(lblSelectedAmt.Caption) = mDeposit Or Val(lblSelectedAmt.Caption) = mWithDraw) Then
'            MsgBox "The Scroll Amounts & the Voucher do not match", vbInformation
'            Exit Sub
'        End If
'        If Val(vsTitleGrid.Tag) > 0 And Val(vsTitleGrid.TextMatrix(1, 0)) > 0 Then
'        '---------------------------------------------------------------------------'
'        'NOTE:-Grid Col(0) contain Selected Scroll Row's ReconciliationID(TokenID)
'        'NOTE:-Tag holds Grid Row ID
'        '    :-Reconciliation Will be done in frmManualReconcile Form
'        '---------------------------------------------------------------------------'
'            mvarManuallyReconciled = False
'            frmManualReconcile.VoucherFlag = False
'            frmManualReconcile.VoucherToBankFlag = True
'            frmManualReconcile.txtVoucherNo.Tag = vsTitleGrid.TextMatrix(1, 0)
'            frmManualReconcile.Show vbModal
'            If mvarManuallyReconciled Then
'                fgBankStatement.Cell(flexcpChecked, Val(vsTitleGrid.Tag), 5) = 1
'                fgBankStatement.Cell(flexcpText, Val(vsTitleGrid.Tag), 6) = ""
'            End If
'        End If
    End Sub

    Private Sub mnuVoucherMutual_Click()
        Dim mCount As Integer
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mFirstRow As Integer
        Dim mVoucherIDTocken As Long
        Dim mTotalDr As Double
        Dim mTotalCr As Double
        Dim mSql As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mFirstRow = 0
        '''''Validations'''''''
        For mCount = 1 To fgVoucherStatement.Rows - 1
            If fgVoucherStatement.Cell(flexcpChecked, mCount, 13) = vbChecked Then
                mTotalDr = mTotalDr + val(fgVoucherStatement.TextMatrix(mCount, 4))
                mTotalCr = mTotalCr + val(fgVoucherStatement.TextMatrix(mCount, 5))
            End If
        Next mCount
        If mTotalDr <> mTotalCr Then
            MsgBox "The Amounts are different", vbInformation
            Exit Sub
        End If
        '''''Saving'''''
        For mCount = 1 To fgVoucherStatement.Rows - 2
            If fgVoucherStatement.RowHidden(mCount) = False And _
                fgVoucherStatement.Cell(flexcpChecked, mCount, 13) = vbChecked Then
                mFirstRow = mFirstRow + 1
                If mFirstRow = 1 Then
                    mVoucherIDTocken = val(fgVoucherStatement.Cell(flexcpText, mCount, 10))
                    MsgBox "Reconciled Successfully"
                End If
                If fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 11) = 1 Then
                    ' Opening Vouchers
                    mSql = " Update faOpeningVouchers"
                    mSql = mSql + " Set tnyReconciled = 5"
                    mSql = mSql + ", dtReconcileDate = getDate()"
                    mSql = mSql + ", numTockenID = mVoucherIDTocken "
                    'mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                    mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mCount, 10))
                Else
                    ' Vouchers
                    mSql = " Update faVouchers "
                    mSql = mSql + " Set tnyReconciled = 5"
                    mSql = mSql + " ,numTockenID = " & mVoucherIDTocken
                    'mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                    mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mCount, 10)) & ";"
                    ' TransactionChild
                    mSql = mSql + "Update faTransactionChild Set dtReconcileDate = getDate(), numTockenID =  " & mVoucherIDTocken & _
                            " From faTransactions Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                            " Where faTransactions.intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mCount, 10)) & " And faTransactionChild.intAccountHeadID = " & mSearchID
                End If
                mCnn.Execute mSql
            End If
        Next mCount
    End Sub

    Private Sub mnuVoucherMutualUnReconcile_Click()
        Dim mCount As Integer
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 11) = 1 Then   ' Opening Vouchers
            mCnn.Execute "Update faOpeningVouchers Set tnyReconciled = Null,dtReconcileDate = Null,numTockenID = Null Where tnyReconciled = 5 And numTockenID = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 8)
        Else                                                                    ' Vouchers & TransactionChild
            mSql = "Update faVouchers Set tnysync=Null,tnyReconciled = Null Where tnyReconciled = 5 And numTockenID = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 8) & ";" & vbNewLine & _
                    "Update faTransactionChild Set tnysync=Null,dtReconcileDate = Null, numTockenID = Null Where numTockenID = " & fgVoucherStatement.TextMatrix(fgVoucherStatement.Row, 8)
            mCnn.Execute mSql
        End If
        
    End Sub

    Private Sub txtBankCrAmt_LostFocus()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        If val(txtBankDrAmt) > 0 Then 'NOTE:-Only need to process if First Amount is given
            For mLoopCount = 1 To fgBankStatement.Rows - 1
                If val(txtBankCrAmt) > 0 Then 'NOTE:-Second Amount is also given;Need to search for an Amount in Range
                    If val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) > 0 Then 'NOTE:-To Decide whether to check the Debit Column or Credit Column
                        If val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) >= val(txtBankDrAmt) And val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) <= val(txtBankCrAmt) Then
                            fgBankStatement.RowHidden(mLoopCount) = False
                        Else
                            fgBankStatement.RowHidden(mLoopCount) = True
                        End If
                    Else 'NOTE:- Check in Credit Column
                        If val(fgBankStatement.Cell(flexcpText, mLoopCount, 4)) >= val(txtBankDrAmt) And val(fgBankStatement.Cell(flexcpText, mLoopCount, 4)) <= val(txtBankCrAmt) Then
                            fgBankStatement.RowHidden(mLoopCount) = False
                        Else
                            fgBankStatement.RowHidden(mLoopCount) = True
                        End If
                    End If
                Else
                    'Note:-Only One Amount is given to search
                    If val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) > 0 Then 'NOTE:-To Decide whether to check the Debit Column or Credit Column
                        If val(fgBankStatement.Cell(flexcpText, mLoopCount, 3)) = val(txtBankDrAmt) Then
                            fgBankStatement.RowHidden(mLoopCount) = False
                        Else
                            fgBankStatement.RowHidden(mLoopCount) = True
                        End If
                    Else 'NOTE:- Check in Credit Column
                        If val(fgBankStatement.Cell(flexcpText, mLoopCount, 4)) = val(txtBankDrAmt) Then
                            fgBankStatement.RowHidden(mLoopCount) = False
                        Else
                            fgBankStatement.RowHidden(mLoopCount) = True
                        End If
                    End If
                End If
            Next mLoopCount
        Else
            For mLoopCount = 1 To fgBankStatement.Rows - 1
                fgBankStatement.RowHidden(mLoopCount) = False
            Next mLoopCount
        End If
    End Sub

    Private Sub txtBankDrAmt_LostFocus()
    '    Dim mStr As String
    '    Dim mLoopCount As Long
    '    Dim mFlag As Boolean
    '
    '    txtBankDrAmt = Trim(txtBankDrAmt)
    '    If Len(txtBankDrAmt) Then
    '            If Trim(txtBankInstNo) <> "" Then
    '                For mLoopCount = 1 To fgBankStatement.Rows - 1
    '                    If InStr(1, fgBankStatement.Cell(flexcpText, mLoopCount, 2), Trim(txtBankInstNo)) Then
    '                        fgBankStatement.RowHidden(mLoopCount) = False
    '                    Else
    '                        fgBankStatement.RowHidden(mLoopCount) = True
    '                    End If
    '                Next mLoopCount
    '            Else
    '                For mLoopCount = 1 To fgBankStatement.Rows - 1
    '                    fgBankStatement.RowHidden(mLoopCount) = False
    '                Next mLoopCount
    '            End If
    '
    '        End If
    '    End If
    '
    '    mStr = Left(txtBankDrAmt, 1)
    '    If Not IsNumeric(mStr) Then
    '        Select Case mStr
    '            Case Is = ">"
    '            Case Is = "<"
    '            Case Else
    '        End Select
    '    End If
    
    End Sub

Private Sub txtBankInstNo_LostFocus()
    Dim mLoopCount As Long
    Dim mFlag As Boolean
    If Trim(txtBankInstNo) <> "" Then
        
        For mLoopCount = 1 To fgBankStatement.Rows - 1
            If InStr(1, fgBankStatement.Cell(flexcpText, mLoopCount, 2), Trim(txtBankInstNo)) Then
                fgBankStatement.RowHidden(mLoopCount) = False
            Else
                fgBankStatement.RowHidden(mLoopCount) = True
            End If
        Next mLoopCount
    Else
        For mLoopCount = 1 To fgBankStatement.Rows - 1
            fgBankStatement.RowHidden(mLoopCount) = False
        Next mLoopCount
    End If
    
End Sub

    Private Sub txtD1_LostFocus()
        '-----------------------------------------------'
        'Note:-Formating Starting Date For Bank Scroll
        '-----------------------------------------------'
        If Trim(txtD1) <> "" Then
            txtD1.Text = CheckDateInMMM(txtD1)
        Else
            txtD1 = DdMmmYy(gbStartingDate)
        End If
    End Sub
    Private Sub txtD2_LostFocus()
        '---------------------------------------------'
        'Note:- Filling Bank Scroll
        '---------------------------------------------'
        If Trim(txtD2) <> "" Then
            txtD2.Text = CheckDateInMMM(txtD2)
        Else
            txtD2 = DdMmmYy(gbEndingDate)
        End If
        Call FillBankStatement
        lblSelectedScrollAmt.Caption = 0
    End Sub
    Private Sub txtD3_LostFocus()
        'Note:- Validate Input Date
        If Trim(txtD3) <> "" Then
            txtD3 = CheckDateInMMM(txtD3)
        Else
            txtD3.Text = DdMmmYy(gbStartingDate)
        End If
    End Sub
    Private Sub txtD4_LostFocus()
        'Note:- Validate Input Date and Fill Grid
        If Trim(txtD4) <> "" Then
            txtD4.Text = CheckDateInMMM(txtD4)
        Else
            txtD4.Text = DdMmmYy(gbEndingDate)
        End If
        Call FillVoucherStatement
    End Sub
    Property Let ManuallyReconciledFlag(mData As Boolean)
        mvarManuallyReconciled = mData
    End Property
    Property Let Remarks(ByVal mData As Variant)
        mvarRemarks = mData
    End Property
    Public Sub ReconcileVouchers()
        Dim mSql As String
        Dim mLoopCount As Long
        Dim mTokenID As Long
        Dim mDt As Date
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mFirstVoucherID As Long
        Dim mVoucherNo As String
        Dim mDifference As Double
        Dim mAmount As Double
        Dim mVoucherID As Double
        Dim mMultipleVouchers As Integer
        Dim mFirstRow As Integer
        Dim mLoop As Integer
        Dim mVoucherOrOpeningVoucher As Integer
        Dim mDateProblemFlag As Variant
        
        mDateProblemFlag = Null
        mMultipleVouchers = 0
        mFirstRow = 0
        
        mTokenID = val(vsTitleGrid.TextMatrix(1, 0))
        If mTokenID <= 0 Then
            MsgBox "Bank Entry not selected to Reconcile!", vbInformation
            Exit Sub
        End If
        If IsDate(vsTitleGrid.TextMatrix(1, 1)) Then
            mDt = CheckDateInMMM(vsTitleGrid.TextMatrix(1, 1))
        Else
            MsgBox "Please Enter a valide Bank date as per scroll!", vbInformation
            Exit Sub
        End If
        If fgBankStatement.Cell(flexcpChecked, fgBankStatement.Row, 5) = vbChecked Then
            MsgBox "Already Reconciled", vbInformation
            Exit Sub
        End If
        objdb.SetConnection mCnn
        '-------------------------------------------------------------------'
        '                              Amount Checking                      '
        '                                New Validation                     '
        '-------------------------------------------------------------------'
        If val(lblSelectedScrollAmt.Caption) = 0 Then                           '''''Ordinary Reconciliation
            mAmount = IIf(Trim(txtVDrAmt.Text) = "", val(Trim(txtVCrAmt.Text)), val(Trim(txtVDrAmt.Text)))
            '''''       Checking Scroll Amount and Total Voucher AMount     ''''
            If Abs(val(mSelectedAmt)) <> IIf(val(Trim(vsTitleGrid.TextMatrix(vsTitleGrid.Row, 3))) = 0, Trim(vsTitleGrid.TextMatrix(vsTitleGrid.Row, 4)), Trim(vsTitleGrid.TextMatrix(vsTitleGrid.Row, 3))) Then
                MsgBox "The The Scroll Amount And Vouchers' Amount must be equal", vbInformation
                Exit Sub
            End If
            For mLoopCount = 1 To fgVoucherStatement.Rows - 2
                If fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 7) = vbChecked And fgVoucherStatement.RowHidden(mLoopCount) = False Then
                    If val(fgVoucherStatement.TextMatrix(mLoopCount, 10)) = 0 Then Exit For
                    If CDate(Format(vsTitleGrid.TextMatrix(1, 1), "dd/mmm/yyyy")) < CDate(Format(fgVoucherStatement.TextMatrix(mLoopCount, 2), "dd/mmm/yyyy")) Then
                        If (MsgBox("Please check the Date, Dou you want to continue with this?", vbQuestion + vbYesNo) = vbYes) Then
                            mDateProblemFlag = 1
                        Else
                            Exit Sub
                        End If
                    End If
                    If Not (vsTitleGrid.TextMatrix(1, 3) = fgVoucherStatement.TextMatrix(mLoopCount, 5) Or vsTitleGrid.TextMatrix(1, 4) = fgVoucherStatement.TextMatrix(mLoopCount, 4)) Then
                        If MsgBox("Debit Or Credit False, Do you want to continue?", vbYesNo + vbQuestion) = vbNo Then   ' Uncommented Because to Save Compound Debit and Credit Amount
                            Exit Sub
                        End If
                    End If
                    If fgVoucherStatement.TextMatrix(mLoopCount, 11) = 0 Then
                        mSql = "Select faTransactions.intTransactionID,tinDebitOrCreditFlag,fltAmount,faTransactions.vchNarration From faTransactions " & _
                                " Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                                " Where intVoucherID = " & val(fgVoucherStatement.TextMatrix(mLoopCount, 10)) & " And intAccountHeadID = " & mSearchID
            
                        Rec.Open mSql, mCnn
                        '''     Checking Voucher Dr/Cr With Scroll Cr/Dr        '''
                        If Not (Rec.EOF And Rec.BOF) Then
                            If Rec!tinDebitOrCreditFlag <> IIf(val(Trim(vsTitleGrid.TextMatrix(1, 3))) <> 0, 0, 1) Then
'                                MsgBox "Voucher Dr/Cr Do not match with Scroll Cr/Dr"
'                                Exit Sub
                            End If
                            If Not (Rec!fltAmount = fgVoucherStatement.TextMatrix(mLoopCount, 4) Or Rec!fltAmount = fgVoucherStatement.TextMatrix(mLoopCount, 5)) Then
                                MsgBox "The Transction Amounts are different"
                                Exit Sub
                            End If
                        Else
                            MsgBox "The Transction not present"
                            Exit Sub
                        End If
                        Rec.Close
                    End If
                End If
            Next mLoopCount
            '-------------------------------------------------------------------'
            '                       Validation Completed                        '
            '-------------------------------------------------------------------'
            For mLoopCount = 1 To fgVoucherStatement.Rows - 1
                If fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 7) = vbChecked And fgVoucherStatement.RowHidden(mLoopCount) = False And val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 10)) <> 0 Then
                    If mFirstVoucherID = 0 Then
                        mFirstVoucherID = val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 10))
                        mVoucherNo = val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 1))
                    End If
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 11)) = 0 Then
                        '-------------------------------------------------------------------'
                        'NOTE:- Current Financial Year's - Voucher Data                     '
                        '       Table faVouchers                                            '
                        '-------------------------------------------------------------------'
                        mSql = " Update faVouchers "
                        mSql = mSql + " Set tnyReconciled = 2"
                        mSql = mSql + ", numTockenID = " & mTokenID
                        mSql = mSql + ", dtRealisationDate = '" & DdMmmYy(mDt) & "'"
                        mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                        mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 10))
                        
                    Else
                        mSql = " Update faOpeningVouchers"
                        mSql = mSql + " Set tnyReconciled = 2"
                        mSql = mSql + ", numTockenID = " & mTokenID
                        mSql = mSql + ", dtRealisationDate = '" & DdMmmYy(mDt) & "'"
                        mSql = mSql + ", dtReconcileDate = getDate()"
                        mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                        mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 10))
                    End If
                    mCnn.Execute mSql
                    '----------------------------------------------------------------------'
                    '               faTransactionChild Updation   - Sinoj                  '
                    '----------------------------------------------------------------------'
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 11)) = 0 Then
                        mSql = "Update faTransactionChild Set dtReconcileDate = getDate(), numTockenID =  " & mTokenID & _
                                " From faTransactions Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                                " Where faTransactions.intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 10)) & " And faTransactionChild.intAccountHeadID = " & mSearchID
                        mCnn.Execute mSql
                    End If
                    fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 6) = vbChecked
                    fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 7) = vbChecked
                    fgVoucherStatement.Cell(flexcpText, mLoopCount, 8) = mTokenID
                End If
            Next mLoopCount
            
            '------------------------------------------------------------------------------'
            'NOTE:- After Updation the Selection should clear other wise                   '
            '       It will cause data error - ( in Token ID)                              '
            '------------------------------------------------------------------------------'
            fgVoucherStatement.Select 1, 7, fgVoucherStatement.Rows - 1, 7
            fgVoucherStatement.Clear 2, 0
            fgVoucherStatement.Select 0, 0
            '------------------------------------------------------------------------------'
            If val(fgBankStatement.Cell(flexcpText, vsTitleGrid.Tag, 3)) > 0 Then
                mDifference = val(fgBankStatement.Cell(flexcpText, vsTitleGrid.Tag, 3)) - val(lblSelectedAmt)
            Else
                mDifference = val(fgBankStatement.Cell(flexcpText, vsTitleGrid.Tag, 4)) - val(lblSelectedAmt)
            End If
            
            mDifference = mDifference * -1
            If mFirstVoucherID <> 0 Then
                mSql = "Update faBankReconciliationEntries Set vchRemarks = '" & mvarRemarks & "' ,"
                mSql = mSql + " intVoucherNo =  " & mVoucherNo
                mSql = mSql + ", tnyReconciled = 2"
                mSql = mSql + ", dtReconcileDate = getDate()"
                mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID)+1, 1) From faBankReconciliationEntries A )"
                mSql = mSql + ", fltDifference = " & mDifference
                If IsNull(mDateProblemFlag) = False Then
                    mSql = mSql + ", tnyDateProblem = " & mDateProblemFlag                  ''' If Problem = True
                End If
                mSql = mSql + " Where intReconciliationID = " & mTokenID
                mCnn.Execute mSql
                If val(vsTitleGrid.Tag) > 0 Then
                    fgBankStatement.Cell(flexcpChecked, vsTitleGrid.Tag, 5) = vbChecked
                    fgBankStatement.Cell(flexcpText, vsTitleGrid.Tag, 6) = mVoucherNo
                    If mDifference <> 0 Then
                        fgBankStatement.Cell(flexcpText, vsTitleGrid.Tag, 7) = mDifference
                    End If
                End If
            End If
        ElseIf val(lblSelectedAmt) > 0 And val(lblSelectedScrollAmt) > 0 Then    ' Mutual Reconciliation
            Dim mVrID As String
            Dim mToID As Double
            mFirstRow = 0
            
            Dim mAmountBScroll As Double
            Dim mAmountVScroll As Double
            
            If val(lblSelectedAmt) <> val(lblSelectedScrollAmt) Then
                MsgBox "The Amounts are not same", vbInformation
                Exit Sub
            End If
            '       To Find the voucherNo to Update faBankReconciliation        --mVrID--
            For mLoop = 1 To fgVoucherStatement.Rows - 1
                If fgVoucherStatement.RowHidden(mLoop) = False Then        'Avoid Hidden Fields
                    If fgVoucherStatement.Cell(flexcpChecked, mLoop, 7) = vbChecked Then    'Verify User Checked or Not
                        If val(fgVoucherStatement.Cell(flexcpText, mLoop, 11)) <> 0 Then 'faOpeningVoucher <> 0
                            mVrID = val(fgVoucherStatement.TextMatrix(mLoop, 1))
                        Else        '   Else faVouchers
                            mVrID = val(fgVoucherStatement.TextMatrix(mLoop, 1))
                            Exit For
                        End If
                    End If
                End If
            Next mLoop
            
            'Update Bank Reconciliation
            mSql = ""
            For mLoop = 1 To fgBankStatement.Rows - 1
                If fgBankStatement.RowHidden(mLoop) = False And _
                        fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked And _
                        fgBankStatement.Cell(flexcpChecked, mLoop, 5) = 2 Then
                    'Total Amount in Scroll
                    mAmountBScroll = mAmountBScroll + (val(fgBankStatement.TextMatrix(mLoop, 4)) - val(fgBankStatement.TextMatrix(mLoop, 3)))
                    
                    mFirstRow = mFirstRow + 1
                    mSql = mSql + " Update faBankReconciliationEntries Set vchRemarks = '" & Trim(mvarRemarks) & "' "
                    mSql = mSql + ", intVoucherNo =  " & mVrID
                    mSql = mSql + ", numTockenID =  " & mTokenID
                    mSql = mSql + ", dtReconcileDate = getDate()"
                    mSql = mSql + ", tnyReconciled = 5"
                    If mFirstRow = 1 Then
                        mToID = fgBankStatement.TextMatrix(mLoop, 0)        ' Finding First TockenID to Update in faVouchers,faOpeningVouchers,faTransactionChild
                        mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID)+1, 1) From faBankReconciliationEntries A )"
                    Else
                        mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID), 1) From faBankReconciliationEntries A )"
                    End If
                    mSql = mSql + " Where intReconciliationID = " & fgBankStatement.TextMatrix(mLoop, 0) & ";" & vbNewLine
                End If
            Next mLoop
            
            '           To Update faVoucher,faTransactionChild,faOpeningVouchers
            For mLoop = 1 To fgVoucherStatement.Rows - 1
                If fgVoucherStatement.RowHidden(mLoop) = False Then        'Avoid Hidden Fields
                    If fgVoucherStatement.Cell(flexcpChecked, mLoop, 7) = vbChecked Then    'Verify User Checked or Not
                        'Total Amount in Scroll
                        mAmountVScroll = mAmountVScroll + (val(fgVoucherStatement.TextMatrix(mLoop, 4)) - val(fgVoucherStatement.TextMatrix(mLoop, 5)))
                        If val(fgVoucherStatement.Cell(flexcpText, mLoop, 11)) <> 0 Then 'faOpeningVoucher <> 0
                            mSql = mSql + " Update faOpeningVouchers"
                            mSql = mSql + " Set tnyReconciled = 5"
                            mSql = mSql + ", numTockenID = " & mToID
                            mSql = mSql + ", dtRealisationDate = '" & mDt & "'"
                            mSql = mSql + ", dtReconcileDate = getDate()"
                            mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                            mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoop, 10)) & ";" & vbNewLine
                        Else        '   Else faVouchers And faTransactionChild
                            mSql = mSql + " Update faVouchers "
                            mSql = mSql + " Set tnyReconciled = 5"
                            mSql = mSql + ", numTockenID = " & mTokenID
                            mSql = mSql + ", dtRealisationDate = ' " & mDt & "'"
                            mSql = mSql + ", vchRemarks = '" & mvarRemarks & "'"
                            mSql = mSql + " Where intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoop, 10)) & ";" & vbNewLine
                            
                            mSql = mSql + " Update faTransactionChild Set dtReconcileDate = getDate(), numTockenID =  " & mToID
                            mSql = mSql + " From faTransactions Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID "
                            mSql = mSql + " Where faTransactions.intVoucherID = " & val(fgVoucherStatement.Cell(flexcpText, mLoop, 10)) & " And faTransactionChild.intAccountHeadID = " & mSearchID & ";" & vbNewLine
                        End If
                    End If
                End If
            Next mLoop
            If mAmountBScroll <> mAmountVScroll Then
                MsgBox "The Amounts are not Equal Here it is Exiting"
                Exit Sub
            End If
            mCnn.Execute mSql
            
''            For mLoop = 1 To fgBankStatement.Rows - 1
''                If fgBankStatement.RowHidden(mLoop) = False And fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked Then
''                    If val(fgBankStatement.TextMatrix(mLoop, 0)) < val(vsTitleGrid.TextMatrix(1, 0)) Then
''                        MsgBox "Please Select the smallest from the list First"
''                        Exit Sub
''                    End If
''                    If Format(fgBankStatement.TextMatrix(mLoop, 1), "dd/mmm/yyyy") <> mBankEntryDate Then
''                        MsgBox "The Dates are different please check"
'''                        Exit Sub
''                    End If
''                End If
''            Next mLoop
''            If lblSelectedAmt <> lblSelectedScrollAmt Then
''                MsgBox "The Amounts are not same", vbInformation
''                Exit Sub
''            End If
''            With fgVoucherStatement
''                For mLoop = 1 To .Rows - 1
''                    If .RowHidden(mLoop) = False And _
''                        .Cell(flexcpChecked, mLoop, 7) = vbChecked Then
''                        mVoucherOrOpeningVoucher = .TextMatrix(mLoop, 11)
''                        mVoucherID = .TextMatrix(mLoop, 10)
''                        mMultipleVouchers = mMultipleVouchers + 1
''                    End If
''                Next mLoop
''            End With
''
''            If mMultipleVouchers <> 1 Then
''                MsgBox "Please select only one voucher to reconcile with 2 or more bank scroll", vbInformation
''                Exit Sub
''            End If
''
''            For mLoop = 1 To fgBankStatement.Rows - 1
''                If fgBankStatement.RowHidden(mLoop) = False And _
''                    fgBankStatement.Cell(flexcpChecked, mLoop, 5) = 2 And _
''                        fgBankStatement.Cell(flexcpChecked, mLoop, 8) = vbChecked Then
''                    mFirstRow = mFirstRow + 1
''                    mSql = "Update faBankReconciliationEntries Set vchRemarks = '" & Trim(mvarRemarks) & "' "
''                    mSql = mSql + ", intVoucherNo =  " & mVoucherID
''                    mSql = mSql + ", numTockenID =  " & mTokenID
''                    mSql = mSql + ", dtReconcileDate = getDate()"
''                    mSql = mSql + ", tnyReconciled = 5"
''                    If mFirstRow = 1 Then
''                        mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID)+1, 1) From faBankReconciliationEntries A )"
''                    Else
''                        mSql = mSql + ", intMaxID = ( Select Isnull(Max(A.intMaxID), 1) From faBankReconciliationEntries A )"
''                    End If
''                    mSql = mSql + " Where intReconciliationID = " & fgBankStatement.TextMatrix(mLoop, 0) & ";" & vbNewLine
''
''                    If mFirstRow = 1 Then
''                        If mVoucherOrOpeningVoucher = 0 Then
''                            mSql = mSql + "Update faTransactionChild Set dtReconcileDate = getDate(), numTockenID =  " & fgBankStatement.TextMatrix(mLoop, 0) & _
''                                    " From faTransactions Inner Join faTransactionChild On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
''                                    " Where faTransactions.intVoucherID = " & mVoucherID & " And faTransactionChild.intAccountHeadID = " & mSearchID
''                        Else
''                            mSql = mSql + " Update faOpeningVouchers Set vchRemarks = '" & Trim(mvarRemarks) & "' "
''                            mSql = mSql + ",numTockenID =  " & fgBankStatement.TextMatrix(mLoop, 0)
''                            mSql = mSql + ", dtReconcileDate = getDate()"
''                            mSql = mSql + ", tnyReconciled = 5"
''                            mSql = mSql + " Where intVoucherID = " & mVoucherID
''                        End If
''                    End If
''                mCnn.Execute mSql
''                frmBankReconcilationProcess.ManuallyReconciledFlag = True
''                frmBankReconcilationProcess.vsTitleGrid.Clear 1, 1
''                frmBankReconcilationProcess.vsTitleGrid.Tag = ""
'            End If
'        Next mLoop
            '------------------------------------------------------------------------------'
            'NOTE:- After Updation the Selection should clear other wise                   '
            '       It will cause data error - ( in Token ID)                              '
            '------------------------------------------------------------------------------'
            fgVoucherStatement.Select 1, 7, fgVoucherStatement.Rows - 1, 7
            fgVoucherStatement.Clear 2, 0
            fgVoucherStatement.Select 0, 0
            
            fgBankStatement.Select 1, 8, fgBankStatement.Rows - 1, 7
            fgBankStatement.Clear 2, 0
            fgBankStatement.Select 0, 0
            '------------------------------------------------------------------------------'
        End If
        lblSelectedAmt.Caption = "0.00"
        lblSelectedScrollAmt = "0.00"
        mSelectedAmt = 0
        mSelectedScroll = 0
        mvarManuallyReconciled = True
        vsTitleGrid.Clear 1, 1
        vsTitleGrid.Tag = ""
        
        'Call FillVoucherStatement
        'Call FillBankStatement
    End Sub

    Private Sub AutoReconcile()
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim Recv As New ADODB.Recordset
        Dim mSql As String
        Dim mLoopCount As Long
        Dim mLoopChildCount As Long
        Dim mInstrumentNo  As String
        Dim mBankAmt As String
        Dim mDebitFlag As Integer
        Dim mD1 As Date
        Dim mD2 As Date
        
        mSql = "Select intReconciliationID,dtBankEntryDate,vchChequeNo, fltDrAmount,fltCrAmount,tnyReconciled, intVoucherNo, fltDifference FROM faBankReconciliationEntries"
        If IsDate(txtD1) Then mD1 = txtD1 Else mD1 = gbStartingDate
        If IsDate(txtD2) Then mD2 = txtD2 Else mD2 = gbEndingDate
        mSql = mSql + " Where dtBankEntryDate Between '" & DdMmmYy(mD1) & "' and '" & DdMmmYy(mD2) & "'"
        mSql = mSql + " And intBankAccountHeadID = " & mSearchID
        mSql = mSql + " And tnyReconciled is Null "
        mSql = mSql + " Order By dtBankEntryDate "
        
        objdb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
            
        FileInitialize
        While Not Rec.EOF
            If Trim(Rec!vchChequeNo) = "" Then
                GoTo SkipNextRecord:
            End If
            If IsNumeric(Rec!fltDrAmount) Then
                mBankAmt = Rec!fltDrAmount
                mDebitFlag = 0
            Else
                mBankAmt = Rec!fltCrAmount
                mDebitFlag = 1
            End If
            mSql = ""
            mSql = mSql + "       SELECT"
            mSql = mSql + "          dbo.faTransactionChild.intTransactionID     , "
            mSql = mSql + "           dbo.faTransactions.intVoucherID         , "
            mSql = mSql + "           dbo.faVouchers.intBookNo, "
            mSql = mSql + "           dbo.faVouchers.intVoucherNo, "
            mSql = mSql + "           dbo.faVouchers.vchInstrumentNo, "
            mSql = mSql + "           dbo.faVouchers.tnyVoucherTypeID, "
            mSql = mSql + "           dbo.faVouchers.tnyReconciled, "
            mSql = mSql + "           dbo.faVouchers.numTockenID, "
            
            mSql = mSql + "           dbo.faVouchers.dtDate, "
            mSql = mSql + "           dbo.faTransactions.vchGroup         , "
            mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then "
            mSql = mSql + "               dbo.faTransactionChild.fltAmount * -1 "
            mSql = mSql + "           End fltCrAmount, "
            mSql = mSql + "           Case when tinDebitOrCreditFlag = 1 then "
            mSql = mSql + "              dbo.faTransactionChild.fltAmount "
            mSql = mSql + "           End fltDrAmount, "
            mSql = mSql + "           dbo.faTransactionChild.tinDebitOrCreditFlag , "
            mSql = mSql + "           dbo.faTransactions.vchNarration, "
            mSql = mSql + "           0 tnyOpeningFlag"
            mSql = mSql + "        FROM        dbo.faTransactionChild      "
            mSql = mSql + "        RIGHT JOIN  dbo.faTransactions  ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID"
            mSql = mSql + "        LEFT  JOIN  dbo.faVouchers          ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID"
            mSql = mSql + "        WHERE   ( "
            mSql = mSql + "                 dbo.faTransactionChild.intAccountHeadID = " & mSearchID
            mSql = mSql + "                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL) "
            mSql = mSql + "                 AND faTransactions.intTransactionID > 0"
            'mSQL = mSQL + "                 AND faTransactions.dtTransactionDate <= '" & DdMmmYy(Rec!dtBankEntryDate) & "' "
            mSql = mSql + "                 AND tnyReconciled is Null "
            mSql = mSql + "                 AND ( dbo.faVouchers.vchInstrumentNo = '" & Rec!vchChequeNo & "'"
            mSql = mSql + "                     OR dbo.faVouchers.vchInstrumentNo = '0" & Rec!vchChequeNo & "')"
            mSql = mSql + "                 AND dbo.faTransactionChild.fltAmount = " & val(mBankAmt)
            mSql = mSql + "                 AND dbo.faTransactionChild.tinDebitOrCreditFlag = " & mDebitFlag
            mSql = mSql + "                 )"
            
            Recv.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
            If Not (Recv.BOF And Recv.EOF) Then
                Print #gbFileNO, Rec!intReconciliationID, Rec!vchChequeNo, mBankAmt, mDebitFlag, Recv!vchInstrumentNo, Recv!fltDrAmount, Recv!fltCrAmount, Recv!tinDebitOrCreditFlag, mDebitFlag
                
                mSql = " Update faVouchers "
                mSql = mSql + " Set tnyReconciled = 2"
                mSql = mSql + ", numTockenID = " & Rec!intReconciliationID
                mSql = mSql + ", dtRealisationDate = '" & DdMmmYy(Rec!dtBankEntryDate) & "'"
                mSql = mSql + ", vchRemarks = 'Auto-Reconciliation'"
                mSql = mSql + " Where intVoucherID = " & Recv!intVoucherID
                mCnn.Execute mSql
                
                mSql = "Update faBankReconciliationEntries Set vchRemarks = 'Auto-Reconciliation' ,"
                mSql = mSql + " intVoucherNo =  " & Recv!intVoucherID
                mSql = mSql + ", tnyReconciled = 2"
                mSql = mSql + ", fltDifference = 0"
                mSql = mSql + " Where intReconciliationID = " & Rec!intReconciliationID
                mCnn.Execute mSql
                
            End If
            Recv.Close
SkipNextRecord:
            Rec.MoveNext
        Wend
        Close #gbFileNO
        ShellPad
        Rec.Close
        
    End Sub
    
Private Sub txtVCrAmt_LostFocus()
    Dim mLoopCount As Long
    Dim mFlag As Boolean
    If val(txtVDrAmt) > 0 Then 'NOTE:-Only need to process if First Amount is given
'        fgVoucherStatement.Cell(flexcpChecked, 1, 7, fgVoucherStatement.Rows - 1, 7) = False
'        fgVoucherStatement.Cell(flexcpChecked, 1, 13, fgVoucherStatement.Rows - 1, 13) = False
'        lblSelectedAmt.Caption = "0.00"
        For mLoopCount = 1 To fgVoucherStatement.Rows - 1
            If val(txtVCrAmt) > 0 Then 'NOTE:-Second Amount is also given;Need to search for an Amount in Range
                If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) > 0 Then 'NOTE:-To Decide whether to check the Debit Column or Credit Column
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) >= val(txtVDrAmt) And val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) <= val(txtVCrAmt) Then
                        fgVoucherStatement.RowHidden(mLoopCount) = False
                    Else
                        fgVoucherStatement.RowHidden(mLoopCount) = True
                    End If
                Else 'NOTE:- Check in Credit Column
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 5)) >= val(txtVDrAmt) And val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 5)) <= val(txtVCrAmt) Then
                        fgVoucherStatement.RowHidden(mLoopCount) = False
                    Else
                        fgVoucherStatement.RowHidden(mLoopCount) = True
                    End If
                End If
            Else
                'Note:-Only One Amount is given to search
                If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) > 0 Then 'NOTE:-To Decide whether to check the Debit Column or Credit Column
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) = val(txtVDrAmt) Then
                        fgVoucherStatement.RowHidden(mLoopCount) = False
                    Else
                        fgVoucherStatement.RowHidden(mLoopCount) = True
                    End If
                Else 'NOTE:- Check in Credit Column
                    If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 5)) = val(txtVDrAmt) Then
                        fgVoucherStatement.RowHidden(mLoopCount) = False
                    Else
                        fgVoucherStatement.RowHidden(mLoopCount) = True
                    End If
                End If
            End If
        Next mLoopCount
    Else
        For mLoopCount = 1 To fgVoucherStatement.Rows - 1
            fgVoucherStatement.RowHidden(mLoopCount) = False
        Next mLoopCount
    End If
End Sub

    Private Sub txtVInstNo_LostFocus()
        Dim mLoopCount As Long
        Dim mFlag As Boolean
        Dim mAmt As Double
        lblSelectedAmt.Caption = ""
        Me.MousePointer = vbHourglass
        mFlag = False
'        fgVoucherStatement.Cell(flexcpChecked, 1, 7, fgVoucherStatement.Rows - 1, 7) = False
'        fgVoucherStatement.Cell(flexcpChecked, 1, 13, fgVoucherStatement.Rows - 1, 13) = False
        
        If val(vsTitleGrid.TextMatrix(1, 0)) <> 0 Then
            For mLoopCount = 1 To fgVoucherStatement.Rows - 1
                ' Checking the Tocken IDs
                If val(vsTitleGrid.TextMatrix(1, 0)) = val(fgVoucherStatement.TextMatrix(mLoopCount, 8)) Then
                    fgVoucherStatement.RowHidden(mLoopCount) = False
                    fgVoucherStatement.Cell(flexcpBackColor, mLoopCount, 0, mLoopCount, 6) = &HD6FFD6
                    mFlag = True
                Else
                    fgVoucherStatement.RowHidden(mLoopCount) = True
                End If
            Next mLoopCount
        End If
        If mFlag Then       ' If tocken ID found Exiting
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If Trim(txtVInstNo) <> "" Then
            For mLoopCount = 1 To fgVoucherStatement.Rows - 1
                If InStr(1, UCase(fgVoucherStatement.Cell(flexcpText, mLoopCount, 3)), UCase(Trim(txtVInstNo))) Then
                    fgVoucherStatement.RowHidden(mLoopCount) = False
                    If fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 6) = vbChecked Then
                        If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) > 0 Then
                            mAmt = mAmt + val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4))
                        Else
                            mAmt = mAmt - val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 5))
                        End If
                    End If
                Else
                    fgVoucherStatement.RowHidden(mLoopCount) = True
                End If
            Next mLoopCount
            lblSelectedAmt.Caption = Abs(mAmt)
        ElseIf vsTitleGrid.Cell(flexcpChecked, 1, 5) = vbChecked Then
            For mLoopCount = 1 To fgVoucherStatement.Rows - 1
                If fgVoucherStatement.TextMatrix(mLoopCount, 8) = vsTitleGrid.TextMatrix(1, 0) Then
                    fgVoucherStatement.RowHidden(mLoopCount) = False
                    If fgVoucherStatement.Cell(flexcpChecked, mLoopCount, 6) = vbChecked Then
                        If val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4)) > 0 Then
                            mAmt = mAmt + val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 4))
                        Else
                            mAmt = mAmt - val(fgVoucherStatement.Cell(flexcpText, mLoopCount, 5))
                        End If
                    End If
                Else
                    fgVoucherStatement.RowHidden(mLoopCount) = True
                End If
            Next mLoopCount
            lblSelectedAmt.Caption = Abs(mAmt)
        Else
            For mLoopCount = 1 To fgVoucherStatement.Rows - 1
                fgVoucherStatement.RowHidden(mLoopCount) = False
            Next mLoopCount
        End If
        Me.MousePointer = vbDefault
    End Sub

    Private Sub ReconciliationReport(mAcHeadID As Variant, Optional mDt1 As Variant = Null, Optional mDt2 As Variant = Null)
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        
        Dim Rec As New ADODB.Recordset
        Dim RecBank As New ADODB.Recordset
        
        Dim mSql As String
        Dim mDate As Date
        
        Dim mD1 As Date
        Dim mD2 As Date
        
        Dim mDrAmt As Double
        Dim mCrAmt As Double
        
        Dim mDrRAmt As Double
        Dim mCrPAmt As Double
        
        Dim mDrBAmt As Double
        Dim mCrBAmt As Double
        
        
        Dim mBankBalance As Double
        Dim mBankPassBookBalance As Double
        Dim mInput As Variant
        Dim objAc As New clsAccounts
        
        objAc.SetAccountID mAcHeadID
        If objAc.AccountHeadID <= 0 Then
            MsgBox "No Account head is Selected!", vbInformation
            Exit Sub
        End If
        
        If Not IsDate(mDt1) Then
            mD1 = gbStartingDate
        Else
            mD1 = mDt1
        End If
        If Not IsDate(mD2) Then
            mD2 = gbEndingDate
        Else
            mD2 = mDt2
        End If
        
        objdb.SetConnection mCnn
        mInput = Array(mSearchID, mD1, mD2)
        Set Rec = objdb.ExecuteSP("spGetClosingBalance", mInput, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            mBankBalance = Rec!NetBalance
        Else
            MsgBox "Bank Balance Didn't able to find as per the date!!", vbInformation
            'Exit Sub
        End If
        Rec.Close
        
        'mBankBalance = 1002900 '26141223.58
        mBankPassBookBalance = 0 '24558717.71

        mSql = ""
        mSql = mSql + "       SELECT      dbo.faOpeningVouchers.intTransactionID intTransactionID    ,"
        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherID      intVoucherID   ,"
        mSql = mSql + "           Null intBookNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.intVoucherNo intVoucherNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.vchInstrumentNo vchInstrumentNo,"
        mSql = mSql + "           dbo.faOpeningVouchers.tnyVoucherTypeID tnyVoucherTypeID,"
        mSql = mSql + "           dbo.faOpeningVouchers.tnyReconciled tnyReconciled,"
        mSql = mSql + "           dbo.faOpeningVouchers.numTockenID numTockenID,"
        mSql = mSql + "           dbo.faOpeningVouchers.dtDate dtDate,"
        mSql = mSql + "           Null vchGroup         ,"
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then"
        mSql = mSql + "               dbo.faOpeningVouchers.fltAmount * -1"
        mSql = mSql + "           End fltCrAmount,"
        mSql = mSql + "       Case when tinDebitOrCreditFlag = 1 then"
        mSql = mSql + "              dbo.faOpeningVouchers.fltAmount"
        mSql = mSql + "           End fltDrAmount,"
        mSql = mSql + "       dbo.faOpeningVouchers.tinDebitOrCreditFlag tinDebitOrCreditFlag,"
        mSql = mSql + "           dbo.faOpeningVouchers.vchNarration vchNarration,"
        mSql = mSql + "           1 tnyOpeningFlag, "
        mSql = mSql + "           0 fltAmount "
        mSql = mSql + "       From dbo.faOpeningVouchers"
        mSql = mSql + "       Where tnyReconciled is Null"
        mSql = mSql + "       Union All "
        
        
        
        'mSQL = ""
        mSql = mSql + "       SELECT"
        mSql = mSql + "          dbo.faTransactionChild.intTransactionID     , "
        mSql = mSql + "           dbo.faTransactions.intVoucherID         , "
        mSql = mSql + "           dbo.faVouchers.intBookNo, "
        mSql = mSql + "           dbo.faVouchers.intVoucherNo, "
        mSql = mSql + "           dbo.faVouchers.vchInstrumentNo, "
        mSql = mSql + "           dbo.faVouchers.tnyVoucherTypeID, "
        mSql = mSql + "           dbo.faVouchers.tnyReconciled, "
        mSql = mSql + "           dbo.faVouchers.numTockenID, "
        
        mSql = mSql + "           dbo.faVouchers.dtDate, "
        mSql = mSql + "           dbo.faTransactions.vchGroup         , "
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 0 then "
        mSql = mSql + "               dbo.faTransactionChild.fltAmount * -1 "
        mSql = mSql + "           End fltCrAmount, "
        mSql = mSql + "           Case when tinDebitOrCreditFlag = 1 then "
        mSql = mSql + "              dbo.faTransactionChild.fltAmount "
        mSql = mSql + "           End fltDrAmount, "
        mSql = mSql + "           dbo.faTransactionChild.tinDebitOrCreditFlag , "
        mSql = mSql + "           dbo.faTransactions.vchNarration, "
        mSql = mSql + "           0 tnyOpeningFlag, "
        mSql = mSql + "           dbo.faVouchers.fltAmount "
        mSql = mSql + "        FROM        dbo.faTransactionChild      "
        mSql = mSql + "        INNER JOIN  dbo.faTransactions  ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID"
        mSql = mSql + "        INNER  JOIN  dbo.faVouchers          ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID"
        
        mSql = mSql + "        WHERE   ( "
        mSql = mSql + "                 dbo.faTransactionChild.intAccountHeadID = " & mSearchID
        mSql = mSql + "                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL) "
        mSql = mSql + "                 AND faTransactions.intTransactionID > 0"
        mSql = mSql + "                 AND faTransactions.dtTransactionDate Between '" & DdMmmYy(mD1) & "' AND '" & DdMmmYy(mD2) & "'"
        mSql = mSql + "                 AND (faVouchers.tnyReconciled is Null )"
        mSql = mSql + "                 )"
        mSql = mSql + "     Order By   dtDate,tnyVoucherTypeID, faVouchers.intVoucherID"
        
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            'FileInitialize
            'Print #gbFileNO, "-----------------------------------------------------------------------------------"
            'Print #gbFileNO, " Date           Voucher No   Type          Debit Amt      Credit Amt    Voucher Amt"
            'Print #gbFileNO, "-----------------------------------------------------------------------------------"
            While Not Rec.EOF
                Select Case Rec!tnyVoucherTypeID
                    Case Is = 10: mDrRAmt = mDrRAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                    Case Is = 20: mCrPAmt = mCrPAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                    Case Else:
                        If IsNumeric(Rec!fltDrAmount) Then
                            mDrRAmt = mDrRAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                        ElseIf IsNumeric(Rec!fltCrAmount) Then
                            mCrPAmt = mCrPAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                        End If
                End Select
                mDrAmt = mDrAmt + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
                mCrAmt = mCrAmt + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount)
                
                Rec.MoveNext
            Wend
        End If
        
           
        
        mSql = "        Select intReconciliationID, "
        mSql = mSql + "       dtBankEntryDate,"
        mSql = mSql + "       intVoucherID,"
        mSql = mSql + "       intVoucherNo,"
        mSql = mSql + "       vchChequeNo,"
        mSql = mSql + "       dtChequeDate,"
        mSql = mSql + "       fltDrAmount,"
        mSql = mSql + "       fltCrAmount,"
        mSql = mSql + "       vchParticulars,"
        mSql = mSql + "       vchRemarks"
        mSql = mSql + " From faBankReconciliationEntries"
        mSql = mSql + " Where dtBankEntryDate <= '" & DdMmmYy(mD2) & "'"
        mSql = mSql + " And tnyReconciled is Null"
        mSql = mSql + " And intBankAccountHeadID = " & mSearchID
        
        mSql = mSql + " Order By dtBankEntryDate "
        
        RecBank.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If Not (RecBank.EOF And RecBank.BOF) Then
            While Not RecBank.EOF
                If Not IsNull(RecBank!fltDrAmount) Then
                    mDrBAmt = mDrBAmt + RecBank!fltDrAmount
                Else
                    mCrBAmt = mCrBAmt + RecBank!fltCrAmount
                End If
                RecBank.MoveNext
            Wend
        End If
        
        
        
        FileInitialize
        Print #gbFileNO,
        Print #gbFileNO, PadC("BANK RECONCILIATION STATEMENT", 80)
        Print #gbFileNO, PadC("as " & DdMmmYy(mD2), 80)
        Print #gbFileNO, PadC(objAc.AccountHead, 80)
        Print #gbFileNO, PadC(String(80, "="), 80)
        Print #gbFileNO, "       Closing Balance as per Bank Book : Rs. " & PadL(Format(mBankBalance, "0.00"), 12)
        Print #gbFileNO, " Add:-"
        Print #gbFileNO, "        Cheque Issued but not presented : Rs. " & PadL(Format(Abs(mCrPAmt), "0.00"), 12)
        Print #gbFileNO, "              Directly Credited by Bank : Rs. " & PadL(Format(mCrBAmt, "0.00"), 12)
        Print #gbFileNO, " Less:-"
        Print #gbFileNO, "        Cheque Deposited but not collected : Rs. " & PadL(Format(mDrRAmt, "0.00"), 12)
        Print #gbFileNO, "               Directly Debited by Bank : Rs. " & PadL(Format(mDrBAmt, "0.00"), 12)
        Print #gbFileNO, "Closing Balance as per Bank's Statement : Rs. " & PadL(Format(mBankBalance + Abs(mCrPAmt) + mCrBAmt - mDrRAmt - mDrBAmt, "0.00"), 12)
        
        
        
        
        
        '======================================================================'
        ' 1. CHEQUE ISSUED BUT NOT PRESENTED
        '======================================================================'
        Rec.MoveFirst
        If Not (Rec.EOF And Rec.BOF) Then
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO, PadC("Bank Reconciliation Statement", 80)
            Print #gbFileNO, PadC("as " & DdMmmYy(mD2), 80)
            Print #gbFileNO, PadC(objAc.AccountHead, 80)
            
            Print #gbFileNO, PadC(String(80, "="), 80)
            Print #gbFileNO,
            Print #gbFileNO, "List of Cheque Issued but not presented"
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "  Date        V.No      Cheque No.        Amount"
            Print #gbFileNO, PadC(String(80, "-"), 80)
            While Not Rec.EOF
                If IsNumeric(Rec!fltCrAmount) Then
                    Print #gbFileNO, "  " & DdMmmYy(Rec!dtDate); " ";
                    If IsNull(Rec!intVoucherNo) Then
                        Print #gbFileNO, PadL(" ", 12);
                    Else
                        Print #gbFileNO, PadL(Trim(str(Rec!intVoucherNo)), 12);
                    End If
                    
                    If IsNull(Rec!vchInstrumentNo) Then
                        Print #gbFileNO, PadR(" ", 12);
                    Else
                        Print #gbFileNO, PadR(Rec!vchInstrumentNo, 12);
                    End If
                    Print #gbFileNO, PadL(Format(Abs(Rec!fltCrAmount), "0.00"), 12)
                End If
                Rec.MoveNext
            Wend
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "                                   " & PadL(Format(Abs(mCrPAmt), "0.00"), 14)
            Print #gbFileNO, PadC(String(80, "="), 80)
        End If
        'Rec.Close
    
    
        
        '======================================================================'
        ' 2. DIRECTLY CREDITED BY BANK
        '======================================================================'
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, PadC("Bank Reconciliation Statement", 80)
        Print #gbFileNO, PadC("as " & DdMmmYy(mD2), 80)
        Print #gbFileNO, PadC(objAc.AccountHead, 80)
        
        Print #gbFileNO, PadC(String(80, "="), 80)
        Print #gbFileNO,
        Print #gbFileNO, "Directly Credited by Bank"
        Print #gbFileNO, PadC(String(80, "-"), 80)
        Print #gbFileNO, "  Date        V.No      Cheque No.        Amount"
        Print #gbFileNO, PadC(String(80, "-"), 80)
        RecBank.MoveFirst
        If Not (RecBank.EOF And RecBank.BOF) Then
            While Not RecBank.EOF
                If Not IsNull(RecBank!fltDrAmount) Then
                    Print #gbFileNO, "  " & DdMmmYy(RecBank!dtBankEntryDate); " ";
                    If IsNull(RecBank!intVoucherNo) Then
                        Print #gbFileNO, PadL(" ", 12);
                    Else
                        Print #gbFileNO, PadL(Trim(str(RecBank!intVoucherNo)), 12);
                    End If
                    If IsNull(RecBank!vchChequeNo) Then
                        Print #gbFileNO, PadR(" ", 12);
                    Else
                        Print #gbFileNO, PadR(RecBank!vchChequeNo, 12);
                    End If
                    Print #gbFileNO, PadL(Format(RecBank!fltDrAmount, "0.00"), 12)
                End If
                RecBank.MoveNext
            Wend
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "                                   " & PadL(Format(mCrBAmt, "0.00"), 14)
            Print #gbFileNO, PadC(String(80, "="), 80)
        End If
    
        '======================================================================'
        ' 3. CHEQUE DEPOSITED BUT NOT COLLECTED
        '======================================================================'
        
        Rec.MoveFirst
        If Not (Rec.EOF And Rec.BOF) Then
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
                
            Print #gbFileNO, PadC("Bank Reconciliation Statement", 80)
            Print #gbFileNO, PadC("as " & DdMmmYy(mD2), 80)
            Print #gbFileNO, PadC(objAc.AccountHead, 80)
            
            Print #gbFileNO, PadC(String(80, "="), 80)
            Print #gbFileNO,
            Print #gbFileNO, "List of Cheque Deposited But Not Collected"
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "  Date        V.No      Cheque No.        Amount"
            Print #gbFileNO, PadC(String(80, "-"), 80)
            While Not Rec.EOF
                If IsNumeric(Rec!fltCrAmount) Then
                        Print #gbFileNO, "  " & DdMmmYy(Rec!dtDate); " ";
                        If IsNull(Rec!intVoucherNo) Then
                            Print #gbFileNO, PadL(" ", 11);
                        Else
                            Print #gbFileNO, PadL(Trim(str(Rec!intVoucherNo)), 11);
                        End If
                        If IsNull(Rec!vchInstrumentNo) Then
                            Print #gbFileNO, PadR(" ", 12);
                        Else
                            Print #gbFileNO, PadR(Rec!vchInstrumentNo, 12);
                        End If
                        Print #gbFileNO, PadL(Format(Abs(Rec!fltCrAmount), "0.00"), 12)
                End If
                Rec.MoveNext
            Wend
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "                                   " & PadL(Format(Abs(mDrRAmt), "0.00"), 14)
            Print #gbFileNO, PadC(String(80, "="), 80)
        End If
        'Rec.Close
        
        '======================================================================'
        ' 4. DIRECTLY DEBITED BY BANK
        '======================================================================'
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, PadC("Bank Reconciliation Statement", 80)
        Print #gbFileNO, PadC("as " & DdMmmYy(mD2), 80)
        Print #gbFileNO, PadC(objAc.AccountHead, 80)
        
        Print #gbFileNO, PadC(String(80, "="), 80)
        Print #gbFileNO,
        Print #gbFileNO, "Directly Credited by Bank"
        Print #gbFileNO, PadC(String(80, "-"), 80)
        Print #gbFileNO, "  Date        V.No      Cheque No.        Amount"
        Print #gbFileNO, PadC(String(80, "-"), 80)
        RecBank.MoveFirst
        If Not (RecBank.EOF And RecBank.BOF) Then
            While Not RecBank.EOF
                If Not IsNull(RecBank!fltCrAmount) Then
                    Print #gbFileNO, "  " & DdMmmYy(RecBank!dtBankEntryDate); " ";
                    If IsNull(RecBank!intVoucherNo) Then
                        Print #gbFileNO, PadL(" ", 12);
                    Else
                        Print #gbFileNO, PadL(Trim(str(RecBank!intVoucherNo)), 12);
                    End If
                    If IsNull(RecBank!vchChequeNo) Then
                        Print #gbFileNO, PadR(" ", 12);
                    Else
                        Print #gbFileNO, PadR(RecBank!vchChequeNo, 12);
                    End If
                    Print #gbFileNO, PadL(Format(RecBank!fltCrAmount, "0.00"), 12)
                End If
                RecBank.MoveNext
            Wend
            Print #gbFileNO, PadC(String(80, "-"), 80)
            Print #gbFileNO, "                                   " & PadL(Format(mDrBAmt, "0.00"), 14)
            Print #gbFileNO, PadC(String(80, "="), 80)
        End If
        Close #gbFileNO
        ShellPad
    
     End Sub
    
    Private Sub menuAllVisible(val As Boolean)
        mnuManuallyReconcile.Visible = val
        mnuUnReconcile.Visible = val
        mnuVoucherMutual.Visible = val
        mnuVoucherMutualUnReconcile.Visible = val
    End Sub


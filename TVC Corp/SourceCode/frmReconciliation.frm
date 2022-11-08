VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReconciliation 
   BackColor       =   &H00F5F8F8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reconciliation Form"
   ClientHeight    =   11205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16035
   FillColor       =   &H00D6FFD6&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11205
   ScaleWidth      =   16035
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCloseReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2475
      Picture         =   "frmReconciliation.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   47
      Top             =   2190
      Visible         =   0   'False
      Width           =   225
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   6225
      Left            =   2220
      TabIndex        =   45
      Top             =   2130
      Visible         =   0   'False
      Width           =   13590
      lastProp        =   500
      _cx             =   23971
      _cy             =   10980
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F5F8F8&
      Caption         =   "List"
      Height          =   945
      Left            =   120
      TabIndex        =   42
      Top             =   8520
      Width           =   1980
      Begin VB.CheckBox chkReconcile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reconciled"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   570
         Width           =   1095
      End
      Begin VB.CheckBox chkUnReconcile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Unreconciled"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F5F8F8&
      Height          =   6270
      Left            =   12555
      TabIndex        =   15
      Top             =   2355
      Width           =   3225
      Begin VB.CommandButton cmdReconcile 
         Caption         =   "RECONCILE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1590
         TabIndex        =   34
         Top             =   5025
         Width           =   1365
      End
      Begin VB.TextBox txtRefVoucherNo 
         Height          =   345
         Left            =   1590
         TabIndex        =   33
         Top             =   4635
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtNote 
         Height          =   330
         Left            =   330
         TabIndex        =   31
         Top             =   4275
         Width           =   2640
      End
      Begin VB.TextBox txtRealisationDate 
         Height          =   345
         Left            =   1590
         TabIndex        =   29
         Top             =   2940
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ENTER to Reconcile  ESC to Clear]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   405
         TabIndex        =   36
         Top             =   3315
         Width           =   2550
      End
      Begin VB.Label lblRefVoucherNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REF.[Voucher]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   345
         TabIndex        =   32
         Top             =   4695
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   345
         TabIndex        =   30
         Top             =   4050
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   2955
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label lblRealisationDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REALISATION DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   510
         Left            =   390
         TabIndex        =   28
         Top             =   2835
         Width           =   1125
      End
      Begin VB.Label lblInstrumentDate 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14-APR-2013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   27
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label lblInstrumentNo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "656564"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   26
         Top             =   1950
         Width           =   1380
      End
      Begin VB.Label lblInstrumentType 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHECK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   25
         Top             =   1620
         Width           =   1380
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "150000.00  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   24
         Top             =   1290
         Width           =   1380
      End
      Begin VB.Label lblVoucherDate 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15-APR-2013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   23
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label lblVoucherNo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   22
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INST. DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   21
         Top             =   2355
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INST. NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   20
         Top             =   2010
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INST.TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   19
         Top             =   1665
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   18
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label lblVoucher 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOUCHER NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   17
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   390
         TabIndex        =   16
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   585
      TabIndex        =   14
      Top             =   10545
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "LOCK "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6135
      TabIndex        =   13
      Top             =   8700
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton cmdShowReconciledItems 
      Caption         =   "Show Reconciled Items"
      Height          =   540
      Left            =   105
      TabIndex        =   12
      Top             =   9555
      Width           =   1980
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   16005
      TabIndex        =   1
      Top             =   0
      Width           =   16035
      Begin VB.CommandButton cmdClear 
         Caption         =   "CLEAR"
         Height          =   450
         Left            =   13590
         TabIndex        =   48
         Top             =   270
         Width           =   705
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         Height          =   450
         Left            =   14310
         TabIndex        =   3
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANK RECONCILIATION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   255
         TabIndex        =   2
         Top             =   240
         Width           =   3360
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6525
      Left            =   2280
      TabIndex        =   0
      Top             =   2130
      Width           =   10230
      _cx             =   18045
      _cy             =   11509
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReconciliation.frx":022C
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
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      Height          =   10230
      Left            =   0
      ScaleHeight     =   10170
      ScaleWidth      =   2190
      TabIndex        =   4
      Top             =   975
      Width           =   2250
      Begin VB.CommandButton cmdReport 
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   90
         TabIndex        =   46
         Top             =   6780
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "REFERESH"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   570
         TabIndex        =   41
         Top             =   8925
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdDirectBankTrn 
         Caption         =   "DIRECT BANK TRANSACTIONS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   39
         Top             =   1425
         Width           =   1980
      End
      Begin VB.CommandButton cmdPreviousVouchers 
         Caption         =   "Previous Vouchers"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   38
         Top             =   2070
         Width           =   1980
      End
      Begin VB.CommandButton cmdPreviousVourchersNotInSaankhya 
         Caption         =   "  Previous Vouchers    [NOT IN SAANKHYA]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   37
         Top             =   2715
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   2235
      ScaleHeight     =   1050
      ScaleWidth      =   10275
      TabIndex        =   5
      Top             =   975
      Width           =   10305
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   6405
         Top             =   150
         Width           =   3045
      End
      Begin VB.Label lblBankStatementAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   255
         Left            =   8055
         TabIndex        =   11
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label lblBankBookAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   255
         Left            =   8055
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblBankStatement 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "BANK STATEMENT"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         TabIndex        =   9
         Top             =   585
         Width           =   1380
      End
      Begin VB.Label lblBankBook 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "BANK BOOK"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         TabIndex        =   8
         Top             =   285
         Width           =   915
      End
      Begin VB.Label lblBankName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NATIONALIZED BANK"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   2355
      End
      Begin VB.Label lblHeadCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[450250101]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   195
         Width           =   1395
      End
   End
   Begin VB.Label lblDiffAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " 0.00 "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   10155
      TabIndex        =   40
      Top             =   8655
      Width           =   2160
   End
   Begin VB.Line Line2 
      X1              =   13095
      X2              =   15750
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Label lblLastDate 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JANUARY-2013"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   13095
      TabIndex        =   35
      Top             =   1320
      Width           =   2670
   End
End
Attribute VB_Name = "frmReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mRowSelected As Integer
    Private mdtLastDate As Variant
    Private mintBankAccountHeadID As Variant
    Private mintReconID As Variant
    Private mBankBalance As Variant
    Private mPassBookBalance As Variant
    Private mReconciliationStarted As Boolean
    Private mReconStatus As Boolean  ' RECONCILIATION FINISHED OR NOT
    Private mFlag       As Integer  ' used in Fillgrid to filter search 1- Reconciled 2 UnReconciled
    Dim mFirstTimeReconFlag As Boolean ' SELECTED BANK DOING RECONCILIATION FIRST TIME OR NOT
    
    Private Sub CheckReconTally(ByVal Rec As Recordset)
        Dim mReconAmt As Double
        
        If Rec.State <> 1 Then
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim mArrIn As Variant
            mArrIn = Array(mintReconID)
            Set Rec = objdb.ExecuteSP("spGetUnReconAmounts", mArrIn)
        End If
        mReconAmt = mBankBalance
            
        While Not Rec.EOF
            If Rec!tnyTypeID = 3 Then                                        '[3] [P] Cheques issued but not presented into bank
                mReconAmt = mReconAmt + IIf(IsNull(Rec!CrAmt), 0, Rec!CrAmt)
            ElseIf Rec!tnyTypeID = 2 Then                                    '[2] [Cr]Directly Credited by Bank
                mReconAmt = mReconAmt + IIf(IsNull(Rec!DrAmt), 0, Rec!DrAmt)
            ElseIf Rec!tnyTypeID = 1 Then                                    '[1] [R]Cheques deposited but not Collected
                mReconAmt = mReconAmt - IIf(IsNull(Rec!DrAmt), 0, Rec!DrAmt)
            ElseIf Rec!tnyTypeID = 4 Then                                    '[4] [Dr]Directly Debited by Bank
                mReconAmt = mReconAmt - IIf(IsNull(Rec!CrAmt), 0, Rec!CrAmt)
            End If
            Rec.MoveNext
        Wend
        
        If mReconAmt = mPassBookBalance Then
            If mReconStatus Then
                cmdLock.Enabled = False
            Else
                cmdLock.Visible = True
            End If
            
            lblDiffAmt.ForeColor = &H8000&
            lblDiffAmt.Caption = "RECONCILIATION TALLY"
            cmdReport.Visible = True
        Else
            cmdReport.Visible = False
            picCloseReport.Visible = False
            cmdLock.Visible = False
            lblDiffAmt.ForeColor = &H80&
            lblDiffAmt.Caption = Format(mReconAmt - mPassBookBalance, "0.00  ")
        End If
        Set Rec = Nothing
        
    End Sub
    
    Private Sub CheckFirstTimeReconciliation()
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        
        If IsNumeric(mintBankAccountHeadID) Then
            mSql = "SELECT Count(intReconID) intReconFlag FROM faBankReconcile Where tnyReconStatus = 1 AND intBankAccountHeadID = " & mintBankAccountHeadID
            objdb.SetConnection mCnn
            Set Rec = mCnn.Execute(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                If Rec!intReconFlag = 0 Then
                    mFirstTimeReconFlag = True
                Else
                    mFirstTimeReconFlag = False
                End If
            End If
        End If
    End Sub
    
    Private Sub DisplayVoucher()
        Dim mRow As Integer
        mRow = vsGrid.Row
        Call ClearVoucherFields
        If mRow > 0 Then
            lblVoucherNo.Caption = vsGrid.TextMatrix(mRow, 2)
            lblVoucherDate.Caption = vsGrid.TextMatrix(mRow, 3)
            lblAmount.Caption = vsGrid.TextMatrix(mRow, 4)
            lblInstrumentType.Caption = vsGrid.TextMatrix(mRow, 10)
            lblInstrumentNo.Caption = vsGrid.TextMatrix(mRow, 5)
            lblInstrumentDate.Caption = vsGrid.TextMatrix(mRow, 6)
            txtRealisationDate.Text = vsGrid.TextMatrix(mRow, 7)
            
            If vsGrid.TextMatrix(mRow, 1) = "CR" Or vsGrid.TextMatrix(mRow, 1) = "DR" Then
                lblRefVoucherNo.Visible = True
                txtRefVoucherNo.Visible = True
            Else
                lblRefVoucherNo.Visible = False
                txtRefVoucherNo.Visible = False
            End If
            
            If val(vsGrid.TextMatrix(mRow, 8)) = 1 Then
                cmdReconcile.Caption = "UNRECONCILE"
            Else
                cmdReconcile.Caption = "RECONCILE"
            End If
            If txtRealisationDate.Enabled Then
                txtRealisationDate.SetFocus
            End If
        End If
    End Sub
    
    
    Private Sub ClearVoucherFields()
        lblVoucherNo.Caption = ""
        lblVoucherDate.Caption = ""
        lblAmount.Caption = ""
        lblInstrumentType.Caption = ""
        lblInstrumentNo.Caption = ""
        lblInstrumentDate.Caption = ""
        txtRealisationDate.Text = ""
        txtNote.Text = ""
        txtRefVoucherNo.Text = ""
        txtRefVoucherNo.Tag = ""
        cmdReconcile.Caption = "RECONCILE"
    End Sub
    
    Private Sub FillGrid()
        If Not IsDate(mdtLastDate) Then
            Exit Sub
        End If
    
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mSql As String
        Dim mDt1 As Date
        Dim mDt2 As Date
        Dim mRow As Integer
        Dim mStrDate As String
        Dim mRealizationDate As String
        Dim mStr As String
        
        mDt2 = mdtLastDate
        mDt1 = DateAdd("m", -1, mDt2)
        mDt1 = DateAdd("d", 1, mDt1)
      
        mSql = "SELECT faBankReconcileChild.intVoucherID," & vbCrLf
        mSql = mSql + " Case WHEN faBankReconcileChild.tnyVoucherTypeID = 10 THEN 'R'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 20 THEN 'P'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 30 THEN 'C'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 40 THEN 'JV'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 11 THEN 'R'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 21 THEN 'P'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 50 THEN 'CR'" & vbCrLf
        mSql = mSql + "      WHEN faBankReconcileChild.tnyVoucherTypeID = 60 THEN 'DR'" & vbCrLf
        
        mSql = mSql + " END vchGroup," & vbCrLf
        mSql = mSql + " faBankReconcileChild.vchVoucherNo," & vbCrLf
        mSql = mSql + " faBankReconcileChild.dtVoucherDate," & vbCrLf
        mSql = mSql + " CASE WHEN NOT numDrAmount IS NULL THEN numDrAmount" & vbCrLf
        mSql = mSql + "      ELSE numCrAmount END numAmount," & vbCrLf
        mSql = mSql + " Left(ISNULL(LTrim(RTrim(faBankReconcileChild.vchInstrumentNo)),''),19) vchInstrumentNo ," & vbCrLf
        mSql = mSql + " faBankReconcileChild.dtInstrumentDate," & vbCrLf
        mSql = mSql + " faVouchers.intInstrumentTypeID," & vbCrLf
        mSql = mSql + " Replace(vchInstrumentType,' ','') vchInstrumentType," & vbCrLf
        mSql = mSql + " faVouchers.intVoucherID," & vbCrLf
        mSql = mSql + " faBankReconcileChild.intTransactionID," & vbCrLf
        mSql = mSql + " intSerialNo," & vbCrLf
        mSql = mSql + " intReconChdID, " & vbCrLf
        mSql = mSql + " ISNULL(tnyFlag,0) tnyFlag, " & vbCrLf
        mSql = mSql + " dtRealization, " & vbCrLf
        mSql = mSql + " ISNULL(vchNote,'') vchNote, " & vbCrLf
        mSql = mSql + " intLinkReconChdID, " & vbCrLf
        mSql = mSql + " intRefVoucherID " & vbCrLf
        mSql = mSql + " From faBankReconcileChild" & vbCrLf
        mSql = mSql + " LEFT JOIN faVouchers ON faVouchers.intVoucherID = faBankReconcileChild.intVoucherID" & vbCrLf
        mSql = mSql + " LEFT JOIN faInstrumentTypes ON faInstrumentTypes.intInstrumentTypeID =  faVouchers.intInstrumentTypeID" & vbCrLf
        mSql = mSql + " Where IsNull(tnyCancelFlag, 0) = 0" & vbCrLf
        'mSql = mSql + "     AND dtVoucherDate BETWEEN '" & DdMmmYy(mDt1) & "' AND '" & DdMmmYy(mDt2) & "'" & vbCrLf
        mSql = mSql + " AND intReconID = " & mintReconID
        If mFlag = 1 Then
            mSql = mSql + " AND faBankReconcileChild.tnyFlag =1"
        ElseIf mFlag = 2 Then
            mSql = mSql + " AND faBankReconcileChild.tnyFlag is null"
        End If
      
        mSql = mSql + " Order by dtDate, faVouchers.intVoucherID " & vbCrLf
    
        If objdb.SetConnection(mCnn) Then
            Rec.CursorLocation = adUseClient
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly, adCmdText
            vsGrid.Rows = 1
            If Not (Rec.BOF And Rec.EOF) Then
                mRow = 1
                'vsGrid.Rows = Rec.RecordCount + 1
                While Not Rec.EOF
                    mStrDate = IIf(IsDate(Rec!dtInstrumentDate), Rec!dtInstrumentDate, "")
                    If IsDate(mStrDate) Then
                        mStrDate = DdMmmYy(CDate(mStrDate))
                    End If
                    mRealizationDate = IIf(IsDate(Rec!dtRealization), Rec!dtRealization, "")
                    If IsDate(mRealizationDate) Then
                        mRealizationDate = DdMmmYy(CDate(mRealizationDate))
                    End If
                    mStr = mRow & vbTab & Rec!vchGroup & vbTab & Rec!vchVoucherNo & vbTab & DdMmmYy(Rec!dtVoucherDate) & vbTab
                    mStr = mStr & Format(Rec!numAmount, "0.00") & vbTab & Rec!vchInstrumentNo & vbTab & mStrDate & vbTab
                    mStr = mStr & mRealizationDate & vbTab & Rec!tnyFlag & vbTab & Rec!intInstrumentTypeID & vbTab & Rec!vchInstrumentType & vbTab & Rec!intVoucherID & vbTab
                    mStr = mStr & "" & vbTab & Rec!intTransactionID & vbTab & Rec!intSerialNo & vbTab & Rec!intReconChdID & vbTab
                    mStr = mStr & Rec!vchNote & vbTab & Rec!intLinkReconChdID & vbTab & Rec!intRefVoucherID
                    vsGrid.AddItem mStr, mRow
                    If Rec!tnyFlag Then
                        vsGrid.Cell(flexcpBackColor, mRow, 0, mRow, 15) = &HD6FFD6
                    End If
                    mRow = mRow + 1
                    Rec.MoveNext
                Wend
            End If
            Rec.Close
        End If
        Call CheckReconTally(Rec)
        
    End Sub
    
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        
        If Not IsDate(mdtLastDate) Then
            DisableButtons (True)
            vsGrid.Rows = 1
        End If
        Call CheckFirstTimeReconciliation
        
        
        Dim objBank As New clsBank
        objBank.SetBankInfoByAccID CInt(mintBankAccountHeadID)  'mintBankAccountHeadID
        If objBank.BankID > 0 Then
            lblHeadCode.Caption = "[ " & objBank.BankAccountHeadCode & " ]"
            lblBankName.Caption = UCase(objBank.BankName)
        Else
            lblHeadCode.Caption = ""
            lblBankName.Caption = ""
        End If
    
        lblBankBookAmount.Caption = Format(mBankBalance, "0.00")
        lblBankStatementAmount.Caption = Format(mPassBookBalance, "0.00")
        lblLastDate.Caption = UCase(Format(mdtLastDate, "mmmm")) & "-" & Year(mdtLastDate)
        Call ClearVoucherFields
        cmdPreviousVouchers.Visible = mFirstTimeReconFlag  'mReconciliationStarted
        cmdPreviousVourchersNotInSaankhya.Visible = mFirstTimeReconFlag
    End Sub
    
    Private Sub DisableButtons(mFlag As Boolean)
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is CommandButton Then
                mCrl.Enable = mFlag
            End If
        Next
    End Sub
    
    Private Sub chkReconcile_Click()
        If chkReconcile.value = vbChecked Then
            mFlag = 1
            If chkUnReconcile.value = vbChecked Then
                mFlag = 0
            End If
            Call FillGrid
        Else
            mFlag = 0
            If chkUnReconcile.value = vbChecked Then
                mFlag = 1
                
            End If
            Call FillGrid
        End If
    End Sub

    Private Sub chkUnReconcile_Click()
        If chkUnReconcile.value = vbChecked Then
            mFlag = 2
            If chkReconcile.value = vbChecked Then
                mFlag = 0
            End If
            Call FillGrid
        Else
            mFlag = 0
            If chkReconcile.value = vbChecked Then
                mFlag = 2
            End If
            Call FillGrid
        End If
    End Sub

    Private Sub cmdClear_Click()
        Call Form_Load
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    
    Private Sub cmdDirectBankTrn_Click()
        frmReconDirectBankTrn.LastDate = mdtLastDate
        frmReconDirectBankTrn.BankAccountHeadID = mintBankAccountHeadID
        frmReconDirectBankTrn.ReconID = mintReconID
        frmReconDirectBankTrn.Show vbModal
    End Sub

Private Sub cmdLock_Click()
    Dim objLdgr As New clsAccounts
    Dim mAmt As Double
    Dim mStr As String
    'Private mdtLastDate As Variant
    'Private mintBankAccountHeadID As Variant
    If IsNumeric(mintBankAccountHeadID) Then
        If IsDate(mdtLastDate) Then
            mAmt = objLdgr.GetLedgerBalance(CInt(mintBankAccountHeadID), mdtLastDate)
            If val(lblBankBookAmount.Caption) <> mAmt Then
                mStr = " " + vbCrLf
                mStr = mStr + " BANK BOOK BALANCE IS CHANGED AFTER YOU STARTED " + vbCrLf
                mStr = mStr + " RECONCILIATION " + vbCrLf
                mStr = mStr + " NOW THE BANK BOOK BALANCE AS ON " & DdMmmYy(CDate(mdtLastDate)) & " IS " + vbCrLf
                mStr = mStr + " Rs." & Format(mAmt, "0.00") + vbCrLf
                mStr = mStr + " PLEASE GO TO THE START PAGE AND EDIT THE BALANCE TO START RECONCILIATION" + vbCrLf
                mStr = mStr + " THANK YOU !" + vbCrLf
                MsgBox mStr, vbInformation
                Exit Sub
            End If
            
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            mSql = "UPDATE faBankReconcile SET tnyReconStatus = 1 WHERE intReconID = " & mintReconID
            If objdb.SetConnection(mCnn) Then
                mCnn.Execute mSql
                mReconStatus = True
                
                If CDate(mdtLastDate) < "01-Apr-2008" Then
                    MsgBox "Error in Closing Date", vbInformation
                    Exit Sub
                End If
                
                mSql = "UPDATE faBanks SET dtReconEndDate = '" & DdMmmYy(CDate(mdtLastDate)) & "' WHERE intAccountHeadID = " & mintBankAccountHeadID
                mCnn.Execute mSql

            End If
            Set mCnn = Nothing
            Set objdb = Nothing
            cmdLock.Enabled = False
            cmdPreviousVouchers.Visible = False
            cmdPreviousVourchersNotInSaankhya.Visible = False
            Call CheckFirstTimeReconciliation
            
            
        End If
    End If
End Sub

    Private Sub cmdPreviousVouchers_Click()
        If mintBankAccountHeadID > 0 Then
            If mintReconID > 0 Then
                frmReconVoucherList.ReconID = mintReconID
                frmReconVoucherList.BankAccountHeadID = mintBankAccountHeadID
                frmReconVoucherList.LastDate = mdtLastDate
                frmReconVoucherList.Show vbModal
                frmReconVoucherList.ZOrder (0)
            Else
                MsgBox "Unexpected Error: Unable to connect to Reconciliation Register", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "No Back is Selected!", vbInformation
        End If
    End Sub
    
    Private Sub cmdPreviousVourchersNotInSaankhya_Click()
        frmReconAddPriorVouchers.LastDate = mdtLastDate
        frmReconAddPriorVouchers.BankAccountHeadID = mintBankAccountHeadID
        frmReconAddPriorVouchers.ReconID = mintReconID
        frmReconAddPriorVouchers.Show vbModal
    End Sub
    
    Private Sub cmdReconcile_Click()
        If lblAmount.Caption = "" Then
        Else
            If cmdReconcile.Caption = "UNRECONCILE" Then
                Dim objdb As New clsDB
                Dim mArrIn As Variant
                
                mArrIn = Array(vsGrid.TextMatrix(mRowSelected, 15), _
                            Null, _
                            Null)
                objdb.ExecuteSP "spReconcile", mArrIn
                
                vsGrid.TextMatrix(mRowSelected, 7) = ""
                vsGrid.TextMatrix(mRowSelected, 8) = 0
                vsGrid.TextMatrix(mRowSelected, 12) = ""
                cmdUpdate.Enabled = True
                vsGrid.Cell(flexcpBackColor, mRowSelected, 0, mRowSelected, 15) = vbDefault
                Call ClearVoucherFields
            Else
                Call txtRealisationDate_KeyPress(vbKeyReturn)
            End If
        End If
    End Sub

    Private Sub cmdRefresh_Click()
        Call FillGrid
    End Sub

''''    Private Sub cmdUpdate_Click()
'''''        Dim mLoop As Integer
'''''        Dim mCnn As New ADODB.Connection
'''''        Dim Rec As New ADODB.Recordset
'''''        Dim objDB As New clsDB
'''''        Dim arrInput As Variant
'''''
'''''        Me.MousePointer = vbHourglass
'''''        objDB.SetConnection mCnn
'''''        For mLoop = 1 To vsGrid.Rows - 1
'''''            If val(vsGrid.TextMatrix(mLoop, 12)) = 1 Then
'''''                If IsDate(vsGrid.TextMatrix(mLoop, 7)) Then
'''''                    ' UPDATE TRANSACTION CHILD WITH REALIZATION DATE
'''''                    arrInput = Array(val(vsGrid.TextMatrix(mLoop, 13)), val(vsGrid.TextMatrix(mLoop, 14)), _
'''''                                     val(vsGrid.TextMatrix(mLoop, 7)), val(vsGrid.TextMatrix(mLoop, 15)))
'''''                    objDB.ExecuteSP "spUpdateReconStatusToTrChild", arrInput, , , mCnn, adCmdStoredProc
'''''                Else
'''''                    'UNDO RECONCILIATION
'''''                    arrInput = Array(val(vsGrid.TextMatrix(mLoop, 13)), val(vsGrid.TextMatrix(mLoop, 14)))
'''''                    objDB.ExecuteSP "spUndoReconStatusInTrChild", arrInput, , , mCnn, adCmdStoredProc
'''''                End If
'''''            End If
'''''            vsGrid.Cell(flexcpData, mLoop, 12) = ""
'''''        Next mLoop
'''''        Call FillGrid
'''''        Me.MousePointer = vbDefault
''''    End Sub
    


Private Sub cmdReport_Click()
            Dim rptFileName As String
            Dim arrInput As Variant
            Set arrInput = Nothing
            Dim Rpt As New CRAXDRT.Report
            Dim mApp As New CRAXDRT.Application
            Dim mLoop As Long
            Dim mYear   As Integer
            crvReport.Visible = True
            picCloseReport.Visible = True
            'mvarRptFileName = App.Path & "..\Reports\rptLedgerView.rpt"
            Debug.Print App.Path & "\Reports\rptLedgerView.rpt"
            
            rptFileName = App.Path & "\Reports\rptReconciliation.rpt"
            
            Screen.MousePointer = vbHourglass
            crvReport.DisplayToolbar = True
            crvReport.Top = 2055
            crvReport.Width = 12885
            crvReport.Height = 6555
            picCloseReport.Visible = True
            
            Set Rpt = Nothing
            mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
            Set Rpt = mApp.OpenReport(rptFileName, 1)
            If Month(mdtLastDate) > 3 Then
                mYear = Year(mdtLastDate)
            Else
                 mYear = Year(mdtLastDate) - 1
            End If
            'Rpt.ParameterFields.Item(1).ClearCurrentValueAndRange
            Rpt.ParameterFields.Item(1).AddCurrentValue mintBankAccountHeadID
            Rpt.ParameterFields.Item(2).AddCurrentValue Month(mdtLastDate)
            Rpt.ParameterFields.Item(3).AddCurrentValue mYear
'
'            If IsArray(arrInput) Then
'                For mLoop = LBound(arrInput) To UBound(arrInput)
'                    Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
'                    Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
'                Next mLoop
'            End If
            Screen.MousePointer = vbDefault
            
            crvReport.ReportSource = Rpt
            crvReport.Refresh
            'crvReport.Left = 0
            'crvReport.Top = 0
            
            crvReport.ViewReport
            crvReport.Zoom (1)
End Sub

    Private Sub cmdShowReconciledItems_Click()
        mFlag = 1
        Call FillGrid
        mFlag = 0
    End Sub

Private Sub Command1_Click()

End Sub

    Private Sub Form_Activate()
        Call FillGrid
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            ClearVoucherFields
            mRowSelected = 0
            vsGrid.Select mRowSelected, 0, mRowSelected, 10
        End If
    End Sub
    
    Private Sub Form_Load()
        Call FormInitialize
    End Sub
        
    Private Sub Form_Unload(Cancel As Integer)
        frmReconBankList.cmdClickBankListGrid.value = True
    End Sub

Private Sub Picture4_Click()

End Sub

Private Sub picCloseReport_Click()
        crvReport.Visible = False
        picCloseReport.Visible = False
End Sub

    Private Sub txtRealisationDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            
            
            Dim mStr As String
            If Not IsDate(txtRealisationDate) Then
                txtRealisationDate.Text = ""
            End If
            
            If Len(Trim(txtRealisationDate)) Then
                txtRealisationDate.Text = CheckDateInMMM(txtRealisationDate.Text)
            End If
            
            If IsDate(txtRealisationDate) Then
                If CDate(txtRealisationDate) > mdtLastDate Then
                    mStr = "Realization date could not be greater than " & DdMmmYy(CDate(mdtLastDate))
                    MsgBox mStr, vbInformation
                    Call txtRealisationDate_GotFocus
                    txtRealisationDate.SetFocus
                    Exit Sub
                End If
                If vsGrid.TextMatrix(mRowSelected, 1) = "CR" Or vsGrid.TextMatrix(mRowSelected, 1) = "DR" Then
                    If val(txtRefVoucherNo.Tag) <= 0 Then
                        MsgBox "Specify the Voucher against which the Transaction is Reconcilied!", vbInformation
                        txtRefVoucherNo.SetFocus
                        Exit Sub
                    End If
                End If
                
                '@intReconChdID  As BigInt,
                '@tnyFlag    As tinyint,
                '@dtRealization  AS smalldatetime,
                '@vchNote    As varchar(200) = Null,
                '@intRefVoucherID As Bigint =Null
                
                Dim objdb As New clsDB
                Dim mArrIn As Variant
                Dim Rec As New ADODB.Recordset
                
                mArrIn = Array(vsGrid.TextMatrix(mRowSelected, 15), _
                                1, _
                                CDate(txtRealisationDate.Text), _
                                IIf(Trim(txtNote.Text) = "", Null, Trim(txtNote.Text)), _
                                IIf(val(txtRefVoucherNo.Tag) > 0, val(txtRealisationDate.Tag), Null))
                
                Set Rec = objdb.ExecuteSP("spReconcile", mArrIn)
                Call CheckReconTally(Rec)
                'Rec.Close
                
                vsGrid.TextMatrix(mRowSelected, 7) = txtRealisationDate.Text
                vsGrid.TextMatrix(mRowSelected, 8) = 1
                vsGrid.TextMatrix(mRowSelected, 12) = 1
                cmdUpdate.Enabled = True
                vsGrid.Cell(flexcpBackColor, mRowSelected, 0, mRowSelected, 15) = &HD6FFD6
            End If
            
            KeyAscii = 0
            If mRowSelected < vsGrid.Rows - 1 Then
                mRowSelected = mRowSelected + 1
            Else
                If vsGrid.Rows >= 1 Then
                    mRowSelected = 1
                End If
            End If
            vsGrid.Select mRowSelected, 0, mRowSelected, 11
            vsGrid.ShowCell mRowSelected, 0
            Call DisplayVoucher
        End If
    End Sub
    
    Private Sub txtRealisationDate_LostFocus()
        If Len(Trim(txtRealisationDate)) Then
            If IsDate(txtRealisationDate) Then
                txtRealisationDate.Text = CheckDateInMMM(txtRealisationDate.Text)
            Else
                txtRealisationDate.Text = ""
            End If
        End If
    End Sub
    
    Private Sub txtRefVoucherNo_LostFocus()
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        
        If Trim(txtRefVoucherNo.Text) <> "" Then
            mSql = "SELECT * FROM faVouchers WHERE intVoucherNO = '" & txtRefVoucherNo.Text & "'"
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn, adOpenStatic
            If Not (Rec.BOF And Rec.EOF) Then
                txtRefVoucherNo.Tag = Rec!intVoucherID
                txtRefVoucherNo.Text = Rec!intVoucherNo
            Else
                txtRefVoucherNo.Tag = ""
                txtRefVoucherNo.Text = ""
            End If
            Rec.Close
        Else
            txtRefVoucherNo.Tag = ""
            txtRefVoucherNo.Text = ""
        End If
    End Sub
    
    Private Sub vsGrid_DblClick()
        mRowSelected = vsGrid.Row
        Call DisplayVoucher
        cmdReconcile.Enabled = Not mReconStatus
        txtRealisationDate.Enabled = Not mReconStatus
    End Sub
 
    Private Sub txtRealisationDate_GotFocus()
        If txtRealisationDate.Text <> "" Then
            txtRealisationDate.Text = CheckDateInMMM(txtRealisationDate.Text)
            txtRealisationDate.SelStart = 0
            txtRealisationDate.SelLength = Len(txtRealisationDate)
        End If
    End Sub
    
    
    Public Property Get LastDate() As Variant
        LastDate = mdtLastDate
    End Property
    
    Public Property Let LastDate(mData As Variant)
        mdtLastDate = mData
    End Property
    
    Public Property Get BankAccountHeadID() As Variant
        BankAccountHeadID = mintBankAccountHeadID
    End Property
    Public Property Let BankAccountHeadID(mData As Variant)
        mintBankAccountHeadID = mData
    End Property
     
    Public Property Get ReconID() As Variant
        ReconID = mintReconID
    End Property
    
    Public Property Let ReconID(mData As Variant)
        mintReconID = mData
    End Property
    
    Public Property Get BankBalance() As Variant
        BankBalance = mBankBalance
    End Property
    
    Public Property Let BankBalance(mData As Variant)
        mBankBalance = mData
    End Property
    
    Public Property Get PassBookBalance() As Variant
        PassBookBalance = mPassBookBalance
    End Property
    
    Public Property Let PassBookBalance(mData As Variant)
        mPassBookBalance = mData
    End Property
    
    Public Property Let ReconciliationStarted(mData As Variant)
        mReconciliationStarted = mData
    End Property
    
    Public Property Let ReconcileStatus(mData As Variant)
        mReconStatus = mData
    End Property
   
    
Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    
    If mReconStatus Then Exit Sub ' ALREADY LOCKED - RECONCILIATION COMPLETED
    
    objdb.SetConnection mCnn
    If mCnn.State = 0 Then Exit Sub ' CONNECTION FAILED
    If KeyCode = vbKeyDelete Then
    If mRowSelected > 0 Then
        If vsGrid.TextMatrix(mRowSelected, 1) = "CR" Or vsGrid.TextMatrix(mRowSelected, 1) = "DR" Then
            If MsgBox("Do you want to remove this entry?", vbYesNo + vbDefaultButton2) = vbYes Then
                mSql = "DELETE FROM faBankReconcileChild WHERE intLinkReconChdID IS NULL AND intReconChdID = " & val(vsGrid.TextMatrix(mRowSelected, 15))
                mCnn.Execute mSql
                Call FillGrid
            
            End If
        End If
    End If
    End If
End Sub

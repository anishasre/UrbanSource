VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTransactionChild 
      Height          =   1125
      Left            =   90
      TabIndex        =   18
      Top             =   2190
      Width           =   7755
      Begin VB.ComboBox cmbAccountHead 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5700
         TabIndex        =   23
         Top             =   180
         Width           =   1965
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   345
         Left            =   6060
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Caption         =   "Add"
         Height          =   345
         Left            =   3660
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "Credit"
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   720
         Width           =   1305
      End
      Begin VB.OptionButton optDebit 
         Caption         =   "Debit"
         Height          =   405
         Left            =   1680
         TabIndex        =   19
         Top             =   690
         Width           =   1485
      End
      Begin VB.Label Label7 
         Caption         =   "Account Head"
         Height          =   345
         Left            =   150
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "External Account Head"
         Height          =   315
         Left            =   3870
         TabIndex        =   25
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame fraTransactionType 
      Height          =   1995
      Left            =   90
      TabIndex        =   3
      Top             =   150
      Width           =   7755
      Begin VB.ComboBox cmbBudgetCentre 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1995
      End
      Begin VB.ComboBox cmbGroup 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1560
         Width           =   1995
      End
      Begin VB.ComboBox cmbExternalModule 
         Height          =   315
         Left            =   5670
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1095
         Width           =   1875
      End
      Begin VB.ComboBox cmbExternalApplication 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1995
      End
      Begin VB.ComboBox cmbFund 
         Height          =   315
         Left            =   5670
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   585
         Width           =   1875
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1770
         TabIndex        =   5
         Top             =   150
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5670
         TabIndex        =   4
         Top             =   150
         Width           =   1845
      End
      Begin VB.Label Label5 
         Caption         =   "Group"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1590
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "External Module"
         Height          =   285
         Left            =   3900
         TabIndex        =   16
         Top             =   1110
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "External Application"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1110
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Fund"
         Height          =   285
         Left            =   3900
         TabIndex        =   14
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Budget Centre"
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Transaction Type"
         Height          =   375
         Left            =   150
         TabIndex        =   12
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label9 
         Caption         =   "New Transaction Type"
         Height          =   405
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Width           =   1635
      End
   End
   Begin VB.Frame fraGrid 
      Height          =   2835
      Left            =   90
      TabIndex        =   0
      Top             =   3420
      Width           =   7755
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   6180
         TabIndex        =   1
         Top             =   2310
         Width           =   1185
      End
      Begin VSFlex8LCtl.VSFlexGrid vsTransactionType 
         Height          =   1755
         Left            =   270
         TabIndex        =   2
         Top             =   360
         Width           =   7005
         _cx             =   12356
         _cy             =   3096
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAddTransactionType.frx":0000
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    lSubPopulateCombos
End Sub
Private Sub lSubPopulateCombos()
    PopulateList cmbBudgetCentre, "SELECT vchBudgetCentre,intBudgetCentreID FROM faBudgetCentres", , True, True, True
End Sub

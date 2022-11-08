VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSn_WrBillListOfTransactionDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Transaction Details"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   Icon            =   "frmSn_WrBillListOfTransactionDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGeneratePV 
      Caption         =   "Generate Pmt.Voucher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9315
      TabIndex        =   19
      Top             =   6285
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbKWAInstitutionType 
      Height          =   315
      Left            =   3825
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6300
      Width           =   1350
   End
   Begin VB.CommandButton cmdSearchKWAInstitution 
      Caption         =   "..."
      Height          =   285
      Left            =   7140
      TabIndex        =   18
      Top             =   6315
      Width           =   300
   End
   Begin VB.TextBox txtKWAInstitution 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   6300
      Width           =   1050
   End
   Begin VB.CommandButton cmdPVGeneratedList 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search PV Gnd.List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1545
      Width           =   2010
   End
   Begin VB.CommandButton cmdSearchPOGeneratedList 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search PO Gnd.List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1545
      Width           =   2025
   End
   Begin VB.CommandButton cmdSearchVerifiedList 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Vrfd.List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5940
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1545
      Width           =   1725
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   6315
      Width           =   1290
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   15
      TabIndex        =   33
      Top             =   -60
      Width           =   11835
      Begin VB.CommandButton cmdSearchOffice 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3045
         TabIndex        =   1
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox txtInstitution 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   525
         Width           =   1500
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10290
         TabIndex        =   7
         Top             =   525
         Width           =   1155
      End
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10290
         TabIndex        =   6
         Top             =   210
         Width           =   1155
      End
      Begin VB.ComboBox cmbInstitutionType 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   2010
      End
      Begin VB.ComboBox cmbInstitutionSubType 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   525
         Width           =   2010
      End
      Begin VB.TextBox txtConsumerNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7860
         TabIndex        =   4
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7860
         TabIndex        =   5
         Top             =   525
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearchCaretaker 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3045
         TabIndex        =   0
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtCaretaker 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   210
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   11460
         TabIndex        =   36
         Top             =   555
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   15925249
         CurrentDate     =   40258
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   11460
         TabIndex        =   37
         Top             =   210
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   15925249
         CurrentDate     =   40258
      End
      Begin VB.Label lblInstitution 
         AutoSize        =   -1  'True
         Caption         =   "Office/Institution"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   45
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9630
         TabIndex        =   43
         Top             =   555
         Width           =   645
      End
      Begin VB.Label lblFromDate 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9435
         TabIndex        =   42
         Top             =   210
         Width           =   840
      End
      Begin VB.Label lblInstitutionType 
         AutoSize        =   -1  'True
         Caption         =   "Inst.Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3825
         TabIndex        =   41
         Top             =   225
         Width           =   765
      End
      Begin VB.Label lblInstitutionSubType 
         AutoSize        =   -1  'True
         Caption         =   "Inst.Sub Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3465
         TabIndex        =   40
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label lblConsumerNo 
         AutoSize        =   -1  'True
         Caption         =   "Consumer No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6735
         TabIndex        =   39
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7305
         TabIndex        =   38
         Top             =   555
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Care Taker"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   615
         TabIndex        =   35
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   15
      TabIndex        =   24
      Top             =   840
      Width           =   11835
      Begin VB.CommandButton cmdSearchSection 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11085
         TabIndex        =   11
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9570
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearchSubDivision 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8340
         TabIndex        =   10
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtSubDivision 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearchDivision 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5265
         TabIndex        =   9
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtDivision 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   1500
      End
      Begin VB.CommandButton cmdSearchCircle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2325
         TabIndex        =   8
         Top             =   225
         Width           =   315
      End
      Begin VB.TextBox txtCircle 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label lblCircle 
         AutoSize        =   -1  'True
         Caption         =   "Circle"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   285
         TabIndex        =   28
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblSection 
         AutoSize        =   -1  'True
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8910
         TabIndex        =   27
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblDivision 
         AutoSize        =   -1  'True
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3120
         TabIndex        =   26
         Top             =   255
         Width           =   600
      End
      Begin VB.Label lblSubDivision 
         AutoSize        =   -1  'True
         Caption         =   "SubDivision"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5880
         TabIndex        =   25
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdVerifyBill 
      Caption         =   "Verify Bills"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   21
      Top             =   6285
      Width           =   1275
   End
   Begin VB.CommandButton cmdGeneratePO 
      Caption         =   "Generate Pmt.Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9585
      TabIndex        =   23
      Top             =   6285
      Width           =   2265
   End
   Begin VB.CommandButton cmdViewBill 
      Caption         =   "View Bill"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   45
      TabIndex        =   20
      Top             =   6285
      Width           =   1230
   End
   Begin VB.CommandButton cmdNewBill 
      Caption         =   "New Bill"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   135
      TabIndex        =   22
      Top             =   6630
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4665
      TabIndex        =   12
      Top             =   1545
      Width           =   1230
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   4335
      Left            =   15
      TabIndex        =   16
      Top             =   1905
      Width           =   11820
      _cx             =   20849
      _cy             =   7646
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorAlternate=   16777215
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
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSn_WrBillListOfTransactionDetails.frx":1CCA
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KWA Inst.Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2610
      TabIndex        =   56
      Top             =   6300
      Width           =   1185
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KWA Institution"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4785
      TabIndex        =   55
      Top             =   6300
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PV Generated"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3255
      TabIndex        =   53
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   210
      Left            =   3000
      TabIndex        =   52
      Top             =   1605
      Width           =   240
   End
   Begin VB.Label lblPO 
      AutoSize        =   -1  'True
      Caption         =   "PO Generated"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1665
      TabIndex        =   51
      Top             =   1590
      Width           =   1185
   End
   Begin VB.Label lblVrfd 
      AutoSize        =   -1  'True
      Caption         =   "Verified"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   450
      TabIndex        =   50
      Top             =   1590
      Width           =   645
   End
   Begin VB.Label lblForward 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   210
      Left            =   1395
      TabIndex        =   49
      Top             =   1605
      Width           =   240
   End
   Begin VB.Label lblVerified 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   210
      Left            =   180
      TabIndex        =   48
      Top             =   1605
      Width           =   240
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7455
      TabIndex        =   47
      Top             =   6360
      Width           =   495
   End
End
Attribute VB_Name = "frmSn_WrBillListOfTransactionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    'tnyStatus in snWrBillDetails
    '0-Default, CareTaker.Vrfd.-1, CareTaker.Fwd.-2, AccountsClerk.Vrfd.-3, PaymentOrder-4, PaymentVoucher-5, Despatch-6, AC.Rjd-9
    '*********************************************************************************************'
    '                                   Form to list all the Water Bills                          '
    '*********************************************************************************************'
    Private Sub Calculate()
        Dim mTotal       As Double
        Dim mCount       As Integer
        Dim mSelect      As Boolean
        
        mSelect = False
        For mCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mCount, 21) = 1 Then 'And vsGrid.TextMatrix(mCount, 22) = 0 Then
                If val(vsGrid.TextMatrix(mCount, 7)) <> 0 Then
                    mTotal = mTotal + Format(val(vsGrid.TextMatrix(mCount, 7)), "0.00")
                    txtTotal.Text = Format(mTotal, "0.00")
                    mSelect = True
        '                Else
        '                    txtTotal.Text = Format(0, "0.00")
                End If
            End If
        Next
        If mSelect = False Then
            txtTotal.Text = Format(0, "0.00")
        End If
    End Sub
        
    Private Sub FillvsGrid(mType As Integer)
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim rec         As New ADODB.Recordset
        Dim mSQL        As String
        Dim mRowCount   As Integer
        Dim mCareTaker  As String
        Dim mFromDate   As String
        Dim mToDate     As String
        
        On Error GoTo err
        objDb.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
                
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        mRowCount = 1
        
        If txtCaretaker.Tag = "" Then
            mCareTaker = "%"
        Else
            mCareTaker = txtCaretaker.Tag
        End If
        
        If txtFromDate.Text = "" Then
            mFromDate = CheckDateInMMM(Date)
        Else
            mFromDate = CheckDateInMMM(txtFromDate.Text)
        End If
        If txtToDate.Text = "" Then
            mToDate = CheckDateInMMM(Date)
        Else
            mToDate = CheckDateInMMM(txtToDate.Text)
        End If
        
        mSQL = "Select snWrBillMastersOfficeInstitution.chvName As OfficeName,snWrBillMastersOfficeInstitution.intID As OfficeID,snWrBillMastersOfficeInstitution.intWardNo As Ward,snWrBillMastersOfficeInstitution.numWardID As WardID,snWrBillDetails.chvConsumerNo As ConsumerNo,snWrBillMastersSection.chvName As Section,snWrBillMastersSection.intID As SectionID,snWrBillDetails.dtBillDate As BillDate,snWrBillDetails.dtBillDueDate As DueDate,snWrBillDetails.fltTotAmount As Amount,snWrBillDetails.intBillID As BillID,snWrBillDetails.intConnID As ConnectionID,snWrBillDetails.chvBillNo As BillNo ,snWrBillDetails.chvRemarks As Remarks,snWrBillConnections.intCareTakerID As CareTakerID,snWrBillMastersCareTakers.chvName As CareTaker,snWrBillDetails.dtBillFromDate[BillFromDate],snWrBillDetails.dtBillToDate[BillToDate],snWrBillDetails.numLastReading[LastReading],snWrBillDetails.numCurrentReading[CurrentReading],snWrBillDetails.tnyStatus[Status],"
        mSQL = mSQL + " snWrBillDetails.vchPayOrderNo[PayOrderNo],snWrBillDetails.intPayOrderID[PayOrderID],snWrBillDetails.intVoucherNo[VoucherNo],snWrBillDetails.intVoucherID[VoucherID],snWrBillDetails.intKWAInstitutionID[KWAInstitutionID],snWrBillDetails.intKWAInstitutionTypeID[KWAInstitutionTypeID] From snWrBillDetails"
        mSQL = mSQL + " Left Join snWrBillConnections On snWrBillDetails.intConnID = snWrBillConnections.intID"
        mSQL = mSQL + " Left Join snWrBillMastersCareTakers On snWrBillConnections.intCareTakerID = snWrBillMastersCareTakers.intID"
        mSQL = mSQL + " Left Join snWrBillMastersOfficeInstitution On snWrBillConnections.intOfficeID = snWrBillMastersOfficeInstitution.intID"
        mSQL = mSQL + " Left Join snWrBillMastersSection On snWrBillConnections.numSection = snWrBillMastersSection.intID"
        mSQL = mSQL + " Left Join snWrBillKWAInstitutionTypes On snWrBillDetails.intKWAInstitutionTypeID = snWrBillKWAInstitutionTypes.intInstitutionTypeID"
        mSQL = mSQL + " Where snWrBillConnections.intCareTakerID Like '" & mCareTaker & "'"
        mSQL = mSQL + " And snWrBillDetails.dtBillDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
        mSQL = mSQL + " And snWrBillDetails.tnyStatus = " & mType
        mSQL = mSQL + " And snWrBillDetails.numAccClerkSeatID = " & gbSeatID
        If txtInstitution.Tag <> "" Then
            mSQL = mSQL + " And snWrBillConnections.intOfficeID = " & txtInstitution.Tag
        End If
        If cmbInstitutionType.ListIndex > 0 Then
            mSQL = mSQL + " And snWrBillMastersOfficeInstitution.intInstTypeID = " & cmbInstitutionType.ItemData(cmbInstitutionType.ListIndex)
        End If
        If cmbInstitutionSubType.ListIndex > 0 Then
            mSQL = mSQL + " And snWrBillMastersOfficeInstitution.intInstSubTypeID = " & cmbInstitutionSubType.ItemData(cmbInstitutionSubType.ListIndex)
        End If
        If txtConsumerNo.Text <> "" Then
            mSQL = mSQL + " And snWrBillDetails.chvConsumerNo = " & txtConsumerNo.Text
        End If
        If txtBillNo.Text <> "" Then
            mSQL = mSQL + " And snWrBillDetails.chvBillNo = " & txtBillNo.Text
        End If
        If txtCircle.Tag <> "" Then
            mSQL = mSQL + " And intKWAInstitutionTypeID = 1"
            mSQL = mSQL + " And intKWAInstitutionID = " & txtCircle.Tag
        End If
        If txtDivision.Tag <> "" Then
            mSQL = mSQL + " And intKWAInstitutionTypeID = 2"
            mSQL = mSQL + " And intKWAInstitutionID = " & txtDivision.Tag
        End If
        If txtSubDivision.Tag <> "" Then
            mSQL = mSQL + " And intKWAInstitutionTypeID = 3"
            mSQL = mSQL + " And intKWAInstitutionID = " & txtSubDivision.Tag
        End If
        If txtSection.Tag <> "" Then
            mSQL = mSQL + " And intKWAInstitutionTypeID = 4"
            mSQL = mSQL + " And intKWAInstitutionID = " & txtSection.Tag
        End If
        rec.Open mSQL, mCnn
        While Not rec.EOF
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.TextMatrix(mRowCount, 0) = mRowCount
            vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(rec!OfficeName), "", rec!OfficeName)
            vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(rec!ConsumerNo), "", rec!ConsumerNo)
            vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(rec!Section), "", rec!Section)
            vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(rec!BillDate), "", rec!BillDate)
            vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(rec!dueDate), "", rec!dueDate)
            If mType = 4 Then 'Payment Order Generated
                vsGrid.ColWidth(6) = 1350
                vsGrid.TextMatrix(0, 6) = "Pmnt.Odr.No."
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(rec!PayOrderNo), "", rec!PayOrderNo)
                vsGrid.TextMatrix(mRowCount, 23) = IIf(IsNull(rec!PayOrderNo), "", rec!PayOrderID)
                vsGrid.TextMatrix(mRowCount, 24) = IIf(IsNull(rec!PayOrderNo), "", rec!PayOrderNo)
                vsGrid.ColWidth(1) = 1890
                vsGrid.ColWidth(3) = 1750
            ElseIf mType = 5 Then 'Payment Voucher Generated
                vsGrid.ColWidth(6) = 1350
                vsGrid.TextMatrix(0, 6) = "VoucherNo"
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(rec!PayOrderNo), "", rec!VoucherNo)
                vsGrid.TextMatrix(mRowCount, 25) = IIf(IsNull(rec!PayOrderNo), "", rec!VoucherID)
                vsGrid.TextMatrix(mRowCount, 26) = IIf(IsNull(rec!PayOrderNo), "", rec!VoucherNo)
                vsGrid.ColWidth(1) = 1890
                vsGrid.ColWidth(3) = 1750
            Else
                vsGrid.ColWidth(6) = 0
                vsGrid.ColWidth(1) = 2500
                vsGrid.ColWidth(3) = 2500
            End If
            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(rec!Amount), "", rec!Amount)
            vsGrid.TextMatrix(mRowCount, 8) = IIf(IsNull(rec!BillID), "", rec!BillID)
            vsGrid.TextMatrix(mRowCount, 9) = IIf(IsNull(rec!OfficeID), "", rec!OfficeID)
            vsGrid.TextMatrix(mRowCount, 10) = IIf(IsNull(rec!WardID), "", rec!WardID)
            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(rec!SectionID), "", rec!SectionID)
            vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(rec!ConnectionID), "", rec!ConnectionID)
            vsGrid.TextMatrix(mRowCount, 13) = IIf(IsNull(rec!BillNo), "", rec!BillNo)
            vsGrid.TextMatrix(mRowCount, 14) = IIf(IsNull(rec!Remarks), "", rec!Remarks)
            vsGrid.TextMatrix(mRowCount, 15) = IIf(IsNull(rec!CareTakerID), "", rec!CareTakerID)
            vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(rec!CareTaker), "", rec!CareTaker)
            vsGrid.TextMatrix(mRowCount, 17) = IIf(IsNull(rec!BillFromDate), "", rec!BillFromDate)
            vsGrid.TextMatrix(mRowCount, 18) = IIf(IsNull(rec!BillToDate), "", rec!BillToDate)
            vsGrid.TextMatrix(mRowCount, 19) = IIf(IsNull(rec!LastReading), "", rec!LastReading)
            vsGrid.TextMatrix(mRowCount, 20) = IIf(IsNull(rec!CurrentReading), "", rec!CurrentReading)
            If Not IsNull(rec!Status) Then
                If rec!Status = 5 Then 'Payment Voucher Generated
                    vsGrid.Cell(flexcpChecked, mRowCount, 21) = vbChecked
                Else
                    vsGrid.Cell(flexcpChecked, mRowCount, 21) = vbUnchecked
                End If
            End If
            vsGrid.TextMatrix(mRowCount, 22) = IIf(IsNull(rec!Status), 0, rec!Status)
            If mType = 2 Then 'Forwarded By CT
                vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 21) = &H80000005
            ElseIf mType = 3 Then 'Verified By AC
                vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 21) = &HC0E0FF
            ElseIf mType = 4 Then 'Payment Order Generated
                vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 21) = &HC0FFC0
            ElseIf mType = 5 Then 'Payment Voucher Generated
                vsGrid.Cell(flexcpBackColor, mRowCount, 0, , 21) = &HC0C0C0
            End If
            vsGrid.TextMatrix(mRowCount, 27) = IIf(IsNull(rec!KWAInstitutionID), "", rec!KWAInstitutionID)
            vsGrid.TextMatrix(mRowCount, 28) = IIf(IsNull(rec!KWAInstitutionTypeID), "", rec!KWAInstitutionTypeID)
            rec.MoveNext
            mRowCount = mRowCount + 1
        Wend
        rec.Close
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmbKWAInstitutionType_Click()
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
    End Sub

    Private Sub cmdGeneratePO_Click()
        On Error GoTo err
        If vsGrid.Rows > 1 Then
            If cmbKWAInstitutionType.ListIndex < 1 Then
                MsgBox "Please select the KWA Institution Type", vbInformation
                cmbKWAInstitutionType.SetFocus
                Exit Sub
            End If
            If txtKWAInstitution.Text = "" Then
                MsgBox "Please select the KWA Institution", vbInformation
                cmdSearchKWAInstitution.SetFocus
                Exit Sub
            End If
            If val(txtTotal.Text) > 0 Then
                If MsgBox("Are you sure to generate the PO for these Bills?", vbYesNo, "Confirm Verification") = vbYes Then
                    Call SetPaymentOrder
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub UpdatePOStatus()
        Dim mCnn As New ADODB.Connection
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim i As Integer
        
        '*********************************************************************************************'
        '              Procedure to Update the PO Generated Status in DB_iSaankhyaMasters             '
        '*********************************************************************************************'
        On Error GoTo err
        If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
            For i = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, i, 21) = 1 Then
'                    If vsGrid.TextMatrix(i, 22) = 1 Then
                mSQL = "Update snWrBillDetails"
                mSQL = mSQL + " Set tnyStatus = 4,"
                mSQL = mSQL + " intPayOrderID = " & frmPaymentOrder.PayOrderID & ","
                mSQL = mSQL + " vchPayOrderNo = '" & frmPaymentOrder.PayOrderNo & "',"
                mSQL = mSQL + " intKWAInstitutionTypeID = " & cmbKWAInstitutionType.ItemData(cmbKWAInstitutionType.ListIndex) & ","
                mSQL = mSQL + " intKWAInstitutionID = " & txtKWAInstitution.Tag
                mSQL = mSQL + " Where intBillID = " & vsGrid.TextMatrix(i, 8)
                mCnn.Execute mSQL
                frmPaymentOrder.WaterBillPOMode = False
'                    End If
                End If
            Next
        End If
        MsgBox "Payment Order Generated", vbInformation
        Call FillvsGrid(3)
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub UpdatePVStatus()
        Dim mCnn As New ADODB.Connection
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim i As Integer
        '*********************************************************************************************'
        '              Procedure to Update the PV Generated Status in DB_iSaankhyaMasters             '
        '*********************************************************************************************'
        On Error GoTo err
        If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
            For i = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, i, 21) = 1 Then
'                    If vsGrid.TextMatrix(i, 22) = 1 Then
                mSQL = "Update snWrBillDetails"
                mSQL = mSQL + " Set tnyStatus = 5,"
                mSQL = mSQL + " intVoucherID = " & frmIntegratedPayments.VoucherID & ","
                mSQL = mSQL + " intVoucherNo ='" & frmIntegratedPayments.VoucherNo & "'"
                mSQL = mSQL + " Where intBillID = " & vsGrid.TextMatrix(i, 8)
                mCnn.Execute mSQL
                frmIntegratedPayments.WaterBillPVMode = False
'                    End If
                End If
            Next
        End If
        'MsgBox "Payment Voucher Generated", vbInformation
        Call FillvsGrid(4)
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub cmdGeneratePV_Click()
        On Error GoTo err
        If vsGrid.Rows > 1 Then
            If cmbKWAInstitutionType.ListIndex < 1 Then
                MsgBox "Please select the KWA Institution Type", vbInformation
                cmbKWAInstitutionType.SetFocus
                Exit Sub
            End If
            If txtKWAInstitution.Text = "" Then
                MsgBox "Please select the KWA Institution", vbInformation
                cmdSearchKWAInstitution.SetFocus
                Exit Sub
            End If
            If val(txtTotal.Text) > 0 Then
                If MsgBox("Are you sure to generate the PV for these Bills?", vbYesNo, "Confirm Verification") = vbYes Then
                    Call SetPaymentVoucher
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdNewBill_Click()
        On Error GoTo err
        frmSn_WrBillDetails.intBillId = 0
        frmSn_WrBillDetails.txtConsumerNo = val(vsGrid.TextMatrix(vsGrid.row, 2))
        frmSn_WrBillDetails.txtConsumerNo.Tag = val(vsGrid.TextMatrix(vsGrid.row, 12)) 'intConnId
        frmSn_WrBillDetails.Show vbModal
        Call FillvsGrid(0)
        cmdVerifyBill.Enabled = True
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdPVGeneratedList_Click()
        On Error GoTo err
        Call FillvsGrid(5)
        txtTotal.Text = ""
        txtTotal.Tag = ""
        Call Calculate
        cmbKWAInstitutionType.ListIndex = 0
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
        cmdGeneratePV.Visible = False
        cmdGeneratePO.Enabled = False
        cmdVerifyBill.Enabled = False
        cmdNewBill.Enabled = False
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearch_Click()
        On Error GoTo err
        Call FillvsGrid(2)
        txtTotal.Text = ""
        txtTotal.Tag = ""
        cmbKWAInstitutionType.ListIndex = 0
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
        cmdGeneratePV.Visible = False
        cmdGeneratePO.Enabled = False
        cmdVerifyBill.Enabled = True
        cmdNewBill.Enabled = True
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchCaretaker_Click()
        On Error GoTo err
        intWrBillSearchID = 5
        txtCaretaker.Text = ""
        txtCaretaker.Tag = ""
        frmSn_WrBillSearchName.Show 1
        If Not gbSearchStr = "" Then
            txtCaretaker.Text = gbSearchStr
            txtCaretaker.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSearchCircle_Click()
        On Error GoTo err
        intWrBillSearchID = 8
        intWrBillCircleID = 0
        intWrBillDivisionID = 0
        intWrBillSubDivisionID = 0
        txtCircle.Text = ""
        txtCircle.Tag = ""
        txtDivision.Text = ""
        txtDivision.Tag = ""
        txtSubDivision.Text = ""
        txtSubDivision.Tag = ""
        txtSection.Text = ""
        txtSection.Tag = ""
        frmSn_WrBillSearchName.Show vbModal
        If Not gbSearchStr = "" Then
            txtCircle.Text = gbSearchStr
            txtCircle.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchDivision_Click()
        On Error GoTo err
        intWrBillSearchID = 9
        intWrBillCircleID = 0
        intWrBillDivisionID = 0
        intWrBillSubDivisionID = 0
        txtDivision.Text = ""
        txtDivision.Tag = ""
        txtSubDivision.Text = ""
        txtSubDivision.Tag = ""
        txtSection.Text = ""
        txtSection.Tag = ""
        If txtCircle.Tag <> "" Then
            intWrBillCircleID = txtCircle.Tag
        End If
        frmSn_WrBillSearchName.Show vbModal
        If Not gbSearchStr = "" Then
            txtDivision.Text = gbSearchStr
            txtDivision.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchKWAInstitution_Click()
        On Error GoTo err
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
        intWrBillSearchID = 0
        If cmbKWAInstitutionType.ListIndex > 0 Then
            If cmbKWAInstitutionType.ItemData(cmbKWAInstitutionType.ListIndex) = 1 Then
                intWrBillSearchID = 8
            ElseIf cmbKWAInstitutionType.ItemData(cmbKWAInstitutionType.ListIndex) = 2 Then
                intWrBillSearchID = 9
            ElseIf cmbKWAInstitutionType.ItemData(cmbKWAInstitutionType.ListIndex) = 3 Then
                intWrBillSearchID = 10
            ElseIf cmbKWAInstitutionType.ItemData(cmbKWAInstitutionType.ListIndex) = 4 Then
                intWrBillSearchID = 11
            End If
            frmSn_WrBillSearchName.Show 1
            If Not gbSearchStr = "" Then
                txtKWAInstitution.Text = gbSearchStr
                txtKWAInstitution.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        End If
        
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchPOGeneratedList_Click()
        On Error GoTo err
        Call FillvsGrid(4)
        txtTotal.Text = ""
        txtTotal.Tag = ""
        cmbKWAInstitutionType.ListIndex = 0
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
        cmdGeneratePV.Visible = True
        cmdGeneratePO.Enabled = False
        cmdVerifyBill.Enabled = False
        cmdNewBill.Enabled = False
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchOffice_Click()
        On Error GoTo err
        intWrBillSearchID = 12
        intWrBillCaretakerID = 0
        txtInstitution.Text = ""
        txtInstitution.Tag = ""
        If txtCaretaker.Tag <> "" Then
            intWrBillCaretakerID = txtCaretaker.Tag
        End If
        frmSn_WrBillSearchName.Show 1
        If Not gbSearchStr = "" Then
            txtInstitution.Text = gbSearchStr
            txtInstitution.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchSection_Click()
        On Error GoTo err
        intWrBillSearchID = 11
        intWrBillCircleID = 0
        intWrBillDivisionID = 0
        intWrBillSubDivisionID = 0
        txtSection.Text = ""
        txtSection.Tag = ""
        If txtCircle.Tag <> "" Then
            intWrBillCircleID = txtCircle.Tag
        End If
        If txtDivision.Tag <> "" Then
            intWrBillDivisionID = txtDivision.Tag
        End If
        If txtSubDivision.Tag <> "" Then
            intWrBillSubDivisionID = txtSubDivision.Tag
        End If
        frmSn_WrBillSearchName.Show vbModal
        If Not gbSearchStr = "" Then
            txtSection.Text = gbSearchStr
            txtSection.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchSubDivision_Click()
        On Error GoTo err
        intWrBillSearchID = 10
        intWrBillCircleID = 0
        intWrBillDivisionID = 0
        intWrBillSubDivisionID = 0
        txtSubDivision.Text = ""
        txtSubDivision.Tag = ""
        txtSection.Text = ""
        txtSection.Tag = ""
        If txtCircle.Tag <> "" Then
            intWrBillCircleID = txtCircle.Tag
        End If
        If txtDivision.Tag <> "" Then
            intWrBillDivisionID = txtDivision.Tag
        End If
        frmSn_WrBillSearchName.Show vbModal
        If Not gbSearchStr = "" Then
            txtSubDivision.Text = gbSearchStr
            txtSubDivision.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdSearchVerifiedList_Click()
        On Error GoTo err
        Call FillvsGrid(3)
        txtTotal.Text = ""
        txtTotal.Tag = ""
        cmbKWAInstitutionType.ListIndex = 0
        txtKWAInstitution.Text = ""
        txtKWAInstitution.Tag = ""
        cmdGeneratePV.Visible = False
        cmdGeneratePO.Enabled = True
        cmdVerifyBill.Enabled = False
        cmdNewBill.Enabled = False
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdVerifyBill_Click()
        Dim mCnn As New ADODB.Connection
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim i As Integer
        Dim j As Integer
        
        On Error GoTo err
        j = 0
        If MsgBox("Are you sure to Verify these Bills?", vbYesNo, "Confirm Verification") = vbYes Then
            If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
                For i = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpChecked, i, 21) = vbChecked Then
                        mSQL = "Update snWrBillDetails"
                        mSQL = mSQL + " Set tnyStatus = 3,"
                        mSQL = mSQL + " numAccClerkID = " & gbUserID & ","
                        mSQL = mSQL + " numAccClerkSeatID = " & gbSeatID & ","
                        mSQL = mSQL + " dtAccClerkVerifiedDate =  '" & DdMmmYy(gbTransactionDate) & "'"
                        mSQL = mSQL + " Where intBillID = " & vsGrid.TextMatrix(i, 8)
                        mCnn.Execute mSQL
                        j = j + 1
                    End If
                Next
            End If
            MsgBox j & " Bills Verified", vbInformation
        End If
        Call FillvsGrid(2)
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmdViewBill_Click()
        On Error GoTo err
        If vsGrid.row > 0 Then
            frmSn_WrBillDetails.intBillId = val(vsGrid.TextMatrix(vsGrid.row, 8)) 'BillID
            frmSn_WrBillDetails.txtCaretaker.Text = vsGrid.TextMatrix(vsGrid.row, 16)
            frmSn_WrBillDetails.txtCaretaker.Tag = val(vsGrid.TextMatrix(vsGrid.row, 15))
            frmSn_WrBillDetails.txtOfficeInst.Text = vsGrid.TextMatrix(vsGrid.row, 1)
            frmSn_WrBillDetails.txtOfficeInst.Tag = val(vsGrid.TextMatrix(vsGrid.row, 9))
            frmSn_WrBillDetails.txtConsumerNo = val(vsGrid.TextMatrix(vsGrid.row, 2))
            frmSn_WrBillDetails.txtConsumerNo.Tag = val(vsGrid.TextMatrix(vsGrid.row, 12)) 'intConnId
            frmSn_WrBillDetails.txtBillNo.Text = val(vsGrid.TextMatrix(vsGrid.row, 13))
            frmSn_WrBillDetails.dtBillDate.value = CDate(vsGrid.TextMatrix(vsGrid.row, 4))
            frmSn_WrBillDetails.dtBillDueDate.value = CDate(vsGrid.TextMatrix(vsGrid.row, 5))
            frmSn_WrBillDetails.txtRemarks.Text = vsGrid.TextMatrix(vsGrid.row, 14)
            frmSn_WrBillDetails.dtpBillFrom = CDate(vsGrid.TextMatrix(vsGrid.row, 17))
            frmSn_WrBillDetails.dtpBillTo = CDate(vsGrid.TextMatrix(vsGrid.row, 18))
            frmSn_WrBillDetails.txtReading1.Text = val(vsGrid.TextMatrix(vsGrid.row, 19))
            frmSn_WrBillDetails.txtReading2.Text = val(vsGrid.TextMatrix(vsGrid.row, 20))
            If vsGrid.TextMatrix(vsGrid.row, 22) > 0 Then
                frmSn_WrBillDetails.cmdSave.Enabled = False
            End If
            frmSn_WrBillDetails.Show vbModal
            'Call FillvsGrid(0)
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub dtpFromDate_CloseUp()
        txtFromDate.Text = dtpFromDate.value
    End Sub
    
    Private Sub dtpToDate_CloseUp()
        txtToDate.Text = dtpToDate.value
    End Sub

    Private Sub Form_Activate()
        On Error GoTo err
        Me.Width = 11970
        Me.Height = 7155
        Me.Left = 0
        Me.Top = 0
        If frmPaymentOrder.WaterBillPOMode = True Then
            UpdatePOStatus
        End If
        If frmIntegratedPayments.WaterBillPVMode = True Then
            UpdatePVStatus
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub Form_Load()
        On Error GoTo err
        Dim mSQL As String
        
        mSQL = "Select vchInstitutionType,intInstitutionTypeID From snWrBillKWAInstitutionTypes"
        PopulateList cmbKWAInstitutionType, mSQL, , True, True, True, enuSourceString.iSaankhyaMasters
        txtFromDate.Text = CDate(Date) - 30
        txtToDate.Text = CDate(Date) + 30
        Call FillvsGrid(2)
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub txtFromDate_LostFocus()
        If txtFromDate.Text <> "" Then
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub
    
    Private Sub txtToDate_LostFocus()
        If txtToDate.Text <> "" Then
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        End If
    End Sub

    Private Sub vsGrid_Click()
        Dim mCount      As Double
        Dim mPayOrderNo As Variant
        Dim mStatus     As Variant
        Dim mSQL        As String
        Dim objDb       As New clsDB
        Dim rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        
        If vsGrid.row > 0 Then
            If vsGrid.col = 21 Then
                cmbKWAInstitutionType.ListIndex = 0
                txtKWAInstitution.Text = ""
                txtKWAInstitution.Tag = ""
                If vsGrid.TextMatrix(vsGrid.row, 22) < 5 Then
                    vsGrid.Editable = flexEDKbdMouse
                    mPayOrderNo = vsGrid.TextMatrix(vsGrid.row, 6)
                    mStatus = vsGrid.Cell(flexcpChecked, vsGrid.row, 21)
                    If val(mPayOrderNo) > 0 Then
                        If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
                            mSQL = "Select intKWAInstitutionID,intKWAInstitutionTypeID,vchInstitutionType,intID,chvName From snWrBillDetails"
                            mSQL = mSQL + " Inner Join snWrBillKWAInstitutionTypes On snWrBillDetails.intKWAInstitutionTypeID = snWrBillKWAInstitutionTypes.intInstitutionTypeID"
                            If vsGrid.TextMatrix(vsGrid.row, 28) = 1 Then
                                mSQL = mSQL + " Inner Join snWrBillMastersCircle On snWrBillDetails.intKWAInstitutionID = snWrBillMastersCircle.intID"
                            ElseIf vsGrid.TextMatrix(vsGrid.row, 28) = 2 Then
                                mSQL = mSQL + " Inner Join snWrBillMastersDivision On snWrBillDetails.intKWAInstitutionID = snWrBillMastersDivision.intID"
                            ElseIf vsGrid.TextMatrix(vsGrid.row, 28) = 3 Then
                                mSQL = mSQL + " Inner Join snWrBillMastersSubDivision On snWrBillDetails.intKWAInstitutionID = snWrBillMastersSubDivision.intID"
                            ElseIf vsGrid.TextMatrix(vsGrid.row, 28) = 4 Then
                                mSQL = mSQL + " Inner Join snWrBillMastersSection On snWrBillDetails.intKWAInstitutionID = snWrBillMastersSection.intID"
                            End If
                            mSQL = mSQL + " Where vchPayOrderNo = " & mPayOrderNo
                            rec.Open mSQL, mCnn
                            If Not (rec.EOF And rec.BOF) Then
                                cmbKWAInstitutionType.Text = IIf(IsNull(rec!vchInstitutionType), "", rec!vchInstitutionType)
                                txtKWAInstitution.Text = IIf(IsNull(rec!chvName), "", rec!chvName)
                                txtKWAInstitution.Tag = IIf(IsNull(rec!intID), "", rec!intID)
                            End If
                            rec.Close
                        End If
                        txtTotal.Tag = mPayOrderNo
                        For mCount = 1 To vsGrid.Rows - 1
                            If vsGrid.TextMatrix(mCount, 22) = 4 Then
                                If vsGrid.TextMatrix(mCount, 6) = mPayOrderNo Then
                                    vsGrid.Cell(flexcpChecked, mCount, 21) = mStatus
                                Else
                                    vsGrid.Cell(flexcpChecked, mCount, 21) = 2
                                End If
                            End If
                        Next
                    End If
                    Call Calculate
                Else
                    vsGrid.Editable = flexEDNone
                End If
            Else
                vsGrid.Editable = flexEDNone
            End If
        End If
    End Sub
        
    Private Sub SetOfficer()
        Dim mCnn    As New ADODB.Connection
        Dim mSQL    As String
        Dim objDb   As New clsDB
        Dim rec     As New ADODB.Recordset
        
        '*********************************************************************************************'
        '              Procedure to get the KWA Institution Officer from DB_iSaankhyaMasters          '
        '*********************************************************************************************'
        On Error GoTo err
        If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
            mSQL = "Select * From snWrBillMasters" + cmbKWAInstitutionType.Text
            mSQL = mSQL + " Where intID = " & txtKWAInstitution.Tag
            rec.Open mSQL, mCnn
            If Not (rec.EOF And rec.BOF) Then
                With frmPaymentOrder
                    .txtName.Text = IIf(IsNull(rec!vchOfficerName), "", rec!vchOfficerName)
                    .txtMainPlace.Text = IIf(IsNull(rec!chvLocation), "", rec!chvLocation)
                    .txtPhone.Text = IIf(IsNull(rec!intPhoneNo), "", rec!intPhoneNo)
                End With
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub SetPaymentOrder()
        Dim objAccounts         As New clsAccounts
        Dim objFunction         As New clsFunction
        Dim objFunctionary      As New clsFunctionary
        Dim objTransactionType  As New clsTransactionType
        
        '*********************************************************************************************'
        '              Procedure to set the fields for PO Generation                                  '
        '*********************************************************************************************'
        On Error GoTo err
        With frmPaymentOrder
            .WaterBillPOMode = False
            .ModuleID = 75
            .txtPayOrder.Tag = ""
            .txtPayOrder.Text = ""
            .txtFunctionary.Tag = 4
            objFunctionary.SetFunctionaryByID (4)
            .txtFunctionary.Text = objFunctionary.FunctionaryName
            .txtFunction.Tag = 6
            objFunction.SetFunctionByID (6)
            .txtFunction.Text = objFunction.FunctionName
            .txtTransactionType.Tag = 1101
            objTransactionType.SetTransactionType (1101)
            .txtTransactionType.Text = objTransactionType.TransactionType
            .txtDrHeadCode.Tag = 394
            .txtCrHeadCode.Tag = 394
            objAccounts.SetAccounts (394)
            .txtDrHeadCode.Text = objAccounts.AccountCode
            .txtDrAccountHead.Text = objAccounts.AccountHead
            .txtCrHeadCode.Text = objAccounts.AccountCode
            .txtCrAccountHead.Text = objAccounts.AccountHead
            .txtDrAmount.Text = val(txtTotal.Text)
            .txtCrAmount.Text = val(txtTotal.Text)
            Call SetOfficer
        End With
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub SetPaymentVoucher()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim rec As New ADODB.Recordset
        Dim mSQL As String
        
        '*********************************************************************************************'
        '              Procedure to set the fields for PV Generation                                  '
        '*********************************************************************************************'
        On Error GoTo err
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSQL = "Select tnyStatus,tnyCancelled From faPayOrder Where vchPayOrderNo = " & txtTotal.Tag
            rec.Open mSQL, mCnn
            If Not (rec.EOF And rec.BOF) Then
                If rec!tnyStatus = 0 Then
                    MsgBox "Please Approve the Payment Order before making Payment", vbInformation
                    'MsgBox "Payment Voucher is already generated for this Payment Order", vbInformation
                    Exit Sub
                End If
            
                'If Not IsNull(Rec!intVoucherID) Then
                If rec!tnyStatus = 2 Then
                    MsgBox "Payment Voucher is already generated for this Payment Order", vbInformation
                    Exit Sub
                End If
                
                If IsNull(rec!tnyCancelled) = False Then
                    If rec!tnyCancelled = 1 Then
                        MsgBox "The payorder is Cancelled", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            rec.Close
        End If
        With frmIntegratedPayments
            .PayOrderNo = txtTotal.Tag

            .WaterBillPVMode = False
            .Visible = True
        End With
        Exit Sub
err:
        MsgBox err.Description
    End Sub

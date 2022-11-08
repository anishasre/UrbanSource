VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchReceiptVouchers 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " List of Receipt Vouchers"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleMode       =   0  'User
   ScaleWidth      =   11970.31
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Filters"
      Height          =   405
      Left            =   9300
      TabIndex        =   58
      Top             =   6060
      Width           =   1155
   End
   Begin VB.ComboBox cmbCounterSeats 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSearchReceiptVouchers.frx":0000
      Left            =   480
      List            =   "frmSearchReceiptVouchers.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   5820
      Width           =   1800
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2670
      Left            =   30
      TabIndex        =   55
      Top             =   2865
      Width           =   11745
      _cx             =   20717
      _cy             =   4710
      Appearance      =   2
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
      Cols            =   13
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
      TextStyleFixed  =   1
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
   Begin VB.ListBox lstMasters 
      BackColor       =   &H00E8F2E8&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   36
      Top             =   555
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame fraAccountHead 
      BackColor       =   &H00DAF2F2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1260
      Left            =   6930
      TabIndex        =   30
      Top             =   -75
      Width           =   4905
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3300
         TabIndex        =   33
         Top             =   840
         Width           =   1470
      End
      Begin VB.TextBox txtBank 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1005
         TabIndex        =   32
         Top             =   840
         Width           =   1740
      End
      Begin VB.CommandButton cmdSearchInstrument 
         Caption         =   "..."
         Height          =   285
         Left            =   4455
         TabIndex        =   28
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Height          =   285
         Left            =   4455
         TabIndex        =   31
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtAccountHead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1005
         TabIndex        =   24
         Top             =   210
         Width           =   3420
      End
      Begin VB.TextBox txtInstrument 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1005
         TabIndex        =   27
         Top             =   510
         Width           =   3420
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   5
         Left            =   2790
         TabIndex        =   35
         Top             =   870
         Width           =   495
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   34
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/cHead"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   23
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   7
         Left            =   30
         TabIndex        =   26
         Top             =   540
         Width           =   915
      End
   End
   Begin VB.Frame fraReceiptNo 
      BackColor       =   &H00DAF2F2&
      Height          =   1245
      Left            =   15
      TabIndex        =   22
      Top             =   -60
      Width           =   6900
      Begin VB.ListBox lstTransactionType 
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         Top             =   615
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
         Height          =   285
         Left            =   6510
         TabIndex        =   52
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1635
         TabIndex        =   51
         Top             =   615
         Width           =   4845
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2865
         TabIndex        =   49
         Top             =   210
         Width           =   1500
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1050
         TabIndex        =   25
         Top             =   210
         Width           =   1500
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Transaction Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   13
         Left            =   135
         TabIndex        =   53
         Top             =   615
         Width           =   1485
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   3
         Left            =   2670
         TabIndex        =   50
         Top             =   210
         Width           =   165
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   29
         Top             =   225
         Width           =   405
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   2040
      Top             =   6225
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   3
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   405
      Left            =   10545
      TabIndex        =   20
      Top             =   6060
      Width           =   1155
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridTransactions 
      Height          =   2025
      Left            =   7230
      TabIndex        =   21
      Top             =   6570
      Visible         =   0   'False
      Width           =   5970
      _cx             =   10530
      _cy             =   3572
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
      Rows            =   1
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
   Begin VB.Frame fraDemandDetails 
      BackColor       =   &H00DAF2F2&
      Height          =   1605
      Left            =   -15
      TabIndex        =   37
      Top             =   1215
      Visible         =   0   'False
      Width           =   11820
      Begin VB.TextBox txtRefNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   50
         TabIndex        =   48
         Top             =   1215
         Width           =   1770
      End
      Begin VB.TextBox txtWardNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   3
         TabIndex        =   43
         Top             =   540
         Width           =   1770
      End
      Begin VB.TextBox txtDoorNo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         MaxLength       =   5
         TabIndex        =   45
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox txtDoorNo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   46
         Top             =   870
         Width           =   660
      End
      Begin VB.ComboBox cmbDZone 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmSearchReceiptVouchers.frx":0004
         Left            =   720
         List            =   "frmSearchReceiptVouchers.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   195
         Width           =   1800
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4065
         MaxLength       =   100
         TabIndex        =   1
         Top             =   210
         Width           =   2145
      End
      Begin VB.TextBox txtHouse 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4650
         MaxLength       =   100
         TabIndex        =   7
         Top             =   540
         Width           =   2820
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4650
         MaxLength       =   100
         TabIndex        =   9
         Top             =   855
         Width           =   2820
      End
      Begin VB.TextBox txtLocalPlace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4650
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1170
         Width           =   2820
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8745
         MaxLength       =   100
         TabIndex        =   13
         Top             =   225
         Width           =   2820
      End
      Begin VB.TextBox txtInit1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6210
         MaxLength       =   1
         TabIndex        =   2
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInit2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6525
         MaxLength       =   1
         TabIndex        =   3
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInit3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtInit4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7155
         MaxLength       =   1
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtPost 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8745
         MaxLength       =   50
         TabIndex        =   15
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10650
         MaxLength       =   6
         TabIndex        =   17
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8745
         MaxLength       =   30
         TabIndex        =   19
         Top             =   855
         Width           =   1635
      End
      Begin VB.ComboBox cmbSeat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8745
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1170
         Width           =   2595
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RefNo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   255
         TabIndex        =   47
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ward No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   75
         TabIndex        =   42
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   120
         TabIndex        =   44
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   330
         TabIndex        =   40
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nam&E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   3630
         TabIndex        =   0
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House/Office"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   3660
         TabIndex        =   6
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   4185
         TabIndex        =   8
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   3795
         TabIndex        =   10
         Top             =   1215
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   7950
         TabIndex        =   12
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   8400
         TabIndex        =   14
         Top             =   585
         Width           =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   10425
         TabIndex        =   16
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   8025
         TabIndex        =   18
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forward To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   7845
         TabIndex        =   39
         Top             =   1245
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   90
      TabIndex        =   57
      Top             =   5880
      Width           =   330
   End
End
Attribute VB_Name = "frmSearchReceiptVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'* Application ID           : 115                                                        *
'* Application Name         : Saankhya Double Entry                                      *
'* Screen id                : Receipts                                                   *
'* Version No               : Ver 1.0.0                                                  *
'* Form Designed By         : Aswathi                                                    *
'* Created on               :                                                            *
'* Coding By                : Aiby                                                       *
'* Coded on                 : 25-Dec-2007                                                *
'* Reviewed By              :                                                            *
'* Reviewed on              :                                                            *
'* Purpose                  : Receipt Vouchers for demand based transactions             *
'*                                                                                       *
'*                                                                                       *
'* Name of Database         : DB_Finance                                                 *
'* Name of Table(s)         : faTransactions, faTransactionChild                         *
'* Look up Table(s)         : faTransactionType, faTransactionChild, faAccountHeads      *
'* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *
'*                                                                                       *
'*                                                                                       *
'*=======================================================================================*
'* | Number  | Modification Date |   Modified By         |   Name of function/Variable   *
'* |---------|-------------------|-----------------------|-------------------------------*
'* |         |                   |                       |                               *
'* |         |                   |                       |                               *
'*=======================================================================================*
' Notes :-
'       cmdSearchAccountHead.Tag Keeps GroupID of AccountHead Type
'       GroupID=1-> Cash   GroupID=2 ->  Bank
'----------------------------------------------------------------------------------------'
    
    Dim mVoucherID As Variant
    Dim mDefaultTransactionTypeID   As Long
    Dim mDefaultAccountHeadCode     As String
    Dim mDefaultInstrumentID        As Long
    Dim mDefaultBankID              As Long
    Dim mDefaultBankHeadCode        As String
    Dim mDefaultZoneID              As Double
    Dim mBuildingID                 As Double
    Dim mSubLedgerID                As Variant
    Dim mUserSessions               As Integer
    Dim mTransactionType            As Long
    Dim mDrAccountHeadID            As Long
    
    Dim mFineRate                   As Single
    Dim mAcHeadCodePTaxArrear       As String
    Dim mAcHeadCodeFine             As String
    Dim mAcHeadCodeRoundOff         As String
    Dim mAcHeadCodeAdvance          As String
    Dim mNumberOfSelections         As Integer
    '---------------------------------------------------------'
    ' For Address details to Save faVoucherAddress            '
    '---------------------------------------------------------'
    Dim intWardNo           As Variant
    Dim intDoorNo           As Variant
    Dim vchDoorNo2          As Variant
    Dim vchName           As Variant
    Dim vchInit1            As Variant
    Dim vchInit2            As Variant
    Dim vchInit3            As Variant
    Dim vchInit4            As Variant
    
    Dim vchHouseName        As Variant
    Dim vchStreetName       As Variant
    Dim vchLocalPlace       As Variant
    Dim vchMainPlace        As Variant
    Dim vchPostOffice       As Variant
    Dim vchDistrict         As Variant
    Dim vchPinNumber        As Variant
    Dim vchPhone            As Variant
    Private mvarSubLedgerID As Variant
    
    Dim mStartingReceiptNo  As Variant       ' Keeps value on every session
    Dim mGrandTotal         As Variant
    Dim mSkipFlag           As Boolean      ' To control AutoFill Text Behaviour or TransactionType Text Box
    Dim mKeyCode            As Long
    Dim mBkSpaceFlag        As Boolean
        
    Private Sub FormatGrid()
        Dim mWidth As Long
        mWidth = vsGrid.Width - 450
        vsGrid.Cols = 5
        
        vsGrid.ColWidth(0) = mWidth * 8 / 100
        vsGrid.ColWidth(1) = mWidth * 15 / 100
        vsGrid.ColWidth(2) = mWidth * 30 / 100
        vsGrid.ColWidth(3) = mWidth * 15 / 100
        vsGrid.ColWidth(4) = mWidth * 32 / 100
        
        vsGrid.TextMatrix(0, 0) = "Date"
        vsGrid.TextMatrix(0, 1) = "Receipt No"
        vsGrid.TextMatrix(0, 2) = "Transaction Type"
        vsGrid.TextMatrix(0, 3) = "Amount"
        vsGrid.TextMatrix(0, 4) = "Name"
        
    End Sub
    
    Private Sub FillReceipts()
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mRow As Long
        Dim mWhere As String
        If Not IsDate(txtDate.Text) Then
            MsgBox "Enter a valid Date", vbInformation
            Exit Sub
        End If
        
        mWhere = " Where dtDate Between '" & txtDate.Text & "' AND '" & txtToDate.Text & "'"
        If Val(txtTransactiontype.Tag) > 0 Then
            mWhere = mWhere + " AND faVouchers.intTransactionTypeID = " & Val(txtTransactiontype.Tag)
        End If
        mSql = "        Select dtDate, intVoucherNo, vchTransactionType, fltAmount, vchName + ' ' + Isnull(vchInit1,'') + ' ' + Isnull(vchInit2,'') vchName, "
        mSql = mSql + " faCounters.intCounterNo From faVouchers Inner Join "
        mSql = mSql + " faTransactionType On faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID Left Join "
        mSql = mSql + " faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID Left Join "
        mSql = mSql + " faCounters On faCounters.intCounterID = faVouchers.intCounterID "
        mSql = mSql + mWhere + " Order By faVouchers.intCounterID, intVoucherNo "
        
        objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        vsGrid.Clear 1, 0
        If Not (Rec.BOF And Rec.EOF) Then
            'vsGrid.Clip = Rec.GetRows
            While Not Rec.EOF
                mRow = mRow + 1
                If mRow > vsGrid.Rows - 1 Then
                    vsGrid.Rows = vsGrid.Rows + 50
                End If

                vsGrid.TextMatrix(mRow, 0) = Rec!dtDate
                vsGrid.TextMatrix(mRow, 1) = Rec!intVoucherNo
                vsGrid.TextMatrix(mRow, 2) = Rec!vchTransactionType
                vsGrid.TextMatrix(mRow, 3) = Rec!fltAmount
                vsGrid.TextMatrix(mRow, 4) = Rec!vchName
                Rec.MoveNext
            Wend
        End If
    End Sub
    Private Sub ClearAddressVariables()
        intWardNo = Null
        vchName = ""
        vchHouseName = ""
        vchStreetName = ""
        vchMainPlace = ""
        vchPostOffice = ""
        vchDistrict = ""
        vchPinNumber = ""
        
        txtWardNo.Text = ""
        txtDoorNo1.Text = ""
        txtDoorNo2.Text = ""
        txtName.Text = ""
        txtInit1.Text = ""
        txtInit2.Text = ""
        txtInit3.Text = ""
        txtInit4.Text = ""
        txtHouse.Text = ""
        txtStreet.Text = ""
        txtLocalPlace.Text = ""
        txtMainPlace.Text = ""
        txtPost.Text = ""
        txtPin.Text = ""
        txtPhone.Text = ""
        
    End Sub
    
    Private Sub SaveAddressInVariables()
        If Trim(txtName.Text) <> "" Then
            vchName = Trim(txtName.Text)
            If Trim(txtInit1) <> "" Then vchName = vchName + "." + Trim(txtInit1)
            If Trim(txtInit2) <> "" Then vchName = vchName + "." + Trim(txtInit2)
            If Trim(txtInit3) <> "" Then vchName = vchName + "." + Trim(txtInit3)
            If Trim(txtInit4) <> "" Then vchName = vchName + "." + Trim(txtInit4)
        End If
        vchHouseName = Trim(txtHouse)
        vchStreetName = Trim(txtStreet)
        vchMainPlace = Trim(txtMainPlace)
        vchPostOffice = Trim(txtPost)
        'vchDistrict = Trim(txtDistrict)
        vchPinNumber = Trim(txtPin)
    End Sub
    Private Sub LoadAddressVariable()
        Dim mStr As String
        mStr = vchName
        mStr = Token(mStr, ".")
        
        txtName.Text = vchName
        txtInit1.Text = ""
        txtInit2.Text = ""
        txtInit3.Text = ""
        txtInit4.Text = ""
        txtHouse.Text = vchHouseName
        txtStreet.Text = vchStreetName
        txtMainPlace.Text = vchMainPlace
        txtPost.Text = vchPostOffice
        'txtDistrict.Text = vchDistrict
        txtPin.Text = vchPinNumber
    End Sub
    Private Sub ShowAddressInParty()
        'txtPayee.Text = vchName
        'txtHouse.Text = vchHouseName
        'txtAddress.Text = vchStreetName & Chr(13)
        'txtAddress.Text = txtAddress.Text & vchMainPlace & Chr(13)
        'txtAddress.Text = txtAddress.Text & vchPostOffice & Chr(13)
        'txtAddress.Text = txtAddress.Text & vchDistrict & " - " & vchPinNumber
    End Sub
    Private Sub ShowFrames(mIndex As Long)
        Select Case mIndex
            Case 2
                fraDemandDetails.Visible = False
                'fraSubLedger.Visible = True
                'fraParty.Visible = False
            Case 3
                fraDemandDetails.Visible = False
               ' fraSubLedger.Visible = False
                'fraParty.Visible = True
            Case Else
                fraDemandDetails.Visible = True
                'fraSubLedger.Visible = False
                'fraParty.Visible = False
        End Select
    End Sub
    Private Sub PrintReceipt(intVoucherID As Double)
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        Open gbFileName For Output As #gbFileNO
        'FileInitialize
        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
        objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(120); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(31); IIf(IsNull(Rec!dtDate), "", Rec!dtDate); Tab(65); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(120); IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            ' Line 15 Next
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(55); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            Print #gbFileNO,
            Print #gbFileNO,
            
            
            ' Line 18 Next
            
            Rec.MoveFirst
            While Not Rec.EOF
                mLoop = mLoop + 1
                
                '==================================================================='
                ' Counter Foil
                '==================================================================='
                Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(12); mstrYear;
                    
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(27); PadL(Format(Rec!fltAmount, "0.00"), 9);
                Else
                    Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                End If
                
                
                '==================================================================='
                ' Receipt Area
                '==================================================================='
                Print #gbFileNO, Tab(48); PadL(CStr(mLoop), 2);
                Print #gbFileNO, Tab(56); PadR(Rec!vchAlias, 41);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(98); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(98); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(98); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(98); mstrYear;
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(109); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Else
                    Print #gbFileNO, Tab(126); PadL(Format(Rec!fltAmount, "0.00"), 9)
                End If
                'Print #gbFileNO, Tab(26); PadL(Trim(str(mLoop)), 3); Tab(31); Rec!vchAccountHeadCode; Tab(40); PadR(IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias), 20); Rec!tnyPeriodID; Tab(70); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Rec.MoveNext
            Wend
            Rec.MoveFirst
            
            For mCount = mLoop + 1 To 10
                Print #gbFileNO,
            Next mCount
            Print #gbFileNO,
                            
            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(117); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            
            Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
            Print #gbFileNO,
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                Print #gbFileNO, Tab(11); objCounter.CounterNo;
                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
            End If
            objUser.SetUser (Rec!intUserID)
            If objUser.UserID > -1 Then
                Print #gbFileNO, Tab(11); objUser.UserName;
                Print #gbFileNO, Tab(61); objUser.UserName
            End If
        End If
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        
        Close #gbFileNO
        ShellPad
        'Shell "Print " & gbFileName
        'Kill gbFileName
    End Sub
    Private Sub StoreAddress()
        vchName = Trim(txtName.Text)
        If Trim(txtInit1) <> "" Then vchName = vchName & "." & Trim(txtInit1)
        If Trim(txtInit2) <> "" Then vchName = vchName & "." & Trim(txtInit2)
        If Trim(txtInit3) <> "" Then vchName = vchName & "." & Trim(txtInit3)
        If Trim(txtInit4) <> "" Then vchName = vchName & "." & Trim(txtInit4)
        
        vchHouseName = Trim(txtHouse)
        vchStreetName = Trim(txtStreet)
        vchMainPlace = Trim(txtMainPlace)
        vchPostOffice = Trim(txtPost)
        'vchDistrict = Trim(txtDistrict)
        vchPinNumber = Trim(txtPin)
    
    End Sub
    
    Private Sub ValuesForHiddenColumns()
        On Error Resume Next
        If vsGrid.Row = 0 Then Exit Sub
        If vsGrid.TextMatrix(vsGrid.Row, 2) = "" Then  ' Year
            vsGrid.TextMatrix(vsGrid.Row, 7) = gbFinancialYearID
        Else
            vsGrid.TextMatrix(vsGrid.Row, 7) = vsGrid.TextMatrix(vsGrid.Row, 2)
        End If
        
        If vsGrid.TextMatrix(vsGrid.Row, 3) = "" Then  'Period
            vsGrid.TextMatrix(vsGrid.Row, 8) = 3
        Else
            vsGrid.TextMatrix(vsGrid.Row, 8) = vsGrid.TextMatrix(vsGrid.Row, 3)
        End If
        
        If vsGrid.TextMatrix(vsGrid.Row, 7) < gbFinancialYearID Then  ' Arrear Flag
            vsGrid.TextMatrix(vsGrid.Row, 9) = 1
        Else
            vsGrid.TextMatrix(vsGrid.Row, 9) = 0
        End If
        If Val(vsGrid.TextMatrix(vsGrid.Row, 4)) > 0 Then   'Arrear Amount
            vsGrid.TextMatrix(vsGrid.Row, 11) = Val(vsGrid.TextMatrix(vsGrid.Row, 4))
        Else                                          'Current Amount
            vsGrid.TextMatrix(vsGrid.Row, 11) = Val(vsGrid.TextMatrix(vsGrid.Row, 5))
        End If
    End Sub
    
    Private Sub LockForm(mLockFlag As Boolean)
        'cmdSave.Enabled = mLockFlag
    
        fraReceiptNo.Enabled = mLockFlag
        fraAccountHead.Enabled = mLockFlag
        
    End Sub
    
    Private Sub FillGridYear()
        Dim mLoop As Integer
        Dim mItem As String
        mItem = ""
        'For mLoop = 1991 To gbFinancialYearID
        For mLoop = gbFinancialYearID + 1 To 1991 Step -1
            mItem = mItem & "|#" & mLoop & ";" & CStr(mLoop) & "-" & CStr(mLoop + 1)
        Next
        vsGrid.ColComboList(2) = mItem
        
        mItem = "#0; "
        mItem = mItem & "|#" & 1 & "; First Half"
        mItem = mItem & "|#" & 2 & "; Second Half"
        mItem = mItem & "|#" & 3 & "; Full Year"
        vsGrid.ColComboList(3) = mItem
    End Sub
    
    Private Sub FillAccountHeads()
        'Call gFillVSGrid(vsGrid, 1, "spGetAccHead4Receipts", enuSourceString.Saankhya)
    End Sub
    
    Private Function CalculatePTaxFine(numBuildingID As Double, mYearID As Long, mPeriodID As Long) As Double
            '        'CalculatePTaxFine(numBuildingID As Double, numDemandID As Double) As Double
            '        Dim objDB As New clsDB
            '        Dim mCnn As New ADODB.Connection
            '        Dim RecIDemand As New ADODB.Recordset
            '        Dim RecAdv As New ADODB.Recordset
            '        Dim mSQL As String
            '
            '        Dim mAdvAmt As Double
            '        Dim mFineAmt As Double
            '        Dim mTotalFine As Double
            '        Dim mPTAmt As Double
            '        Dim mPTRate As Single
            '        Dim mFromDate As Date
            '        Dim mToDate As Date
            '        Dim mNote As String
            '
            '        mAdvAmt = 0
            '        mFineAmt = 0
            '        mTotalFine = 0
            '        mPTAmt = 0
            '        mPTRate = 1
            '        objDB.SetConnection mCnn
            '
            '        mSQL = ""
            '        mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount, numSubLedgerID"
            '        mSQL = mSQL + " From faIDemandChild Inner Join"
            '        mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID"
            '        mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
            '        mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxArrearHeadCode & "'"
            '        mSQL = mSQL + " And faIDemandTbl.numSubLedgerID = " & numBuildingID
            '        mSQL = mSQL + " And ( faIDemandTbl.intYearID < " & mYearID
            '        mSQL = mSQL + " Or ( faIDemandTbl.intYearID = " & mYearID & " AND faIDemandTbl.tnyPeriodID = " & mPeriodID & " ) )"
            '
            '        'mSQL = mSQL + " And faIDemandTbl.numDemandID <= " & numDemandID
            '        RecIDemand.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
            '        If Not (RecIDemand.EOF And RecIDemand.BOF) Then
            '            mSQL = ""
            '            mSQL = mSQL + " Select faIDemandChild.numDemandID, faIDemandChild.dtOnDate, faIDemandChild.fltAmount"
            '            mSQL = mSQL + " From faIDemandChild Inner Join"
            '            mSQL = mSQL + " faIDemandTbl On faIDemandTbl.numDemandID = faIDemandChild.numDemandID "
            '            mSQL = mSQL + " Where faIDemandTbl.tnyStatus = 0 And faIDemandTbl.intTransactionTypeID = " & mPTaxTransactionTypeID
            '            mSQL = mSQL + " And faIDemandChild.vchAccountHeadCode = '" & mPTaxAdvanceCollected & "' And faIDemandTbl.numSubLedgerID = " & RecIDemand!numSubLedgerID
            '            RecAdv.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
            '        Else
            '            CalculatePTaxFine = 0
            '            Return
            '        End If
            '        While Not RecIDemand.EOF
            '            mPTAmt = RecIDemand!fltAmount
            '            mFromDate = RecIDemand!dtOnDate
            '            '->
            '            mNote = mNote + DdMmmYy(mFromDate) + "  PTax : " + Format(mPTAmt, "0.00") + vbCrLf
            '            While Not RecAdv.EOF
            '                If mAdvAmt <= 0 Then
            '                    mAdvAmt = RecAdv!fltAmount
            '                    mToDate = RecAdv!dtOnDate
            '                    '->
            '                    mNote = mNote + DdMmmYy(mToDate) + "   Adv : " + Format(mPTAmt, "0.00") + vbCrLf
            '                    GoTo CalculatFine:
            '                Else
            ' CalculatFine:
            '                    mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
            '                    '->
            '                    mNote = mNote + Str(mFineAmt) & DdMmmYy(mFromDate) & "  " & DdMmmYy(mToDate) & Str(mPTAmt) & Str(mPTRate)
            '                    mTotalFine = mTotalFine + mFineAmt
            '                    If mAdvAmt >= mFineAmt Then
            '                        mAdvAmt = mAdvAmt - mFineAmt
            '                        mFineAmt = 0
            '                    Else
            '                        mFineAmt = mFineAmt - mAdvAmt
            '                        mAdvAmt = 0
            '                    End If
            '                    If mAdvAmt >= mPTAmt Then
            '                        mAdvAmt = mAdvAmt - mPTAmt
            '                        mPTAmt = 0
            '                    Else
            '                        mPTAmt = mPTAmt - mAdvAmt
            '                        mAdvAmt = 0
            '                    End If
            '                    If mAdvAmt > 0 Then
            '                        GoTo ReadNextDemand:
            '                    End If
            '                    If mPTAmt > 0 Then
            '                        mFromDate = mToDate
            '                    End If
            '                    RecAdv.MoveNext
            '                End If
            '            Wend
            '            If mPTAmt > 0 Then
            '                mToDate = gbTransactionDate
            '                mFineAmt = CalculateFine(mFromDate, mToDate, mPTAmt, mPTRate)
            '                mTotalFine = mTotalFine + mFineAmt
            '            End If
            '
            ' ReadNextDemand:
            '            RecIDemand.MoveNext
            '        Wend
            '        RecIDemand.Close
            '        Set RecIDemand = Nothing
            '        CalculatePTaxFine = mTotalFine
    End Function
    
    Private Sub SelectDemands()
        Dim mLoop As Long
        Dim mDemandID As Double
        Dim mAmount As Double
        For mLoop = 1 To vsGrid.Rows - 1
            If mDemandID = Val(vsGrid.TextMatrix(mLoop, 10)) Then
                mAmount = mAmount + Val(vsGrid.TextMatrix(mLoop, 11))
            Else
                mDemandID = Val(vsGrid.TextMatrix(mLoop, 10))
            End If
            'vsGrid.Cell(flexcpChecked, mLoop, 12) = 2 'vbUnchecked
        Next mLoop
    End Sub
    
    Public Sub DisplayBuildingDetails()
        'Dim arrInput As Variant
        'Dim Rec As New ADODB.Recordset
        'Dim objDb As New clsDB
        'Dim mCnn As New ADODB.Connection
        '
        ''arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
        '            Val(txtWard.Tag), _
        '            Val(txtHouseNo1), _
        '            Trim(txtHouseNo2))
        '
        'arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
        '            Val(txtWard.Tag), _
        '            Val(txtHouseNo1), _
        '            Trim(txtHouseNo2))
        '
        ' If objDb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
        '    'Set Rec = objDB.ExecuteSP("spGetBuildingDetails", arrInput, , , mCnn, adCmdStoredProc)
        '    Set Rec = objDb.ExecuteSP("spSanGetSearchBuildingList", arrInput, , , mCnn, adCmdStoredProc)
        '    If Not (Rec.BOF And Rec.EOF) Then
        '        mBuildingID = Rec!numBuildingID
        '        txtBuildingNo.Text = Rec!numBuildingID
        '        txtWard.Text = Rec!intWardNo
        '        txtHouseNo1.Text = Rec!intDoorNo1
        '        txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
        '
        '        vchName = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
        '        'vchName = vchName & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
        '        'vchName = vchName & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
        '        'vchName = vchName & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
        '        'vchName = vchName & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
        '        vchHouseName = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
        '        'vchStreetName = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
        '        'vchMainPlace = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
        '        vchMainPlace = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
        '        'vchPostOffice = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
        '        'vchDistrict = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
        '        'vchPinNumber = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
        '
        '        txtAddress.Text = vchName
        '        txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName
        '        txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName
        '        txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace), vchMainPlace & ", ", "")
        '        txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice
        '        txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict
        '        txtAddress.Text = txtAddress.Text & " - " & vchPinNumber
        '    Else
        '        mBuildingID = -1
        '    End If
        'End If
    End Sub
    
    Private Sub DisplayBuildingTaxDemands(mBuildingID As Double)
        
    End Sub
    Private Sub ListPostingHeadsInGridForGeneralReceipts()
        
    End Sub
    
    Private Sub ListPostingHeadsInGrid(mTransactionType As Long, Optional mGroupID As Variant = Null)
               
    End Sub
    
    Public Sub Calculate()
'
    End Sub
    
    Private Sub ListMasters(mMasterType As Integer)
        Dim mSql As String
        lstMasters.Tag = mMasterType
        Select Case mMasterType
            Case Is = 1 ' Transaction Type
                mSql = "Select vchTransactionType, intTransactionTypeID, intGroupID From faTransactionType Where intGroupID = 10 Order By vchTransactionType"
                Call PopulateList(lstMasters, mSql, "Property Tax", False, , True)
                lstMasters.Top = 1000
                lstMasters.Left = 250
                lstMasters.Height = 2000
                lstMasters.Visible = True
                lstMasters.Width = txtTransactiontype.Width + 1500
                lstMasters.SetFocus
            Case Is = 2 ' Instruments
                mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes"
                Call PopulateList(lstMasters, mSql, "", , , True)
                lstMasters.Left = 8340
                lstMasters.Top = 450
                lstMasters.Height = 2000
                lstMasters.Width = txtInstrument.Width
                lstMasters.Visible = True
                lstMasters.SetFocus
            Case Is = 3 ' Account Heads
                If mMasterType Then
                    mSql = "Select vchBankName, intBankID From faBanks Order By vchBankName"
                    Call PopulateList(lstMasters, mSql, "", , , True)
                    lstMasters.Left = 8340
                    lstMasters.Top = txtAccountHead.Top - 25
                    lstMasters.Height = 2000
                    lstMasters.Visible = True
                    lstMasters.SetFocus
                Else
                    mSql = "Select vchBankName, intBankID From faBanks Order By vchBankName"
                    Call PopulateList(lstMasters, mSql, "", , , True)
                    lstMasters.Left = 8340
                    lstMasters.Top = 1380
                End If
        End Select
        '           Newly Added         '
        'txtInstNo.Text = ""
        txtBank.Text = ""
        'txtDated.Text = ""
        txtPlace.Text = ""
    End Sub
    
    Private Sub FormInitialize()
        Dim Rec As New ADODB.Recordset
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mOutput As Variant
        Dim mStr As String
        
        'fraSubLedger.Visible = True
        Call ShowFrames(1)
        Call LockForm(True)
        Call ClearAddressVariables
        
        mUserSessions = -1
        mBuildingID = -1
        objDb.SetConnection mCnn
        
        '--------------------------------------------------------------------------------------------'
        ' Blocked Due to Error
        '--------------------------------------------------------------------------------------------'
        '        Set Rec = GetRecordSet("spGetNewReceiptNoAndBookNo " & gbFinancialYearID & ", 10 ", adOpenStatic, adLockOptimistic, mCnn)
        '        If Not (Rec.EOF Or Rec.BOF) Then
        '            txtReceiptNo.Text = Format(Rec!intVoucherNo, "0000")
        '            txtBookNo.Text = Format(Rec!intBookNo, "0000")
        '        End If
        '--------------------------------------------------------------------------------------------'
        
        mVoucherID = Null
        txtDate.Text = DdMmmYy(gbDate)
        txtToDate.Text = DdMmmYy(gbDate)
        txtTransactiontype.Text = ""
        txtTransactiontype.Tag = ""
       
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        txtInstrument.Text = ""
        txtInstrument.Tag = ""
       
        txtBank.Text = ""
        txtPlace.Text = ""
        txtRefNo.Text = ""
        txtRefNo.Tag = ""
       
        vchName = ""
        vchHouseName = ""
        vchStreetName = ""
        vchMainPlace = ""
        vchPostOffice = ""
        vchDistrict = ""
        vchPinNumber = ""
        vsGrid.Rows = 1
        vsGrid.Rows = 15
        Call SetDefaultSettings
        
        
    End Sub
    
    Private Sub ClearAddress()
        
    End Sub
    Private Sub FillZone()
        'Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Order By chvZoneNameEnglish", , True, True, True, DBMaster)
    End Sub
    Private Sub FillCounterSeats()
        Call PopulateList(cmbCounterSeats, "Select chvSeatTitle, Right(numSeatID, Len(numSeatID)-5) From GL_Seats Where intGroupID > 98 Order By chvSeatTitle", gbSeatName, True, , True, DBMaster)
    End Sub
    Private Sub FillTransactionTypes()
        Dim mSql As String
        mSql = "Select vchTransactionType, intTransactionTypeID, intGroupID From faTransactionType Where intGroupID = 10 Order By vchTransactionType"
        Call PopulateList(lstTransactionType, mSql, , True, True, True)
    End Sub

    Private Sub SetDefaultSettings()
        Dim objTranType As New clsTransactionType
        Dim objAc As New clsAccounts
        Dim objInstruments As New clsInstruments
        Dim objBank As New clsBank
        Dim mLoopCount As Long
        
        mFineRate = 1  ' Fine Rate = 1.%'
        mAcHeadCodePTaxArrear = "431100200"
        mAcHeadCodeFine = "140200000"
        mAcHeadCodeRoundOff = "00000"
        
        mDefaultTransactionTypeID = gbDefaultTransactionTypeID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultTransactionTypeID"))
        mDefaultAccountHeadCode = gbAcHeadCodeCash 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultAccountHeadCode")
        mDefaultInstrumentID = gbInstrumentCash 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultInstumentID"))
        mDefaultBankID = gbDefaultBankID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID"))
        mDefaultZoneID = gbnumZonalID 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultZone"))
        
        If mDefaultZoneID > 0 Then
            'For mLoopCount = 0 To cmbZone.ListCount - 1
            '    If cmbZone.ItemData(mLoopCount) = mDefaultZoneID Then
            '        cmbZone.ListIndex = mLoopCount
            '        Exit For
            '    End If
            'Next
        End If
        
        objTranType.SetTransactionType (mDefaultTransactionTypeID)
        objAc.SetAccountCode (mDefaultAccountHeadCode)
        objInstruments.SetInstrumentType (mDefaultInstrumentID)
        objBank.SetBankInfo (mDefaultBankID)
        
        'txtTransactionType.Text = objTranType.TransactionType
        'txtTransactionType.Tag = objTranType.TransactionTypeID
        'Call txtTransactionType_DblClick
        txtAccountHead.Text = objAc.AccountHead & " [ " & objAc.AccountCode & " ]"
        txtAccountHead.Tag = objAc.AccountHeadID
        cmdSearchAccountHead.Tag = objAc.GroupID    ' Cash Type
        If Not IsNull(objInstruments.InstrumentTypeID) Then
            txtInstrument.Text = objInstruments.InstrumentType
            txtInstrument.Tag = objInstruments.InstrumentTypeID
        End If
        mDefaultBankHeadCode = objBank.BankAccountHeadCode
        mAcHeadCodeAdvance = 350410101
    End Sub
    
    Private Sub cmbZone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
        

    Private Sub cmdCancel_Click()
    '''Select Case mUserSessions
    '''    Case Is = 1
    '''        Call FormInitialize
    '''    Case Is = -1
    '''        Unload Me
    '''End Select
    Unload Me
    
    End Sub
    

    Private Sub cmdFind_Click()
        Call DisplayBuildingDetails
        Call DisplayBuildingTaxDemands(mBuildingID)
    End Sub
    
    Private Sub cmdNew_Click()
        Call FormInitialize
    End Sub
    
    Private Sub cmdClear_Click()
        Call FormInitialize
    End Sub
    Private Sub cmdSearch_Click()
        Call FillReceipts
    End Sub

    Private Sub cmdSearchAccountHead_Click()
            Dim mSql As String
            If Val(txtInstrument.Tag) > 0 Then
                Select Case Val(txtInstrument.Tag)
                    Case 5, 6, 8 '[Cheque]
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
                    Case 7  '[Treasury Bills]
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE  tinHiddenFlag = 0 AND faAccountHeads.vchAccountHeadCode Like '45045%'"
                    Case Else
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE  tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faCash
                End Select
                frmSearchAccountHeads.SQLString = mSql
                frmSearchAccountHeads.Show vbModal
                txtAccountHead.SetFocus
            End If
    End Sub
    Private Sub cmdSearchDemandNo_Click()
        
    End Sub
    Private Sub cmdSearchInstrument_Click()
        Call ListMasters(2)
    End Sub
    Private Sub cmdSearchTransactionType_Click()
        Call ListMasters(1)
    End Sub
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        txtDate.Text = DdMmmYy(gbTransactionDate)
        Call Calculate
    End Sub
    Private Sub Form_GotFocus()
        txtDate.Text = DdMmmYy(gbTransactionDate)
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 And Shift = 2 Then
            Call MsgBox("Search!", vbInformation)
        End If
        If KeyCode = vbKeyF8 Then
            frmPropertyTaxCalculator.Show vbModal
        End If
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FormatGrid
        Call FillZone
        Call FillCounterSeats
        Call FormInitialize
        Call FillTransactionTypes
        
        vsGrid.ColComboList(0) = "|..."
        cmbDZone.AddItem "Main Office"
        cmbDZone.ItemData(cmbDZone.NewIndex) = 1
        cmbDZone.ListIndex = 0
        
        Call FillReceipts
        
    End Sub
    
    Private Sub lstMasters_DblClick()
        Select Case Val(lstMasters.Tag)
            Case 1
                txtTransactiontype.Text = lstMasters.Text
                txtTransactiontype.Tag = lstMasters.ItemData(lstMasters.ListIndex)
                Call txtTransactionType_KeyPress(13)
            Case 2
                txtInstrument.Text = lstMasters.Text
                txtInstrument.Tag = lstMasters.ItemData(lstMasters.ListIndex)
                Call txtInstrument_LostFocus
            Case 3
                Dim objBank As New clsBank
                objBank.SetBankInfo (lstMasters.ItemData(lstMasters.ListIndex))
                If objBank.BankID > 0 Then
                    txtAccountHead.Text = objBank.BankName & " [ " & objBank.BankAccountHeadCode & " ]"
                    txtAccountHead.Tag = objBank.BankAccountHeadID
                End If
        End Select
        
        If lstMasters.Text = "Property Tax" Then
            'txtDemandNo.Visible = False
            'txtDemandPrefix.Visible = False
            'cmdSearchDemandNo.Visible = False
        Else
            'txtDemandNo.Visible = True
            'txtDemandPrefix.Visible = True
            'cmdSearchDemandNo.Visible = True
        End If
        lstMasters.Tag = ""
        lstMasters.Visible = False
    End Sub
    Private Sub lstMasters_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            lstMasters_DblClick
            KeyAscii = 0
        End If
    End Sub
    Private Sub lstMasters_LostFocus()
        lstMasters.Visible = False
    End Sub
    Private Sub txtAccountHead_GotFocus()
        If gbSearchID > 0 Then
            Dim objBank As New clsBank
            objBank.SetBankInfoByAccID (gbSearchID)
            gbSearchID = -1
            gbSearchStr = ""
            'objBank.SetBankInfo (lstMasters.ItemData(lstMasters.ListIndex))
            If objBank.BankID > 0 Then
                txtAccountHead.Text = objBank.BankName & " [ " & objBank.BankAccountHeadCode & " ]"
                txtAccountHead.Tag = objBank.BankAccountHeadID
            End If
        End If
    End Sub
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtAddress_DblClick()
        
    End Sub

    Private Sub txtBank_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        Call PressTabKey
        End If
    End Sub

Private Sub txtBank_LostFocus()
    txtBank.Text = FormatIntoProperCase(txtBank)
End Sub

    Private Sub txtBuildingNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtDated_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtDated_LostFocus()
        
    End Sub
    Private Sub txtDistrict_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtDemandNo_LostFocus()
        Call DisplayDemandDetails
    End Sub
    
    Private Sub txtDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtDate_LostFocus()
        txtDate.Text = CheckDateInMMM(txtDate.Text)
    End Sub
    
    Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDoorNo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
            Exit Sub
        End If
        If KeyAscii = Asc(" ") Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDoorNo2_LostFocus()
        txtDoorNo2.Text = UCase(txtDoorNo2.Text)
    End Sub
    
    Private Sub txtHouse_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
        
    Private Sub txtHouse_LostFocus()
        txtHouse.Text = FormatIntoProperCase(txtHouse)
    End Sub

    Private Sub txtHouseNo1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtHouseNo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtHouseNo3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtInit1_LostFocus()
        txtInit1.Text = FormatIntoProperCase(txtInit1)
    End Sub

    Private Sub txtInit2_LostFocus()
        txtInit2.Text = FormatIntoProperCase(txtInit2)
    End Sub
        
    Private Sub txtInit3_LostFocus()
        txtInit3.Text = FormatIntoProperCase(txtInit3)
    End Sub
    
    Private Sub txtInit4_LostFocus()
        txtInit4.Text = FormatIntoProperCase(txtInit4)
    End Sub
    
    Private Sub txtInstNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtInstrument_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtInstrument_LostFocus()
        Dim objBk As New clsBank
        '------------------------------------------------------------------------------'
        ' When Instrument Cheque is selected then it Set's the Default Bank            '
        '------------------------------------------------------------------------------'
        If Val(txtInstrument.Tag) = gbInstrumentCheque Or Val(txtInstrument.Tag) = 8 Then '8=Bank Pay_in_Slip
            If cmdSearchAccountHead.Tag = 2 Then
                If Val(txtAccountHead.Tag) > 0 Then
                    objBk.SetBankInfo (Val(txtAccountHead.Tag))
                    If objBk.BankID > 0 Then
                        txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                    Else
                        GoTo DefaultBank:
                    End If
                Else
DefaultBank:        objBk.SetBankInfo (mDefaultBankID)
                    txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                    txtAccountHead.Tag = objBk.BankAccountHeadID
                End If
            Else
                objBk.SetBankInfo (mDefaultBankID)
                txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                txtAccountHead.Tag = objBk.BankAccountHeadID
            End If
            cmdSearchAccountHead.Tag = 2
        Else
            If Val(cmdSearchAccountHead.Tag) <> 1 Then
                Dim objAcc As New clsAccounts
                objAcc.SetAccountCode (mDefaultAccountHeadCode)
                If objAcc.AccountHeadID > 0 Then
                    txtAccountHead.Text = objAcc.AccountHead & " [ " & objAcc.AccountCode & " ]"
                    txtAccountHead.Tag = objAcc.AccountHeadID
                Else
                    txtAccountHead.Text = ""
                    txtAccountHead.Tag = ""
                End If
                
                cmdSearchAccountHead.Tag = 1
            End If
        End If
    End Sub
    
    Private Sub txtLocalPlace_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtLocalPlace_LostFocus()
        txtLocalPlace.Text = FormatIntoProperCase(txtLocalPlace)
    End Sub
    Private Sub txtMainPlace_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtMainPlace_LostFocus()
        txtMainPlace.Text = FormatIntoProperCase(txtMainPlace)
    End Sub
    Private Sub txtName_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtInit1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtInit2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtInit3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtInit4_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtName_LostFocus()
        txtName.Text = FormatIntoProperCase(txtName)
    End Sub

    Private Sub txtPin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    
    Private Sub txtPlace_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        Call PressTabKey
        End If
    End Sub

    Private Sub txtPlace_LostFocus()
        txtPlace.Text = FormatIntoProperCase(txtPlace)
    End Sub

    Private Sub txtPost_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
        
    Private Sub txtPost_LostFocus()
        txtPost.Text = FormatIntoProperCase(txtPost)
    End Sub

Private Sub txtRefNo_LostFocus()
    txtRefNo.Text = UCase(txtRefNo.Text)
End Sub

    Private Sub txtStreet_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
        
    Private Sub txtStreet_LostFocus()
        txtStreet.Text = FormatIntoProperCase(txtStreet)
    End Sub

    Private Sub txtTotal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
        
    Private Sub txtToDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
        End If
    End Sub

    Private Sub txtTodate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate)
        Call FillReceipts
    End Sub
    
    Private Sub txtTransactionType_Change()
        Dim mIndex As Long
        Dim mStr As String
        If Not mSkipFlag Then
            If mKeyCode = 8 Or mKeyCode = 46 Or txtTransactiontype.Text = "" Then 'Tab or delete
                'Flcontrol = 1
                'FormResize
            End If
            If Not (mBkSpaceFlag Or mKeyCode = 40 Or mKeyCode = 38) Then
                
                With lstTransactionType
                    mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal txtTransactiontype.Text)
                    If mIndex >= 0 Then
                        .ListIndex = mIndex
                    End If
                End With
        
                If mIndex >= 0 Then
                    mStr = txtTransactiontype.Text
                    txtTransactiontype.Text = lstTransactionType.List(mIndex)
                    txtTransactiontype.SelStart = Len(mStr)
                    
                    If Len(txtTransactiontype.Text) - Len(mStr) > 0 Then
                        txtTransactiontype.SelLength = Len(txtTransactiontype.Text) - Len(mStr)
                    End If
                End If
            End If
        End If
        mSkipFlag = False
    
    End Sub
    
    Private Sub txtTransactionType_DblClick()
       Call txtTransactionType_KeyPress(13)
    End Sub
    
    Private Sub txtTransactionType_GotFocus()
        If Trim(txtTransactiontype.Text) = "" Then
            ListMasters (1)
            lstMasters.Refresh
        End If
    End Sub
        
    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then ' vbKeyDelete Then
            txtTransactiontype.Text = ""
            lstTransactionType.ListIndex = -1
        End If
        If KeyCode = 40 Then 'Down Arrow
            If lstTransactionType.ListIndex > -1 Then
                lstTransactionType.ListIndex = (lstTransactionType.ListIndex + 1) Mod lstTransactionType.ListCount
                txtTransactiontype.Text = lstTransactionType.Text
            End If
        ElseIf KeyCode = 38 Then 'Uparrow
            If lstTransactionType.ListIndex > -1 Then
                If lstTransactionType.ListIndex = 0 Then
                    lstTransactionType.ListIndex = lstTransactionType.ListCount - 1
                    txtTransactiontype.Text = lstTransactionType.Text
                Else
                    lstTransactionType.ListIndex = (lstTransactionType.ListIndex - 1) Mod lstTransactionType.ListCount
                    txtTransactiontype.Text = lstTransactionType.Text
                End If
            End If
        End If
        If KeyCode = vbKeyBack Then
            mBkSpaceFlag = True
        Else
            mBkSpaceFlag = False
        End If
    End Sub
    
    Private Sub txtTransactionType_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
        End If
    End Sub
    
    Private Sub txtTransactionType_LostFocus()
        
        Dim mIndex As Long
lblStart:
        With lstTransactionType
            mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal txtTransactiontype.Text)
            If mIndex >= 0 Then
                txtTransactiontype.Tag = lstTransactionType.ItemData(mIndex)
            Else
                txtTransactiontype.Text = ""
                txtTransactiontype.Tag = ""
                'GoTo lblStart:
            End If
        End With
        Select Case Val(txtTransactiontype.Tag)
            Case Is = gbTransactionTypePTax
                On Error Resume Next
                If gbLinkWithPropertyTax Then
                    frmPropertyTax.Show vbModal
                End If
            Case Is = 2
                'frmProfessionalTax.Visible = True
                'frmProfessionTaxSearch.ZOrder (0)
            Case Is = gbTransactionTypeRentOnBuilding
                'frmRentOnLandBuildings.Show vbModal
            Case Is = 9999
                Call FillAccountHeads
                Call FillGridYear
            Case Else
        End Select
        
    End Sub
    
    Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
            If vsGrid.Row > 1 Then
'                If vsGrid.TextMatrix(vsGrid.Row - 1, 0) = "" Or _
'                   (Val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 And _
'                   Val(vsGrid.TextMatrix(vsGrid.Row - 1, 5)) <= 0) Then
'                   Cancel = True
'                   Exit Sub
'                End If
            End If
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                    vsGrid.TextMatrix(Row, 1) = objAccHead.AccountHead
                    vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
                vsGrid.Col = vsGrid.Col + 2
                vsGrid.Redraw = flexRDDirect
                gbSearchStr = ""
            End If
    End Sub
    
    Private Sub vsGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
        If OldRow >= vsGrid.Rows - 1 Then
            vsGrid.Rows = vsGrid.Rows + 5
        End If
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim mSql As String
        
        If Val(txtTransactiontype.Tag) > 0 Then
            mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
            mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
            mSql = mSql + " Where intTransactionTypeID = " & Val(txtTransactiontype.Tag) & " Order By faTransactionTypeChild.intOrder"
            frmSearchAccountHeads.SQLString = mSql '"Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        Else
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        End If
        frmSearchAccountHeads.Show vbModal
    End Sub
    
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        'If vsGrid.Row > 0 Then
        If Row > 0 Then
            
'            If Col = 1 And vsGrid.ComboIndex > -1 Then
'                Set objAccHead = New clsAccounts
'                If objAccHead.FindAccountByHead(Trim(vsGrid.ComboItem)) Then
'                vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
'                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
'                End If
'            ElseIf vsGrid.Col = 4 Then
'                vsGrid.TextMatrix(Row, 4) = Format(Val(vsGrid.TextMatrix(Row, 4)), "#0")
'                If Val(vsGrid.TextMatrix(Row, 4)) > 0 Then
'                vsGrid.TextMatrix(Row, 5) = ""
'                End If
'                Call Calculate
'            ElseIf vsGrid.Col = 5 Then
'                vsGrid.TextMatrix(Row, 5) = Format(Val(vsGrid.TextMatrix(Row, 5)), "#0")
'                If Val(vsGrid.TextMatrix(Row, 5)) > 0 Then
'                vsGrid.TextMatrix(Row, 4) = ""
'                End If
'                Call Calculate
'            End If
            'Call ValuesForHiddenColumns
        End If
    End Sub
    
    Private Sub vsGrid_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
        If Col = 0 Then
            If KeyCode >= Asc("0") And KeyCode <= Asc("9") Or KeyCode = vbKeyBack Then
                
            ElseIf KeyCode = Asc(vbTab) Or KeyCode = 13 Then
                gbSearchStr = vsGrid.Cell(flexcpText, Row, Col)
                vsGrid.Cell(flexcpText, Row, Col) = ""
                Call vsGrid_BeforeEdit(Row, Col, False)
                
            Else
                KeyCode = 0
            End If
        End If
    End Sub
    
    Private Sub vsGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
        ''----------------------------------------------------------------------------'
        '' Selection and Deselection of Demands in Grid only permits for 3 demands    '
        '' 3 Demands = 6 Rows in Receipt in the case of Property Tax                  '
        '' Selection must be periodicity Order                                        '
        ''----------------------------------------------------------------------------'
        'Dim mLoop As Long
        'If Row > 0 Then
        '    If vsGrid.Cell(flexcpChecked, Row, Col) = 2 Then
        '        If mNumberOfSelections < 3 Then
        '            If Row = 1 Or vsGrid.Cell(flexcpChecked, Row - 1, Col) = vbChecked Then
        '                vsGrid.Cell(flexcpChecked, Row, Col) = vbChecked
        '                mNumberOfSelections = mNumberOfSelections + 1 'IIf(Row Mod 2 = 0, 1, 0)
        '            Else
        '                Cancel = True
        '            End If
        '        Else
        '            Cancel = True
        '        End If
        '    Else ' Already  Checked
        '        If vsGrid.Cell(flexcpChecked, Row - 1, Col) = 1 Then
        '        For mLoop = Row To vsGrid.Rows - 1
        '            If vsGrid.TextMatrix(Row, 10) <> vsGrid.TextMatrix(mLoop, 10) Then
        '                If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
        '                    Cancel = True
        '                End If
        '                mNumberOfSelections = mNumberOfSelections - 1
        '                Exit For
        '            End If
        '        Next mLoop
        '        Else
        '            Cancel = True
        '        End If
        '    End If
        'End If
    
    End Sub
    
    Public Sub DisplayDemandDetails()
        
    End Sub
    
    Public Property Let SubLedgerID(mSubLedgerID As Double)
        mvarSubLedgerID = mSubLedgerID
    End Property
    
    Public Property Get SubLedgerID() As Double
        SubLedgerID = mvarSubLedgerID
    End Property

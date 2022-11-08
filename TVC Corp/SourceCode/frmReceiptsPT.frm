VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReceiptsPT 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frmReceiptsPT.frx":0000
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ControlBox      =   0   'False
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
   Begin VB.TextBox txtAdvance 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8070
      TabIndex        =   57
      Top             =   4695
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.ListBox lstMasters 
      BackColor       =   &H00C0EBF0&
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
      Left            =   285
      TabIndex        =   36
      Top             =   705
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
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
      Height          =   1245
      Left            =   6330
      TabIndex        =   53
      Top             =   15
      Width           =   5355
      Begin VB.CommandButton cmdSearchInstrument 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   45
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   54
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
         Left            =   1020
         TabIndex        =   41
         Top             =   210
         Width           =   3870
      End
      Begin VB.TextBox txtInstNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1020
         TabIndex        =   47
         Top             =   810
         Width           =   2055
      End
      Begin VB.TextBox txtDated 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3645
         TabIndex        =   49
         Top             =   810
         Width           =   1245
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
         Left            =   1020
         TabIndex        =   44
         Top             =   510
         Width           =   3870
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
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inst. No"
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
         Index           =   9
         Left            =   315
         TabIndex        =   46
         Top             =   825
         Width           =   675
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dated"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   48
         Top             =   840
         Width           =   525
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
         Left            =   75
         TabIndex        =   43
         Top             =   540
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000018&
      Height          =   1230
      Left            =   3270
      TabIndex        =   37
      Top             =   30
      Width           =   3030
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   210
         Width           =   1665
      End
      Begin VB.TextBox txtReceiptNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   810
         Width           =   1665
      End
      Begin VB.TextBox txtBookNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No"
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
         Index           =   4
         Left            =   210
         TabIndex        =   52
         Top             =   825
         Width           =   960
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book No"
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
         Left            =   420
         TabIndex        =   51
         Top             =   540
         Width           =   750
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
         Left            =   765
         TabIndex        =   50
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8070
      TabIndex        =   33
      Top             =   4395
      Width           =   1710
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   8070
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   5025
      Width           =   3450
   End
   Begin VB.TextBox txtTotalCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9795
      TabIndex        =   31
      Top             =   4095
      Width           =   1725
   End
   Begin VB.TextBox txtTotalArrear 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8070
      TabIndex        =   30
      Top             =   4095
      Width           =   1710
   End
   Begin VB.Frame framParty 
      BackColor       =   &H80000018&
      Caption         =   "Address"
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
      Height          =   2580
      Left            =   90
      TabIndex        =   29
      Top             =   4080
      Width           =   6285
      Begin VB.TextBox txtHouseNo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   8
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtInitial4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5835
         MaxLength       =   1
         TabIndex        =   16
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtInitial3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5535
         MaxLength       =   1
         TabIndex        =   15
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtInitial2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5250
         MaxLength       =   1
         TabIndex        =   14
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtInitial1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4965
         MaxLength       =   1
         TabIndex        =   13
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtHouseName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1320
         Width           =   4995
      End
      Begin VB.TextBox txtHouseNo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3270
         MaxLength       =   4
         TabIndex        =   7
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtPayee 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   12
         Top             =   960
         Width           =   3765
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   1155
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1665
         Width           =   4995
      End
      Begin VB.TextBox txtWard 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         MaxLength       =   4
         TabIndex        =   5
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtBuildingNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtHouseNo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4890
         MaxLength       =   20
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.ComboBox cmbZone 
         Height          =   315
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2325
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "-->"
         Height          =   315
         Left            =   5730
         TabIndex        =   10
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   480
         TabIndex        =   19
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   645
         TabIndex        =   17
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   705
         TabIndex        =   11
         Top             =   1020
         Width           =   405
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   2670
         TabIndex        =   6
         Top             =   660
         Width           =   585
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " Ward"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   675
         TabIndex        =   4
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   3405
         TabIndex        =   2
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   315
         TabIndex        =   0
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   1245
      Left            =   75
      TabIndex        =   28
      Top             =   30
      Width           =   3165
      Begin VB.TextBox txtTransactionType 
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
         Height          =   285
         Left            =   150
         TabIndex        =   23
         Top             =   510
         Width           =   2535
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
         Height          =   285
         Left            =   2715
         TabIndex        =   24
         Top             =   510
         Width           =   315
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
         TabIndex        =   22
         Top             =   210
         Width           =   1485
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   2040
      Top             =   6225
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CanceL"
      Height          =   405
      Left            =   10365
      TabIndex        =   27
      Top             =   6105
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   9315
      TabIndex        =   26
      Top             =   6105
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   405
      Left            =   8265
      TabIndex        =   25
      Top             =   6105
      Width           =   1005
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2430
      Left            =   75
      TabIndex        =   21
      Top             =   1320
      Width           =   11775
      _cx             =   20770
      _cy             =   4286
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   15400959
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   15400959
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReceiptsPT.frx":008A
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   2
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
   Begin VSFlex8LCtl.VSFlexGrid vsGridTransactions 
      Height          =   2025
      Left            =   540
      TabIndex        =   35
      Top             =   4290
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
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance to be C/f"
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   6705
      TabIndex        =   58
      Top             =   4770
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   210
      Left            =   7200
      TabIndex        =   56
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   210
      Left            =   7230
      TabIndex        =   55
      Top             =   5025
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Total"
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   7665
      TabIndex        =   34
      Top             =   4155
      Width           =   360
   End
End
Attribute VB_Name = "frmReceiptsPT"
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
    Dim mDefaultTransactionTypeID   As Long
    Dim mDefaultAccountHeadCode     As String
    Dim mDefaultInstrumentID        As Long
    Dim mDefaultBankID              As Long
    Dim mDefaultBankHeadCode        As String
    Dim mDefaultZoneID              As Double
    Dim mBuildingID                 As Double
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
    Dim vchName_3        As String
    Dim vchHouseName_4   As String
    Dim vchStreetName_5  As String
    Dim vchMainPlace_6   As String
    Dim vchPostOffice_7  As String
    Dim vchDistrict_8    As String
    Dim vchPinNumber_9   As String
    Private Sub PrintReceipt()
        Dim mLoop As Long
        FileInitialize
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO, Tab(10); txtReceiptNo.Text; Tab(45); txtReceiptNo.Text
            Print #gbFileNO, Tab(10); txtBookNo.Text; Tab(20); txtDate; Tab(45); txtBookNo.Text; Tab(55); txtDate
            Print #gbFileNO, Tab(10); txtPayee; " "; txtInitial1; " "; txtInitial2; Tab(45); txtPayee; " "; txtInitial1; " "; txtInitial2
            Print #gbFileNO, Tab(10); txtHouseNo1; "/"; txtHouseNo2; ", "; Trim(txtHouseName);
            Print #gbFileNO, Tab(45); txtHouseNo1; "/"; txtHouseNo2; ", "; Trim(txtHouseName)
            Print #gbFileNO, Tab(10); vchStreetName_5; Tab(45); vchStreetName_5
            Print #gbFileNO, Tab(10); vchMainPlace_6; Tab(45); vchMainPlace_6
            Print #gbFileNO, Tab(10); vchPostOffice_7; Tab(45); vchMainPlace_6
            Print #gbFileNO, Tab(10); vchDistrict_8; " - "; vchPinNumber_9; Tab(45); vchDistrict_8; " - "; vchPinNumber_9
            Print #gbFileNO,
            Print #gbFileNO,
            Print #gbFileNO,
            
            For mLoop = 1 To vsGrid.Rows - 1
                Print #gbFileNO, PadL(Trim(Str(mLoop)), 2); vsGrid.TextMatrix(mLoop, 0); "  ";
                Print #gbFileNO, Right(vsGrid.TextMatrix(mLoop, 7), 2) & "-" & Right(Str(Val(vsGrid.TextMatrix(mLoop, 7)) + 1), 2);
                Print #gbFileNO, PadL(Trim(vsGrid.TextMatrix(mLoop, 7)), 6)
                
                
                
                Print #gbFileNO, Tab(45); PadL(Trim(Str(mLoop)), 2); vsGrid.TextMatrix(mLoop, 0); "  ";
                Print #gbFileNO, Right(vsGrid.TextMatrix(mLoop, 7), 2) & "-" & Right(Str(Val(vsGrid.TextMatrix(mLoop, 7)) + 1), 2)
            Next mLoop
            
        Close #gbFileNO
        ShellPad
    End Sub
    Private Sub ValuesForHiddenColumns()
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
    
    Private Sub FillGridYear()
        Dim mLoop As Integer
        Dim mItem As String
        mItem = ""
        For mLoop = 1991 To gbFinancialYearID
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
        Call gFillVSGrid(vsGrid, 1, "spGetAccHead4Receipts", enuSourceString.Saankhya)
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
            'CalculatFine:
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
            'ReadNextDemand:
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
        Dim arrInput As Variant
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        
        arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
                    Val(txtWard.Text), _
                    Val(txtHouseNo1), _
                    Trim(txtHouseNo2))
                    
        If objDB.CreateNewConnection(mCnn, SanchayaLite) Then
            Set Rec = objDB.ExecuteSP("spGetBuildingDetails", arrInput, , , mCnn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                mBuildingID = Rec!numBuildingID
                txtBuildingNo.Text = Rec!numBuildingID
                txtWard.Text = Rec!intWardNO
                txtHouseNo1.Text = Rec!intDoorNo1
                txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
                
                vchName_3 = IIf(IsNull(Rec!chvName), "", Rec!chvName)
                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
                vchHouseName_4 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
                vchStreetName_5 = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
                vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                vchPostOffice_7 = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
                vchDistrict_8 = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
                vchPinNumber_9 = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
                
                txtAddress.Text = vchName_3
                txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName_4
                txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
                txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
                txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice_7
                txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
                txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
            Else
                mBuildingID = -1
            End If
        End If
    End Sub
    
    Private Sub DisplayBuildingTaxDemands(mBuildingID As Double)
        Dim arrInput    As Variant
        Dim Rec         As New ADODB.Recordset
        Dim objDB       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim mRows       As Long
        Dim objAcc      As New clsAccounts
        Dim mArrearFlag As Integer
        Dim mAmtArrear  As Double
        Dim mAmtCurrent As Double
        Dim mFineFromDate As Date
        Dim mNoOfMonths As Integer
        Dim mFineAmt    As Double
        Dim mAmt        As Double
        
        vsGrid.Rows = 1
        mNumberOfSelections = 0
        arrInput = Array(mBuildingID)
        If objDB.SetConnection(mCnn) Then
            
            Rec.CursorLocation = adUseClient
            Set Rec = objDB.ExecuteSP("spGetPropertyTaxDemands", arrInput, , , mCnn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                vsGrid.Rows = 1
                mRows = 1
                vsGrid.MergeCells = flexMergeFree
                While Not Rec.EOF
                    vsGrid.Rows = vsGrid.Rows + 1
                    objAcc.SetAccountID (Rec!intAccountHeadID)
                    vsGrid.Cell(flexcpText, mRows, 0) = Rec!vchAccountHeadCode
                    vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                    vsGrid.Cell(flexcpText, mRows, 2) = Str(Rec!intYearID) & " - " & Str(Rec!intYearID + 1)
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: vsGrid.Cell(flexcpText, mRows, 3) = "Ist Half"
                        Case Is = 2: vsGrid.Cell(flexcpText, mRows, 3) = "IInd Half"
                        Case Is = 3: vsGrid.Cell(flexcpText, mRows, 3) = "Full Year"
                    End Select
                    vsGrid.MergeCol(12) = True
                    vsGrid.Cell(flexcpText, mRows, 12) = Rec!numDemandID
                    '-------------------------------------------------------'
                    ' To Restrict only 3 selections at a time               '
                    '-------------------------------------------------------'
                    If mNumberOfSelections < 3 Then
                        vsGrid.Cell(flexcpChecked, mRows, 12) = vbChecked
                    Else
                        vsGrid.Cell(flexcpChecked, mRows, 12) = 2
                    End If
                    vsGrid.Cell(flexcpText, mRows, 6) = Rec!intAccountHeadID
                    vsGrid.Cell(flexcpText, mRows, 7) = Rec!intYearID
                    vsGrid.Cell(flexcpText, mRows, 8) = Rec!tnyPeriodID
                    vsGrid.Cell(flexcpText, mRows, 9) = Rec!tnyArrearFlag
                    vsGrid.Cell(flexcpText, mRows, 10) = Rec!numDemandID
                    vsGrid.Cell(flexcpText, mRows, 11) = Rec!fltAmount
                    mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), 0, Rec!tnyArrearFlag)
                    If mArrearFlag Then
                        mAmtArrear = mAmtArrear + Rec!fltAmount
                        vsGrid.Cell(flexcpText, mRows, 4) = Rec!fltAmount
                    Else
                        mAmtCurrent = mAmtCurrent + Rec!fltAmount
                        vsGrid.Cell(flexcpText, mRows, 5) = Rec!fltAmount
                    End If
                    '-------------------------------------------------'
                    ' Calculating Fine Amount and Storing in Grid     '
                    '-------------------------------------------------'
                    If Rec!vchAccountHeadCode = mAcHeadCodePTaxArrear Then ' Case of Prpoerty Tax Arrear - Calculating Fine
                        mFineFromDate = DateSerial(Rec!intYearID, IIf(Rec!tnyPeriodID = 1, 1, 9), 1)
                        mNoOfMonths = DateDiff("m", mFineFromDate, gbTransactionDate)
                        mFineAmt = mFineAmt + (Rec!fltAmount * mFineRate / 100) * mNoOfMonths
                        vsGrid.Cell(flexcpText, mRows, 13) = (Rec!fltAmount * mFineRate / 100)
                    End If
                    If mNumberOfSelections < 3 Then
                        mNumberOfSelections = mNumberOfSelections + IIf(mRows Mod 2 = 0, 1, 0)
                    End If
                    
                    mRows = mRows + 1
                    Rec.MoveNext
                Wend
                '-------------------------------------------------'
                ' Head will be added for total Fine Amount        '
                '-------------------------------------------------'
                If mFineAmt > 0 Then
                    objAcc.SetAccountCode (mAcHeadCodeFine)
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
                    vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                    vsGrid.Cell(flexcpText, mRows, 2) = Str(gbFinancialYearID) & " - " & Str(gbFinancialYearID + 1)
                    vsGrid.Cell(flexcpText, mRows, 3) = ""
                    vsGrid.Cell(flexcpText, mRows, 5) = mFineAmt
                    vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
                    vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
                    vsGrid.Cell(flexcpText, mRows, 9) = 0
                    vsGrid.Cell(flexcpText, mRows, 10) = 0
                    vsGrid.Cell(flexcpText, mRows, 11) = mFineAmt
                End If
                mAmtCurrent = mAmtCurrent + mFineAmt
                txtTotalArrear.Text = Format(mAmtArrear, "0.00")
                txtTotalCurrent.Text = Format(mAmtCurrent, "0.00")
                txtTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
            End If
            Rec.Close
            Set mCnn = Nothing
        End If
    End Sub
    Private Sub ListPostingHeadsInGridForGeneralReceipts()
        Dim mLoopCount As Integer
        Dim objDB As clsDB
        vsGridTransactions.Rows = 1
        
        '------------------------------------------------------------------'
        ' Posting of Cash or Bank
        '------------------------------------------------------------------'
        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = mDrAccountHeadID
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(Val(txtTotal), "0.00")
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = "" 'RecTransactionHeads!intPostingHeadID
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
        
        For mLoopCount = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoopCount, 0) <> "" Then
                vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = mLoopCount + 1
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = vsGrid.TextMatrix(mLoopCount, 6)
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 0
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Val(vsGrid.TextMatrix(mLoopCount, 11))
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = mDrAccountHeadID
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
            Else
                Exit For
            End If
        Next mLoopCount
    End Sub
    
    Private Sub ListPostingHeadsInGrid(mTransactionType As Long, Optional mGroupID As Variant = Null)
                Dim mSQL As String
                Dim RecTransactionHeads As New ADODB.Recordset
                Dim mLoopCount As Long
                Dim mLoop As Long
                Dim mAmt As Double
                vsGridTransactions.Rows = 1
                mSQL = "Select * From faTransactionTypeChild Where intTransactionTypeID = " & mTransactionType & " Order By intOrder"
                Set RecTransactionHeads = GetRecordSet("spGetTransactionTypePostingHeads " & mTransactionType)
                For mLoopCount = 1 To vsGrid.Rows - 1
                    While Not RecTransactionHeads.EOF
                        If vsGrid.Cell(flexcpChecked, mLoopCount, 12) = 1 Then
                            If RecTransactionHeads!intAccountHeadID = Val(vsGrid.TextMatrix(mLoopCount, 6)) Then
                                If Val(vsGrid.TextMatrix(mLoopCount, 9)) Then    ' Arrear Flag = True
                                    mAmt = Val(vsGrid.TextMatrix(mLoopCount, 4)) ' Amount from the Arrear Column
                                Else
                                    mAmt = Val(vsGrid.TextMatrix(mLoopCount, 5)) ' Amount from the Current Column
                                End If
                                '--------------------------------------------------------------------------------'
                                ' Check whether the posting head is already there in the Transaction (Hidden)    '
                                ' Grid. If found add the Amount. Other wise add a new row in the Grid            '                                                 '
                                '--------------------------------------------------------------------------------'
                                For mLoop = 1 To vsGridTransactions.Rows - 1
                                    If RecTransactionHeads!intPostingHeadID = Val(vsGridTransactions.TextMatrix(mLoop, 1)) Then
                                        vsGridTransactions.TextMatrix(mLoop, 3) = Val(vsGridTransactions.TextMatrix(mLoop, 3)) + mAmt
                                        Exit For
                                    End If
                                Next mLoop
                                If mLoop = vsGridTransactions.Rows Then         ' Not found in Grid - Add as new row
                                     vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                                     vsGridTransactions.TextMatrix(mLoop, 0) = IIf(IsNull(RecTransactionHeads!intPostingHeadOrder), RecTransactionHeads!intOrder, RecTransactionHeads!intPostingHeadOrder)
                                     vsGridTransactions.TextMatrix(mLoop, 1) = RecTransactionHeads!intPostingHeadID
                                     vsGridTransactions.TextMatrix(mLoop, 2) = IIf(RecTransactionHeads!tinDebitOrCredit, 0, 1)
                                     vsGridTransactions.TextMatrix(mLoop, 3) = mAmt
                                     vsGridTransactions.TextMatrix(mLoop, 4) = ""
                                     vsGridTransactions.TextMatrix(mLoop, 5) = ""
                                     vsGridTransactions.TextMatrix(mLoop, 6) = ""
                                     vsGridTransactions.TextMatrix(mLoop, 7) = ""
                                     vsGridTransactions.TextMatrix(mLoop, 8) = ""
                                End If
                                vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intOrder
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intAccountHeadID
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = RecTransactionHeads!tinDebitOrCredit
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = mAmt
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = RecTransactionHeads!intPostingHeadID
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = Val(vsGrid.TextMatrix(mLoopCount, 10))
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                                GoTo STEP1:
                            End If
                        End If
                        RecTransactionHeads.MoveNext
                    Wend
STEP1:
                    RecTransactionHeads.MoveFirst
                Next
                
                '------------------------------------------------------------------'
                ' Posting of Cash or Bank
                '------------------------------------------------------------------'
                While Not RecTransactionHeads.EOF
                    If RecTransactionHeads!intAccountHeadID = mDrAccountHeadID Then
                        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intOrder
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intAccountHeadID
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = RecTransactionHeads!tinDebitOrCredit
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(Val(txtTotal), "0.00")
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = RecTransactionHeads!intPostingHeadID
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                        
                        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intPostingHeadOrder
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intPostingHeadID
                        If RecTransactionHeads!tinDebitOrCredit Then
                            vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 0
                        Else
                            vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 1
                        End If
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(Val(txtTotal), "0.00")
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                    End If
                    RecTransactionHeads.MoveNext
                Wend
                
                '------------------------------------------------------------------'
                ' Amount carry forward to Advance Head                             '
                '------------------------------------------------------------------'
                If Val(txtAdvance.Text) > 0 Then
                RecTransactionHeads.MoveFirst
                While Not RecTransactionHeads.EOF
                    If RecTransactionHeads!vchAccountHeadCode = mAcHeadCodeAdvance Then
                        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intOrder
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intAccountHeadID
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = RecTransactionHeads!tinDebitOrCredit
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(Val(txtAdvance), "0.00")
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = RecTransactionHeads!intPostingHeadID
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                        
                        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intPostingHeadOrder
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intPostingHeadID
                        If RecTransactionHeads!tinDebitOrCredit Then
                            vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 0
                        Else
                            vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 1
                        End If
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(Val(txtAdvance), "0.00")
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                    End If
                    RecTransactionHeads.MoveNext
                Wend
                End If
                
                If vsGridTransactions.Rows > 1 Then
                   vsGridTransactions.Select 1, 0, vsGridTransactions.Rows - 1, 0
                   vsGridTransactions.Sort = flexSortNumericAscending
                End If
    End Sub
    
    Public Sub Calculate()
        Dim mAmtArrear As Double
        Dim mAmtCurrent As Double
        Dim mCount As Long
        For mCount = 1 To vsGrid.Rows - 1
            If Val(vsGrid.TextMatrix(mCount, 4)) Then
                mAmtArrear = mAmtArrear + Val(vsGrid.Cell(flexcpText, mCount, 4))
            Else
                mAmtCurrent = mAmtCurrent + Val(vsGrid.Cell(flexcpText, mCount, 5))
            End If
        Next
        txtTotalArrear.Text = Format(mAmtArrear, "0.00")
        txtTotalCurrent.Text = Format(mAmtCurrent, "0.00")
        txtTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
        If Val(txtAdvance.Text) Then
            txtAdvance.Visible = True
            lblAdvance.Visible = True
        Else
            txtAdvance.Visible = False
            txtAdvance.Visible = False
            txtAdvance.Text = ""
        End If
    
    End Sub
    
    Private Sub ListMasters(mMasterType As Integer)
        Dim mSQL As String
        lstMasters.Tag = mMasterType
        Select Case mMasterType
            Case Is = 1 ' Transaction Type
                
                mSQL = "Select vchTransactionType, intTransactionTypeID, intGroupID From faTransactionType Where intGroupID = 10 And intExternalModuleID = 1"
                Call PopulateList(lstMasters, mSQL, "Property Tax", False, , True)
                lstMasters.Top = 1000
                lstMasters.Left = 550
                lstMasters.Height = 2000
                lstMasters.Visible = True
                lstMasters.SetFocus
            Case Is = 2 ' Instruments
                mSQL = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes"
                Call PopulateList(lstMasters, mSQL, "", , , True)
                lstMasters.Left = 8340
                lstMasters.Top = 450
                lstMasters.Height = 2000
                lstMasters.Width = txtInstrument.Width
                lstMasters.Visible = True
                lstMasters.SetFocus
            Case Is = 3 ' Account Heads
                If mMasterType Then
                    mSQL = "Select vchBankName, intBankID From faBanks Order By vchBankName"
                    Call PopulateList(lstMasters, mSQL, "", , , True)
                    lstMasters.Left = 8340
                    lstMasters.Top = txtAccountHead.Top - 25
                    lstMasters.Height = 2000
                    lstMasters.Visible = True
                    lstMasters.SetFocus
                Else
                    mSQL = "Select vchBankName, intBankID From faBanks Order By vchBankName"
                    Call PopulateList(lstMasters, mSQL, "", , , True)
                    lstMasters.Left = 8340
                    lstMasters.Top = 1380
                End If
            
                
        End Select
    End Sub
    Private Sub FormInitialize()
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        mUserSessions = -1
        mBuildingID = -1
        objDB.SetConnection mCnn
        Set Rec = GetRecordSet("spGetNewReceiptNoAndBookNo " & gbFinancialYearID & ", 10 ", adOpenStatic, adLockOptimistic, mCnn)
        If Not (Rec.EOF Or Rec.BOF) Then
            txtReceiptNo.Text = Format(Rec!intVoucherNo, "0000")
            txtBookNo.Text = Format(Rec!intBookNo, "0000")
        End If
        txtDate.Text = DdMmmYy(gbDate)
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        txtAdvance.Text = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        txtInstrument.Text = ""
        txtInstrument.Tag = ""
        txtDated.Text = ""
        txtInstNo.Text = ""
        txtBuildingNo.Text = ""
        txtWard.Text = ""
        txtHouseNo1.Text = ""
        txtHouseNo2.Text = ""
        txtPayee.Text = ""
        txtInitial1.Text = ""
        txtInitial2.Text = ""
        txtInitial3.Text = ""
        txtInitial4.Text = ""
        txtHouseName.Text = ""
        txtAddress.Text = ""
        txtTotalArrear.Text = ""
        txtTotalCurrent.Text = ""
        txtTotal.Text = ""

        txtDescription.Text = ""
       
        vchName_3 = ""
        vchHouseName_4 = ""
        vchStreetName_5 = ""
        vchMainPlace_6 = ""
        vchPostOffice_7 = ""
        vchDistrict_8 = ""
        vchPinNumber_9 = ""
        vsGrid.Rows = 1
        vsGrid.Rows = 15
        Call SetDefaultSettings
    End Sub
    Private Sub ClearAddress()
        mUserSessions = -1
        txtBuildingNo.Text = ""
        txtWard.Text = ""
        txtHouseNo1.Text = ""
        txtHouseNo2.Text = ""
        txtPayee.Text = ""
        txtInitial1.Text = ""
        txtInitial2.Text = ""
        txtInitial3.Text = ""
        txtInitial4.Text = ""
        txtHouseName.Text = ""
        txtAddress.Text = ""
    End Sub
    Private Sub FillZone()
        Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Order By chvZoneNameEnglish", , True, True, True, DBMaster)
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
        
        mDefaultTransactionTypeID = Val(ReadIniFile(gbSaankhyaINI, "Receipt PTax", "DefaultTransactionTypeID"))
        mDefaultAccountHeadCode = ReadIniFile(gbSaankhyaINI, "Receipt PTax", "DefaultAccountHeadCode")
        mDefaultInstrumentID = Val(ReadIniFile(gbSaankhyaINI, "Receipt PTax", "DefaultInstumentID"))
        mDefaultBankID = Val(ReadIniFile(gbSaankhyaINI, "Receipt PTax", "DefaultBankID"))
        mDefaultZoneID = Val(ReadIniFile(gbSaankhyaINI, "Receipt PTax", "DefaultZone"))
        
        If mDefaultZoneID > 0 Then
            For mLoopCount = 0 To cmbZone.ListCount - 1
                If cmbZone.ItemData(mLoopCount) = mDefaultZoneID Then
                    cmbZone.ListIndex = mLoopCount
                    Exit For
                End If
            Next
        End If
        
        objTranType.SetTransactionType (mDefaultTransactionTypeID)
        objAc.SetAccountCode (mDefaultAccountHeadCode)
        objInstruments.SetInstrumentType (mDefaultInstrumentID)
        objBank.SetBankInfo (mDefaultBankID)
        
        txtTransactionType.Text = objTranType.TransactionType
        txtTransactionType.Tag = objTranType.TransactionTypeID
        txtAccountHead.Text = objAc.AccountHead & " [ " & objAc.AccountCode & " ]"
        txtAccountHead.Tag = objAc.AccountHeadID
        cmdSearchAccountHead.Tag = objAc.GroupID    ' Cash Type
        
        txtInstrument.Text = objInstruments.InstrumentType
        txtInstrument.Tag = objInstruments.InstrumentTypeID
        mDefaultBankHeadCode = objBank.BankAccountHeadCode
        mAcHeadCodeAdvance = 350410101
        
    End Sub
    Private Sub cmbZone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub cmdCancel_Click()
        Select Case mUserSessions
            Case Is = 1
                Call FormInitialize
            Case Is = -1
                Unload Me
        End Select
    End Sub
    
    Private Sub cmdFind_Click()
        Call DisplayBuildingDetails
        Call DisplayBuildingTaxDemands(mBuildingID)


'        Dim arrInput As Variant
'        Dim Rec As New ADODB.Recordset
'        Dim objDb As New clsDB
'        Dim mCnn As New ADODB.Connection
'
'        arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
'                    Val(txtWard.Text), _
'                    Val(txtHouseNo1), _
'                    Trim(txtHouseNo2))
'        If objDb.CreateNewConnection(mCnn, SanchayaLite) Then
'            Call ClearAddress
'            Set Rec = objDb.ExecuteSP("spGetBuildingDetails", arrInput, , , mCnn, adCmdStoredProc)
'            If Not (Rec.BOF And Rec.EOF) Then
'                mUserSessions = 1
'                mBuildingID = Rec!numBuildingID
'                txtBuildingNo.Text = Rec!numBuildingID
'                txtWard.Text = Rec!intWardNO
'                txtHouseNo1.Text = Rec!intDoorNo1
'                txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
'                'txtHouseNo3.Text = IIf(IsNull(Rec!chvDoorNo3), "", Rec!chvDoorNo3)
'                txtPayee.Text = IIf(IsNull(Rec!chvName), "", Rec!chvName & " ")
'                txtInitial1.Text = IIf(IsNull(Rec!chvInitial1), "", Rec!chvInitial1)
'                txtInitial2.Text = IIf(IsNull(Rec!chvInitial2), "", Rec!chvInitial2)
'                txtInitial3.Text = IIf(IsNull(Rec!chvInitial3), "", Rec!chvInitial3)
'                txtInitial4.Text = IIf(IsNull(Rec!chvInitial4), "", Rec!chvInitial4)
'                txtHouseName.Text = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
'
'                vchName_3 = IIf(IsNull(Rec!chvName), "", Rec!chvName)
'                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial1), "", "." & Rec!chvInitial1)
'                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial2), "", "." & Rec!chvInitial2)
'                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial3), "", "." & Rec!chvInitial3)
'                vchName_3 = vchName_3 & IIf(IsNull(Rec!chvInitial4), "", "." & Rec!chvInitial4)
'                vchHouseName_4 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
'                vchStreetName_5 = IIf(IsNull(Rec!chvHouseName), "", Rec!chvResStreetName)
'                vchMainPlace_6 = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
'                vchPostOffice_7 = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
'                vchDistrict_8 = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
'                vchPinNumber_9 = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
'
'                txtAddress.Text = IIf(IsNull(Rec!chvResidenceAssNo), "", Rec!chvResidenceAssNo & " - ")
'                txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName_5
'                txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace_6), vchMainPlace_6 & ", ", "")
'                txtAddress.Text = txtAddress.Text & vchPostOffice_7
'                txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict_8
'                txtAddress.Text = txtAddress.Text & " - " & vchPinNumber_9
'                Call DisplayBuildingTaxDemands(mBuildingID)
'            End If
'            Rec.Close
'            Set mCnn = Nothing
'        End If
        
        
    End Sub
    Private Sub cmdNew_Click()
        Call PrintReceipt
    End Sub
    Private Sub cmdSave_Click()
        '==============================================================='
        ' Function    : Revenues :- Property Tax       :9091 0000
        '                           Professional Tax   :9094 0000
        ' Functionary : Revenue Department             :08
        '==============================================================='
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim objFunctions As New clsFunction
        Dim objFunctionaries As New clsFunctionary
        Dim mFunctionaryID  As Variant
        Dim mFunctionID As Variant
        Dim mLoopCount As Long
        Dim mLoop As Long
        
        Dim Rec As New ADODB.Recordset
        
        mTransactionType = Val(txtTransactionType.Tag)
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then ' CREATED NEW CONNECTION
                objFunctionaries.SetFunctionary ("080000")
                mFunctionaryID = objFunctionaries.FunctionaryID
                Select Case mTransactionType
                    Case 2 ' Property Tax
                        objFunctions.SetFunction ("90910000")
                        mFunctionID = objFunctions.FunctionID
                    Case Else
                        mFunctionID = Null
                End Select
                If Val(txtAccountHead.Tag) > 0 Then
                    mDrAccountHeadID = Val(txtAccountHead.Tag)
                Else
                    MsgBox "Error : Cash/Bank AccountHead is not set", vbInformation
                    Exit Sub
                End If
                
                If mTransactionType = 9999 Then
                    Call ListPostingHeadsInGridForGeneralReceipts
                Else
                    Call ListPostingHeadsInGrid(mTransactionType)
                End If
                    
                
                ' Exit Sub
                '-------------------------------------------------------'
                ' faVoucher
                '-------------------------------------------------------'
                Dim mintVoucherID_1                As Double
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                Dim mintTransactionTypeID_4        As Long
                Dim mtnyVoucherTypeID_5            As Integer
                Dim mintVoucherNo_6                As Long
                Dim mintBookNo_7                   As Long
                Dim mdtDate_8                      As Date
                Dim mfltAmount_9                   As Double
                Dim mintInstrumentTypeID_10        As Integer
                Dim mvchInstrumentNo_11            As String
                Dim mdtInstrumentDate_12           As Variant
                Dim mvchDescription_13             As String
                Dim mnumZoneID_14                  As Variant
                Dim mnumWardID_15                  As Double
                Dim mintDoorNoP1_16                As Long
                Dim mvchDoorNoP2_17                As String
                Dim mvchDoorNoP3_18                As String
                Dim mintUserID_19                  As Long
                Dim mintCounterID_20               As Long
                Dim mnumSubLedgerID_21             As Double
                Dim mintKeyID1_22                  As Variant
                Dim mintKeyID2_23                  As Variant
                Dim mintExternalApplicationID_24   As Long
                Dim mintExternalModuleID_25        As Long
                Dim mintFinancialYearID_26         As Long
                
                '@intVoucherID_1     [bigint],
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                mintTransactionTypeID_4 = Val(txtTransactionType.Tag)
                mtnyVoucherTypeID_5 = 10
                mintVoucherNo_6 = Val(txtReceiptNo.Text)
                mintBookNo_7 = Val(txtBookNo.Text)
                mdtDate_8 = gbDate
                mfltAmount_9 = Val(txtTotal.Text)
                mintInstrumentTypeID_10 = Val(txtInstrument.Tag)
                mvchInstrumentNo_11 = Trim(txtInstNo.Text)
                mdtInstrumentDate_12 = IIf(Trim(txtDated) <> "", CheckDateInMMM(txtDated), Null)
                mvchDescription_13 = Trim(txtDescription.Text)
                mnumZoneID_14 = IIf(cmbZone.ItemData(cmbZone.ListIndex) > 0, cmbZone.ItemData(cmbZone.ListIndex), Null)
                mnumWardID_15 = Val(txtWard.Text)
                mintDoorNoP1_16 = Val(txtHouseNo1.Text)
                mvchDoorNoP2_17 = Trim(txtHouseNo2.Text)
                mvchDoorNoP3_18 = Trim(txtHouseNo3.Text)
                mintUserID_19 = gbUserID
                mintCounterID_20 = gbCounterNo
                mnumSubLedgerID_21 = mBuildingID
                mintKeyID1_22 = mDrAccountHeadID
                mintKeyID2_23 = Null
                mintExternalApplicationID_24 = AppID.Saankhya
                mintExternalModuleID_25 = 0
                mintFinancialYearID_26 = gbFinancialYearID
                '========================================='
                ' BEGIN TRANSACTION                       '
                '-----------------------------------------'
                    'mCnn.BeginTrans
                    'On Error GoTo ErrorRollBack:
                '========================================='
                arrInput = Array( _
                -1, _
                gbLocalBodyID, _
                Null, _
                mintTransactionTypeID_4, _
                mtnyVoucherTypeID_5, _
                mintVoucherNo_6, _
                mintBookNo_7, _
                mdtDate_8, _
                mfltAmount_9, _
                mintInstrumentTypeID_10, _
                mvchInstrumentNo_11, _
                mdtInstrumentDate_12, _
                mvchDescription_13, _
                mnumZoneID_14, _
                mnumWardID_15, _
                mintDoorNoP1_16, _
                mvchDoorNoP2_17, _
                mvchDoorNoP3_18, _
                mintUserID_19, _
                mintCounterID_20, _
                mnumSubLedgerID_21, _
                mintKeyID1_22, mintKeyID2_23, mintExternalApplicationID_24, _
                mintExternalModuleID_25, _
                mintFinancialYearID_26)
                
                objDB.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID_1 = arrOutPut(0, 0)
                Else
                    GoTo ErrorRollBack:
                End If
                '-------------------------------------------------------'
                ' faVoucher Address
                '-------------------------------------------------------'
                'Dim intVoucherID_1  As Double
                'Dim intLocalBodyID_2 As Long
                'Dim vchName_3        As String
                'Dim vchHouseName_4   As String
                'Dim vchStreetName_5  As String
                'Dim vchMainPlace_6   As String
                'Dim vchPostOffice_7  As String
                'Dim vchDistrict_8    As String
                'Dim vchPinNumber_9   As String
                arrInput = Array(mintVoucherID_1, gbLocalBodyID, vchName_3, vchHouseName_4, vchStreetName_5, vchMainPlace_6, vchPostOffice_7, vchDistrict_8, vchPinNumber_9)
                objDB.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn
                '-------------------------------------------------------'
                ' faVoucher Child
                '-------------------------------------------------------'
                'Dim mintVoucherID_1         As Double  '
                Dim mintLocalBodyID_2       As Long
                Dim mintSlNo_3              As Long
                Dim mintAccountHeadID_4     As Long
                Dim mtnyDebitOrCredit_5     As Integer
                Dim mintYearID_6            As Long
                Dim mtnyPeriodID_7          As Integer
                Dim mtnyArrearFlag_8        As Integer
                Dim mnumDemandID_9          As Double
                Dim mfltAmount_10           As Double
                For mLoopCount = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                    
                        mintLocalBodyID_2 = gbLocalBodyID
                        mintSlNo_3 = mLoopCount
                        mintAccountHeadID_4 = vsGrid.Cell(flexcpText, mLoopCount, 6)
                        mtnyDebitOrCredit_5 = 0
                        mintYearID_6 = Val(vsGrid.Cell(flexcpText, mLoopCount, 7))
                        mtnyPeriodID_7 = Val(vsGrid.Cell(flexcpText, mLoopCount, 8))
                        mtnyArrearFlag_8 = Val(vsGrid.Cell(flexcpText, mLoopCount, 9))
                        mnumDemandID_9 = Val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                        mfltAmount_10 = Val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                        
                        arrInput = Array( _
                        mintVoucherID_1, _
                        mintLocalBodyID_2, _
                        mintSlNo_3, _
                        mintAccountHeadID_4, _
                        mtnyDebitOrCredit_5, _
                        mintYearID_6, _
                        mtnyPeriodID_7, _
                        mtnyArrearFlag_8, _
                        mnumDemandID_9, _
                        mfltAmount_10 _
                        )
                        objDB.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                    Else
                        Exit For
                    End If
                Next mLoopCount
                '-------------------------------------------------------'
                ' Transactions
                '-------------------------------------------------------'
                Dim intTransactionID_1   As Double
                'Dim mintLocalBodyID_2  As Long
                Dim mintFinancialYearID_3  As Long
                Dim mdtTransactionDate_4   As Date
                Dim mintExternalApplicationID_5    As Long
                Dim mintExternalApplicationModuleID_6  As Long
                Dim mintFunctionID_7   As Variant
                Dim mintFunctionaryID_8   As Variant
                Dim mintFieldID_9 As Variant
                Dim mintFundID_10 As Variant
                Dim mintBudgetCentreID_11  As Variant
                Dim mvchNarration_12   As String
                Dim mintTransactionTypeID_13   As Long
                Dim mintVoucherNo_14   As Long
                Dim mintProcessID_15    As Variant
                Dim mintGroupID_17    As Long
                Dim mvchGroup_16   As String
                Dim mintKeyID_18   As Variant
                Dim mnumSubLedgerID_19    As Double
                'Dim mintUserID_20  As Long
                intTransactionID_1 = -1
                mintLocalBodyID_2 = gbLocalBodyID
                mintFinancialYearID_3 = gbFinancialYearID
                mdtTransactionDate_4 = gbDate
                mintExternalApplicationID_5 = AppID.Saankhya
                mintExternalApplicationModuleID_6 = 0
                mintFunctionID_7 = mFunctionID
                mintFunctionaryID_8 = mFunctionaryID
                mintFieldID_9 = IIf(Val(txtWard) < 1, Null, Val(txtWard))
                mintFundID_10 = Null
                mintBudgetCentreID_11 = Null
                mvchNarration_12 = Trim(txtDescription.Text)
                mintTransactionTypeID_13 = mTransactionType
                mintVoucherNo_14 = mintVoucherID_1
                mintProcessID_15 = Null
                mvchGroup_16 = "R"
                mintGroupID_17 = 10
                mintKeyID_18 = Null
                mnumSubLedgerID_19 = mBuildingID
                'mintUserID_20 = gbUserID
                
                arrInput = Array( _
                intTransactionID_1, _
                mintLocalBodyID_2, _
                mintFinancialYearID_3, _
                mdtTransactionDate_4, _
                mintExternalApplicationID_5, _
                mintExternalApplicationModuleID_6, _
                mintFunctionID_7, _
                mintFunctionaryID_8, _
                mintFieldID_9, _
                mintFundID_10, _
                mintBudgetCentreID_11, _
                mvchNarration_12, _
                mintTransactionTypeID_13, _
                mintProcessID_15, _
                mvchGroup_16, _
                mintGroupID_17, _
                mintKeyID_18, _
                mnumSubLedgerID_19, _
                gbUserID, _
                mintVoucherNo_14)
                
                Set arrOutPut = Nothing
                objDB.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    intTransactionID_1 = arrOutPut(0, 0)
                Else
                    GoTo ErrorRollBack:
                End If
                '-------------------------------------------------------'
                ' Transaction Child
                '-------------------------------------------------------'
                For mLoop = 1 To vsGridTransactions.Rows - 1
                    '@intTransactionID      int     ,
                    '@intSerialNo           int     ,
                    '@intAccountHeadID      int     ,
                    '@fltAmount             float   ,
                    '@tinDebitOrCreditFlag  tinyint ,
                    '@intByAccountHeadID    int ,
                    '@vchNarration          varChar(500)  ,
                    '@intFundID             int
                    arrInput = Array(intTransactionID_1, _
                                    mLoop, _
                                    vsGridTransactions.TextMatrix(mLoop, 1), _
                                    vsGridTransactions.TextMatrix(mLoop, 3), _
                                    vsGridTransactions.TextMatrix(mLoop, 2), _
                                    vsGridTransactions.TextMatrix(mLoop, 4), _
                                    Null, _
                                    Null)
                    objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                Next mLoop
                '-------------------------------------------------------'
                ' Update Demand Table
                '-------------------------------------------------------'
                If mTransactionType = 1 Then
                Dim mDemandID   As Double
                Dim mStatusFlag As Integer
                For mLoop = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked And mDemandID <> vsGrid.Cell(flexcpText, mLoop, 10) Then
                        mDemandID = Val(vsGrid.TextMatrix(mLoop, 10))
                        mStatusFlag = 1
                        arrInput = Array(mDemandID, mStatusFlag, mintVoucherID_1)
                        objDB.ExecuteSP "spUpdateIDemandStatus", arrInput, , , mCnn
                    End If
                Next mLoop
                End If
                '========================================='
                ' TRANSACTION COMMITTING                  '
                '-----------------------------------------'
                    'mCnn.CommitTrans
                    Set mCnn = Nothing
                '========================================='
                Call FormInitialize
                On Error GoTo 0
                
                
                '========================================='
                ' Sharing Data to SanchayaLite            '
                '-----------------------------------------'
                If mTransactionType = 1 Then
                Dim mchvReceiptNO As String
                Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                If Not (Rec.EOF And Rec.BOF) Then
                If objDB.CreateNewConnection(mCnn, SanchayaLite) Then
                    '@numDemandID       [numeric],      -- demand id from saankhya
                    '@intVoucherID      [int],          -- Receipt id from saankhya
                    '@chvReceiptNo      [varchar](20),  -- Receipt number
                    '@dtVoucherDate     [smalldatetime],-- receipt date
                    '@fltCollection     [float]         -- total amount against this demand
                    mDemandID = Rec!numDemandID
                    mintVoucherID_1 = Rec!intVoucherID
                    mchvReceiptNO = Trim(Str(Rec!intBookNo)) & "/" & Trim(Str(Rec!intVoucherNo))
                    mdtDate_8 = Rec!dtDate
                    
                    arrInput = Array(mDemandID, mintVoucherID_1, mchvReceiptNO, mdtDate_8, 0)
                    objDB.ExecuteSP "sp_SankhyaReceiptUpdate_U", arrInput, , , mCnn
                Else
                    MsgBox "Connection Error:", vbInformation
                End If
                End If
                End If
        Else
                Debug.Print "Error in establishing connection with Saankhya DB"
                Exit Sub
        End If
        Exit Sub
ErrorRollBack:
        'mCnn.RollbackTrans
        Set mCnn = Nothing
    End Sub

    Private Sub cmdSearchAccountHead_Click()
        If Val(cmdSearchAccountHead.Tag) = 2 Then
            Call ListMasters(3)
        End If
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
        Call Calculate
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 And Shift = 2 Then
            Call MsgBox("Search!", vbInformation)
        End If
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FillZone
        Call FormInitialize
        vsGrid.ColComboList(0) = "|..."
    End Sub
    Private Sub lstMasters_DblClick()
        Select Case Val(lstMasters.Tag)
            Case 1
                txtTransactionType.Text = lstMasters.Text
                txtTransactionType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
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
                    txtAccountHead.Tag = objBank.BankID
                End If
        End Select
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
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
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
        txtDated.Text = CheckDateInMMM(txtDated.Text)
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
        If Val(txtInstrument.Tag) = gbInstrumentCheque Then
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
    Private Sub txtTotal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtTransactionType_DblClick()
       Call txtTransactionType_KeyPress(13)
    End Sub
    Private Sub txtTransactionType_GotFocus()
        If Trim(txtTransactionType.Text) = "" Then
            ListMasters (1)
            lstMasters.Refresh
        End If
    End Sub
    Private Sub txtTransactionType_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        Select Case Val(txtTransactionType.Tag)
            Case Is = gbTransactionTypePTax
                On Error Resume Next
                frmPropertyTax.Show vbModal
            Case Is = 9999
                Call FillAccountHeads
                Call FillGridYear
                
            Case Else
        End Select
        Call PressTabKey
        End If
    End Sub
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
            If vsGrid.Row > 1 Then
                If vsGrid.TextMatrix(vsGrid.Row - 1, 0) = "" Or _
                   (Val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 And _
                   Val(vsGrid.TextMatrix(vsGrid.Row - 1, 5)) <= 0) Then
                   Cancel = True
                   Exit Sub
                End If
            End If
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                    vsGrid.TextMatrix(Row, 1) = objAccHead.AccountHead
                    vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
                '
                'vsGrid.TextMatrix(vsGrid.Row, 0) = Token(gbSearchStr, " ")
                'vsGrid.TextMatrix(vsGrid.Row, 1) = Trim(gbSearchStr)
                vsGrid.Col = vsGrid.Col + 2
                vsGrid.Redraw = flexRDDirect
                gbSearchStr = ""
            End If
    End Sub
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
    End Sub
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        If vsGrid.Row > 0 Then
            
            If Col = 1 And vsGrid.ComboIndex > -1 Then
                Set objAccHead = New clsAccounts
                If objAccHead.FindAccountByHead(Trim(vsGrid.ComboItem)) Then
                vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
            ElseIf Col = 4 Then
                vsGrid.TextMatrix(Row, 4) = Format(Val(vsGrid.TextMatrix(Row, 4)), "0.00")
                If Val(vsGrid.TextMatrix(Row, 4)) > 0 Then
                vsGrid.TextMatrix(Row, 5) = ""
                End If
                Call Calculate
            ElseIf vsGrid.Col = 5 Then
                vsGrid.TextMatrix(Row, 5) = Format(Val(vsGrid.TextMatrix(Row, 5)), "0.00")
                If Val(vsGrid.TextMatrix(Row, 5)) > 0 Then
                vsGrid.TextMatrix(Row, 4) = ""
                End If
                Call Calculate
            End If
            Call ValuesForHiddenColumns
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
    
    

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReceiptsCounter 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                                                                                             R e c e i p t s"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmReceiptsCounter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleMode       =   0  'User
   ScaleWidth      =   11970.31
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox frameTransaction 
      BackColor       =   &H00C0C0C0&
      Height          =   810
      Left            =   6990
      ScaleHeight     =   750
      ScaleWidth      =   4395
      TabIndex        =   130
      Top             =   5310
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   133
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtCardTransaction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         TabIndex        =   131
         Top             =   315
         Width           =   2100
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1050
         TabIndex        =   132
         Top             =   105
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   15
         Picture         =   "frmReceiptsCounter.frx":1CCA
         Top             =   -45
         Width           =   1095
      End
   End
   Begin VB.Frame fraParty 
      BackColor       =   &H00808080&
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
      Height          =   450
      Left            =   6765
      TabIndex        =   99
      Top             =   6180
      Visible         =   0   'False
      Width           =   525
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
         Height          =   975
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   1605
         Width           =   4995
      End
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
         TabIndex        =   87
         Top             =   585
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
         TabIndex        =   95
         Top             =   930
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
         TabIndex        =   94
         Top             =   930
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
         TabIndex        =   93
         Top             =   930
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
         TabIndex        =   92
         Top             =   930
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
         TabIndex        =   97
         Top             =   1260
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
         TabIndex        =   86
         Top             =   585
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
         TabIndex        =   91
         Top             =   930
         Width           =   3765
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
         TabIndex        =   84
         Top             =   585
         Width           =   1485
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
         TabIndex        =   88
         Top             =   585
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.ComboBox cmbZone 
         Height          =   315
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   240
         Width           =   2325
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "-->"
         Height          =   315
         Left            =   5730
         TabIndex        =   89
         Top             =   585
         Width           =   405
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
         TabIndex        =   96
         Top             =   1269
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
         TabIndex        =   90
         Top             =   990
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
         TabIndex        =   85
         Top             =   645
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
         TabIndex        =   83
         Top             =   645
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
         TabIndex        =   81
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   11445
      Top             =   4230
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   4395
      Width           =   1710
   End
   Begin VB.CheckBox chkRoundOff 
      Caption         =   "Check1"
      Height          =   195
      Left            =   7005
      TabIndex        =   80
      Top             =   5085
      Width           =   210
   End
   Begin VB.TextBox txtRoundOff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5010
      Width           =   690
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
      Left            =   1905
      TabIndex        =   3
      Top             =   135
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2445
      Left            =   30
      TabIndex        =   30
      Top             =   1500
      Width           =   11745
      _cx             =   20717
      _cy             =   4313
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
      BackColor       =   16318457
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16318457
      BackColorAlternate=   16318457
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReceiptsCounter.frx":5661
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
      TextStyleFixed  =   1
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
      Begin VB.Label lblInterruptStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "The request for Interrupted Receipt is Pending for Approval !"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2250
         TabIndex        =   125
         Top             =   1125
         Visible         =   0   'False
         Width           =   7635
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   1965
         Picture         =   "frmReceiptsCounter.frx":584C
         Stretch         =   -1  'True
         Top             =   1125
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.TextBox txtAdvance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   4695
      Visible         =   0   'False
      Width           =   1710
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
      Height          =   1560
      Left            =   6990
      TabIndex        =   105
      Top             =   -75
      Width           =   4845
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3300
         TabIndex        =   29
         Top             =   1110
         Width           =   1470
      End
      Begin VB.TextBox txtBank 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1110
         Width           =   1740
      End
      Begin VB.CommandButton cmdSearchInstrument 
         Caption         =   "..."
         Height          =   285
         Left            =   4455
         TabIndex        =   22
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Height          =   285
         Left            =   4455
         TabIndex        =   19
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
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   210
         Width           =   3420
      End
      Begin VB.TextBox txtInstNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   24
         Top             =   810
         Width           =   1740
      End
      Begin VB.TextBox txtDated 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3300
         TabIndex        =   25
         Top             =   810
         Width           =   1470
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
         Locked          =   -1  'True
         TabIndex        =   21
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
         TabIndex        =   28
         Top             =   1140
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
         TabIndex        =   26
         Top             =   1110
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
         TabIndex        =   17
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
         Left            =   270
         TabIndex        =   23
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
         Left            =   2775
         TabIndex        =   103
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
         Left            =   30
         TabIndex        =   20
         Top             =   540
         Width           =   915
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   5010
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
      Height          =   540
      Left            =   7980
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   76
      Top             =   5325
      Width           =   3450
   End
   Begin VB.TextBox txtTotalCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9705
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   4095
      Width           =   1725
   End
   Begin VB.TextBox txtTotalArrear 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   4095
      Width           =   1710
   End
   Begin VB.Frame fraTransactionType 
      BackColor       =   &H00DAF2F2&
      Height          =   1545
      Left            =   30
      TabIndex        =   98
      Top             =   -60
      Width           =   6930
      Begin VB.CheckBox chkIntrNoSuffix 
         Caption         =   "add Suffix to Interrupted Receipt No "
         Height          =   330
         Left            =   1035
         TabIndex        =   129
         Top             =   1170
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txtIntruptNoSuffix 
         Height          =   285
         Left            =   6525
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   128
         Top             =   855
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtDemandPrefix 
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
         Left            =   1590
         MaxLength       =   14
         TabIndex        =   6
         Text            =   "99999"
         Top             =   540
         Width           =   1080
      End
      Begin VB.CheckBox chkLinkDemand 
         Caption         =   "Check1"
         Height          =   195
         Left            =   180
         TabIndex        =   126
         ToolTipText     =   "Link with Web Interface"
         Top             =   585
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtZone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         TabIndex        =   16
         Top             =   1170
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.TextBox txtReceiptNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   855
         Width           =   1500
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   555
         Width           =   1500
      End
      Begin VB.ListBox lstTransactionType 
         Height          =   255
         Left            =   4260
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtOutDoorStaff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         TabIndex        =   12
         Top             =   855
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.CommandButton cmdSearchDemandNo 
         Caption         =   "..."
         Height          =   285
         Left            =   3885
         TabIndex        =   8
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox txtDemandNo 
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
         Left            =   2820
         MaxLength       =   8
         ScrollBars      =   1  'Horizontal
         TabIndex        =   7
         Text            =   "99999"
         Top             =   540
         Width           =   1050
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
         Left            =   1590
         TabIndex        =   1
         Top             =   210
         Width           =   4905
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         Caption         =   "..."
         Height          =   285
         Left            =   6525
         TabIndex        =   2
         Top             =   195
         Width           =   315
      End
      Begin VB.Label lblInterruptedReceipt 
         BackStyle       =   0  'Transparent
         Caption         =   "Interrupted Receipt"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4230
         TabIndex        =   124
         Top             =   1200
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label lblZone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zonal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1005
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   405
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
         Left            =   4560
         TabIndex        =   9
         Top             =   585
         Width           =   405
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
         Left            =   4005
         TabIndex        =   13
         Top             =   855
         Width           =   960
      End
      Begin VB.Label lblOutDoorStaff 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&OutDoor Staff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   210
         TabIndex        =   11
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2700
         TabIndex        =   116
         Top             =   480
         Width           =   105
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Demand No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   510
         TabIndex        =   5
         Top             =   540
         Width           =   1080
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   75
         TabIndex        =   0
         Top             =   210
         Width           =   1485
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   6825
      Top             =   6615
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   3
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CanceL"
      Height          =   405
      Left            =   10425
      TabIndex        =   78
      Top             =   6165
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   9390
      TabIndex        =   77
      Top             =   6165
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   405
      Left            =   8325
      TabIndex        =   79
      Top             =   6165
      Width           =   1005
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGridTransactions 
      Height          =   2025
      Left            =   7230
      TabIndex        =   100
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
   Begin VB.Frame fraSubLedger 
      BackColor       =   &H80000015&
      Caption         =   "Rend on Land and Buildings"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6330
      TabIndex        =   106
      Top             =   6150
      Width           =   555
      Begin VB.TextBox txtThirdLine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1725
         TabIndex        =   112
         Top             =   1020
         Width           =   3945
      End
      Begin VB.TextBox txtSecondLine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1725
         TabIndex        =   110
         Top             =   690
         Width           =   3945
      End
      Begin VB.TextBox txtFirstLine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1725
         TabIndex        =   108
         Top             =   360
         Width           =   3945
      End
      Begin VB.TextBox txtMemo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   1725
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   1395
         Width           =   3930
      End
      Begin VB.Label lblLessee 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lessee"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1185
         TabIndex        =   114
         Top             =   1365
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building Name "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   113
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lblShop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1350
         TabIndex        =   111
         Top             =   705
         Width           =   345
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1035
         TabIndex        =   109
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.Frame fraReceiptNo 
      BackColor       =   &H00DAF2F2&
      Height          =   1425
      Left            =   4320
      TabIndex        =   101
      Top             =   -60
      Width           =   2640
      Begin VB.TextBox txtBookNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1095
         TabIndex        =   102
         Top             =   510
         Visible         =   0   'False
         Width           =   1500
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
         Left            =   1080
         TabIndex        =   104
         Top             =   540
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame fraDemandDetails 
      BackColor       =   &H00DAF2F2&
      Height          =   2745
      Left            =   30
      TabIndex        =   117
      Top             =   3930
      Visible         =   0   'False
      Width           =   6705
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   5790
         TabIndex        =   127
         Top             =   2355
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox txtBuildingNo 
         Enabled         =   0   'False
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
         Left            =   855
         TabIndex        =   39
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CheckBox chkGroupReceipt 
         Caption         =   "Check1"
         Height          =   195
         Left            =   90
         TabIndex        =   123
         Top             =   2490
         Width           =   195
      End
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
         Left            =   855
         MaxLength       =   50
         TabIndex        =   41
         Top             =   1605
         Width           =   1755
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
         Left            =   855
         MaxLength       =   3
         TabIndex        =   34
         Top             =   540
         Width           =   1755
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
         Left            =   855
         MaxLength       =   5
         TabIndex        =   36
         Top             =   870
         Width           =   1080
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
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   37
         Top             =   870
         Width           =   645
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
         ItemData        =   "frmReceiptsCounter.frx":639D
         Left            =   855
         List            =   "frmReceiptsCounter.frx":639F
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   195
         Width           =   1770
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
         Left            =   3165
         MaxLength       =   100
         TabIndex        =   45
         Top             =   210
         Width           =   2175
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
         Left            =   3750
         MaxLength       =   100
         TabIndex        =   51
         Top             =   540
         Width           =   2850
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
         Left            =   3750
         MaxLength       =   100
         TabIndex        =   53
         Top             =   855
         Width           =   2850
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
         Left            =   3750
         MaxLength       =   100
         TabIndex        =   55
         Top             =   1170
         Width           =   2850
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
         Left            =   3750
         MaxLength       =   100
         TabIndex        =   57
         Top             =   1485
         Width           =   2850
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
         Left            =   5310
         MaxLength       =   1
         TabIndex        =   46
         Top             =   210
         Width           =   345
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
         Left            =   5625
         MaxLength       =   1
         TabIndex        =   47
         Top             =   210
         Width           =   345
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
         Left            =   5940
         MaxLength       =   1
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   210
         Width           =   345
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
         Left            =   6255
         MaxLength       =   1
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   210
         Width           =   345
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
         Left            =   3750
         MaxLength       =   50
         TabIndex        =   59
         Top             =   1800
         Width           =   1650
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
         Left            =   5655
         MaxLength       =   6
         TabIndex        =   61
         Top             =   1800
         Width           =   945
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
         Height          =   315
         Left            =   3750
         MaxLength       =   15
         TabIndex        =   63
         Top             =   2115
         Width           =   1650
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
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1920
         Width           =   1770
      End
      Begin VB.Label lblBuildingNo 
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
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   60
         TabIndex        =   38
         Top             =   1245
         Width           =   795
      End
      Begin VB.Label lblFromReceiptNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   450
         TabIndex        =   122
         Top             =   2475
         Width           =   345
      End
      Begin VB.Label lblToReceiptNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   1770
         TabIndex        =   121
         Top             =   2475
         Width           =   195
      End
      Begin VB.Label lblGroupTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00DAF2F2&
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2985
         TabIndex        =   120
         Top             =   2490
         Width           =   315
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
         Left            =   405
         TabIndex        =   40
         Top             =   1650
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
         Left            =   225
         TabIndex        =   33
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
         Left            =   270
         TabIndex        =   35
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
         Left            =   480
         TabIndex        =   31
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
         Left            =   2730
         TabIndex        =   44
         Top             =   270
         Width           =   435
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
         Left            =   2760
         TabIndex        =   50
         Top             =   585
         Width           =   990
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
         Left            =   3285
         TabIndex        =   52
         Top             =   900
         Width           =   465
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
         Left            =   2895
         TabIndex        =   54
         Top             =   1215
         Width           =   855
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
         Left            =   2955
         TabIndex        =   56
         Top             =   1530
         Width           =   795
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
         Left            =   3405
         TabIndex        =   58
         Top             =   1845
         Width           =   345
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
         Left            =   5430
         TabIndex        =   60
         Top             =   1860
         Width           =   450
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
         Left            =   3030
         TabIndex        =   62
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fwd To"
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
         Left            =   300
         TabIndex        =   42
         Top             =   2010
         Width           =   555
      End
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   8805
      TabIndex        =   67
      Top             =   4425
      Width           =   870
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round off"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   7260
      TabIndex        =   73
      Top             =   5055
      Width           =   705
   End
   Begin VB.Label lblAdminNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin. Note:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   8160
      TabIndex        =   119
      Top             =   5895
      UseMnemonic     =   0   'False
      Width           =   960
   End
   Begin VB.Label lblAdminNoteCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin. Note:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   7125
      TabIndex        =   118
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Adj."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   8745
      TabIndex        =   69
      Top             =   4725
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   8835
      TabIndex        =   71
      Top             =   5055
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Description"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   7170
      TabIndex        =   75
      Top             =   5325
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   7575
      TabIndex        =   64
      Top             =   4140
      Width           =   375
   End
End
Attribute VB_Name = "frmReceiptsCounter"
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
'* | Number  | Modification Date |   Modified By         |   Name of function Variable   *
'* |---------|-------------------|-----------------------|-------------------------------*
'* |         |                   |                       |                               *
'* |         |                   |                       |                               *
'*=======================================================================================*
' Notes :-                                                                               '
'       cmdSearchAccountHead.Tag Keeps GroupID of AccountHead Type                       '
'       GroupID=1-> Cash   GroupID=2 ->  Bank                                            '
'----------------------------------------------------------------------------------------'
    Dim mVoucherID                  As Variant
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
    Dim vchName             As Variant
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
    Dim mRoundOffDecimalPlace As Boolean
    
    Dim mPTaxFormLoadFlag As Boolean
    Dim mRentSearchFormLoadedFlag As Boolean
    
    Private mRePrintFlag As Integer
    Private mvarDemandBasedFlag As Boolean
    
    Dim mZonal As Integer 'Sunil on 22-july-2011
    
    '------------------------------------------------------------
    '                   Added On 24/04/2009
    '------------------------------------------------------------
    Private mPermitType As Integer
    Private mBuildingType As Double
    Private mKMBRAccess As Integer
    Private mBuildingWard As Double
    Private mPoorHomeCess           As Boolean ' To check Whether PoorHome Cess Account Head is Included in Property Tax. Used in Property Tax Integration
    Dim lSoochikaFeildID As Variant
    Dim mSeatPrefix As String
    Dim mReceiptNo  As Double
    Dim mKMBRFlag As Boolean
    '-----------------------------------------------------------
    
    Dim mTimer                      As Integer ' Added for blinking Interrupted Request Status
    Dim mInterruptedModeFlag        As Boolean '
    Dim mInterruptEditMode          As Boolean
    Private mInterruptedModeSoochikaFlag    As Boolean
    Dim mdtDate                     As Date
    Private mAssesmentYearID        As Integer ' Added on 12-aug-2009 by cijith for Sanchaya Zonal Connectivity
    Private mDataBase               As Boolean ' True for DB_FinanceHO And False for DB_Finance Added By Sinoj
    
    
    'Added on 4.9.11 For Zonal integration
    Public mZoneDate                As Date  '' set from frmTransactionTypewiseDemandInbox Form
    '------------------------------------
    '-----Added On 7.Jul/2011 By Anisha---------
    '-----
    Private mDemandMode             As Integer 'To identify 1=Direct 2=Zonal/3=OutDoor /4=JSk Friend Collections
    Private mDemandTrDate           As Variant    'Other than direct Transactions In Acc Clerk's Login dtTransactiondate Should be mDemandTrDate
    '-------------------------------------------
    Private mSoochikaConnected      As Boolean
    Private lSoochikaFileID         As Variant 'Added by Akheel 09.02.11 for Soochika Unicode
    Public mFinewave                As Boolean ' Added On 15-Feb-10 Used To Update the Status of Fine wave
    Dim mGrandTotalValidityFlag     As Boolean ' To Check Grand Total and Grid Amount
    Public mReverseMode             As Boolean ' Added On 30/12/10 Used For Reverse Mode. This variable Will set From fareverse Appoval Form

    Public mPreviousYearMode As Integer
    Public mPreviousYearRequestID As Integer
    
    Public mWebExtractMode As Boolean           ' Added on 19 Oct 2017 to do Project related receipts 1- active
    Public mWebExtractDate As Date
    '-------------------------------------------------------------------------'
    ' INTERRUTED REGISTER MODE                                                '
    '-------------------------------------------------------------------------'
    Dim mInterruptedRegister            As Integer ' Added by Minu for IR Register
    Dim mInterruptedRegisterReceiptNo   As Variant
    Dim mInterruptedRegisterReceiptDate As Date
    Dim mIRBookID                       As Variant
    Dim mInterruptedRegisterID          As Long
    Dim mIRVoucherDate As Variant
    
    Public Function ConnectSoochika(ByRef mCn As ADODB.Connection) As Boolean
        Dim objdb As New clsDB
        If gbLinkWithSoochika = 2 Then
            ConnectSoochika = objdb.CreateNewConnection(mCn, enuSourceString.SoochikaUnicode)
        Else '1
            ConnectSoochika = objdb.CreateNewConnection(mCn, enuSourceString.SOOCHIKA)
        End If
    End Function
    
    Public Property Let DataBaseHO(mVal As Boolean)
        mDataBase = mVal
    End Property
    
    Public Sub CheckInterruptReceiptRequestStatus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If mWebExtractMode = True Then
            Exit Sub
        End If
        
        mStatus = ""
        mSql = "Select tnyStatus,dtReceiptDate, dtRequestDate From faInterruptedRequests"
        mSql = mSql + " Where numUserID =" & gbUserID
        mSql = mSql + " And intCounterID =" & gbCounterID
        mSql = mSql + " And intTypeID = 1"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            mdtDate = Rec!dtReceiptDate
            txtDate.Text = Format(mdtDate, "dd-mmm-yyyy")
        End If
        Rec.Close
        mCnn.Close
        If mStatus <> "" Then
            If mStatus = 1 Then
                lblInterruptedReceipt.Visible = False
                cmdSave.Enabled = False
                cmdNew.Enabled = False
                Timer1.Enabled = True
                mTimer = 0
                mInterruptedModeFlag = False
            End If
            If mStatus = 2 Then
                mInterruptedModeFlag = True
                cmdSave.Enabled = True
                cmdNew.Enabled = True
                Timer1.Enabled = True
                mTimer = 2
            End If
        Else
            lblInterruptedReceipt.Visible = False
            Timer1.Enabled = False
            mTimer = 0
            mInterruptedModeFlag = False
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
    
    Private Sub GroupCalc()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        
        objdb.SetConnection mCnn
        Rec.Open "Select isnull(sum(fltAmount),0) as GroupTotal From faVouchers Where tnyVoucherTypeID=10 and  intVoucherNo BetWeen " & val(lblFromReceiptNo.Caption) & "And " & val(lblToReceiptNo.Caption), mCnn
        lblGroupTotal.Caption = Format(Rec!GroupTotal, "0.00")
        Rec.Close
        mCnn.Close
    End Sub
    Public Sub DisplayTransactionWiseDetails(ByRef intTrTypeID As Integer)   'Added By Sunil
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mCnnFin As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim RecAddress As New ADODB.Recordset
        Dim mSql As String
        Dim arrInput As Variant
        Dim objTranType As New clsTransactionType
        Dim objAc As New clsAccounts
        Dim mCount As Long
        Dim mDemandID As Variant
        Dim objInstruments As New clsInstruments
        Dim arrIn As Variant
        arrInput = Array(mDemandID)
        If mDataBase Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO) = False Then
                MsgBox "Connection Saankhya H O not Present"
                Exit Sub
            End If
        Else
            objdb.SetConnection mCnn
        End If
        
        mSql = "Select faIDemandTbl.*, vchSectionName From faIDemandTbl Inner Join "
        mSql = mSql + " faSection On faSection.intSectionID = faIDemandTbl.intSectionID"
        mSql = mSql + " Where vchDemandNo = '" & txtDemandPrefix & "-" & txtDemandNo & "'"
        mSql = mSql + " AND tnyStatus = 0"
        
        'Note:-Changed on 5-Nov-2009
        
        mSql = " Select * From faIDemandTbl "
        mSql = mSql + " INNER JOIN faIDemandChild ON faIDemandChild.numDemandID = faIDemandTbl.numDemandID "
        mSql = mSql + " WHERE faIDemandChild.intTransactionTypeID = " & intTrTypeID
        mSql = mSql + " AND vchDemandNo = '" & txtDemandPrefix & "-" & txtDemandNo & "'"
        mSql = mSql + " AND faIDemandTbl.tnyStatus = 0"
                
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            On Error Resume Next
            mDemandID = Rec!numDemandID
            txtDemandNo.Tag = mDemandID
    
            objTranType.SetTransactionType Rec!intTransactionTypeID
            If objTranType.TransactionTypeID > 0 Then
                txtTransactionType.Text = frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 1) 'objTranType.TransactionType
                txtTransactionType.Tag = frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4) 'objTranType.TransactionTypeID
            End If
            
            '----------Added On 7-Jun-2011 By Anisha For Implementing Mode Of Transactions
            If IsNull(Rec!intDemandMode) Then
                If txtTransactionType.Tag = gbTransactionTypeZonalCollection Then
                    mDemandMode = 2
                ElseIf txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                    mDemandMode = 3
                ElseIf txtTransactionType.Tag = gbTransactionTypeFriendsCollection Then
                    mDemandMode = 4
                End If
            Else
                mDemandMode = Rec!intDemandMode
            End If
            
            If IsNull(Rec!dtTransactionDate) = False Then
                mDemandTrDate = Format((Rec!dtTransactionDate), "dd-mmm-yyyy")
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    txtDate.Text = Format(mDemandTrDate, "dd-mmm-yyyy")
                End If
            End If
            
            mSql = Rec!vchDemandNo
            txtDemandPrefix.Text = Token(mSql, "-")
            txtDemandNo.Text = mSql
            txtDescription.Text = Rec!vchRemarks
            
            If Not IsNull(Rec!vchAdminNote) Then
                lblAdminNoteCaption.Visible = True
                lblAdminNote.Visible = True
                lblAdminNote.Caption = Rec!vchAdminNote
            Else
                lblAdminNoteCaption.Visible = False
                lblAdminNote.Visible = False
            End If
            
            If Not (IsNull(Rec!intKeyID)) Then
                objAc.SetAccountID Rec!intKeyID
                txtAccountHead.Text = objAc.AccountHead
                txtAccountHead.Tag = objAc.AccountHeadID
            End If
            
            If Rec!intInstrumentTypeID <> gbInstrumentCash Then
                objInstruments.SetInstrumentType (Rec!intInstrumentTypeID)
                If objInstruments.InstrumentTypeID <> gbInstrumentCash Then
                    txtInstrument.Text = objInstruments.InstrumentType
                    txtInstrument.Tag = objInstruments.InstrumentTypeID
                    txtInstNo.Text = Rec!vchInstrumentNo
                    txtDated.Text = DdMmmYy(Rec!dtInstrumentDate)
                    txtBank.Text = Rec!vchDrawnFrom
                    txtPlace.Text = Rec!vchDrawnPlace
                    cmdSearchAccountHead.Tag = objAc.GroupID
                    Call txtInstrument_LostFocus
                Else
                    GoTo LB
                End If
            Else
LB:             txtAccountHead.Text = "Cash [ " & gbAcHeadCodeCash & " ]"
                txtAccountHead.Tag = gbAcHeadIDCash
                txtInstrument.Text = "Cash"
                txtInstrument.Tag = gbInstrumentCash
                txtInstNo.Text = ""
                txtDated.Text = ""
                txtBank.Text = ""
                txtPlace.Text = ""
            End If
            
            If Rec!intTransactionTypeID = gbTransactionTypeOutDoor Then
                lblOutDoorStaff(8).Caption = "&OutDoor Staff"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("snPDE_ODStaff", "chvEmployeeName", "numUserID", Rec!intKeyID2, SanchayaLite)
            ElseIf Rec!intTransactionTypeID = gbTransactionTypeZonalCollection Then
                lblOutDoorStaff(8).Caption = "Zone"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("GM_Zone", "chvZoneNameEnglish", "numZoneID", Rec!intKeyID2, DBMaster)
                'Note:- Modified For Zonal Integration on 5-Sep-2009
                cmbDZone.Tag = Rec!intKeyID
                cmbDZone.Text = "Main Office"
                cmbZone.Locked = True
            Else
                lblOutDoorStaff(8).Visible = False
                txtOutDoorStaff.Visible = False
                txtOutDoorStaff.Tag = ""
            End If
 
          ' ------------------------------------------------------------------------------------
                              'Transaction Type Wise Head  ----Added by sunil
          '------------------------------------------------------------------------------------
            Dim mSQLHO As String
            Dim RecHO As New ADODB.Recordset
          
            ' Set Rec = objDB.ExecuteSP("spGetHeadWiseDetails", arrIn, , , mCnn)
            mSQLHO = " Select faIDemandChild.intTransactionTypeID,faIDemandChild.intAccountHeadID as intAccountHeadID ,faIDemandChild.fltAmount from faIDemandTBL"
            mSQLHO = mSQLHO + " Inner JOIN faIDemandChild ON faIDemandChild.numDemandID=faIDemandTBL.numDemandID"
            'mSQLHO = mSQLHO + " where dtDemandDate='" & Format(CDate(frmTransactionTypeWiseDemandInbox.txtDate.Text), "dd/MMM/yyyy") & "'"
            'mSQLHO = mSQLHO + " And faIDemandChild.intTransactionTypeID = " & frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 4)
            
            mSQLHO = mSQLHO + "  Where faIDemandChild.intTransactionTypeID = " & intTrTypeID
            mSQLHO = mSQLHO + " AND vchDemandNo = '" & txtDemandPrefix & "-" & txtDemandNo & "'"
            mSQLHO = mSQLHO + " AND faIDemandTbl.tnyStatus = 0"
            
         
            RecHO.Open mSQLHO, mCnn, adOpenKeyset, adLockOptimistic
             If Not (RecHO.BOF And RecHO.EOF) Then
             While Not RecHO.EOF
                mCount = mCount + 1
                vsGrid.Row = mCount
                objAc.SetAccountID RecHO!intAccountHeadID
                If objAc.AccountHeadID > 0 Then
                vsGrid.TextMatrix(mCount, 6) = objAc.AccountHeadID
                vsGrid.TextMatrix(mCount, 0) = objAc.AccountCode
                vsGrid.TextMatrix(mCount, 1) = objAc.AccountHead
                If InStr(1, objAc.AccountHead, "Arrear") Then
                        vsGrid.TextMatrix(mCount, 4) = Format(RecHO!fltAmount, "0.00")
                        vsGrid.TextMatrix(mCount, 5) = ""
                Else
                       vsGrid.TextMatrix(mCount, 5) = Format(RecHO!fltAmount, "0.00")
                       vsGrid.TextMatrix(mCount, 4) = ""
                End If
                '   vsGrid.TextMatrix(mCount, 2) = Rec!intYearID
                '  vsGrid.TextMatrix(mCount, 3) = Rec!tnyPeriodID
                '   If Rec1!tnyArrearFlag = 1 Then
                '      vsGrid.TextMatrix(mCount, 4) = Rec!fltAmount
                '      vsGrid.TextMatrix(mCount, 5) = ""
                '   Else
                '    vsGrid.TextMatrix(mCount, 4) = ""
               ' vsGrid.TextMatrix(mCount, 5) = RecHO!fltAmount
                '   End If
                vsGrid.TextMatrix(mCount, 10) = mDemandID
                   '   End If
                    Call ValuesForHiddenColumns(vsGrid.Row)
                Else
                        MsgBox "Unknown AccountHead selected or Invalid Demand !", vbInformation
                        FormInitialize
                        Exit Sub
                End If
                    RecHO.MoveNext
            Wend
                Call Calculate
            End If

            
            '-------------------------------------------------------------------'
            ' A d d r e s s   D e t a i l s                         '
            '-------------------------------------------------------------------'
            mSql = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
            RecAddress.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
            If Not (RecAddress.BOF And RecAddress.EOF) Then
                fraParty.Visible = False
                Call ShowFrames(1)
                On Error Resume Next
                txtName.Text = RecAddress!vchName
                txtInit1.Text = RecAddress!vchInit1
                txtInit2.Text = RecAddress!vchInit2
                txtInit3.Text = RecAddress!vchInit3
                txtInit4.Text = RecAddress!vchInit4
                txtOutDoorStaff = txtName.Text
                
                txtHouse.Text = RecAddress!vchHouseName
                txtStreet.Text = RecAddress!vchStreet
                txtLocalPlace.Text = RecAddress!vchLocalPlace
                txtMainPlace.Text = RecAddress!vchMainPlace
                txtPost.Text = RecAddress!vchPost
                txtPin.Text = RecAddress!vchPin
                txtPhone.Text = RecAddress!vchPhone
                txtWardNo.Text = RecAddress!intWardNo
                txtDoorNo1.Text = RecAddress!intDoorNo
                txtDoorNo2.Text = RecAddress!vchDoorNo2
                On Error GoTo 0
            End If
        End If
        Rec.Close
        vsGrid.Editable = flexEDNone
    End Sub
            
    Public Sub DisplayDemandDetails()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim RecAddress As New ADODB.Recordset
        Dim mSql As String
        Dim arrInput As Variant
        Dim objTranType As New clsTransactionType
        Dim objAc As New clsAccounts
        Dim mCount As Long
        Dim mDemandID As Variant
        Dim objInstruments As New clsInstruments
        Dim arrIn As Variant  'Added  by Sunil babu
        Dim RecReq As New ADODB.Recordset
        
        arrInput = Array(mDemandID)
        If mDataBase Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO) = False Then
                MsgBox "Connection Saankhya H O not Present"
                Exit Sub
            End If
        Else
            objdb.SetConnection mCnn
        End If
        mSql = "        Select faIDemandTbl.*, vchSectionName From faIDemandTbl Inner Join "
        mSql = mSql + " faSection On faSection.intSectionID = faIDemandTbl.intSectionID"
        mSql = mSql + " Where vchDemandNo = '" & txtDemandPrefix & "-" & txtDemandNo & "'"
        mSql = mSql + " AND tnyStatus = 0"
        
        'Note:-Changed on 5-Nov-2009
        mSql = "Select * From faIDemandTbl "
        mSql = mSql + " Where vchDemandNo = '" & txtDemandPrefix & "-" & txtDemandNo & "'"
        mSql = mSql + " AND tnyStatus = 0"
        
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
            On Error Resume Next
            
            If Rec!intFinancialYearID <> gbFinancialYearID And Rec!intFinancialYearID = (gbFinancialYearID - 1) Then
                mSql = "SELECT * FROM faPendingTaskRequest WHERE numDemandID = " & Rec!numDemandID
                RecReq.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                If Not (RecReq.BOF And RecReq.EOF) Then
                    mPreviousYearMode = 1
                    txtDate.Text = DdMmmYy(RecReq!dtTransactionDate)
                    
                    GoTo STEP1:
                End If
                RecReq.Close
            End If
            
            If IsDate(Rec!dtExpiryDate) Then
                If Rec!dtExpiryDate < gbTransactionDate Then
                    mSql = "This demand dated " & DdMmmYy(Rec!dtExpiryDate) & " is " & vbCrLf
                    mSql = mSql & " not valid any more"
                    MsgBox mSql, vbInformation
                    Call FormInitialize
                    Exit Sub
                End If
            Else
                MsgBox "Demand Validity Date not specified!", vbInformation
                Exit Sub
            End If
            
STEP1:
            
            mDemandID = Rec!numDemandID
            txtDemandNo.Tag = mDemandID
            
            objTranType.SetTransactionType Rec!intTransactionTypeID
            If objTranType.TransactionTypeID > 0 Then
                txtTransactionType.Text = objTranType.TransactionType
                txtTransactionType.Tag = objTranType.TransactionTypeID
            End If
            
            '
            'Added On 7-Jun-2011 By Anisha For Implementing Mode Of Transactions
            '
            If IsNull(Rec!intDemandMode) Then
                If txtTransactionType.Tag = gbTransactionTypeZonalCollection Then
                    mDemandMode = 2
                ElseIf txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                     mDemandMode = 3
                ElseIf txtTransactionType.Tag = gbTransactionTypeFriendsCollection Then
                     mDemandMode = 4
                Else
                    mDemandMode = 1
                End If
            Else
                mDemandMode = Rec!intDemandMode
            End If
            
            If IsNull(Rec!dtTransactionDate) = False Then
                mDemandTrDate = Format((Rec!dtTransactionDate), "dd-mmm-yyyy")
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    If mPreviousYearMode = 0 Then
                        If CDate(mDemandTrDate) >= gbStartingDate And CDate(mDemandTrDate) <= gbEndingDate Then
                        txtDate.Text = Format(mDemandTrDate, "dd-mmm-yyyy")
                        Else
                            MsgBox "Transaction date entered is in wrong FinancailYear"
                            Exit Sub
                        End If
                    End If
                End If
            Else
                mDemandTrDate = Null
            End If
            '-------------------------------------------------
            'Added On 12/03/10 By Anisha
            'To Avoid Other than Cash Transaction at Jsk/Cash Group
            If gbCounterSectionID = gbJSKSectionID Then
                If Not (Rec!intInstrumentTypeID = gbInstrumentCash Or Rec!intInstrumentTypeID = gbInstrumentCheque _
                 Or Rec!intInstrumentTypeID = 2 Or Rec!intInstrumentTypeID = 3 Or Rec!intInstrumentTypeID = 4 Or Rec!intInstrumentTypeID = 8 Or Rec!intInstrumentTypeID = 6) Then
                     MsgBox " You are not Allowed To Use the Instrument In this Section", vbApplicationModal
                    Exit Sub
                End If
                If val(txtTransactionType.Tag) = gbTransactionTypeBFundSSSFund Or _
                    val(txtTransactionType.Tag) = gbTransactionTypeMoneyOrderReturns Then
                    If (CDate(Rec!dtDemandDate) <> CDate(txtDate.Text)) Then
                        MsgBox "This Demand cannot be taken through this Counter", vbInformation
                        Exit Sub
                    End If
                End If
            ElseIf (gbCounterSectionID <> gbJSKSectionID And (CDate(gbOnlinedate) > mDemandTrDate) And (gbLBPanchayat = 1 Or gbLBType = 4 Or gbLBID = 221)) Then
                '''Skip Instrument Validation
                
            Else
                If Not (Rec!intInstrumentTypeID = 6 Or Rec!intInstrumentTypeID = 7 Or Rec!intInstrumentTypeID = 9 Or Rec!intInstrumentTypeID = 10 Or Rec!intInstrumentTypeID = 11) Then
                    'Added By Vinod :- For B-Fund And Money Order Returns
                    '       74 - Money Order Returns ;  112- B Fund-State Sponsored Scheme Funds
                    If Not (txtTransactionType.Tag = 74 Or txtTransactionType.Tag = 112) Then
                        If mPreviousYearMode <> 1 Then
                            MsgBox " You are not Allowed To Use this Instrument Type ", vbApplicationModal
                            Exit Sub
                        End If
                    Else
                        If (CDate(Rec!dtDemandDate) <> CDate(txtDate.Text)) Then
                            MsgBox "Please change the Transaction Date ( Demand Date is " & CDate(Rec!dtDemandDate) & ")", vbInformation
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            mSql = Rec!vchDemandNo
            txtDemandPrefix.Text = Token(mSql, "-")
            txtDemandNo.Text = mSql
            txtDescription.Text = Rec!vchRemarks
            
            If Not IsNull(Rec!vchAdminNote) Then
                lblAdminNoteCaption.Visible = True
                lblAdminNote.Visible = True
                lblAdminNote.Caption = Rec!vchAdminNote
            Else
                lblAdminNoteCaption.Visible = False
                lblAdminNote.Visible = False
            End If
            
            If Not (IsNull(Rec!intKeyID)) Then
                objAc.SetAccountID Rec!intKeyID
                txtAccountHead.Text = objAc.AccountHead
                txtAccountHead.Tag = objAc.AccountHeadID
            End If
            
            If Rec!intInstrumentTypeID <> gbInstrumentCash Then
                objInstruments.SetInstrumentType (Rec!intInstrumentTypeID)
                If objInstruments.InstrumentTypeID <> gbInstrumentCash Then
                    txtInstrument.Text = objInstruments.InstrumentType
                    txtInstrument.Tag = objInstruments.InstrumentTypeID
                    txtInstNo.Text = Rec!vchInstrumentNo
                    txtDated.Text = DdMmmYy(Rec!dtInstrumentDate)
                    txtBank.Text = Rec!vchDrawnFrom
                    txtPlace.Text = Rec!vchDrawnPlace
                    cmdSearchAccountHead.Tag = objAc.GroupID
                    Call txtInstrument_LostFocus
                Else
                    GoTo LB
                End If
           ElseIf txtTransactionType.Tag = 119 Or txtTransactionType.Tag = 120 Or txtTransactionType.Tag = 121 Or txtTransactionType.Tag = 122 Or txtTransactionType.Tag = 123 Then
                txtInstrument.Text = ""
                txtInstrument.Tag = ""
                txtInstNo.Text = ""
                txtDated.Text = ""
                txtBank.Text = ""
                txtPlace.Text = ""
            
            Else
LB:             txtAccountHead.Text = "Cash [ " & gbAcHeadCodeCash & " ]"
                txtAccountHead.Tag = gbAcHeadIDCash
                txtInstrument.Text = "Cash"
                txtInstrument.Tag = gbInstrumentCash
                txtInstNo.Text = ""
                txtDated.Text = ""
                txtBank.Text = ""
                txtPlace.Text = ""
            End If
            
           
            
            If Rec!intTransactionTypeID = gbTransactionTypeOutDoor Then
                lblOutDoorStaff(8).Caption = "&OutDoor Staff"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("snPDE_ODStaff", "chvEmployeeName", "numUserID", Rec!intKeyID2, SanchayaLite)
            ElseIf Rec!intTransactionTypeID = gbTransactionTypeZonalCollection Then
                lblOutDoorStaff(8).Caption = "Zone"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("GM_Zone", "chvZoneNameEnglish", "numZoneID", Rec!intKeyID2, DBMaster)
                'Note:- Modified For Zonal Integration on 5-Sep-2009
                cmbDZone.Tag = Rec!intKeyID
                cmbDZone.Text = txtOutDoorStaff.Text
            Else
                lblOutDoorStaff(8).Visible = False
                txtOutDoorStaff.Visible = False
                txtOutDoorStaff.Tag = ""
            End If
            '-----------------------------------------------------
            'Added On 19/DEC/2015 By Anisha
            'To DISPLAY OUT DOOR STAFF
            If mDemandMode = 2 Then
                lblOutDoorStaff(8).Caption = "Zone"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("GM_Zone", "chvZoneNameEnglish", "numZoneID", Rec!intKeyID2, DBMaster)
                'Note:- Modified For Zonal Integration on 5-Sep-2009
                cmbDZone.Tag = Rec!intKeyID
                cmbDZone.Text = txtOutDoorStaff.Text
            ElseIf mDemandMode = 3 Then
                lblOutDoorStaff(8).Caption = "&OutDoor Staff"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Visible = True
                txtOutDoorStaff.Tag = Rec!intKeyID2
                txtOutDoorStaff.Text = FindMaster("snPDE_ODStaff", "chvEmployeeName", "numUserID", Rec!intKeyID2, SanchayaLite)

            End If
            '-----------------------------------------------------
            
            '-------------------------------------------------------------------'
            ' A c c o u n t   H e a d s   S e l e c t e d                       '
            '-------------------------------------------------------------------'
            mSql = "Select * From faIDemandChild Where numDemandID = " & mDemandID
            RecChild.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
            If Not (RecChild.BOF And RecChild.EOF) Then
                While Not RecChild.EOF
                    mCount = mCount + 1
                    vsGrid.Row = mCount
                    objAc.SetAccountID RecChild!intAccountHeadID
                    If objAc.AccountHeadID > 0 Then
                        vsGrid.TextMatrix(mCount, 6) = objAc.AccountHeadID
                        vsGrid.TextMatrix(mCount, 0) = objAc.AccountCode
                        vsGrid.TextMatrix(mCount, 1) = objAc.AccountHead
   
                        If objAc.AccountCode = gbAcHeadCodeAdvanceDandO Then
                            Dim mLoop As Integer
                            Dim mItem As String
                            mItem = "#0; "
                            For mLoop = gbFinancialYearID + 5 To 1970 Step -1
                                mItem = mItem & "|#" & mLoop & ";" & CStr(mLoop) & "-" & CStr(mLoop + 1)
                            Next
                            vsGrid.ColComboList(2) = mItem
                        End If
                        
                        vsGrid.TextMatrix(mCount, 2) = RecChild!intYearID
                        vsGrid.TextMatrix(mCount, 3) = RecChild!tnyPeriodID
                        If RecChild!tnyArrearFlag = 1 Then
                            vsGrid.TextMatrix(mCount, 4) = RecChild!fltAmount
                            vsGrid.TextMatrix(mCount, 5) = ""
                        Else
                            vsGrid.TextMatrix(mCount, 4) = ""
                            vsGrid.TextMatrix(mCount, 5) = RecChild!fltAmount
                        End If
                        Call ValuesForHiddenColumns(vsGrid.Row)
                    Else
                        MsgBox "Unknown AccountHead selected or Invalid Demand !", vbInformation
                        FormInitialize
                        Exit Sub
                    End If
                    RecChild.MoveNext
                Wend
                Call Calculate
            End If
            
            '-------------------------------------------------------------------'
            ' A d d r e s s   D e t a i l s                                     '
            '-------------------------------------------------------------------'
            mSql = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
            RecAddress.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
            If Not (RecAddress.BOF And RecAddress.EOF) Then
                fraParty.Visible = False
                Call ShowFrames(1)
                On Error Resume Next
                txtName.Text = RecAddress!vchName
                txtInit1.Text = RecAddress!vchInit1
                txtInit2.Text = RecAddress!vchInit2
                txtInit3.Text = RecAddress!vchInit3
                txtInit4.Text = RecAddress!vchInit4
                txtOutDoorStaff = txtName.Text
                
                txtHouse.Text = RecAddress!vchHouseName
                txtStreet.Text = RecAddress!vchStreet
                txtLocalPlace.Text = RecAddress!vchLocalPlace
                txtMainPlace.Text = RecAddress!vchMainPlace
                txtPost.Text = RecAddress!vchPost
                txtPin.Text = RecAddress!vchPin
                txtPhone.Text = RecAddress!vchPhone
                txtWardNo.Text = RecAddress!intWardNo
                txtDoorNo1.Text = RecAddress!intDoorNo
                txtDoorNo2.Text = RecAddress!vchDoorNo2
                On Error GoTo 0
            End If
            If gbSeatGroupID = gbSeatGroupCashier Then
                If mDemandMode <> 1 Then
                    If mDemandTrDate <> txtDate.Text Then
                        MsgBox "Demand Transaction Date is :" & mDemandTrDate & vbNewLine & "Only Accounts Clerk Can Post Receipt On" & mDemandTrDate & "", vbInformation
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "Demand Does not Exists", vbApplicationModal
            Exit Sub
        End If
        Rec.Close
        vsGrid.Editable = flexEDNone
        
    End Sub
    
        Public Sub DisplayReceiptDetails(mVoucherID As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        Call FormInitialize
'        mSql = "Select numSeatID From faVouchers Where intVoucherNo=" & mVoucherNo
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
'        End If
'        Rec.Close
'        If mSeatID <> "" Then
        'If gbUserTypeID = 1 Then
        mSql = "Select *,faVouchers.intVoucherNo As VoucherNo,vchDoorNoP3 From faVouchers"
        mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
        mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
        mSql = mSql + " Inner Join faTransactions On faVouchers.intVoucherID = faTransactions.intVoucherID"
        mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        mSql = mSql + " Inner Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
        mSql = mSql + " Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
        'mSQL = mSQL + " Or  faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID "
        mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
        mSql = mSql + " Where faVouchers.intVoucherID=" & mVoucherID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtReceiptNo.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            txtReceiptNo.Text = IIf(IsNull(Rec!VoucherNo), "", Rec!VoucherNo)
                
            txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtDate.Text = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
            txtDate.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
            If mInterruptEditMode Then
                mIRVoucherDate = txtDate.Text
            Else
                mIRVoucherDate = Null
            End If
            txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead) & " [ " & IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode) & " ]"
            txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
            txtInstNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            txtDated.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
            txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
            
            If IsNull(Rec!chvZoneNameEnglish) = False Then
                cmbDZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
            End If
            txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
            If Not IsNull(Rec!vchDoorNoP3) Then
                txtIntruptNoSuffix.Visible = True
                txtIntruptNoSuffix.Enabled = False
                txtIntruptNoSuffix.Text = IIf(IsNull(Rec!vchDoorNoP3), "", Rec!vchDoorNoP3)
            End If
            txtRefNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            
            txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            txtInit1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
            txtInit2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
            txtInit3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
            txtInit4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
            txtHouse.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            txtStreet.Text = IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
            txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            txtPost.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            txtPin.Text = IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            txtDescription.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            mSqlAccHeads = "Select * From faVoucherChild"
            mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
            mSqlAccHeads = mSqlAccHeads + " Left Join faPeriodicity On faVoucherChild.tnyPeriodID = faPeriodicity.intPeriodicityID"
            mSqlAccHeads = mSqlAccHeads + " Where intVoucherID=" & txtReceiptNo.Tag
            mSqlAccHeads = mSqlAccHeads + " Order By tnyArrearFlag Desc"
            RecAccHeads.Open mSqlAccHeads, mCnn
            mRowCount = 1
             While Not Rec.EOF
                While Not RecAccHeads.EOF
                    vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchPeriodicity), "", RecAccHeads!vchPeriodicity)
                    'vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!intPeriodicityID), "", RecAccHeads!intPeriodicityID)
                    vsGrid.Cell(flexcpText, mRowCount, 2) = val(RecAccHeads!intYearID) & "-" & val(RecAccHeads!intYearID) + 1
                    mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), "", RecAccHeads!tnyArrearFlag)
                    If mArrearFlag = 0 Then
                        vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    End If
                    If mArrearFlag = 1 Then
                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    End If
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCount = mRowCount + 1
                    RecAccHeads.MoveNext
                Wend
                Rec.MoveNext
            Wend
            RecAccHeads.Close
            Call Calculate
        End If
        Rec.Close
        
        If mInterruptEditMode Then
                cmdNew.Enabled = False
        End If
        Debug.Print "Display Receipt Details Called"
End Sub
    
        Public Sub DisplayReceiptDetailsIREdit(mVoucherID As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        Call FormInitialize
'        mSql = "Select numSeatID From faVouchers Where intVoucherNo=" & mVoucherNo
'        Rec.Open mSql, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
'        End If
'        Rec.Close
'        If mSeatID <> "" Then
        'If gbUserTypeID = 1 Then
        mSql = "Select *,faVouchers.intVoucherNo As VoucherNo,vchDoorNoP3 From faVouchers"
        mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
        mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
        mSql = mSql + " Inner Join faTransactions On faVouchers.intVoucherID = faTransactions.intVoucherID"
        mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        mSql = mSql + " Inner Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
        mSql = mSql + " Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
        'mSQL = mSQL + " Or  faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID "
        mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
        mSql = mSql + " Where faVouchers.intVoucherID=" & mVoucherID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtReceiptNo.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            txtReceiptNo.Text = IIf(IsNull(Rec!VoucherNo), "", Rec!VoucherNo)
                
            txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
            txtDate.Text = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
            txtDate.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
            If mInterruptEditMode Then
                mIRVoucherDate = txtDate.Text
            Else
                mIRVoucherDate = Null
            End If
            txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead) & " [ " & IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode) & " ]"
            txtAccountHead.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
            txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
            txtInstrument.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
            txtInstNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            txtDated.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
            txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
            
            If IsNull(Rec!chvZoneNameEnglish) = False Then
                cmbDZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
            End If
            txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
            If Not IsNull(Rec!vchDoorNoP3) Then
                txtIntruptNoSuffix.Visible = True
                txtIntruptNoSuffix.Enabled = False
                txtIntruptNoSuffix.Text = IIf(IsNull(Rec!vchDoorNoP3), "", Rec!vchDoorNoP3)
            End If
            txtRefNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            
            txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            txtInit1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
            txtInit2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
            txtInit3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
            txtInit4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
            txtHouse.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            txtStreet.Text = IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
            txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            txtPost.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            txtPin.Text = IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            txtDescription.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            mSqlAccHeads = "Select * From faVoucherChild"
            mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
            mSqlAccHeads = mSqlAccHeads + " Left Join faPeriodicity On faVoucherChild.tnyPeriodID = faPeriodicity.intPeriodicityID"
            mSqlAccHeads = mSqlAccHeads + " Where intVoucherID=" & txtReceiptNo.Tag
            mSqlAccHeads = mSqlAccHeads + " Order By tnyArrearFlag Desc"
            RecAccHeads.Open mSqlAccHeads, mCnn
            mRowCount = 1
             While Not Rec.EOF
                While Not RecAccHeads.EOF
                    vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                    vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchPeriodicity), "", RecAccHeads!vchPeriodicity)
                    'vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!intPeriodicityID), "", RecAccHeads!intPeriodicityID)
                    vsGrid.Cell(flexcpText, mRowCount, 2) = val(RecAccHeads!intYearID) & "-" & val(RecAccHeads!intYearID) + 1
                    mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), "", RecAccHeads!tnyArrearFlag)
                    If mArrearFlag = 0 Then
                        vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    End If
                    If mArrearFlag = 1 Then
                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    End If
                    vsGrid.Rows = vsGrid.Rows + 1
                    mRowCount = mRowCount + 1
                    RecAccHeads.MoveNext
                Wend
                Rec.MoveNext
            Wend
            RecAccHeads.Close
            Call Calculate
        End If
        Rec.Close
        
        If mInterruptEditMode Then
                cmdNew.Enabled = False
        End If
        Debug.Print "Display Receipt Details Called"
End Sub
Private Sub PrintReceipt_ForNewFormat_ModifiedByMinu(intVoucherID As Double)
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        Open gbFileName For Output As #gbFileNO
'        Print #gbFileNO, Chr$(27) + Chr$(80)
'        Print #gbFileNO, String(136, "-")
'        Close #gbFileNO
'        Shell "Print " & gbFileName
'------------------------------------------------------------------------------------------------------------'
'-----------------------------------------Printing in 17 CPI-------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String

        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        'FileInitialize
''''        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
''''        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
''''        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Left Join faPeriodicity On  faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
''''        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
''''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

''''''        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
''''''            If Rec.RecordCount > 9 Then
''''''                Rec.Close
''''''                Call PrintSummaryReceiptPTax(intVoucherID)
''''''                Exit Sub
''''''            End If
''''''        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        Print #gbFileNO, Tab(3); gbBold; gbDoubleWidth; "RECEIPT"; Tab(31); gbLBName; " Panchayat"; gbDoubleWidthOff
'        Select Case Rec!intInstrumentTypeID
'        Case Is = 1
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
'        Case Is = 4
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
'            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Is = 5
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
'            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Else
'            Print #gbFileNO,
'        End Select

        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            'Print #gbFileNO, ; gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(65); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff;
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
            Else
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If

            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4

            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff

            'Changed for Sujith by Aiby - 24-Mar-2009

'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                'Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                'Print #gbFileNO,
            End Select
            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            If Not (IsNull(Rec!vchRefNo)) Then
                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
            Else
                Print #gbFileNO,
            End If
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 47 Then
'                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 46)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
            'Print #gbFileNO,

            ' Line 18 Next
            
            
            
            Dim RecPTAX         As New ADODB.Recordset
            Dim mStartingYear   As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear     As Integer
            Dim mEndingPeriod   As Integer
            Dim mNarration      As String
            
            mStartingYear = 2100
            
            
            
            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
                mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                mSql = mSql + " Where intVoucherID =" & intVoucherID '& Rec!intVoucherID
                mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
                mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
                RecPTAX.Open mSql, mCnn
                While Not RecPTAX.EOF
                    mLoop = mLoop + 1
                    Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
                    Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!Amount, "0.00"), 9);
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
                    Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!Amount, "0.00"), 9)
                    RecPTAX.MoveNext
                Wend
                RecPTAX.Close
                While Not Rec.EOF
                    If mStartingYear > Rec!intYearID Then
                        mStartingYear = Rec!intYearID
                        mStartingPeriod = Rec!tnyPeriodID
                    End If
                    If mEndingYear < Rec!intYearID Then
                        mEndingYear = Rec!intYearID
                    End If
                    mEndingPeriod = Rec!tnyPeriodID
                    Rec.MoveNext
                Wend
                'Rec.Close
                Rec.MoveFirst
                Print #gbFileNO,
                mLoop = mLoop + 1
                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
                Print #gbFileNO, mNarration; Tab(54); mNarration
                mLoop = mLoop + 1
                
                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                If mStartingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                ElseIf mStartingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                Else
                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                End If
                
                If mEndingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf )"
                ElseIf mEndingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf )"
                Else
                    mNarration = mNarration & ")"
                End If
                mLoop = mLoop + 1
                Print #gbFileNO, mNarration; Tab(52); mNarration
            Else
'               GoTo LB ' To print Property Tax containing less than 9 rows
'           End If
'            Else
            
'LB:
                mLoop = 0
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
                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
    
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    Else
                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    End If
    
                    '==================================================================='
                    ' Receipt Area
                    '==================================================================='
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(Rec!vchAlias, 46);
                    If Not IsNull(Rec!intYearID) Then
                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                    Else
                        mstrYear = ""
                    End If
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    Else
                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    End If
                    Rec.MoveNext
                Wend
            End If
            Rec.MoveFirst

            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
            Print #gbFileNO, Tab(54); Left(mRupees, 75)

            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)

            'Print #gbFileNO,'Commented By Vinod
            Dim mInward As String
            mInward = ""
            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
            Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
'            Else
'                Print #gbFileNO,
            End If
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 46 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
'            Else
'                Print #gbFileNO,
            End If
            
            
'             objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                Print #gbFileNO, Tab(30); objCounter.CounterNo;
'                Print #gbFileNO, Tab(67); objCounter.CounterNo & " : " & objCounter.CounterDescription
'            End If
'            objUser.SetUser (Rec!intUserID)
'            If objUser.UserID > -1 Then
'                Print #gbFileNO, Tab(27); objUser.UserName;
'                Print #gbFileNO, Tab(67); objUser.UserName
'            End If
            
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName;
                    Print #gbFileNO, Tab(66); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(93); objUser.UserName
                End If
            End If
                

            'Print #gbFileNO,
        End If

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
finishprinting:
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName
    
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
        txtPayee.Text = vchName
        txtHouse.Text = vchHouseName
        txtAddress.Text = vchStreetName & Chr(13)
        txtAddress.Text = txtAddress.Text & vchMainPlace & Chr(13)
        txtAddress.Text = txtAddress.Text & vchPostOffice & Chr(13)
        txtAddress.Text = txtAddress.Text & vchDistrict & " - " & vchPinNumber
    End Sub
    Private Sub ShowFrames(mIndex As Long)
        Select Case mIndex
            Case 2
                fraDemandDetails.Visible = False
                fraSubLedger.Visible = True
                fraParty.Visible = False
            Case 3
                fraDemandDetails.Visible = False
                fraSubLedger.Visible = False
                fraParty.Visible = True
            Case Else
                fraDemandDetails.Visible = True
                fraSubLedger.Visible = False
                fraParty.Visible = False
        End Select
    End Sub
    
    
    '------------------------------------------------------------'
    '                   Added On 24/04/2009                      '
    '           By Cijith Sreedharan For KMBR Integration        '
    '------------------------------------------------------------'
    Private Function SaveSoochika(ByRef mCnnSoochika As ADODB.Connection) As Double
        On Error GoTo err:
        Dim mVarrIn As Variant
        Dim mVarrOut As Variant
        Dim ForwardTo As Double
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        
        ReDim mVarrIn(41)
        Dim lSoochikaCurrentNo As Double
        Dim PermitType As Integer
        
        mVarrIn(0) = 0 'FldCurrentNo.
        mVarrIn(1) = gbTransactionDate 'FldDateOfReceipt.
        mVarrIn(2) = Trim(txtName.Text) 'FldSenderName.
        mVarrIn(3) = mBuildingWard  'Trim(txtWardNo.Text) 'FldWardNo.
        mVarrIn(4) = CStr(txtDoorNo2.Text) 'CStr(txtDoorNo1.Text) + CStr(txtDoorNo2.Text) 'FldHouseNo
        mVarrIn(5) = Trim(txtMainPlace.Text) 'FldLocality
        mVarrIn(6) = val(txtMainPlace.Tag)   '  cboDistrict.itemData(cboDistrict.ListIndex) 'FldDistrict
        mVarrIn(7) = gbSeatID 'Val(gbUserID) 'bntCurrUserId.
        ForwardTo = CDbl(mSeatPrefix + CStr(cmbSeat.ItemData(cmbSeat.ListIndex)))    'CDbl(Trim(mSeatPrefix + CStr(cmbSeat.itemData(cmbSeat.ListIndex))))
        mVarrIn(8) = ForwardTo
        mVarrIn(9) = 1 'intInwardType.
        mVarrIn(10) = 5 'FldPriority
        mVarrIn(11) = gbTransactionDate  'dtmForwardDate
        mVarrIn(12) = "Application for Build Prmit"  'FldRemarks
        mVarrIn(13) = Null 'intAttachmentType
        mVarrIn(14) = Null 'FldManualSummary
        mVarrIn(15) = Null 'FldElectronicsSummary
        mVarrIn(16) = gbDeptID 'intDept
        mVarrIn(17) = Null 'FlgCourFeeStamp
        mVarrIn(18) = Null 'intManualPage
        mVarrIn(19) = Null 'FldOutsideNo
        mVarrIn(20) = Null 'FldRefDate
        mVarrIn(21) = Null 'intRegPost
        mVarrIn(22) = Null 'bitInstflg
        mVarrIn(23) = Null 'fldInstName
        mVarrIn(24) = Null 'fldDesign
        mVarrIn(25) = Trim(txtPost.Text)      'cmbPostOffice.List(cmbPostOffice.ListIndex) 'FldPostOffice
        mVarrIn(26) = Trim(txtPin.Text) 'FldPin
        mVarrIn(27) = Null 'FldEmail
        mVarrIn(28) = txtPhone.Text 'FldPhone
        mVarrIn(29) = Null 'fldReglttoWhom
        mVarrIn(30) = Null 'fldReglttoDesign
        mVarrIn(31) = Null 'fldRegltpoNo
        mVarrIn(32) = Null 'sessionID
        mVarrIn(33) = Null 'intBillRecFlg
        mVarrIn(34) = Null 'intInsideLBFlg
        mVarrIn(35) = txtHouse.Text 'FldHouseName
        mVarrIn(36) = Null 'intCertAddrFlg
        mVarrIn(37) = Null 'intGender
        mVarrIn(38) = val(txtDoorNo1.Text) 'intDoorNo
        mVarrIn(39) = 0 'InwardFlg
        mVarrIn(40) = 117 'Suit
        If mPermitType = 0 Then
            mVarrIn(41) = 292 'Subject Gereral Permit
        Else
            mVarrIn(41) = 296 'Subject OneDay
        End If
        
        objdb.ExecuteSP "spSaveCorpOfficeView_KMBR", mVarrIn, mVarrOut, , mCnnSoochika, adCmdStoredProc
        If IsArray(mVarrOut) Then
           lSoochikaFeildID = mVarrOut(0, 0)
        End If
        
        Set Rec = objdb.ExecuteSP("SELECT FldCurrentNo From TblTappalDetails WHERE FldFileId = " & lSoochikaFeildID, , mVarrOut, , mCnnSoochika, adCmdText)
        If IsArray(mVarrOut) Then
           lSoochikaCurrentNo = mVarrOut(0, 0)
        End If
        SaveSoochika = lSoochikaCurrentNo
        Exit Function
err:
    MsgBox "Saankhya Error Handler: " & Error$
    SaveSoochika = -1
    End Function
    
    Private Function SaveSoochikaInwardDetails(mCnn As ADODB.Connection)
    Dim arrIn As Variant
    Dim Rec As New ADODB.Recordset
    Dim arrOut As Variant
    Dim objdb As New clsDB
    ReDim arrIn(36)
    arrIn(0) = 1 'cmbCorrespondance.ItemData(cmbCorrespondance.ListIndex)
    arrIn(1) = 5 'cmbPriority.ItemData(cmbPriority.ListIndex)
    'If (chkInstitution.value = 1) Then
    '    arrIn(2) = 1
    '    arrIn(3) = "" 'txtInstitutionName.Text
    '    arrIn(4) = "" 'txtInstitutionDesg.Text
    'Else
        arrIn(2) = Null
        arrIn(3) = Null
        arrIn(4) = Null
    'End If
    arrIn(5) = 1 'cmbGender.ItemData(cmbGender.ListIndex)
    arrIn(6) = txtName.Text
    arrIn(7) = txtHouse.Text '  txtHouseName.Text
    arrIn(8) = txtWardNo.Text
    arrIn(9) = txtDoorNo1.Text
    arrIn(10) = txtDoorNo2.Text
    arrIn(11) = txtMainPlace.Text
    If (txtLocalPlace.Text = "") Then
        arrIn(12) = txtMainPlace.Text
    Else
        arrIn(12) = txtLocalPlace.Text
    End If
    arrIn(13) = txtPost.Text
    arrIn(14) = txtPin.Text
    arrIn(15) = val(txtMainPlace.Tag)  'cmbDistrict.ItemData(cmbDistrict.ListIndex)
    arrIn(16) = 32 'cmbState.ItemData(cmbState.ListIndex)
    arrIn(17) = txtPhone.Text  'txtContactNo.Text
    arrIn(18) = "" 'txtContactMail.Text
    arrIn(19) = 292 ' txtSubID.Text
    arrIn(20) = Null
    arrIn(21) = "Application for Building Permit" 'txtSubject.Text
    arrIn(22) = gbTransactionDate  'dtpDeliveryDate.value
    arrIn(23) = 115 'gbSuitID
    arrIn(24) = Null
    arrIn(25) = Null
    
    arrIn(26) = "0" 'txtNoofPages.Text
   ' arrIn(27) = cmbSeat.ItemData(cmbSeat.ListIndex) 'cmbSeatID.Text
    arrIn(27) = CDbl(mSeatPrefix + CStr(cmbSeat.ItemData(cmbSeat.ListIndex))) 'paperless
    arrIn(28) = gbSeatID  'gbnumSeatID
    arrIn(29) = arrIn(27)
    arrIn(30) = 0   'Changed by Renjitha on 29.05.2012
    arrIn(31) = gbLBID
    arrIn(32) = gbLocationID
    arrIn(33) = "Soochika Saankhya inward module[KMBR] "
    arrIn(34) = Null
    ' Changed by soumya VS on 29 May 2015
    If (gbLBID = 167) Then
        arrIn(35) = Null
        arrIn(36) = Null
    End If
    
    Set Rec = objdb.ExecuteSP("SpSaveInwardDetails", arrIn, arrOut, , mCnn, adCmdStoredProc)
    If (IsArray(arrOut) = True) Then
        SaveSoochikaInwardDetails = arrOut(0, 0)
    End If
End Function
Private Function SaveSanketham(ByRef lSoochikaCurrentNo As Variant, ByRef mCnnKMBR As ADODB.Connection) As Boolean
        On Error GoTo err:
        
        Dim mVarrIn As Variant
        Dim mVarrOut As Variant
        Dim objdb As New clsDB
'        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim strName As String
        ReDim mVarrIn(20)
     
'        Dim Rec1 As New ADODB.Recordset
        
        Dim i As Integer
        Dim J As Integer
        Dim ary As Variant
        Dim sp As Integer
        
        Dim N1 As String
        Dim N2 As String
        Dim N3 As String
        Dim N4 As String
        Dim I1 As String
        Dim I2 As String
        Dim I3 As String
        Dim I4 As String
        
        lSoochikaFeildID = lSoochikaCurrentNo
        
        mVarrIn(0) = IIf(Len(lSoochikaFeildID) >= 6, Right(lSoochikaFeildID, 6), lSoochikaFeildID)  'intMain,
        mVarrIn(1) = gbLocalBodyID 'intLB,
        mVarrIn(2) = cmbDZone.ItemData(cmbDZone.ListIndex) 'intZone,
        mVarrIn(3) = val(mBuildingWard)   '  Trim(txtWardNo.Text)   'cboWard.itemData(cboWard.ListIndex) 'intWard,
        mVarrIn(4) = 1 'bitDocuments,
        mVarrIn(5) = mPermitType 'bitOneDayPermit,
        mVarrIn(6) = Trim(txtName.Text) 'chvName,
        mVarrIn(7) = Trim(txtDoorNo1.Text) 'intDoorNo1AddressHn,
        mVarrIn(8) = Trim(txtDoorNo2.Text) 'intDoorNo2AddressHn,
        mVarrIn(9) = val(txtWardNo.Text) 'intWardNoAddress,
        mVarrIn(10) = Trim(txtHouse.Text) 'chvHouseNameAddress,
        mVarrIn(11) = Trim(txtMainPlace.Text) ',
        mVarrIn(12) = val(txtMainPlace.Tag)  'cboDistchvMainPlaceAddressrict.itemData(cboDistrict.ListIndex) 'intDistIdAddress,
        mVarrIn(13) = val(txtPost.Tag)  'cmbPostOffice.itemData(cmbPostOffice.ListIndex) 'intPostOfficeAddress,
        mVarrIn(14) = cmbSeat.Text 'cmbSeat.itemData(cmbSeat.ListIndex) 'chvSeatNoFlnoPart1,
        mVarrIn(15) = val(Right(str(lSoochikaCurrentNo), 6)) 'intCurrentNoFlnoPart2,
        mVarrIn(16) = gbUserID 'intUserId,
        mVarrIn(17) = 1 'intFileStatus,
        mVarrIn(18) = 0 'intProcess,
        mVarrIn(19) = val(vsGrid.TextMatrix(1, 5)) 'Fee
        mVarrIn(20) = lSoochikaFeildID
        
        
        objdb.ExecuteSP "SoochikaIns1", mVarrIn, mVarrOut, , mCnnKMBR, adCmdStoredProc
        '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
        If IsArray(mVarrOut) Then ' Changed by Mahesh.. For what purpose that i dont know!!!
'          Dim lSoochikaFileIDKMBR As Variant
           lSoochikaFileIDKMBR = mVarrOut(0, 0)
        Else
            GoTo err
        End If
        '------------------------------------------------------------------------------------
        
        
        sp = 0
        N1 = ""
        N2 = ""
        N3 = ""
        N4 = ""
        I1 = ""
        I2 = ""
        I3 = ""
        I4 = ""

        strName = CStr(txtName.Text) + " " + CStr(txtInit1.Text) + " " + CStr(txtInit2.Text) + " " + CStr(txtInit3.Text) + " " + CStr(txtInit4.Text)
        
        For i = 1 To Len(Trim(strName))
            If LCase(mID(Trim(strName), i, 1)) = " " Then
                sp = sp + 1
            End If
        Next i
        
        ReDim ary(sp)
        If Trim(strName) <> "" Then
            ary = Split(Trim(strName), " ", , vbTextCompare)
        End If
        
        For J = 0 To sp
            If N1 = "" And Len(ary(J)) > 1 Then
                N1 = ary(J)
            ElseIf I1 = "" And Len(ary(J)) = 1 Then
                I1 = ary(J)
            ElseIf N2 = "" And Len(ary(J)) > 1 Then
                N2 = ary(J)
            ElseIf I2 = "" And Len(ary(J)) = 1 Then
                I2 = ary(J)
            ElseIf N3 = "" And Len(ary(J)) > 1 Then
                N3 = ary(J)
            ElseIf I3 = "" And Len(ary(J)) = 1 Then
                I3 = ary(J)
            ElseIf N4 = "" And Len(ary(J)) > 1 Then
                N4 = ary(J)
            ElseIf I4 = "" And Len(ary(J)) = 1 Then
                I4 = ary(J)
            End If
        Next J
        
        'Note:- Only Require if Demand is Generated From Saankhya
        If chkLinkDemand.Value = 0 Then
            ReDim mVarrIn(12)
            '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
            'mVarrIn(0) = IIf(Len(lSoochikaFeildID) >= 6, Right(lSoochikaFeildID, 6), lSoochikaFeildID) 'intMain,
'            mVarrIn(0) = lSoochikaFeildID 'intMain
             mVarrIn(0) = lSoochikaFileIDKMBR 'mainid
            '-------------------------------------------------------------------------------------
            mVarrIn(1) = 0 'CatId ,
            mVarrIn(2) = 4 'LanguageId,
            mVarrIn(3) = N1 'Name1,
            mVarrIn(4) = N2 'Name2,
            mVarrIn(5) = N3 'Name3,
            mVarrIn(6) = N4 'Name4,
            mVarrIn(7) = I1 'Intial1,
            mVarrIn(8) = I2 'Intial2,
            mVarrIn(9) = I3 'Intial3,
            mVarrIn(10) = I4 'Intial4,
            mVarrIn(11) = cmbDZone.ItemData(cmbDZone.ListIndex)
            mVarrIn(12) = gbLocalBodyID
            objdb.ExecuteSP "NameIns", mVarrIn, , , mCnnKMBR, adCmdStoredProc
                
            ReDim mVarrIn(18)
            
            '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
            'mVarrIn(0) = IIf(Len(lSoochikaFeildID) >= 6, Right(lSoochikaFeildID, 6), lSoochikaFeildID)  'intMain,
            'mVarrIn(0) = lSoochikaFeildID 'intMain,
            mVarrIn(0) = lSoochikaFileIDKMBR  'intMain,
            '------------------------------------------------------------------------------------
            mVarrIn(1) = 0 'tnyTypeId,
            mVarrIn(2) = val(txtWardNo.Text) 'intHouseNoWard,
            mVarrIn(3) = CStr(txtDoorNo2.Text) 'chvHouseNo,
            mVarrIn(4) = val(txtMainPlace.Tag) 'intDistrict,
            mVarrIn(5) = Trim(txtHouse.Text) 'chvHouseName,
            mVarrIn(6) = Null   'txtResAssNo.Text 'chvResAssocNo,
            mVarrIn(7) = Null   '  txtResAssoName.Text 'chvResAssoc,
            mVarrIn(8) = Trim(txtLocalPlace.Text) 'chvLandMark,
            mVarrIn(9) = Trim(txtStreet.Text) 'chvStreetName,
            mVarrIn(10) = Trim(txtMainPlace.Text) 'chvMainPlace,
            mVarrIn(11) = val(txtPost.Tag)  'cmbPostOffice.itemData(cmbPostOffice.ListIndex) 'intPostOfficeId,
            mVarrIn(12) = val(txtPin.Text) 'intPincode,
            mVarrIn(13) = Trim(txtPhone.Text) 'nmPhoneNo,
            mVarrIn(14) = Null  '  txtMobileNo.Text 'nmMobileNo,
            mVarrIn(15) = 0 'tnyCatId
            mVarrIn(16) = gbLocalBodyID
            mVarrIn(17) = cmbDZone.ItemData(cmbDZone.ListIndex)
            mVarrIn(18) = val(txtDoorNo1.Text)
            objdb.ExecuteSP "AddressIns", mVarrIn, , , mCnnKMBR, adCmdStoredProc
        End If
        
        ReDim mVarrIn(11)
        Dim lReceiptID As Double
        Dim chvFileNo As String
        
        mVarrIn(0) = 0  'ReceiptId
        '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
        'mVarrIn(1) = IIf(Len(lSoochikaFeildID) >= 6, Right(lSoochikaFeildID, 6), lSoochikaFeildID)  'Main
       ' mVarrIn(1) = lSoochikaFeildID 'Main   lSoochikaFileID
       mVarrIn(1) = lSoochikaFileIDKMBR 'Main
       '--------------------------------------------------------------------------------------
        mVarrIn(2) = gbLocalBodyID 'Lb
    '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
        mVarrIn(3) = cmbSeat.Text & " /" & lSoochikaFileIDKMBR & " /" & Year(Now) 'FileNo
'       mVarrIn(3) = cmbseatlSoochikaFeildID 'FileNo
    '----------------------------------------------------------------------------------------
        mVarrIn(4) = gbCounterID 'Counter
        mVarrIn(5) = Null 'TransationId
        mVarrIn(6) = mVoucherID 'VoucherId
        mVarrIn(7) = mReceiptNo 'ReceiptNo
        mVarrIn(8) = Trim(txtDate.Text)   'txtReceiptDate.Text 'dtReceipt
        mVarrIn(9) = val(vsGrid.TextMatrix(1, 5)) 'Amount
        mVarrIn(10) = val(txtInstrument.Tag) 'cboInstrument.itemData(cboInstrument.ListIndex) 'CreditId
        mVarrIn(11) = cmbDZone.ItemData(cmbDZone.ListIndex)
        objdb.ExecuteSP "Receipt_InsTR", mVarrIn, mVarrOut, , mCnnKMBR, adCmdStoredProc

        If IsArray(mVarrOut) Then
           lReceiptID = mVarrOut(0, 0)
        End If
        
        ReDim mVarrIn(8)
        mVarrIn(0) = 0 'intId
        'mVarrIn(1) = mVoucherID 'intReceiptId
        mVarrIn(1) = lReceiptID 'intReceiptId
        mVarrIn(2) = gbLocalBodyID 'intLBId
        '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
        'mVarrIn(3) = IIf(Len(lSoochikaFeildID) >= 6, Right(lSoochikaFeildID, 6), lSoochikaFeildID)  'intMainId
'        mVarrIn(3) = lSoochikaFeildID 'intMainId
        mVarrIn(3) = lSoochikaFileIDKMBR 'intMainId
        '-------------------------------------------------------------------------------------
        mVarrIn(4) = val(vsGrid.TextMatrix(1, 6)) 'intDebitId
        mVarrIn(5) = val(vsGrid.TextMatrix(1, 5))   'Left(Trim(fgAccHead.TextMatrix(1, 6)), 4) 'fltAmount
        mVarrIn(6) = gbFinancialYearID   'Left(Trim(fgAccHead.TextMatrix(1, 5)), 4) 'intFinancialyear
        mVarrIn(7) = Null 'intPeriod
        mVarrIn(8) = cmbDZone.ItemData(cmbDZone.ListIndex)
        objdb.ExecuteSP "Receipt_InsTC", mVarrIn, , , mCnnKMBR, adCmdStoredProc
      '-------------------CHANGED BY SYALIMA ON 29/01/2014----------------------------------
        'Split And Update MainID ZoneID,FileNo to KMBR Tables
        ReDim mVarrIn(3)
        mVarrIn(0) = val(Trim(txtDemandPrefix.Text))
'        mVarrIn(0) = Array(Trim(txtDemandPrefix.Text))
        mVarrIn(1) = lSoochikaFileIDKMBR
        mVarrIn(2) = cmbDZone.ItemData(cmbDZone.ListIndex)
        mVarrIn(3) = cmbSeat.Text & " /" & lSoochikaFileIDKMBR & " /" & Year(Now)
        
        objdb.ExecuteSP "SPLITUpdateByEfileNo", mVarrIn, , , mCnnKMBR, adCmdStoredProc



      
               ' 'Note:- sp_UpdateMainID
'        mVarrIn = Array(Trim(txtDemandPrefix.Text), lSoochikaFileIDKMBR, lSoochikaCurrentNo, cmbDZone.ItemData(cmbDZone.ListIndex))
''        mVarrIn = Array(Trim(txtDemandPrefix.Text), lSoochikaFileID, lSoochikaCurrentNo)
'        objDB.ExecuteSP "UpdateMainID", mVarrIn, , , mCnnKMBR, adCmdStoredProc

'----------------------------------------------------------------------------------------------

        SaveSanketham = True
        Exit Function
err:
        MsgBox "Saankhya Error Handler: " & Error$
        SaveSanketham = False
        Exit Function
    End Function

    
    
Private Sub PrintReceipt20(intVoucherID As Double)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        
        'FileInitialize
''''        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
''''        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
''''        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Left Join faPeriodicity On  faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
''''        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'objDb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic
        
        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptPTax(intVoucherID)
                Exit Sub
            End If
        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(77)  ' Set to 12 CPI
        Print #gbFileNO,
        Print #gbFileNO,
        
        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(26); gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(76); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
            Else
                Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(134); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, gbBold; Tab(17); mName; Tab(67); mName
            
            'Changed for Sujith by Aiby - 24-Mar-2009
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(67); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(67); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
                If Len(mStr1) < 51 Then
                    mStr1 = mStr1 & String(52 - Len(mStr1), " ")
                Else
                    mStr1 = PadR(mStr1, 50)
                End If
                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            Print #gbFileNO, mStr1; Tab(60); mStr2
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
                    Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(28); PadL(Format(Rec!fltAmount, "0.00"), 9);
                Else
                    Print #gbFileNO, Tab(38); PadL(Format(Rec!fltAmount, "0.00"), 9);
                End If
                
                '==================================================================='
                ' Receipt Area
                '==================================================================='
                Print #gbFileNO, Tab(55); PadL(CStr(mLoop), 2);
                Print #gbFileNO, Tab(60); PadR(Rec!vchAlias, 58);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(120); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(120); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(120); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(120); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(137); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Else
                    Print #gbFileNO, Tab(149); PadL(Format(Rec!fltAmount, "0.00"), 9)
                End If
                Rec.MoveNext
            Wend
            Rec.MoveFirst
            
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
            Print #gbFileNO, Tab(37); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(149); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            
            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
            
            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            Print #gbFileNO, Tab(12); Left(mRupees, 38);
            Print #gbFileNO, Tab(70); Left(mRupees, 85)
            
            Print #gbFileNO, Tab(12); mID$(mRupees, 39, 70);
            Print #gbFileNO, Tab(70); mID$(mRupees, 86, 90);
            
            Print #gbFileNO,
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 43);
            'Print #gbFileNO, Tab(66); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            Print #gbFileNO, Tab(66); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 94)
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 43 Then
                Print #gbFileNO, Tab(7); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 44, 43);
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 94 Then
                Print #gbFileNO, Tab(66); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 95, 94)
'            Else
'                Print #gbFileNO,
            End If
            
'            objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                Print #gbFileNO, Tab(11); objCounter.CounterNo;
'                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
'            End If
'            objUser.SetUser (Rec!intUserID)
'            If objUser.UserID > -1 Then
'                Print #gbFileNO, Tab(11); objUser.UserName;
'                Print #gbFileNO, Tab(61); objUser.UserName
'            End If
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(11); objCounter.CounterNo; Tab(15); objUser.UserName;
                    Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(100); objUser.UserName
                End If
            End If
            Print #gbFileNO,
        End If
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName
    End Sub
    
Private Function PrintReceipt(intVoucherID As Double) As Integer
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        Open gbFileName For Output As #gbFileNO
'        Print #gbFileNO, Chr$(27) + Chr$(80)
'        Print #gbFileNO, String(136, "-")
'        Close #gbFileNO
'        Shell "Print " & gbFileName
'------------------------------------------------------------------------------------------------------------'
'-----------------------------------------Printing in 17 CPI-------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String
        Dim mPrintLBName As Boolean
        
        '***To print Local Body Name in Receipt***'
        '*******Added On 31/3/11 By Vinod MV******'
        '*******PRint LB name in PrePrinted Slip if this Varible is true
        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
            mPrintLBName = True
        Else
            mPrintLBName = False
        End If
        '*****************************************'
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
    
        'FileInitialize

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
''                Call PrintSummaryReceiptPTax(intVoucherID)
                Call PrintSummaryReceiptPTaxRes(intVoucherID)
                Exit Function
            End If
        ElseIf Rec!intTransactionTypeID = gbTransactionTypeRentOnBuilding Or Rec!intTransactionTypeID = gbTransactionTypeRentOnLand Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptRLB(intVoucherID)
                Exit Function
            End If
        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80) ' Set to 10 CPI
        Print #gbFileNO,
        Print #gbFileNO,
        If mPrintLBName = True Then
            Print #gbFileNO, gbBold; gbDoubleWidth; Tab(28); gbLBTitle; gbDoubleWidthOff; gbBoldOff
        Else
            Print #gbFileNO,
        End If

        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            'For GSTIN in Receipt on 18/12/2018
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(76); "CASH"; gbDoubleWidthOff
             Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(60); "CASH GST:"; LTrim(RTrim(gbGSTIN)); gbDoubleWidthOff; gbBoldOff;
        Case Is = 4
            'For GSTIN in Receipt on 18/12/2018
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "D.Draft"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(76); "Demand Draft"; gbDoubleWidthOff
            Print #gbFileNO, Tab(31); gbDoubleWidth; "D.Draft"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(55); "D.D GST:"; LTrim(RTrim(gbGSTIN)); gbDoubleWidthOff; gbBoldOff;
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'            Else
'                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
            End If
        Case Is = 5
        'For GSTIN in Receipt on 18/12/2018
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(76); "CHEQUE"; gbDoubleWidthOff
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(50); IIf(gbPrinterMode = 9, gbLBName, ""); Tab(60); "CHQ GST:"; LTrim(RTrim(gbGSTIN)); gbDoubleWidthOff; gbBoldOff;
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + IIf(IsNull(Rec!dtInstrumentDate), "", DdMmmYy(Rec!dtInstrumentDate))
            End If
        Case Else
            Print #gbFileNO,
        End Select


'        Commented by syalima on 20/12/2018 START
'
'        If Not (Rec.EOF And Rec.BOF) Then
'            ' Line 6
'            Print #gbFileNO, Tab(18); gbBold; gbDoubleWidth; Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); Tab(65); Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff
'            ' Changed for KMBR By Cijith Sreedharan
'            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
'                If mKMBRFlag Or mSoochikaConnected Then
'                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
'                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'                Else
'                    Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'                End If
'            Else
'                Print #gbFileNO, Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(116); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
'            End If
'
'            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
'            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
'            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
'            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
'            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
'
'
''            Print #gbFileNO, gbBold; Tab(17); mName; Tab(67); mName; Tab(87); gbBold; "GSTIN : "; Tab(96); gbGSTIN;
'            'Commented by syalima on 20/12/2018
'            'Print #gbFileNO, gbBold; Tab(17); mName; Tab(67); mName; Tab(87); gbBold;
'            Print #gbFileNO, gbBold; Tab(17); Left(mName, 34) & "."; Tab(67); mName; Tab(87); gbBold; gbDoubleWidthOff;
'
'            'Changed for Sujith by Aiby - 24-Mar-2009
'
'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
'
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(67); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(67); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
'            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
'            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
'
'            ' --------------------------------------------------------------------------------- '
'            ' To Print Check Number and DD Number Printing Phone Number is Commented
'            ' --------------------------------------------------------------------------------- '
'            Select Case Rec!intInstrumentTypeID
'            Case Is = 1
'                Print #gbFileNO,
'            Case Is = 4
'                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                If Not IsNull(Rec!dtInstrumentDate) Then
'                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                End If
'                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Is = 5
'                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                If Not IsNull(Rec!dtInstrumentDate) Then
'                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'                End If
'                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
'            Case Else
'                Print #gbFileNO,
'            End Select
'
'            ' Line 15 Next
'            'Changed its Possition- Requested by Sujith on 24-Mar-2009
'            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
'
'            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
'                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 45 Then
'                    mStr1 = mStr1 & String(45 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 44)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
'            Print #gbFileNO, mStr1; Tab(60); mStr2
'            Print #gbFileNO,
'
'            ' Line 18 Next
'            Rec.MoveFirst
'            While Not Rec.EOF
'                mLoop = mLoop + 1
'                '==================================================================='
'                ' Counter Foil
'                '==================================================================='
'                Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
'                If Not IsNull(Rec!intYearId) Then
'                    mstrYear = CStr(Rec!intYearId) & "-" & Right(CStr(Rec!intYearId + 1), 2)
'                Else
'                    mstrYear = ""
'                End If
'                Select Case Rec!tnyPeriodID
'                    Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
'                    Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
'                    Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
'                    Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchperiodicity), "", Rec!vchperiodicity), 3);
'
'                End Select
'
'                If Rec!intYearId < gbFinancialYearID Then
'                    Print #gbFileNO, Tab(24); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                Else
'                    Print #gbFileNO, Tab(31); PadL(Format(Rec!fltAmount, "0.00"), 9);
'                End If
'
'                '==================================================================='
'                ' Receipt Area
'                '==================================================================='
'                Print #gbFileNO, Tab(46); PadL(CStr(mLoop), 2);
'                Print #gbFileNO, Tab(49); PadR(Rec!vchAlias, 55);
'                If Not IsNull(Rec!intYearId) Then
'                    mstrYear = CStr(Rec!intYearId) & "-" & Right(CStr(Rec!intYearId + 1), 2)
'                Else
'                    mstrYear = ""
'                End If
'                Select Case Rec!tnyPeriodID
'                    Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
'                    Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
'                    Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
'                    Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchperiodicity), "", Rec!vchperiodicity), 3);
'                End Select
'
'                If Rec!intYearId < gbFinancialYearID Then
'                    Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                Else
'                    Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
'                End If
'                Rec.MoveNext
'            Wend
'            Rec.MoveFirst
'
'            For mCount = mLoop + 1 To 9
'                Print #gbFileNO,
'            Next mCount
'            If Rec!fltAdvAmtAdj > 0 Then
'                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
'            Else
''                Print #gbFileNO,'Commented By Vinod
'            End If
'            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
'
'            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
'            Print #gbFileNO, Tab(130); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
'
'            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
'            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
'
'            mRupees = Rupees(Rec!TotalAmt)
'            If Len(mRupees) < 186 Then
'                mRupees = mRupees + String(185 - Len(mRupees), " ")
'            End If
'            Print #gbFileNO, Tab(12); Left(mRupees, 32);
'            Print #gbFileNO, Tab(60); Left(mRupees, 75)
'
'            Print #gbFileNO, Tab(12); mID$(mRupees, 33, 32);
'            Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)
'
'            'Print #gbFileNO,'Commented By Vinod
'            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 37);
'            Print #gbFileNO, Tab(60); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 75)
'
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 37 Then
'                Print #gbFileNO, Tab(7); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 38, 37);
'            Else
'                Print #gbFileNO,
'            End If
'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 75 Then
'                Print #gbFileNO, Tab(60); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 76, 75)
''            Else
''                Print #gbFileNO,
'            End If
'
'            objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                objUser.SetUser (Rec!intUserID)
'                If objUser.UserID > -1 Then
'                    Print #gbFileNO, Tab(11); objCounter.CounterNo; Tab(15); objUser.UserName;
'                    Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(100); objUser.UserName
'                End If
'            End If
'
'
'            Print #gbFileNO,
'        End If
'
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO,
'        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
'        Close #gbFileNO
'        'ShellPad
'        Dim mFlag As Integer
'        mFlag = Shell("Print " & gbFileName)
'        Sleep 1000
'        PrintReceipt = mFlag
'
'        'Kill gbFileName
'
'Commented by syalima on 20/12/2018 END

If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(18); gbBold; gbDoubleWidth; Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); Tab(65); Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
            Else
                Print #gbFileNO, Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(116); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If

            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
'            If Len(mName) > 35 Then
'                mName = Left(mName, 34)
'            Else
'                mName = mName
'            End If
            Print #gbFileNO, gbBold; Tab(17); Left(mName, 34) & "."; Tab(67); mName; gbDoubleWidthOff
'            Print #gbFileNO, gbBold; Tab(17); mName; Tab(67); mName; gbDoubleWidthOff

            'Changed for Sujith by Aiby - 24-Mar-2009

            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff; gbDoubleWidthOff

            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(67); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(67); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select

            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
                If Len(mStr1) < 45 Then
                    mStr1 = mStr1 & String(45 - Len(mStr1), " ")
                Else
                    mStr1 = PadR(mStr1, 44)
                End If
                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            Print #gbFileNO, mStr1; Tab(60); mStr2
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
                    Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);

                End Select

                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(24); PadL(Format(Rec!fltAmount, "0.00"), 9);
                Else
                    Print #gbFileNO, Tab(31); PadL(Format(Rec!fltAmount, "0.00"), 9);
                End If

                '==================================================================='
                ' Receipt Area
                '==================================================================='
                Print #gbFileNO, Tab(46); PadL(CStr(mLoop), 2);
                Print #gbFileNO, Tab(49); PadR(Rec!vchAlias, 55);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                End Select

                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Else
                    Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                End If
                Rec.MoveNext
            Wend
            Rec.MoveFirst

            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True); gbDoubleWidthOff;
            Print #gbFileNO, Tab(130); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            Print #gbFileNO, Tab(12); Left(mRupees, 32); gbDoubleWidthOff;
            Print #gbFileNO, Tab(60); Left(mRupees, 75); gbDoubleWidthOff;

            Print #gbFileNO, Tab(12); mID$(mRupees, 33, 32); gbDoubleWidthOff;
            Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85); gbDoubleWidthOff;

            'Print #gbFileNO,'Commented By Vinod
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 37); gbDoubleWidthOff;
            Print #gbFileNO, Tab(60); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 75); gbDoubleWidthOff;

            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 37 Then
                Print #gbFileNO, Tab(7); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 38, 37); gbDoubleWidthOff;
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 75 Then
                Print #gbFileNO, Tab(60); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 76, 75); gbDoubleWidthOff;
'            Else
'                Print #gbFileNO,
            End If

            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(11); objCounter.CounterNo; Tab(15); objUser.UserName; gbDoubleWidthOff;
                    Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(100); objUser.UserName; gbDoubleWidthOff;
                End If
            End If


            Print #gbFileNO,
        End If

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        Close #gbFileNO
        'ShellPad
        Dim mFlag As Integer
        mFlag = Shell("Print " & gbFileName)
        Sleep 1000
        PrintReceipt = mFlag

        
        
      
        
        
    End Function





Private Sub PrintReceiptBlankPaper(intVoucherID As Double)
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        Open gbFileName For Output As #gbFileNO
'        Print #gbFileNO, Chr$(27) + Chr$(80)
'        Print #gbFileNO, String(136, "-")
'        Close #gbFileNO
'        Shell "Print " & gbFileName
'------------------------------------------------------------------------------------------------------------'
'-----------------------------------------Printing in 17 CPI-------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String

        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptPTax(intVoucherID)
                Exit Sub
            End If
        ElseIf Rec!intTransactionTypeID = gbTransactionTypeRentOnBuilding Or Rec!intTransactionTypeID = gbTransactionTypeRentOnLand Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptRLB(intVoucherID)
                Exit Sub
            End If
        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80) ' Set to 10 CPI
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,


        Print #gbFileNO, gbLBName
        Print #gbFileNO, gbLBType
        Print #gbFileNO, "RECEIPT"
        Print #gbFileNO, String(80, "-")
        Print #gbFileNO, "R.No:"; Tab(18); gbBold; gbDoubleWidth; Right(gbLocationID, 2); "/"; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff;
        Print #gbFileNO, Tab(46); "Date : "; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
        
        
        mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
        If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
        If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
        If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
        If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
        
        Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2)
        Print #gbFileNO, "Name      :"; gbBold; mName; gbBoldOff
        Print #gbFileNO, "Address   :";
        Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
        Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
        Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
        Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
        
        Print #gbFileNO, "Demand No :"; Tab(55); "Ref.No : "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
        Print #gbFileNO, "Instrument Type : ";
        Select Case Rec!intInstrumentTypeID
        Case Is = 1
            Print #gbFileNO, Tab(25); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
            End If
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + IIf(IsNull(Rec!dtInstrumentDate), "", DdMmmYy(Rec!dtInstrumentDate))
            End If
        Case Else
            Print #gbFileNO,
        End Select

        Print #gbFileNO, String(80, "-")
        Print #gbFileNO, "Sl.  Head Code   Particulars                   Year     Period     Amount "
        Print #gbFileNO, String(80, "-")
        If Not (Rec.EOF And Rec.BOF) Then
            Rec.MoveFirst
            While Not Rec.EOF
                mLoop = mLoop + 1
                Print #gbFileNO, PadL(CStr(mLoop), 2);
                Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
                Print #gbFileNO, Tab(39); PadR(Rec!vchAlias, 55);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                End Select
                Print #gbFileNO, Tab(24); PadL(Format(Rec!fltAmount, "0.00"), 9);
                
                
                Rec.MoveNext
            Wend
            Rec.MoveFirst

            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46);
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, String(80, "-")
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
            Print #gbFileNO, Tab(70); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            
            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            Print #gbFileNO, Tab(12); Left(mRupees, 32);
            Print #gbFileNO, Tab(12); mID$(mRupees, 33, 32);
            Print #gbFileNO, "Remarks :"; Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 75)
            Print #gbFileNO, String(80, "-")
            
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(11); objCounter.CounterNo; Tab(15); objUser.UserName;
                End If
            End If
            Print #gbFileNO,
        End If

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        Close #gbFileNO
        ShellPad
        'Shell "Print " & gbFileName
        Kill gbFileName
    End Sub
Private Function PrintReceipt_ForNewFormatRes(intVoucherID As Double) As Integer
 ' NEW FORMAT FOR  SAANKHYA SOOCHIKA Modified on 11-Oct-2011 (Aiby)
    '        gbFileNO = FreeFile
    '        gbFileName = "C:\Report.txt"
    '        Open gbFileName For Output As #gbFileNO
    '        Print #gbFileNO, Chr$(27) + Chr$(80)
    '        Print #gbFileNO, String(136, "-")
    '        Close #gbFileNO
    '        Shell "Print " & gbFileName
    '------------------------------------------------------------------------------------------------------------'
    '-----------------------------------------Printing in 17 CPI-------------------------------------------------'
    '------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String
        Dim mInwardNo As String
             
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

                
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
        Dim mLBType As String
        Select Case gbLBType
            
            
            Case Is = 1 ' District
                mLBType = "District Panchayat"
            Case Is = 2 ' Block
                mLBType = "Block Panchayat"
''''            Case Is = 3 ' Block
''''                mLBType = "Muncipality"
''''            Case Is = 4 ' Block
''''                mLBType = "Corporation"
            Case Else
                mLBType = "Grama Panchayat"
        End Select
        
        '1 line
        Print #gbFileNO, Tab(3); gbBold; gbDoubleWidth; "RECEIPT"; Tab(31); gbLBName; " "; mLBType; gbDoubleWidthOff
        '2nd line
        If Not (Rec.EOF And Rec.BOF) Then
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    'Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                
                    'Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));  '3line
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)) '4thlin
            Else
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); '3line
                ' For GSTIN in Receipt on 18/12/2018
'                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(56); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(71); gbBold; "GSTIN:"; Tab(78); gbGSTIN; gbDoubleWidthOff; Tab(97); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            '3rd line
            'syalima start
            'Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff; Tab(86); gbBold; "GSTIN : "; Tab(95); gbGSTIN;
            'syalima END
            'syalima start
            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff;
             ' For GSTIN in Receipt on 18/12/2018
            'Print #gbFileNO, Tab(46); gbBold; "GSTIN : "; Tab(58); gbGSTIN;
            'syalima END
            
            'Changed for Sujith by Aiby - 24-Mar-2009

'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            '4th line
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
           '5th line
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                'Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                'Print #gbFileNO,
            End Select
           '6th line
            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            '7th line
            If Not (IsNull(Rec!vchRefNo)) Then
                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
            Else
                Print #gbFileNO,
            End If
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 47 Then
'                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 46)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            '8th line
            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
            'Print #gbFileNO,

            ' Line 18 Next
            
            Dim RecPTAX         As New ADODB.Recordset
            Dim mStartingYear   As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear     As Integer
            Dim mEndingPeriod   As Integer
            Dim mNarration      As String
            
            mStartingYear = 2100
            
            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                
                If Rec!intTransactionTypeID <> gbTransactionTypePTax Then
                    mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
                    mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                    mSql = mSql + " Where intVoucherID =" & intVoucherID '& Rec!intVoucherID
                    mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
                    mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
                    RecPTAX.Open mSql, mCnn
                    While Not RecPTAX.EOF
                        mLoop = mLoop + 1
                        Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
                        Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!Amount, "0.00"), 9);
                        Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                        Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
                        Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!Amount, "0.00"), 9)
                        RecPTAX.MoveNext
                    Wend
                    RecPTAX.Close
                    While Not Rec.EOF
                        If gbLBPanchayat Then
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        Else
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        End If
                        Rec.MoveNext
                    Wend
                    
                    'Rec.Close
                    Rec.MoveFirst
                    Print #gbFileNO,
                    mLoop = mLoop + 1
                    mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
                    Print #gbFileNO, mNarration; Tab(54); mNarration
                    mLoop = mLoop + 1

                    mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                    If mStartingPeriod = 1 Then
                        mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    ElseIf mStartingPeriod = 2 Then
                        mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    Else
                        mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    End If

                    If mEndingPeriod = 1 Then
                        mNarration = mNarration & " Ist Hf )"
                    ElseIf mEndingPeriod = 2 Then
                        mNarration = mNarration & " IInd Hf )"
                    Else
                        mNarration = mNarration & ")"
                    End If
                    mLoop = mLoop + 1
                    Print #gbFileNO, mNarration; Tab(52); mNarration
                Else
                    Dim mAmtPTaxCurrent As Double
                    Dim mAmtPTaxArrear As Double
                    Dim mAmtLC As Double
                    Dim mAmtPenal As Double
                    Dim mAmtServiceCess As Double
                    Dim mAmtSurcharge As Double
                    Dim mSplServices As Double
                    Dim mSurCentralGovtBuild As Double
                    Dim mNotice As Double
                    Dim mWarantee As Double
                    Dim mAdvance As Double
                    While Not Rec.EOF
                        Select Case Rec!vchAccountHeadCode
                            
                            Case gbAcHeadCodePropertyTaxCurrent, gbAcHeadCodePropertyTax_NonResidential_Current
                                 mAmtPTaxCurrent = mAmtPTaxCurrent + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodePropertyTaxArrear, gbAcHeadCodePropertyTax_NonResidential_Arrear
                                 mAmtPTaxArrear = mAmtPTaxArrear + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeLibraryCess
                                mAmtLC = mAmtLC + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodePenalInterest
                                mAmtPenal = mAmtPenal + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeServicceCessCurrent, gbAcHeadCodeServicceCessCurrentNonR, gbAcHeadCodeServicceCessArrear, gbAcHeadCodeServicceCessArrearNonR
                                mAmtServiceCess = mAmtServiceCess + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSurPTCurrent, gbAcHeadCodeSurPTCurrentNonR, gbAcHeadCodeSurPTArrear, gbAcHeadCodeSurPTArrearNonR
                                mAmtSurcharge = mAmtSurcharge + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSplServicesCurrent, gbAcHeadCodeSplServicesArrear
                                mSplServices = mSplServices + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSurCentralGovtBuildCurrent, gbAcHeadCodeSurCentralGovtBuildArrear
                                mSurCentralGovtBuild = mSurCentralGovtBuild + Format(Rec!fltAmount, "0.00")
                            Case 140400101          '''Notice fee
                                mNotice = mNotice + Format(Rec!fltAmount, "0.00")
                            Case 140400102         '''Warantee fee
                                mWarantee = mWarantee + Format(Rec!fltAmount, "0.00")
    '                        Case gbAcHeadCode, gbAcHeadCodeSurCentralGovtBuildArrear          '''Advance
    '                            mAdvance = mAdvance + Format(Rec!fltAmount, "0.00")
                        
                        End Select
                    Rec.MoveNext
                    Wend
                    '9th line to 18th line
                    If mAmtPTaxCurrent > 0 Then
                        Print #gbFileNO, "Property Tax(Current)"; Tab(26); PadL(Format(mAmtPTaxCurrent, "0.00"), 9); Tab(54); "Receivables for Property Tax(Current)"; Tab(128); PadL(Format(mAmtPTaxCurrent, "0.00"), 9)
                    End If
                    If mAmtPTaxArrear > 0 Then
                        Print #gbFileNO, "Property Tax(Arrears)"; Tab(26); PadL(Format(mAmtPTaxArrear, "0.00"), 9); Tab(54); "Receivables for Property Tax(Arrears)"; Tab(128); PadL(Format(mAmtPTaxArrear, "0.00"), 9)
                    End If
                    If mAmtLC > 0 Then
                     Print #gbFileNO, "Library Cess "; Tab(26); PadL(Format(mAmtLC, "0.00"), 9); Tab(54); "Library Cess Payable"; Tab(128); PadL(Format(mAmtLC, "0.00"), 9)
                    End If
                    If mAmtPenal > 0 Then
                        Print #gbFileNO, "Penal Interest"; Tab(26); PadL(Format(mAmtPenal, "0.00"), 9); Tab(54); "Penal Interest"; Tab(128); PadL(Format(mAmtPenal, "0.00"), 9)
                    End If
                    If mAmtServiceCess > 0 Then
                        Print #gbFileNO, "Service Cess "; Tab(26); PadL(Format(mAmtServiceCess, "0.00"), 9); Tab(54); "Receivables for Service Cess"; Tab(128); PadL(Format(mAmtServiceCess, "0.00"), 9)
                    End If
                    If mAmtSurcharge > 0 Then
                        Print #gbFileNO, "Surcharge"; Tab(26); PadL(Format(mAmtSurcharge, "0.00"), 9); Tab(54); "Receivables for Surcharge"; Tab(128); PadL(Format(mAmtSurcharge, "0.00"), 9)
                    End If
                    If mSplServices > 0 Then
                        Print #gbFileNO, "Special Service"; Tab(26); PadL(Format(mSplServices, "0.00"), 9); Tab(54); "Fees on Buildings for Special Service"; Tab(128); PadL(Format(mSplServices, "0.00"), 9)
                    End If
                    If mSurCentralGovtBuild > 0 Then
                        Print #gbFileNO, " Service Charge"; Tab(26); PadL(Format(mSurCentralGovtBuild, "0.00"), 9); Tab(54); "Service Charge on Central Govt Buildings"; Tab(128); PadL(Format(mSurCentralGovtBuild, "0.00"), 9)
                    End If
                    If mNotice > 0 Then
                        Print #gbFileNO, "Notice Fee"; Tab(26); PadL(Format(mNotice, "0.00"), 9); Tab(54); "NoticeFee"; Tab(128); PadL(Format(mNotice, "0.00"), 9)
                    End If
                    If mWarantee > 0 Then
                        Print #gbFileNO, "Warrant Fee"; Tab(26); PadL(Format(mWarantee, "0.00"), 9); Tab(54); "WarrantFee"; Tab(128); PadL(Format(mWarantee, "0.00"), 9)
                    End If
                   
                End If
                Rec.MoveFirst
                While Not Rec.EOF
                    If gbLBPanchayat Then
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    Else
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    End If
                    Rec.MoveNext
                Wend
                Rec.MoveFirst
        '19th line
                Print #gbFileNO,
                mLoop = mLoop + 1
                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
        '20th and 21st  line
                Print #gbFileNO, mNarration; Tab(54); mNarration
                mLoop = mLoop + 1
                
                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                If mStartingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                ElseIf mStartingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                Else
                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                End If
                
                If mEndingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf )"
                ElseIf mEndingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf )"
                Else
                    mNarration = mNarration & ")"
                End If
                mLoop = mLoop + 1
                Print #gbFileNO, mNarration; Tab(52); mNarration
        Else  ''<< If Rec.RecordCount > 9 Then
            
'               GoTo LB ' To print Property Tax containing less than 9 rows
'               End If
'               Else
            
'LB:
                mLoop = 0
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
                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
    
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    Else
                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    End If
    
                    '==================================================================='
                    ' Receipt Area
                    '==================================================================='
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias), 46);
                    If Not IsNull(Rec!intYearID) Then
                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                    Else
                        mstrYear = ""
                    End If
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    End Select

                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    Else
                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    End If
                    Rec.MoveNext
                Wend
            Rec.MoveFirst
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
      End If
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
            Print #gbFileNO, Tab(54); Left(mRupees, 75)

            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)

            'Print #gbFileNO,'Commented By Vinod
            Dim mInward As String
            If Not IsNull(Rec!numInwardNo) Then
                mInward = Rec!numInwardNo & "/" & Trim(str(Year(Rec!dtDate)))
            Else
                mInward = ""
            End If
            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23)
            If mSoochikaConnected Then
                Print #gbFileNO, IIf(IsNull(frmUSoochikaInward.cmbSeat.Text), "", frmUSoochikaInward.cmbSeat.Text);
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            Else
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            End If
            
            'Print #gbFileNO, Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
            'Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
            End If
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
            End If

            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
            End If
            Print #gbFileNO,
           ' Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName; 'Tab(20); mInward
                    'Print #gbFileNO, Tab(70); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(97); objUser.UserName;
                    Print #gbFileNO, Tab(60); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(77); objUser.UserName & "    " & mInward
                End If
            End If
                

 End If

finishprinting:
        Close #gbFileNO
        'ShellPad
        Dim mFlag As Integer
        Dim X As Integer
        
        
        mFlag = Shell("Print " & gbFileName)
        Sleep 1000
        PrintReceipt_ForNewFormatRes = mFlag
        'Kill gbFileName
    
End Function
Private Function PrintReceipt_ForNewFormatResForUrban(intVoucherID As Double) As Integer  'ADDED BY MINU FOR NEWLY FORMED MUN/CORP
 ' NEW FORMAT FOR  SAANKHYA SOOCHIKA Modified on 11-Oct-2011 (Aiby)
    '        gbFileNO = FreeFile
    '        gbFileName = "C:\Report.txt"
    '        Open gbFileName For Output As #gbFileNO
    '        Print #gbFileNO, Chr$(27) + Chr$(80)
    '        Print #gbFileNO, String(136, "-")
    '        Close #gbFileNO
    '        Shell "Print " & gbFileName
    '------------------------------------------------------------------------------------------------------------'
    '-----------------------------------------Printing in 17 CPI-------------------------------------------------'
    '------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String
        Dim mInwardNo As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
        Dim mLBType As String
        Select Case gbLBType
            
            
''''            Case Is = 1 ' District
''''                mLBType = "District Panchayat"
''''            Case Is = 2 ' Block
''''                mLBType = "Block Panchayat"
            Case Is = 3 ' Block
                mLBType = "Muncipality"
            Case Is = 4 ' Block
                mLBType = "Corporation"
            Case Else
                mLBType = ""
        End Select
        
        '1 line
        Print #gbFileNO, Tab(3); gbBold; gbDoubleWidth; "RECEIPT"; Tab(31); gbLBName; " "; mLBType; gbDoubleWidthOff
        '2nd line
        If Not (Rec.EOF And Rec.BOF) Then
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    'Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    'Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));  '3line
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)) '4thlin
            Else
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); '3line
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If

            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            '3rd line
            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff

            'Changed for Sujith by Aiby - 24-Mar-2009

'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            '4th line
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
           '5th line
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                'Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                'Print #gbFileNO,
            End Select
           '6th line
            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            '7th line
            If Not (IsNull(Rec!vchRefNo)) Then
                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
            Else
                Print #gbFileNO,
            End If
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 47 Then
'                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 46)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            '8th line
            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
            'Print #gbFileNO,

            ' Line 18 Next
            
            Dim RecPTAX         As New ADODB.Recordset
            Dim mStartingYear   As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear     As Integer
            Dim mEndingPeriod   As Integer
            Dim mNarration      As String
            
            mStartingYear = 2100
            
            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                
                If Rec!intTransactionTypeID <> gbTransactionTypePTax Then
                    mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
                    mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                    mSql = mSql + " Where intVoucherID =" & intVoucherID '& Rec!intVoucherID
                    mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
                    mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
                    RecPTAX.Open mSql, mCnn
                    While Not RecPTAX.EOF
                        mLoop = mLoop + 1
                        Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
                        Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!Amount, "0.00"), 9);
                        Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                        Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
                        Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!Amount, "0.00"), 9)
                        RecPTAX.MoveNext
                    Wend
                    RecPTAX.Close
                    While Not Rec.EOF
                        If gbLBPanchayat Then
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        Else
                            If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                                Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                                If mStartingYear > Rec!intYearID Then
                                    mStartingYear = Rec!intYearID
                                    mStartingPeriod = Rec!tnyPeriodID
                                End If
                                If mEndingYear < Rec!intYearID Then
                                    mEndingYear = Rec!intYearID
                                End If
                                mEndingPeriod = Rec!tnyPeriodID
                            End If
                        End If
                        Rec.MoveNext
                    Wend
                    
                    'Rec.Close
                    Rec.MoveFirst
                    Print #gbFileNO,
                    mLoop = mLoop + 1
                    mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
                    Print #gbFileNO, mNarration; Tab(54); mNarration
                    mLoop = mLoop + 1

                    mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                    If mStartingPeriod = 1 Then
                        mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    ElseIf mStartingPeriod = 2 Then
                        mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    Else
                        mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                    End If

                    If mEndingPeriod = 1 Then
                        mNarration = mNarration & " Ist Hf )"
                    ElseIf mEndingPeriod = 2 Then
                        mNarration = mNarration & " IInd Hf )"
                    Else
                        mNarration = mNarration & ")"
                    End If
                    mLoop = mLoop + 1
                    Print #gbFileNO, mNarration; Tab(52); mNarration
                Else
                    Dim mAmtPTaxCurrent As Double
                    Dim mAmtPTaxArrear As Double
                    Dim mAmtLC As Double
                    Dim mAmtPenal As Double
                    Dim mAmtServiceCess As Double
                    Dim mAmtSurcharge As Double
                    Dim mSplServices As Double
                    Dim mSurCentralGovtBuild As Double
                    Dim mNotice As Double
                    Dim mWarantee As Double
                    Dim mAdvance As Double
                    While Not Rec.EOF
                        Select Case Rec!vchAccountHeadCode
                            
                            Case gbAcHeadCodePropertyTaxCurrent, gbAcHeadCodePropertyTax_NonResidential_Current
                                 mAmtPTaxCurrent = mAmtPTaxCurrent + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodePropertyTaxArrear, gbAcHeadCodePropertyTax_NonResidential_Arrear
                                 mAmtPTaxArrear = mAmtPTaxArrear + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeLibraryCess
                                mAmtLC = mAmtLC + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodePenalInterest
                                mAmtPenal = mAmtPenal + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeServicceCessCurrent, gbAcHeadCodeServicceCessCurrentNonR, gbAcHeadCodeServicceCessArrear, gbAcHeadCodeServicceCessArrearNonR
                                mAmtServiceCess = mAmtServiceCess + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSurPTCurrent, gbAcHeadCodeSurPTCurrentNonR, gbAcHeadCodeSurPTArrear, gbAcHeadCodeSurPTArrearNonR
                                mAmtSurcharge = mAmtSurcharge + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSplServicesCurrent, gbAcHeadCodeSplServicesArrear
                                mSplServices = mSplServices + Format(Rec!fltAmount, "0.00")
                            Case gbAcHeadCodeSurCentralGovtBuildCurrent, gbAcHeadCodeSurCentralGovtBuildArrear
                                mSurCentralGovtBuild = mSurCentralGovtBuild + Format(Rec!fltAmount, "0.00")
                            Case 140400101          '''Notice fee
                                mNotice = mNotice + Format(Rec!fltAmount, "0.00")
                            Case 140400102         '''Warantee fee
                                mWarantee = mWarantee + Format(Rec!fltAmount, "0.00")
    '                        Case gbAcHeadCode, gbAcHeadCodeSurCentralGovtBuildArrear          '''Advance
    '                            mAdvance = mAdvance + Format(Rec!fltAmount, "0.00")
                        
                        End Select
                    Rec.MoveNext
                    Wend
                    '9th line to 18th line
                    If mAmtPTaxCurrent > 0 Then
                        Print #gbFileNO, "Property Tax(Current)"; Tab(26); PadL(Format(mAmtPTaxCurrent, "0.00"), 9); Tab(54); "Receivables for Property Tax(Current)"; Tab(128); PadL(Format(mAmtPTaxCurrent, "0.00"), 9)
                    End If
                    If mAmtPTaxArrear > 0 Then
                        Print #gbFileNO, "Property Tax(Arrears)"; Tab(26); PadL(Format(mAmtPTaxArrear, "0.00"), 9); Tab(54); "Receivables for Property Tax(Arrears)"; Tab(128); PadL(Format(mAmtPTaxArrear, "0.00"), 9)
                    End If
                    If mAmtLC > 0 Then
                     Print #gbFileNO, "Library Cess "; Tab(26); PadL(Format(mAmtLC, "0.00"), 9); Tab(54); "Library Cess Payable"; Tab(128); PadL(Format(mAmtLC, "0.00"), 9)
                    End If
                    If mAmtPenal > 0 Then
                        Print #gbFileNO, "Penal Interest"; Tab(26); PadL(Format(mAmtPenal, "0.00"), 9); Tab(54); "Penal Interest"; Tab(128); PadL(Format(mAmtPenal, "0.00"), 9)
                    End If
                    If mAmtServiceCess > 0 Then
                        Print #gbFileNO, "Service Cess "; Tab(26); PadL(Format(mAmtServiceCess, "0.00"), 9); Tab(54); "Receivables for Service Cess"; Tab(128); PadL(Format(mAmtServiceCess, "0.00"), 9)
                    End If
                    If mAmtSurcharge > 0 Then
                        Print #gbFileNO, "Surcharge"; Tab(26); PadL(Format(mAmtSurcharge, "0.00"), 9); Tab(54); "Receivables for Surcharge"; Tab(128); PadL(Format(mAmtSurcharge, "0.00"), 9)
                    End If
                    If mSplServices > 0 Then
                        Print #gbFileNO, "Special Service"; Tab(26); PadL(Format(mSplServices, "0.00"), 9); Tab(54); "Fees on Buildings for Special Service"; Tab(128); PadL(Format(mSplServices, "0.00"), 9)
                    End If
                    If mSurCentralGovtBuild > 0 Then
                        Print #gbFileNO, " Service Charge"; Tab(26); PadL(Format(mSurCentralGovtBuild, "0.00"), 9); Tab(54); "Service Charge on Central Govt Buildings"; Tab(128); PadL(Format(mSurCentralGovtBuild, "0.00"), 9)
                    End If
                    If mNotice > 0 Then
                        Print #gbFileNO, "Notice Fee"; Tab(26); PadL(Format(mNotice, "0.00"), 9); Tab(54); "NoticeFee"; Tab(128); PadL(Format(mNotice, "0.00"), 9)
                    End If
                    If mWarantee > 0 Then
                        Print #gbFileNO, "Warrant Fee"; Tab(26); PadL(Format(mWarantee, "0.00"), 9); Tab(54); "WarrantFee"; Tab(128); PadL(Format(mWarantee, "0.00"), 9)
                    End If
                   
                End If
                Rec.MoveFirst
                While Not Rec.EOF
                    If gbLBPanchayat Then
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    Else
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    End If
                    Rec.MoveNext
                Wend
                Rec.MoveFirst
        '19th line
                Print #gbFileNO,
                mLoop = mLoop + 1
                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
        '20th and 21st  line
                Print #gbFileNO, mNarration; Tab(54); mNarration
                mLoop = mLoop + 1
                
                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                If mStartingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                ElseIf mStartingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                Else
                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                End If
                
                If mEndingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf )"
                ElseIf mEndingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf )"
                Else
                    mNarration = mNarration & ")"
                End If
                mLoop = mLoop + 1
                Print #gbFileNO, mNarration; Tab(52); mNarration
        Else  ''<< If Rec.RecordCount > 9 Then
            
'               GoTo LB ' To print Property Tax containing less than 9 rows
'               End If
'               Else
            
'LB:
                mLoop = 0
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
                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
    
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    Else
                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    End If
    
                    '==================================================================='
                    ' Receipt Area
                    '==================================================================='
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias), 46);
                    If Not IsNull(Rec!intYearID) Then
                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                    Else
                        mstrYear = ""
                    End If
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    End Select

                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    Else
                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    End If
                    Rec.MoveNext
                Wend
            Rec.MoveFirst
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
      End If
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
            Print #gbFileNO, Tab(54); Left(mRupees, 75)

            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)

            'Print #gbFileNO,'Commented By Vinod
            Dim mInward As String
            If Not IsNull(Rec!numInwardNo) Then
                mInward = Rec!numInwardNo & "/" & Trim(str(Year(Rec!dtDate)))
            Else
                mInward = ""
            End If
            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23)
            If mSoochikaConnected Then
                Print #gbFileNO, IIf(IsNull(frmUSoochikaInward.cmbSeat.Text), "", frmUSoochikaInward.cmbSeat.Text);
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            Else
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            End If
            
            'Print #gbFileNO, Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
            'Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
            End If
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
            End If

            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
            End If
            Print #gbFileNO,
           ' Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName; 'Tab(20); mInward
                    'Print #gbFileNO, Tab(70); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(97); objUser.UserName;
                    Print #gbFileNO, Tab(60); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(77); objUser.UserName & "    " & mInward
                End If
            End If
                

 End If

finishprinting:
        Close #gbFileNO
        'ShellPad
        Dim mFlag As Integer
        Dim X As Integer
        
        
        mFlag = Shell("Print " & gbFileName)
        Sleep 1000
        PrintReceipt_ForNewFormatResForUrban = mFlag
        'Kill gbFileName
    
End Function
    
    
Private Sub PrintReceipt_ForNewFormat(intVoucherID As Double)
' NEW FORMAT FOR  SAANKHYA SOOCHIKA Modified on 11-Oct-2011 (Aiby)
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        Open gbFileName For Output As #gbFileNO
'        Print #gbFileNO, Chr$(27) + Chr$(80)
'        Print #gbFileNO, String(136, "-")
'        Close #gbFileNO
'        Shell "Print " & gbFileName
'------------------------------------------------------------------------------------------------------------'
'-----------------------------------------Printing in 17 CPI-------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String
        Dim mInwardNo As String

        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        'FileInitialize
''''        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
''''        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
''''        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Left Join faPeriodicity On  faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
''''        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
''''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

''''''        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
''''''            If Rec.RecordCount > 9 Then
''''''                Rec.Close
''''''                Call PrintSummaryReceiptPTax(intVoucherID)
''''''                Exit Sub
''''''            End If
''''''        End If
        Open gbFileName For Output As #gbFileNO

        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        Dim mLBType As String
        Select Case gbLBType


            Case Is = 1 ' District
                mLBType = "District Panchayat"
            Case Is = 2 ' Block
                mLBType = "Block Panchayat"
            Case Else
                mLBType = "Grama Panchayat"
        End Select


        Print #gbFileNO, Tab(3); gbBold; gbDoubleWidth; "RECEIPT"; Tab(31); gbLBName; " "; mLBType; gbDoubleWidthOff
'        Select Case Rec!intInstrumentTypeID
'        Case Is = 1
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
'        Case Is = 4
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
'            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Is = 5
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
'            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Else
'            Print #gbFileNO,
'        End Select

        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            'Print #gbFileNO, ; gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(65); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff;
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    'Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    'Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            Else
                Print #gbFileNO, gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate));
                Print #gbFileNO, Tab(46); gbBold; gbDoubleWidth; "RECEIPT"; Tab(58); IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbDoubleWidthOff; Tab(86); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If

            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4

            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff

            'Changed for Sujith by Aiby - 24-Mar-2009

'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                'Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                'Print #gbFileNO,
            End Select
            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            If Not (IsNull(Rec!vchRefNo)) Then
                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
            Else
                Print #gbFileNO,
            End If
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 47 Then
'                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 46)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
            'Print #gbFileNO,

            ' Line 18 Next

            Dim RecPTAX         As New ADODB.Recordset
            Dim mStartingYear   As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear     As Integer
            Dim mEndingPeriod   As Integer
            Dim mNarration      As String

            mStartingYear = 2100

            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
                mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                mSql = mSql + " Where intVoucherID =" & intVoucherID '& Rec!intVoucherID
                mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
                mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
                RecPTAX.Open mSql, mCnn
                While Not RecPTAX.EOF
                    mLoop = mLoop + 1
                    Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
                    Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!Amount, "0.00"), 9);
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
                    Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!Amount, "0.00"), 9)
                    RecPTAX.MoveNext
                Wend
                RecPTAX.Close
                While Not Rec.EOF
                    If gbLBPanchayat Then
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Current Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTax_NonResidential_Arrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    Else
                        If Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxCurrent Or _
                            Rec!vchAccountHeadCode = gbAcHeadCodePropertyTaxArrear Then
                            If mStartingYear > Rec!intYearID Then
                                mStartingYear = Rec!intYearID
                                mStartingPeriod = Rec!tnyPeriodID
                            End If
                            If mEndingYear < Rec!intYearID Then
                                mEndingYear = Rec!intYearID
                            End If
                            mEndingPeriod = Rec!tnyPeriodID
                        End If
                    End If
                    Rec.MoveNext
                Wend

                'Rec.Close
                Rec.MoveFirst
                Print #gbFileNO,
                mLoop = mLoop + 1
                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
                Print #gbFileNO, mNarration; Tab(54); mNarration
                mLoop = mLoop + 1

                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                If mStartingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                ElseIf mStartingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                Else
                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                End If

                If mEndingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf )"
                ElseIf mEndingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf )"
                Else
                    mNarration = mNarration & ")"
                End If
                mLoop = mLoop + 1
                Print #gbFileNO, mNarration; Tab(52); mNarration
            Else
'               GoTo LB ' To print Property Tax containing less than 9 rows
'           End If
'            Else

'LB:
                mLoop = 0
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
                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);

                    End Select

                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    Else
                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    End If

                    '==================================================================='
                    ' Receipt Area
                    '==================================================================='
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(Rec!vchAlias, 46);
                    If Not IsNull(Rec!intYearID) Then
                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                    Else
                        mstrYear = ""
                    End If
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    End Select

                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    Else
                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    End If
                    Rec.MoveNext
                Wend
            End If
            Rec.MoveFirst
'            For mCount = mLoop + 1 To 9
'                Print #gbFileNO,
'            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
            Print #gbFileNO, Tab(54); Left(mRupees, 75)

            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)

            'Print #gbFileNO,'Commented By Vinod
            Dim mInward As String
            If Not IsNull(Rec!numInwardNo) Then
                mInward = Rec!numInwardNo & "/" & Trim(str(Year(Rec!dtDate)))
            Else
                mInward = ""
            End If

''''''            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
''''''            Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
''''''
''''''            If mSoochikaConnected Then
''''''                Print #gbFileNO, IIf(IsNull(frmUSoochikaInward.cmbSeat.Text), "", frmUSoochikaInward.cmbSeat.Text);
''''''            End If
''''''            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
''''''                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
''''''            Else
''''''                Print #gbFileNO,
''''''            End If
''''''            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
''''''                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
''''''            End If
''''''
''''''            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 46 Then
''''''                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
''''''            Else
''''''                Print #gbFileNO,
''''''            End If
''''''            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
''''''                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
''''''            End If
''''''


            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23)
            If mSoochikaConnected Then
                Print #gbFileNO, IIf(IsNull(frmUSoochikaInward.cmbSeat.Text), "", frmUSoochikaInward.cmbSeat.Text);
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            Else
                If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                    Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
                End If
                Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            End If

            'Print #gbFileNO, Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
            'Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)

            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
            End If

            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
            End If

'            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 46 Then
'                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
'            Else
'                Print #gbFileNO,
'            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
            End If




'             objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                Print #gbFileNO, Tab(30); objCounter.CounterNo;
'                Print #gbFileNO, Tab(67); objCounter.CounterNo & " : " & objCounter.CounterDescription
'            End If
'            objUser.SetUser (Rec!intUserID)
'            If objUser.UserID > -1 Then
'                Print #gbFileNO, Tab(27); objUser.UserName;
'                Print #gbFileNO, Tab(67); objUser.UserName
'            End If

            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName;
                    Print #gbFileNO, Tab(70); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(97); objUser.UserName; Tab(97); mInward
                End If
            End If


            'Print #gbFileNO,
        End If

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
finishprinting:
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName

End Sub



Private Sub PrintReceipt_ForNewFormat_Vinod(intVoucherID As Double)
'        gbFileNO = FreeFile
'        gbFileName = "C:\Report.txt"
'        Open gbFileName For Output As #gbFileNO
'        Print #gbFileNO, Chr$(27) + Chr$(80)
'        Print #gbFileNO, String(136, "-")
'        Close #gbFileNO
'        Shell "Print " & gbFileName
'------------------------------------------------------------------------------------------------------------'
'-----------------------------------------Printing in 17 CPI-------------------------------------------------'
'------------------------------------------------------------------------------------------------------------'
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        Dim mRupees As String
        Dim mStr1 As String
        Dim mStr2 As String

        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If

        'FileInitialize
''''        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
''''        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
''''        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
''''        mSql = mSql + " Left Join faPeriodicity On  faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
''''        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
''''        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        objdb.SetConnection mCnn
        Rec.CursorLocation = adUseClient
        Rec.Open "spGetPrintVoucher " & intVoucherID, mCnn, adOpenKeyset, adLockOptimistic

''''''        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
''''''            If Rec.RecordCount > 9 Then
''''''                Rec.Close
''''''                Call PrintSummaryReceiptPTax(intVoucherID)
''''''                Exit Sub
''''''            End If
''''''        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO, Chr$(27) + Chr$(80); ' Set to 10 CPI
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        Print #gbFileNO, Tab(20); gbDoubleWidth; "RECEIPT"; Tab(69); "RECEIPT"; gbDoubleWidthOff
'        Select Case Rec!intInstrumentTypeID
'        Case Is = 1
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
'        Case Is = 4
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
'            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Is = 5
'            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
'            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
'        Case Else
'            Print #gbFileNO,
'        End Select

        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            'Print #gbFileNO, ; gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(65); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff;
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
                If mKMBRFlag Or mSoochikaConnected Then
                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(28); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    Print #gbFileNO, Tab(36); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
            Else
                Print #gbFileNO, Tab(9); gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff; IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(92); gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Trim(Rec!intVoucherNo)); gbBoldOff; gbDoubleWidthOff; Tab(110); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If

            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4

            Print #gbFileNO, Tab(9); gbBold; mName; Tab(64); mName; gbBoldOff

            'Changed for Sujith by Aiby - 24-Mar-2009

'            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
'            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(63); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(9); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(63); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
'            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)

            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                'Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                If Not IsNull(Rec!dtInstrumentDate) Then
                    mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                End If
                'Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                'Print #gbFileNO,
            End Select
            Print #gbFileNO, ; gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(15); PadR(mChequeNo, 30);
            Print #gbFileNO, Tab(57); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff;
            Print #gbFileNO, Tab(72); PadR(mChequeNo, 32);
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff

            'Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(62); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            If Not (IsNull(Rec!vchRefNo)) Then
                Print #gbFileNO, Tab(106); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", PadR(Rec!vchRefNo, 28))
            Else
                Print #gbFileNO,
            End If
                mStr1 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                If Len(mStr1) < 47 Then
'                    mStr1 = mStr1 & String(47 - Len(mStr1), " ")
'                Else
'                    mStr1 = PadR(mStr1, 46)
'                End If
'                'mStr1 = mStr1 & String(52 - Len(mStr1), " ")
'                mStr2 = IIf(IsNull(Rec!vchTransactionType), "", "(" & Rec!vchTransactionType & ")")
'                mStr2 = mStr2 & String(90 - Len(mStr2), " ")
            Print #gbFileNO, PadR(mStr1, 46); Tab(57); PadR(mStr1, 78)
            'Print #gbFileNO,

            ' Line 18 Next
            
            
            
            Dim RecPTAX         As New ADODB.Recordset
            Dim mStartingYear   As Integer
            Dim mStartingPeriod As Integer
            Dim mEndingYear     As Integer
            Dim mEndingPeriod   As Integer
            Dim mNarration      As String
            
            mStartingYear = 2100
            
            
            
            'If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                mSql = "Select faVoucherChild.intAccountHeadID,Sum(fltAmount) As Amount,vchAccountHeadCode,vchAlias,tnyArrearFlag From faVoucherChild"
                mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                mSql = mSql + " Where intVoucherID =" & Rec!intVoucherID
                mSql = mSql + " Group By faVoucherChild.intAccountHeadID,vchAccountHeadCode,vchAlias,tnyArrearFlag"
                mSql = mSql + " Order By tnyArrearFlag Desc,vchAccountHeadCode Desc"
                RecPTAX.Open mSql, mCnn
                While Not RecPTAX.EOF
                    mLoop = mLoop + 1
                    Print #gbFileNO, IIf(IsNull(RecPTAX!vchAccountHeadCode), "", RecPTAX!vchAccountHeadCode);
                    Print #gbFileNO, Tab(37); PadL(Format(RecPTAX!Amount, "0.00"), 9);
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(RecPTAX!vchAlias, 46);
                    Print #gbFileNO, Tab(127); PadL(Format(RecPTAX!Amount, "0.00"), 9)
                    RecPTAX.MoveNext
                Wend
                RecPTAX.Close
                While Not Rec.EOF
                    If mStartingYear > Rec!intYearID Then
                        mStartingYear = Rec!intYearID
                        mStartingPeriod = Rec!tnyPeriodID
                    End If
                    If mEndingYear < Rec!intYearID Then
                        mEndingYear = Rec!intYearID
                    End If
                    mEndingPeriod = Rec!tnyPeriodID
                    Rec.MoveNext
                Wend
                'Rec.Close
                Rec.MoveFirst
                Print #gbFileNO,
                mLoop = mLoop + 1
                mNarration = "(Being the " & Rec!vchTransactionType & " Collected for the Period"
                Print #gbFileNO, mNarration; Tab(54); mNarration
                mLoop = mLoop + 1
                
                mNarration = " of" & str(mStartingYear) & "-" & Trim(Right(str(mStartingYear + 1), 2))
                If mStartingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                ElseIf mStartingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                Else
                    mNarration = mNarration & " to " & str(mEndingYear) & "-" & Trim(Right(str(mEndingYear + 1), 2))
                End If
                
                If mEndingPeriod = 1 Then
                    mNarration = mNarration & " Ist Hf )"
                ElseIf mEndingPeriod = 2 Then
                    mNarration = mNarration & " IInd Hf )"
                Else
                    mNarration = mNarration & ")"
                End If
                mLoop = mLoop + 1
                Print #gbFileNO, mNarration; Tab(52); mNarration
            Else
'               GoTo LB ' To print Property Tax containing less than 9 rows
'           End If
'            Else
            
'LB:
                mLoop = 0
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
                        Case Else:   Print #gbFileNO, Tab(12); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
    
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(26); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    Else
                        Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                    End If
    
                    '==================================================================='
                    ' Receipt Area
                    '==================================================================='
                    Print #gbFileNO, Tab(54); PadL(CStr(mLoop), 2);
                    Print #gbFileNO, Tab(58); PadR(Rec!vchAlias, 46);
                    If Not IsNull(Rec!intYearID) Then
                        mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                    Else
                        mstrYear = ""
                    End If
                    Select Case Rec!tnyPeriodID
                        Case Is = 1: Print #gbFileNO, Tab(106); mstrYear & "/1Hf";
                        Case Is = 2: Print #gbFileNO, Tab(106); mstrYear & "/2Hf";
                        Case Is = 3: Print #gbFileNO, Tab(106); mstrYear & "/F";
                        Case Else:   Print #gbFileNO, Tab(106); mstrYear & "/" & PadR(IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity), 3);
                    End Select
    
                    If Rec!intYearID < gbFinancialYearID Then
                        Print #gbFileNO, Tab(118); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    Else
                        Print #gbFileNO, Tab(127); PadL(Format(Rec!fltAmount, "0.00"), 9)
                    End If
                    Rec.MoveNext
                Wend
            End If
            Rec.MoveFirst

            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 15); Tab(54); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 20);
            Else
'                Print #gbFileNO,'Commented By Vinod
            End If
            Print #gbFileNO, Tab(25); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(116); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"

            Print #gbFileNO, Tab(25); "Total :"; Tab(36); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(116); "Total :"; Tab(128); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)

            'Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            'Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)

            mRupees = Rupees(Rec!TotalAmt)
            If Len(mRupees) < 186 Then
                mRupees = mRupees + String(185 - Len(mRupees), " ")
            End If
            'Print #gbFileNO, Tab(12); Left(mRupees, 34);
            Print #gbFileNO, Tab(54); Left(mRupees, 75)

            'Print #gbFileNO, Tab(12); mID$(mRupees, 33, 34);
            'Print #gbFileNO, Tab(50); mID$(mRupees, 76, 85)

            'Print #gbFileNO,'Commented By Vinod
            Dim mInward As String
            mInward = "1234567890123456"
            Print #gbFileNO, mInward; Tab(27); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 23);
            Print #gbFileNO, Tab(64); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 73)
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 23 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 24, 23);
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 73 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 74, 83)
'            Else
'                Print #gbFileNO,
            End If
            
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 46 Then
                Print #gbFileNO, Tab(27); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 47, 23);
            Else
                Print #gbFileNO,
            End If
            If Len(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription))) > 156 Then
                Print #gbFileNO, Tab(54); mID$(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 157, 83)
'            Else
'                Print #gbFileNO,
            End If
            
            
'             objCounter.SetCounter (Rec!intCounterID)
'            If objCounter.CounterID > 0 Then
'                Print #gbFileNO, Tab(30); objCounter.CounterNo;
'                Print #gbFileNO, Tab(67); objCounter.CounterNo & " : " & objCounter.CounterDescription
'            End If
'            objUser.SetUser (Rec!intUserID)
'            If objUser.UserID > -1 Then
'                Print #gbFileNO, Tab(27); objUser.UserName;
'                Print #gbFileNO, Tab(67); objUser.UserName
'            End If
            
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                objUser.SetUser (Rec!intUserID)
                If objUser.UserID > -1 Then
                    Print #gbFileNO, Tab(27); objCounter.CounterNo; Tab(31); objUser.UserName;
                    Print #gbFileNO, Tab(66); objCounter.CounterNo & " : " & objCounter.CounterDescription; Tab(93); objUser.UserName
                End If
            End If
                

            'Print #gbFileNO,
        End If

        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName
    
End Sub
Private Sub PrintReceipt080410(intVoucherID As Double)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        Dim mStrInWard As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        
        'FileInitialize
        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'objDb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        
        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptPTax(intVoucherID)
                Exit Sub
            End If
        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        
        Select Case Rec!intInstrumentTypeID
        
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(65); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(65); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
            End If
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(65); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            If Not IsNull(Rec!dtInstrumentDate) Then
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
            End If
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(22); gbBold; gbDoubleWidth; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(65); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff; gbDoubleWidthOff
            ' Changed for KMBR By Cijith Sreedharan
            'Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mSoochikaConnected Then
            
                If mKMBRFlag Or mSoochikaConnected Then
                    'Print #gbFileNO, Style("INWARD No", True); "    "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(80); Style("INWARD No", True); "      "; Style(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), True); Tab(130); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                    mStrInWard = PadR(IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo), 6)
                    Print #gbFileNO, gbBold + gbDoubleWidth & "Inw No: "; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(26); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(50); gbBold + gbDoubleWidth & "Inw No:"; mStrInWard; gbBoldOff + gbDoubleWidthOff; Tab(104); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                Else
                    Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
                End If
            Else
                Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            End If
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(67); Style(mName, True)
            
            'Changed for Sujith by Aiby - 24-Mar-2009
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(67); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(67); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(67); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(67); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(67); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
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
                Rec.MoveNext
            Wend
            Rec.MoveFirst
            
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
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
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        Close #gbFileNO
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName
    End Sub
    Private Sub PrintReceipt_B4Change(intVoucherID As Double)
        'Changed at Tvm - Fort Zonal By Aiby
        ' 17-Mar-2009
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        
        'FileInitialize
        mSql = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
        mSql = mSql + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
        mSql = mSql + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
        mSql = mSql + " Where faVouchers.intVoucherID = " & intVoucherID
        objdb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        
        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptPTax(intVoucherID)
                Exit Sub
            End If
        End If
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        
        Select Case Rec!intInstrumentTypeID
        
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(120); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(65); IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            
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
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
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
            
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
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
        'ShellPad
        Shell "Print " & gbFileName
        Kill gbFileName
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
    
    Private Sub ValuesForHiddenColumns(ByVal mRw As Integer)
        On Error Resume Next
        If mRw = 0 Then Exit Sub
        If vsGrid.TextMatrix(mRw, 2) = "" Then  ' Year
            vsGrid.TextMatrix(mRw, 7) = gbFinancialYearID
        Else
            vsGrid.TextMatrix(mRw, 7) = vsGrid.TextMatrix(mRw, 2)
        End If
        
        If vsGrid.TextMatrix(mRw, 3) = "" Then  'Period
            vsGrid.TextMatrix(mRw, 8) = 3
        Else
            vsGrid.TextMatrix(mRw, 8) = vsGrid.Cell(flexcpText, mRw, 3)
            
        End If
        
        If vsGrid.TextMatrix(mRw, 7) < gbFinancialYearID Then  ' Arrear Flag
            vsGrid.TextMatrix(mRw, 9) = 1
        Else
            vsGrid.TextMatrix(mRw, 9) = 0
        End If
        If val(vsGrid.TextMatrix(mRw, 4)) > 0 Then   'Arrear Amount
            vsGrid.TextMatrix(mRw, 11) = val(vsGrid.TextMatrix(mRw, 4))
        Else                                          'Current Amount
            vsGrid.TextMatrix(mRw, 11) = val(vsGrid.TextMatrix(mRw, 5))
        End If
    End Sub
    
    Private Sub LockForm(mLockFlag As Boolean)
        cmdSave.Enabled = mLockFlag
        fraTransactionType.Enabled = mLockFlag
        fraReceiptNo.Enabled = mLockFlag
        fraAccountHead.Enabled = mLockFlag
        fraSubLedger.Enabled = mLockFlag
        fraParty.Enabled = mLockFlag
        txtDescription.Enabled = mLockFlag
    End Sub
    
    Private Sub FillGridYear()
        Dim mLoop As Integer
        Dim mItem As String
        'mItem = ""
        mItem = "#0; "
        'For mLoop = 1991 To gbFinancialYearID
        For mLoop = gbFinancialYearID + 1 To 1971 Step -1
            mItem = mItem & "|#" & mLoop & ";" & CStr(mLoop) & "-" & CStr(mLoop + 1)
        Next
        vsGrid.ColComboList(2) = mItem
    
        
        'Note:- Filling Month
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        objdb.SetConnection mCnn
        mSql = "Select * From faPeriodicity"
        mItem = "#0; "
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                mItem = mItem & "|#" & Rec!intPeriodicityID & "; " & Rec!vchPeriodicity
                Rec.MoveNext
            Wend
        End If
        vsGrid.ColComboList(3) = mItem
        
        'mItem = "#0; "
        'mItem = mItem & "|#" & 1 & "; First Half"
        'mItem = mItem & "|#" & 2 & "; Second Half"
        'mItem = mItem & "|#" & 3 & "; Full Year"
        'vsGrid.ColComboList(3) = mItem
        
        
        
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
            If mDemandID = val(vsGrid.TextMatrix(mLoop, 10)) Then
                mAmount = mAmount + val(vsGrid.TextMatrix(mLoop, 11))
            Else
                mDemandID = val(vsGrid.TextMatrix(mLoop, 10))
            End If
            'vsGrid.Cell(flexcpChecked, mLoop, 12) = 2 'vbUnchecked
        Next mLoop
    End Sub
    
    Public Sub DisplayBuildingDetails()
        Dim arrInput As Variant
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        
        'arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
                    Val(txtWard.Tag), _
                    Val(txtHouseNo1), _
                    Trim(txtHouseNo2))
                    
       '-------------------------------------------------------------'
        ' Changed to Change the Stored Procedure spGetBuildingDetails '
        ' New Stored Procedure is spSanGetSearchBuildingList          '
        '-------------------------------------------------------------'
        'arrInput = Array(cmbZone.ItemData(cmbZone.ListIndex), _
        '     mWardID, _
        '     Val(txtHouseNo1), _
        '     Trim(txtHouseNo2))
        '
        
        '--------------------------------------'
        ' Additions                            '
        '--------------------------------------'
        Dim numBuildingID   As Variant
        Dim numZoneID       As Variant
        Dim intAssessmentYear As Variant
        Dim intWardNo       As Variant
        Dim intDoorNo1      As Variant
        Dim chvDoorNo2      As Variant
        Dim chvName         As Variant
        Dim chvResHName     As Variant
        
        If IsNumeric(mvarSubLedgerID) Then
            numBuildingID = mvarSubLedgerID
            mSubLedgerID = numBuildingID
        Else
            numBuildingID = Null
            mSubLedgerID = Null
        End If
        ''Added on 5 Dec 2019 By anisha to avoid fetch data from local db while web is enabled
        If frmPropertyTax.mDemandWeb = True Then
            mSubLedgerID = SubLedgerID
            Exit Sub
        End If
        'If Trim(txtBuildingNo) <> "" Then numBuildingID = Val(txtBuildingNo) Else numBuildingID = Null
        'If cmbZone.ListIndex > -1 Then numZoneID = cmbZone.ItemData(cmbZone.ListIndex) Else numZoneID = Null
        'If cmbAssessmentYear.ListIndex > -1 Then intAssessmentYear = cmbAssessmentYear.ItemData(cmbAssessmentYear.ListIndex) Else intAssessmentYear = Null
        'If cmbWard.ListIndex > -1 Then intWardNo = cmbWard.ItemData(cmbWard.ListIndex) Else intWardNo = Null
        'If Trim(txtHouseNo1) <> "" Then intDoorNo1 = Val(txtHouseNo1) Else intDoorNo1 = Null
        'If Trim(txtHouseNo2) <> "" Then chvDoorNo2 = Trim(txtHouseNo2) Else chvDoorNo2 = Null
         
        chvName = Null
        chvResHName = Null
        arrInput = Array(numBuildingID, _
            numZoneID, _
            intAssessmentYear, _
            intWardNo, _
            intDoorNo1, _
            chvDoorNo2, _
            chvName, _
            chvResHName)
        
        mBuildingID = -1
        If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
            'Set Rec = objDB.ExecuteSP("spGetBuildingDetails", arrInput, , , mCnn, adCmdStoredProc)
            Set Rec = objdb.ExecuteSP("spSanGetSearchBuildingList", arrInput, , , mCnn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                mBuildingID = Rec!numBuildingID
                txtBuildingNo.Text = Rec!numBuildingID
''''''''                txtWard.Text = Rec!intWardNo
''''''''                txtHouseNo1.Text = Rec!intDoorNo1
''''''''                txtHouseNo2.Text = IIf(IsNull(Rec!chvDoorNo2), "", Rec!chvDoorNo2)
''''''''
''''''''                vchName = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
''''''''                vchHouseName = IIf(IsNull(Rec!chvHouseName), "", Rec!chvHouseName)
''''''''                vchStreetName = IIf(IsNull(Rec!chvResStreetName), "", Rec!chvResStreetName)
''''''''                'vchLocalPlace = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
''''''''                vchMainPlace = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
''''''''
''''''''                txtName.Text = vchName
''''''''                txtHouse.Text = vchHouseName
''''''''                txtStreet.Text = vchStreetName
''''''''                txtLocalPlace.Text = vchLocalPlace
                
                'vchMainPlace = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
                'vchPostOffice = IIf(IsNull(Rec!chvPostoffice), "", Rec!chvPostoffice)
                'vchDistrict = IIf(IsNull(Rec!chvDistrict), "", Rec!chvDistrict)
                'vchPinNumber = IIf(IsNull(Rec!chvPinnumber), "", Rec!chvPinnumber)
                
                'txtAddress.Text = vchName
                'txtAddress.Text = txtAddress.Text & vbCrLf & vchHouseName
                'txtAddress.Text = txtAddress.Text & vbCrLf & vchStreetName
                'txtAddress.Text = txtAddress.Text & vbCrLf & IIf(Len(vchMainPlace), vchMainPlace & ", ", "")
                'txtAddress.Text = txtAddress.Text & vbCrLf & vchPostOffice
                'txtAddress.Text = txtAddress.Text & vbCrLf & vchDistrict
                'txtAddress.Text = txtAddress.Text & " - " & vchPinNumber
                
                'FileNO
                txtRefNo.Text = IIf(IsNull(Rec!chvRefNo), "", Rec!chvRefNo) 'Rec!chvCorpFileNo)
                txtDescription.Text = IIf(IsNull(Rec!chvRemarks), "", Rec!chvRemarks)
                
            Else
                mBuildingID = -1
            End If
        End If
        
    End Sub
    
    Private Sub DisplayBuildingTaxDemands(mBuildingID As Double)
        Dim arrInput    As Variant
        Dim Rec         As New ADODB.Recordset
        Dim objdb       As New clsDB
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
        If objdb.SetConnection(mCnn) Then
            
            Rec.CursorLocation = adUseClient
            Set Rec = objdb.ExecuteSP("spGetPropertyTaxDemands", arrInput, , , mCnn, adCmdStoredProc)
            If Not (Rec.BOF And Rec.EOF) Then
                vsGrid.Rows = 1
                mRows = 1
                vsGrid.MergeCells = flexMergeFree
                While Not Rec.EOF
                    vsGrid.Rows = vsGrid.Rows + 1
                    objAcc.SetAccountID (Rec!intAccountHeadID)
                    vsGrid.Cell(flexcpText, mRows, 0) = Rec!vchAccountHeadCode
                    vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
                    vsGrid.Cell(flexcpText, mRows, 2) = str(Rec!intYearID) & " - " & str(Rec!intYearID + 1)
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
                    vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
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
        Dim objdb As clsDB
        Dim mTotAmt As Variant
        Dim mGrandTot As Variant
        mGrandTotalValidityFlag = False
        vsGridTransactions.Rows = 1
        
        '------------------------------------------------------------------'
        ' Posting of Cash or Bank
        '------------------------------------------------------------------'
        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = mDrAccountHeadID
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 1
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(val(txtTotal), "0.00")
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = "" 'RecTransactionHeads!intPostingHeadID
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
        
        For mLoopCount = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoopCount, 0) <> "" And vsGrid.Cell(flexcpText, mLoopCount, 14) <> 1 Then
                vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                
                'Debug.Print mLoopCount, vsGrid.TextMatrix(mLoopCount, 6), Val(vsGrid.TextMatrix(mLoopCount, 11))
                
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = mLoopCount + 1
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = vsGrid.TextMatrix(mLoopCount, 6)
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = 0
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = val(vsGrid.TextMatrix(mLoopCount, 11))
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 4) = mDrAccountHeadID
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 6) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 7) = ""
                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 8) = ""
                If mRoundOffDecimalPlace Then
                    mTotAmt = mTotAmt + val(Format(val(vsGrid.TextMatrix(mLoopCount, 11)), "#0"))
                Else
                    mTotAmt = mTotAmt + val(vsGrid.TextMatrix(mLoopCount, 11))
                End If
            Else
                'Exit For
            End If
        Next mLoopCount
        mTotAmt = Format(mTotAmt, "0.00")
        mGrandTot = Format(val(txtGrandTotal.Text), "0.00")
        If mTotAmt <> mGrandTot Then
            mGrandTotalValidityFlag = True
        End If
    End Sub
    
    Private Sub ListPostingHeadsInGrid(mTransactionType As Long, Optional mGroupID As Variant = Null)
                Dim mSql As String
                Dim RecTransactionHeads As New ADODB.Recordset
                Dim mLoopCount As Long
                Dim mLoop As Long
                Dim mAmt As Double
                vsGridTransactions.Rows = 1
                mSql = "Select * From faTransactionTypeChild Where intTransactionTypeID = " & mTransactionType & " Order By intOrder"
                Set RecTransactionHeads = GetRecordSet("spGetTransactionTypePostingHeads " & mTransactionType)
                For mLoopCount = 1 To vsGrid.Rows - 1
                    While Not RecTransactionHeads.EOF
                        If vsGrid.Cell(flexcpChecked, mLoopCount, 12) = 1 Then
                            If RecTransactionHeads!intAccountHeadID = val(vsGrid.TextMatrix(mLoopCount, 6)) Then
                                If val(vsGrid.TextMatrix(mLoopCount, 9)) Then    ' Arrear Flag = True
                                    mAmt = val(vsGrid.TextMatrix(mLoopCount, 4)) ' Amount from the Arrear Column
                                Else
                                    mAmt = val(vsGrid.TextMatrix(mLoopCount, 5)) ' Amount from the Current Column
                                End If
                                '--------------------------------------------------------------------------------'
                                ' Check whether the posting head is already there in the Transaction (Hidden)    '
                                ' Grid. If found add the Amount. Other wise add a new row in the Grid            '                                                 '
                                '--------------------------------------------------------------------------------'
                                For mLoop = 1 To vsGridTransactions.Rows - 1
                                    If RecTransactionHeads!intPostingHeadID = val(vsGridTransactions.TextMatrix(mLoop, 1)) Then
                                        vsGridTransactions.TextMatrix(mLoop, 3) = val(vsGridTransactions.TextMatrix(mLoop, 3)) + mAmt
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
                                vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 5) = val(vsGrid.TextMatrix(mLoopCount, 10))
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
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(val(txtTotal), "0.00")
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
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(val(txtTotal), "0.00")
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
                If val(txtAdvance.Text) > 0 Then
                RecTransactionHeads.MoveFirst
                While Not RecTransactionHeads.EOF
                    If RecTransactionHeads!vchAccountHeadCode = mAcHeadCodeAdvance Then
                        vsGridTransactions.Rows = vsGridTransactions.Rows + 1
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 0) = RecTransactionHeads!intOrder
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 1) = RecTransactionHeads!intAccountHeadID
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 2) = RecTransactionHeads!tinDebitOrCredit
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(val(txtAdvance), "0.00")
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
                        vsGridTransactions.TextMatrix(vsGridTransactions.Rows - 1, 3) = Format(val(txtAdvance), "0.00")
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
        Dim mAdvAmt As Variant
        
        For mCount = 1 To vsGrid.Rows - 1
'            If vsGrid.Cell(flexcpText, mCount, 0) = gbAcHeadCodeAdvancePTax And mvarDemandBasedFlag Then
             If val(vsGrid.Cell(flexcpText, mCount, 14)) = 1 Then   ' Changed for Advance calculation For all
                If vsGrid.Cell(flexcpText, mCount, 10) <> "" Then   ''Checking the Demand ID is Null By Sinoj
                    If val(vsGrid.Cell(flexcpText, mCount, 4)) > 0 Then
                        mAdvAmt = mAdvAmt + val(vsGrid.Cell(flexcpText, mCount, 4))
                    Else
                        mAdvAmt = mAdvAmt + val(vsGrid.Cell(flexcpText, mCount, 5))
                    End If
                    GoTo NextRow
                End If
            End If
            If val(vsGrid.TextMatrix(mCount, 4)) Then
                mAmtArrear = mAmtArrear + val(vsGrid.Cell(flexcpText, mCount, 4))
            Else
                mAmtCurrent = mAmtCurrent + val(vsGrid.Cell(flexcpText, mCount, 5))
            End If
NextRow:
        Next
        txtTotalArrear.Text = Format(mAmtArrear, "0.00")
        txtTotalCurrent.Text = Format(mAmtCurrent, "0.00")
        txtGrandTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
        If mAdvAmt > 0 Then
            txtAdvance = Format(mAdvAmt, "0.00")
        Else
            txtAdvance = ""
        End If
        txtTotal.Text = val(txtGrandTotal) - val(txtAdvance)
        txtRoundOff.Text = Format(RoundOffAdjustment(val(txtTotal)), "0.00")
        txtTotal.Text = Format(val(txtTotal) + val(txtRoundOff), "0.00")
        
        If val(txtAdvance.Text) Then
            txtAdvance.Visible = True
            lblAdvance.Visible = True
        Else
            txtAdvance.Visible = False
            txtAdvance.Visible = False
            txtAdvance.Text = ""
        End If
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
                lstMasters.Width = txtTransactionType.Width + 1500
                lstMasters.SetFocus
            Case Is = 2 ' Instruments
                'Note:- Only in JSK the Cash and Bank Account Transactions Record
                If gbTransactionDate < DateValue(gbOnlinedate) And gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    'Modified By Anisha On 31.10.2011 To Exclude Directly Debited To Bank Instrument
'                    mSQL = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes "  ' Old Query Commented on 30.10.2011
                    mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID<>10"
                Else
                    If gbCounterSectionID = gbJSKSectionID Then
                    'Modified By sunil for new Instrumet Type CardPaymets with ID=11
                        mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (1,2,3,4,5,8,9,11)" ' & gbInstrumentCash & "Or intInstrumentTypeID =" & gbInstrumentCheque
                    Else
                       ' mSQL = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (6,7,9,10,11)" '<>" & gbInstrumentCash  ' Old Query Commented on 30.10.2011
                        
'                        'Modified By Anisha On 31.10.2011 To Exclude Directly Debited To Bank Instrument
'                        'mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (6,7,9,11)" '<>" & gbInstrumentCash
                        If gbLBPanchayat = 1 Then
                            Select Case txtTransactionType.Tag
                                Case 119, 120, 121, 122, 123:
                                    mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (4,5,6,7,9,11)" '<>" & gbInstrumentCash
                                Case Else
                                    mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (6,7,9,11)" '<>" & gbInstrumentCash
                            End Select
                        Else
                            'Modified By Anisha On 31.10.2011 To Exclude Directly Debited To Bank Instrument
                            mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID In (6,7,9,11)" '<>" & gbInstrumentCash
                        End If
                    End If
                End If
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
        'txtBank.Text = ""
        'txtDated.Text = ""
        'txtPlace.Text = ""
    End Sub
    
    Private Sub FormInitialize()
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mOutput As Variant
        Dim mStr As String
        txtTransactionType.Enabled = True
        cmdSave.Enabled = True
        cmdPrint.Enabled = True
        mGrandTotalValidityFlag = False
        mIRVoucherDate = Null
        
        Call ShowFrames(1)
        Call LockForm(True)
        Call CheckInterruptReceiptRequestStatus
        Call FillGridYear
        
        mUserSessions = -1
        mBuildingID = -1
       ' mSubLedgerID = Null
        mPTaxFormLoadFlag = False
        
        objdb.SetConnection mCnn
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
        txtDate.Text = DdMmmYy(gbTransactionDate)
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        txtDemandPrefix.Text = ""
        txtDemandPrefix.Tag = ""
        txtDemandNo.Text = ""
        txtDemandNo.Tag = ""
        txtOutDoorStaff.Text = ""
        txtOutDoorStaff.Tag = ""
        lblOutDoorStaff(8).Caption = ""
        txtOutDoorStaff.Visible = False
        txtOutDoorStaff.Enabled = False
        lblOutDoorStaff(8).Enabled = False
        txtAdvance.Text = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        txtInstrument.Text = ""
        txtInstrument.Tag = ""
        txtDated.Text = ""
        txtInstNo.Text = ""
        
        lblAdminNote.Caption = ""
        'lblAdminNoteCaption.Caption = True
        txtBank.Text = ""
        txtPlace.Text = ""
        
        cmdSearchAccountHead.Enabled = True
        
        If chkGroupReceipt.Value = 0 Then
            Call ClearAddressVariables
            txtBuildingNo.Text = ""
            txtWard.Text = ""
            txtWard.Tag = ""
            txtRefNo.Text = ""
            txtRefNo.Tag = ""
            txtHouseNo1.Text = ""
            txtHouseNo2.Text = ""
            txtPayee.Text = ""
            txtInitial1.Text = ""
            txtInitial2.Text = ""
            txtInitial3.Text = ""
            txtInitial4.Text = ""
            txtHouseName.Text = ""
            'txtAddress.Text = ""
            vchName = ""
            vchHouseName = ""
            vchStreetName = ""
            vchMainPlace = ""
            vchPostOffice = ""
            vchDistrict = ""
            vchPinNumber = ""
        End If
        
        txtTotalArrear.Text = ""
        txtTotalCurrent.Text = ""
        txtTotal.Text = ""
        txtRoundOff.Text = ""
        txtDescription.Text = ""
        txtGrandTotal.Text = ""
        
        vsGrid.Editable = flexEDKbdMouse
        vsGrid.Rows = 1
        vsGrid.Rows = 13
        Call SetDefaultSettings
        fraSubLedger.Visible = False
        lblLocation.Caption = "Location"
        lblShop.Caption = "Shop"
        lblName.Caption = "Building Name"
        lblLessee.Caption = "Lessee"
        txtFirstLine.Text = ""
        txtSecondLine.Text = ""
        txtThirdLine.Text = ""
        txtMemo.Text = ""
        mvarDemandBasedFlag = False
        mRentSearchFormLoadedFlag = False
        chkIntrNoSuffix.Value = vbUnchecked
        If mInterruptedModeFlag Then
            'If gbLBPanchayat Then
                If IsNull(gbOnlinedate) = False Then
                    If CDate(mdtDate) < DateValue(gbOnlinedate) Then
                        chkIntrNoSuffix.Visible = True
                        txtDate.Text = Format(mdtDate, "dd-mmm-yyyy")
                        'txtIntruptNoSuffix.Visible = True
                    End If
                End If
            'End If
            If gbManualReceiptNewBool = False Then
                arrInput = Array(gbCounterID)
            Else
                arrInput = Array(gbCounterID, 1)        '---- Second Parameter "1" for New Interrrupted Recipt Format Sinoj
            End If
      
''''            Set Rec = objDb.ExecuteSP("spGetNextReceiptNoInterrupted", arrInput, arrOutPut, , mcnn, adCmdStoredProc)
''''            If IsNull(arrOutPut(0, 0)) Then
''''                MsgBox "Check whether there is any Active Manual Book issued to this Counter", vbInformation
''''                cmdSave.Enabled = False
''''                Exit Sub
''''            Else
''''                If IsArray(arrOutPut) Then
''''                    txtReceiptNo.Text = arrOutPut(0, 0)
''''                    If val(txtReceiptNo.Text) <= 0 Then cmdSave.Enabled = False
''''                End If
''''            End If
''''            Rec.Close


            '-------------------------------------------------------------------------------'
            ' CHECKING INTERRUPTED REGISTER MODE [MINU]  ADDED FOR IR REGISTER BY MINU
            '-------------------------------------------------------------------------------'
            'NOTE : 1:= NEW RECEIPT/CANCEL RECEIPT
            '       3:= RECEIPT WITH SUFIX No
           If mInterruptedRegister = 1 Or mInterruptedRegister = 3 Then
                txtReceiptNo.Text = mInterruptedRegisterReceiptNo   ' Passing from frmInterruptReceiptRegister as Property
                txtDate.Text = mInterruptedRegisterReceiptDate      ' ""
                mdtDate = mInterruptedRegisterReceiptDate
                chkIntrNoSuffix.Visible = False
                txtIntruptNoSuffix.Visible = True
            Else
                ' OLD INTERRUPTED RECEIPT MODE
                Set Rec = objdb.ExecuteSP("spGetNextReceiptNoInterrupted", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
                If IsNull(arrOutPut(0, 0)) Then
                    MsgBox "Check whether there is any Active Manual Book issued to this Counter", vbInformation
                    cmdSave.Enabled = False
                    Exit Sub
                ElseIf IsArray(arrOutPut) Then
                    txtReceiptNo.Text = arrOutPut(0, 0)
                    If val(txtReceiptNo.Text) <= 0 Then cmdSave.Enabled = False
                End If
                Rec.Close
            End If
           '-------------------------------------------------------------------------------'
 
        Else
            arrInput = Array(gbCounterID, val(txtInstrument.Tag), gbFinancialYearID)
            Set Rec = objdb.ExecuteSP("spGetNextReceiptNo", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
            If IsArray(arrOutPut) Then
                txtReceiptNo.Text = arrOutPut(0, 0)
            End If
            Rec.Close
        End If
        
        'Added On 4.9.11 by Sunil
        If mZonal = 1 Then
           mdtDate = Format(mZoneDate, "dd/mmm/yyyy")  'Format(mDemandTrDate, "dd/mmm/yyyy")
        ElseIf mWebExtractMode = True Then
            mdtDate = Format(mWebExtractDate, "dd/mmm/yyyy")
        End If
        If Len(txtReceiptNo) > 6 Then
            mStr = Left(txtReceiptNo, 6)
            mStr = mStr + "-" + Right(txtReceiptNo, Len(txtReceiptNo) - 6)
'            mStr = Right(txtReceiptNo, 5)
'            mStr = Left(txtReceiptNo, Len(txtReceiptNo) - 5) + "-" + mStr
            txtReceiptNo.Text = mStr
        End If
        If gbPDEMode Then
            vsGrid.Rows = 50
        End If
        txtDemandPrefix.Width = 1080
        chkLinkDemand.Visible = False
        lblZone.Visible = False
        txtZone.Visible = False
        Dim index As Integer
        
        For index = 0 To cmbDZone.ListCount - 1
            If cmbDZone.ItemData(index) = gbLocationID Then
                cmbDZone.ListIndex = index
            End If
        Next
                
        frmReceiptsCounter.txtWardNo.Enabled = True
        frmReceiptsCounter.txtDoorNo1.Enabled = True
        frmReceiptsCounter.txtDoorNo2.Enabled = True
        frmReceiptsCounter.txtName.Enabled = True
        frmReceiptsCounter.txtInit1.Enabled = True
        frmReceiptsCounter.txtInit2.Enabled = True
        frmReceiptsCounter.txtInit3.Enabled = True
        frmReceiptsCounter.txtInit4.Enabled = True
        frmReceiptsCounter.txtHouse.Enabled = True
        frmReceiptsCounter.txtStreet.Enabled = True
        frmReceiptsCounter.txtLocalPlace.Enabled = True
        frmReceiptsCounter.txtMainPlace.Enabled = True
        frmReceiptsCounter.txtPost.Enabled = True
        frmReceiptsCounter.txtPin.Enabled = True
        frmReceiptsCounter.txtPhone.Enabled = True
        
        txtDate.Text = DdMmmYy(mdtDate)
        lblBuildingNo.Caption = "Building No"
        
        frameTransaction.Visible = False
        txtCardTransaction.Text = ""
        
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
        Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone WHERE intLBID =" & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
        Call PopulateList(cmbDZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone WHERE intLBID =" & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster)
    End Sub
    
    
    Private Sub FillSeats()
        ''Added by Cijith ?!! - Aiby 25-05-2009 Integration with KMBR
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim mQuery As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.DBMaster
        mQuery = "Select Left(Convert( VarChar(10),numSeatID), 6) As Prefix From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        Rec.Open mQuery, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mSeatPrefix = IIf(IsNull(Rec!Prefix), "", Rec!Prefix)
        End If
        Rec.Close
        
        mSql = "Select chvSeatTitle, Right(Convert( VarChar(10),numSeatID), 5) From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        'mSQL = "Select chvSeatTitle, numSeatID From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        Call PopulateList(cmbSeat, mSql, , True, True, True, enuSourceString.DBMaster)
    End Sub
    
    Private Sub FillTransactionTypes()
        Dim mSql As String
        mSql = "Select vchTransactionType, intTransactionTypeID, intGroupID From faTransactionType Where intGroupID = 10 Or (intGroupID=20 And intTransactionTypeID in (1141,1151,1161,1171,1181,1191)) Order By vchTransactionType"
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
    
        mDefaultTransactionTypeID = gbDefaultTransactionTypeID  'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultTransactionTypeID"))
        mDefaultAccountHeadCode = gbAcHeadCodeCash              'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultAccountHeadCode")
        mDefaultInstrumentID = gbInstrumentCash                 'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultInstumentID"))
        mDefaultBankID = gbDefaultBankID                        'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID"))
        mDefaultZoneID = gbnumZonalID                           'Val(ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultZone"))
        
        If mDefaultZoneID > 0 Then
            For mLoopCount = 0 To cmbZone.ListCount - 1
                If cmbZone.ItemData(mLoopCount) = mDefaultZoneID Then
                    cmbZone.ListIndex = mLoopCount
                    Exit For
                End If
            Next
        End If
        
        objTranType.SetTransactionType (mDefaultTransactionTypeID)
        If gbCounterSectionID = gbJSKSectionID Then
            objAc.SetAccountCode (mDefaultAccountHeadCode)
            objInstruments.SetInstrumentType (mDefaultInstrumentID)
        Else
'            objAc.SetAccountID (mDefaultBankID)
'            objInstruments.SetInstrumentType (gbInstrumentCheque)
        End If
        
        objBank.SetBankInfo (mDefaultBankID)
        txtTransactionType.Text = objTranType.TransactionType
        txtTransactionType.Tag = objTranType.TransactionTypeID
        Call txtTransactionType_DblClick
        
    txtAccountHead.Text = objAc.AccountHead & " [ " & objAc.AccountCode & " ]"
        txtAccountHead.Tag = objAc.AccountHeadID
        cmdSearchAccountHead.Tag = objAc.GroupID
        If Not IsNull(objInstruments.InstrumentTypeID) Then
            txtInstrument.Text = objInstruments.InstrumentType
            txtInstrument.Tag = objInstruments.InstrumentTypeID
        End If
        mDefaultBankHeadCode = objBank.BankAccountHeadCode
        mAcHeadCodeAdvance = 350410101
    End Sub



    Private Sub chkIntrNoSuffix_Click()
        Dim Rec             As New ADODB.Recordset
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim arrInput        As Variant
        Dim arrOutPut       As Variant
        Dim mSql            As String
        Dim mString         As String
            If chkIntrNoSuffix.Value = vbChecked Then
            txtIntruptNoSuffix.Visible = True
'''''''''''             arrInput = Array(gbCounterID)
'''''''''''             Set Rec = objDB.ExecuteSP("spGetNextReceiptNoInterrupted", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
'''''''''''                If IsNull(arrOutPut(0, 0)) Then
'''''''''''                    MsgBox "Check whether there is any Active Manual Book issued to this Counter", vbInformation
'''''''''''                    cmdSave.Enabled = False
'''''''''''                    Exit Sub
'''''''''''                Else
'''''''''''                    If IsArray(arrOutPut) Then
'''''''''''                        txtReceiptNo.Text = arrOutPut(0, 0) - 1
'''''''''''                        If val(txtReceiptNo.Text) <= 0 Then cmdSave.Enabled = False
'''''''''''                    End If
'''''''''''                End If
'''''''''''                Rec.Close
'''''''''''                mSQL = "Select vchDoorNoP3 From faVouchers Where intVoucherNo=" & txtReceiptNo.Text
'''''''''''                Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
'''''''''''                mString = "A"
'''''''''''                If Not (Rec.BOF And Rec.EOF) Then
'''''''''''                    While Not (Rec.EOF)
'''''''''''                       If Rec!vchDoorNoP3 <> "A" Then
'''''''''''                            mString = mString & "," & Rec!vchDoorNoP3
'''''''''''                       End If
'''''''''''                       Rec.MoveNext
'''''''''''                    Wend
'''''''''''                End If
'''''''''''                txtIntruptNoSuffix.ToolTipText = mString
'''''''''''            Rec.Close
        ''''''Modified To Autogenerate Suffix
          arrInput = Array(gbCounterID)
          Set Rec = objdb.ExecuteSP("spGetNextReceiptNoInterrupted", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
          If IsNull(arrOutPut(0, 0)) Then
                MsgBox "Check whether there is any Active Manual Book issued to this Counter", vbInformation
                cmdSave.Enabled = False
                Exit Sub
          Else
                If IsArray(arrOutPut) Then
                    'txtReceiptNo.Text = arrOutPut(0, 0) - 1
                    '''---------------------------------------------
                    '''To Get Previously Entered Interrupt ReceiptNo of same Counter Book
                    '''
                    mSql = "Select Top 1 intVoucherNo,* from FaVouchers V"
                    mSql = mSql + " Where tnyVoucherTypeID = 10 And tnyVoucherGroupID = 4 And ( tnyCancelFlag is Null or tnyCancelFlag <> 4 ) And intBookNo=(Select Min(intBookID) From faInterruptedReceiptBooks Where intCounterID = " & gbCounterID & " And tnyClosed <> 1) Order By intVoucherID Desc"
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.EOF And Rec.BOF) Then
                        If arrOutPut(0, 0) - 1 = Rec!intVoucherNo Then
                            txtReceiptNo.Text = arrOutPut(0, 0) - 1
                        Else
                            txtReceiptNo.Text = Rec!intVoucherNo
                        End If
                    Else
                        MsgBox "No Interrupted Receipt Generated For This Counter ", vbInformation
                        chkIntrNoSuffix.Value = vbUnchecked
                        Exit Sub
                    End If
                    Rec.Close
                    mSql = "Select top 1 vchDoorNoP3 From faVouchers Where tnyVoucherGroupID=4 And intVoucherNo=" & txtReceiptNo.Text & " Order By intVoucherID Desc"
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.BOF And Rec.EOF) Then
                         If Rec!vchDoorNoP3 = "Z" Then
                            MsgBox "Suffic for the Voucher reached its Maximum.. ", vbInformation
                            chkIntrNoSuffix.Value = vbUnchecked
                            Exit Sub
                         ElseIf IsNull(Rec!vchDoorNoP3) Then
                            txtIntruptNoSuffix.Text = CStr("B")
                         Else
                            txtIntruptNoSuffix.Text = Chr(Asc(Rec!vchDoorNoP3) + 1)
                         End If
                    Else
                        If txtReceiptNo.Text <> "" Then
                            
                        End If
                    End If
                    If val(txtReceiptNo.Text) <= 0 Then cmdSave.Enabled = False
                End If
          End If
        Else
             txtIntruptNoSuffix.Text = ""
             txtIntruptNoSuffix.Visible = False
             
             arrInput = Array(gbCounterID)
             Set Rec = objdb.ExecuteSP("spGetNextReceiptNoInterrupted", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
                If IsNull(arrOutPut(0, 0)) Then
                    MsgBox "Check whether there is any Active Manual Book issued to this Counter", vbInformation
                    cmdSave.Enabled = False
                    Exit Sub
                Else
                    If IsArray(arrOutPut) Then
                        txtReceiptNo.Text = arrOutPut(0, 0)
                        If val(txtReceiptNo.Text) <= 0 Then cmdSave.Enabled = False
                    End If
                End If
            Rec.Close
        End If
    End Sub
    Private Sub chkLinkDemand_Click()
       If chkLinkDemand.Value = 1 Then   'Note: If gbLinkWithDandOPFA = 1 Then
            txtDemandPrefix.Width = 2280
        Else
            txtDemandPrefix.Width = 1080
        End If
    End Sub
    Private Sub chkRoundOff_Click()
        mRoundOffDecimalPlace = chkRoundOff.Value
    End Sub
    Private Sub cmbZone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    
    Private Sub cmdCancel_Click()
        Dim mPFlag As Integer
        Dim mNewVoucherID As Double
        
        Select Case mUserSessions
            Case Is = 1
                Call FormInitialize
            Case Is = -1
                If Not IsNull(mVoucherID) Then
                    If mInterruptedModeFlag Then
                        MsgBox "Interrupted Receipts Are not Allowed to Cancel", vbApplicationModal
                        'Exit Sub
                        Unload Me
                    Else
                        If MsgBox("Cancel Receipt?", vbYesNo) = vbYes Then
                            'If val(txtInstrument.Tag) <> 6 Then
                                If Not IsNull(mVoucherID) Then
                                    'Added by sunil for zonal collection
                                    If mZonal = 1 Then
                                        frmCancelReceipt.ZonalCollection = 1
                                    End If
                                    mRePrintFlag = 0
                                    frmCancelReceipt.ReceiptNO = txtReceiptNo.Text
                                    frmCancelReceipt.LoadMode = 1
                                    frmCancelReceipt.InstrumentTypeID = txtInstrument.Tag
                                    frmCancelReceipt.Show vbModal
                                Else
                                    MsgBox "Sorry, No Receipts in System Memory", vbInformation
                                End If
                                
                                If mRePrintFlag = 1 Then
                                    If gbLBPanchayat = 1 Then                'ADDED BY MINU ON 26/09/2011
                                        If gbLBID = 266666 Then
                                            PrintReceipt (mVoucherID)
                                        Else
                                            'PrintReceipt_ForNewFormat (mVoucherID)
                                            'Dim mPFlag As Integer
                                            
                                            mNewVoucherID = mVoucherID
                                            mPFlag = PrintReceipt_ForNewFormatRes(mNewVoucherID)
                                            Kill "C:\Report.txt"
                                            
                                        End If
                                    Else
                                        mNewVoucherID = mVoucherID
                                        mPFlag = PrintReceipt(mNewVoucherID)
                                        Kill "C:\Report.txt"
                                    End If
                                    'PrintReceipt (mVoucherID)
                                End If
                            'Else
                            '    MsgBox "Letter of Allotment Vouchers cant cancel...Only Reverse..."
                            '    Exit Sub
                            'End If
                        Else
                            Unload Me
                        End If
                    End If
                Else
                    Unload Me
                End If
        End Select
    End Sub
    Private Sub cmdFind_Click()
        Call DisplayBuildingDetails
        Call DisplayBuildingTaxDemands(mBuildingID)
    End Sub
    Private Sub cmdNew_Click()
        If Me.InterruptEditMode Then
            cmdNew.Enabled = False
            Exit Sub
        ElseIf mWebExtractMode = True Then
            cmdNew.Enabled = False
            
            Exit Sub
        End If
        
        Call FormInitialize
    End Sub
    
    Private Sub cmdPrint_Click()
        Call cmdSave_Click
    End Sub

    Private Sub cmdSave_Click()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim objFunctions As New clsFunction
        Dim objFunctionaries As New clsFunctionary
        Dim mFunctionaryID  As Variant
        Dim mFunctionId As Variant
        Dim mLoopCount As Long
        Dim mLoop As Long
        Dim Rec As New ADODB.Recordset
        Dim mDemandID As Variant
        Dim objAc As New clsAccounts
        Dim lSoochikaCurrentNo As Variant
        Dim mBoolGiveAdvanceToSanchaya As Boolean
        Dim mAdvAmtToSanchaya As Variant
        Dim mSql As String
        Dim mSourceOfFundID As Integer
        Dim mPayOrderNo As Variant
        Dim ss As Variant
        Dim mTransactionDate As Date
        Dim mYearID As Integer
        Dim mPFlag As Integer
        Dim mStr As String
        
        
        If mPreviousYearMode = 1 Then
            If IsDate(txtDate) Then
                mTransactionDate = CDate(txtDate.Text)
                If Not (mTransactionDate >= DateAdd("yyyy", -1, gbStartingDate) And mTransactionDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                    MsgBox "Transaction Date found mismatch!", vbInformation
                    Exit Sub
                End If
                mYearID = gbFinancialYearID - 1
            Else
                MsgBox "Transaction Date found mismatch!", vbInformation
                Exit Sub
            End If
        ElseIf mInterruptedModeFlag = True Then
            mYearID = gbFinancialYearID
            mTransactionDate = DdMmmYy(txtDate.Text)
        ElseIf mWebExtractMode = True Then
            mYearID = gbFinancialYearID
            mTransactionDate = DdMmmYy(mWebExtractDate)
        Else
            mYearID = gbFinancialYearID
            mTransactionDate = gbTransactionDate
        End If
        
        
         '''Added Anisha
        If CDate(mTransactionDate) <= GetLastReconDate(val(txtAccountHead.Tag)) Then
            mStr = ""
            mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
            mStr = mStr + " No new Transaction is allowed to Enter during the period."
            MsgBox mStr, vbInformation
            txtAccountHead.Tag = -1
            txtAccountHead.Text = ""
            Exit Sub
        End If
        
        If gbLBPanchayat = 1 Then
            If mTransactionDate < gbRPOnlinedate Then
                mSql = "Date must be greater than the ONLINE date" & vbCrLf
                mSql = mSql & "[" & DdMmmYy(CDate(gbRPOnlinedate)) & "]"
                MsgBox mSql, vbInformation
                cmdSave.Enabled = False
                Exit Sub
            End If
        End If
        
        If mInterruptEditMode Then
            Call funUpdate
            Exit Sub
        End If
        
        If (val(txtTotal)) <= 0 Then
            MsgBox "Please Check the Amount...!", vbInformation
            Exit Sub
        End If
        
        If Trim(txtName.Text) = "" Then
            MsgBox "Please Enter the name of Person who is remitting the Amount", vbInformation
            txtName.SetFocus
            Exit Sub
        End If
        
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            If txtAccountHead.Tag = gbAcHeadIDCash Then
                MsgBox "Cash Not Allowed in Accountant's Login", vbInformation, "Saankhya"
                Exit Sub
            End If
        End If
        
        If vsGrid.Rows > 1 Then
            If vsGrid.TextMatrix(1, 0) = "" Or _
               (val(vsGrid.TextMatrix(1, 4)) <= 0 And _
                val(vsGrid.TextMatrix(1, 5)) <= 0) Then
                MsgBox "No Item has been entered in the Grid..!", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "No Item has been entered in the Grid..!", vbInformation
        End If
        If val(txtInstrument.Tag) = 0 Then
            MsgBox "Please Select An Instrument", vbInformation
            txtInstrument.SetFocus
            Exit Sub
        End If
        If val(txtInstrument.Tag) = 5 Then
            If Trim(txtInstNo) = "" Then
                MsgBox "Please Enter the Cheque No.", vbInformation
                txtInstNo.SetFocus
                Exit Sub
            End If
            If Not IsDate(txtDated) Then
                MsgBox "Please Enter the Cheque Date", vbInformation
                txtDated.SetFocus
                Exit Sub
            End If
            If Trim(txtBank) = "" Then
                MsgBox "Please Enter the Name of Bank/Branch, Who issued the cheque...", vbInformation
                txtBank.SetFocus
                Exit Sub
            End If
            If Trim(txtPlace) = "" Then
                MsgBox "Please Enter the place of Bank issued the Cheque..", vbInformation
                txtPlace.SetFocus
                Exit Sub
            End If
        ElseIf val(txtInstrument.Tag) = 1 Then
            If txtAccountHead.Tag <> gbAcHeadIDCash Then
                MsgBox "Please select proper Account Head", vbInformation
                cmdSearchAccountHead.SetFocus
                Exit Sub
            End If
        End If
        If val(txtTransactionType.Tag) < 1 Then
            MsgBox "Please Select TransactionType..", vbInformation
            txtTransactionType.SetFocus
            Exit Sub
        Else
            If mWebExtractDate = True Then
                    ''1141,1151,1161,1171,1181,1191 project payment
                If val(txtTransactionType.Tag) = 1141 Or val(txtTransactionType.Tag) = 1151 Or val(txtTransactionType.Tag) = 1161 Or _
                    val(txtTransactionType.Tag) = 1171 Or val(txtTransactionType.Tag) = 1181 Or val(txtTransactionType.Tag) = 1191 Then
                Else
                    MsgBox "please select proper Transaction type", vbApplicationModal
                    Exit Sub
                End If
            Else
                mTransactionType = val(txtTransactionType.Tag)
            End If
        End If
        
        'added On 23/06/2012 By anisha
            'To Avoid the use of refund of payment for normal demand ()
        If mTransactionType = gbTransactionTypeRefundOfPayment Then     ' Payment Order Cancellation
            If val(txtDemandNo.Tag) > 0 Then
                objdb.SetConnection mCnn
                Rec.Open "Select intKeyID2 From faIDemandTBL Where numDemandID = " & val(txtDemandNo.Tag), mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If Not IsNull(Rec!intKeyID2) Then
                    'mPayOrderNo = IIf(IsNull(Rec!intKeyID2), -1, Rec!intKeyID2)
                        mPayOrderNo = IIf(IsNull(Rec!intKeyID2), -1, Rec!intKeyID2)
                    Else
                        MsgBox "This Transaction type is used for the demand " & vbNewLine & " that Done through Payment Order Cancellation", vbApplicationModal
                        Exit Sub
                    End If
                End If
                Rec.Close
                Rec.Open "Select * From intKeyID2 = " & mPayOrderNo, mCnn
                If (Rec.EOF And Rec.BOF) Then
                    MsgBox "This Transaction type is used for the demand Done" & vbNewLine & " through Payment Order Cancellation", vbApplicationModal
                    Exit Sub
                End If
                Rec.Close
                Set mCnn = Nothing
            End If
        End If
            '------------
        If mZonal <> 1 Then             'Added by sunil
            If mTransactionType = gbTransactionTypePTax Then
                If mDemandMode <= 1 Then
                    If val(txtWardNo) < 1 Then
                        MsgBox "Enter the Ward No ", vbInformation
                        txtWardNo.SetFocus
                        Exit Sub
                    End If
                    
                    If val(txtDoorNo1) < 1 Then
                        MsgBox "Enter the valid Door No ", vbInformation
                        txtDoorNo1.SetFocus
                        Exit Sub
                    End If
                End If
            ElseIf mTransactionType = gbTransactionTypeOutDoor Then
                mSubLedgerID = val(txtOutDoorStaff.Tag)
        End If
        End If
        mDemandID = -1
        If val(txtDemandNo.Tag) > 0 Then
            mDemandID = txtDemandNo.Tag
        End If
        mPayOrderNo = -1
        '-------------------------------------------------------------------------------'
        ' S E R V E R   D A T E   V A R I F I C A T I O N                               '
        '-------------------------------------------------------------------------------'
        If mInterruptedModeFlag = False Then
            objdb.SetConnection mCnn
            Set Rec = mCnn.Execute("Select GetDate()")
            If IsDate(Rec.Fields(0)) Then
                mdtDate = DdMmmYy(Rec.Fields(0))
            Else
                MsgBox "Didn't able to Access Server Date", vbInformation
                Exit Sub
            End If
            Rec.Close
            Set mCnn = Nothing
        End If
        
        '-------------------------------------------------------------------------------'
        ' INTERRUPTED MODE - Checking Last Receipt Number                               '
        '-------------------------------------------------------------------------------'
        If mInterruptedModeFlag = True Then
            '--------Interrupt Receipt Book---------
            Dim mBookId As Integer
            mSql = "Select * From faInterruptedReceiptBooks Where tnyClosed<>1 And intCounterID=" & gbCounterID & " And intFinancialYearID=" & mYearID
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn
            If Not (Rec.BOF And Rec.EOF) Then
                If IsNull(Rec!intBookID) Then
                    MsgBox "Book not Issued for this Counter"
                    Exit Sub
                ElseIf mIRBookID = Rec!intBookID Then
                    mBookId = Rec!intBookID
                Else
                    MsgBox "This Book is Issued to Another Counter", vbInformation
                    Exit Sub
                End If
            ' MODIFIED BY ANJU ON 19/Nov/2015
            Else
                MsgBox "This Book is Issued to Another Counter", vbInformation
                Exit Sub
            End If
            Rec.Close
            Set mCnn = Nothing
            '-------------------------------------------------------------------------'
            ' INTERRUTED REGISTER MODE                                                '
            '-------------------------------------------------------------------------'
            If mInterruptedRegister <> 1 Then  'ADDED BY MINU ON 29/11/2012 FOR IR REGISTER
                objdb.SetConnection mCnn
                mSql = "Select intRecCount,intCount, B.intBookID From ("
                mSql = mSql + " Select Count(A.intBookID) intRecCount, A.intBookID  From"
                mSql = mSql + " (Select intBookNo intBookID"
                mSql = mSql + " From faVouchers"
                mSql = mSql + " Where tnyVoucherGroupID = 4 And IsNull(tnyStatus,0)<>4 And intBookNo = " & mBookId & " And intFinancialYearID = " & mYearID & " Group By intVoucherNo, intBookNo "
                mSql = mSql + " Union All"
                mSql = mSql + " Select intBookID"
                mSql = mSql + " From faInterruptedCancelledReceipts"
                mSql = mSql + " Where intBookID = " & mBookId & ") A Group by intBookID )"
                mSql = mSql + " B  INNER JOIN"
                mSql = mSql + " faInterruptedReceiptBooks ON faInterruptedReceiptBooks.intBOOKID = B.intBookID"
    
                Rec.Open mSql, mCnn, adOpenForwardOnly, adLockReadOnly
                If Not (Rec.BOF And Rec.EOF) Then
                    mSql = "If this Receipt is continuation of Receipt No:" & txtReceiptNo.Text & vbCrLf
                    mSql = mSql + "Please tick the "
                    If Rec!intRecCount >= Rec!intCount Then
                        If chkIntrNoSuffix.Value = 0 Then
                            mSql = "If this Receipt is continuation of Receipt No:" & txtReceiptNo.Text & vbCrLf
                            mSql = mSql + "Please tick the check box 'Add Suffix to Interrupted Receipt No"
                            MsgBox mSql, vbInformation
                            chkIntrNoSuffix.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
                Rec.Close
                Set mCnn = Nothing
            End If
            '-------------------------------------------------------------------------'
            ' END:::: INTERRUTED REGISTER MODE                                        '
            '-------------------------------------------------------------------------'
        End If
        
        
        '-------------------------------- For Give Advance to Sanchaya------------------'   Sinoj
        mBoolGiveAdvanceToSanchaya = False
        mAdvAmtToSanchaya = 0
        '-------------------------------------------------------------------------------'
        
        '-------------------------------------------------------------------------------'
        ' C h e c k    I t e m s   i n    G r i d    C o r r e c t l y     F i l l e d  '
        '-------------------------------------------------------------------------------'
        Dim mEmptyRow As Integer
        mEmptyRow = 9999
        For mLoopCount = 1 To vsGrid.Rows - 1
           If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then      ' (1)Account Head Code
                If mEmptyRow < mLoopCount Then                        ' (4)Checks any Previos Row is incomplete
                    MsgBox "Row is not completed!", vbInformation
                    vsGrid.Row = mEmptyRow
                    Exit Sub
                End If
            If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 Then ' (2)Valid Column
            If val(vsGrid.Cell(flexcpText, mLoopCount, 11)) > 0 Then  ' (3)Amount>0
                If mEmptyRow < mLoopCount Then                        ' (4)Checks any Previos Row is incomplete
                    MsgBox "Row is not completed!", vbInformation
                    vsGrid.Row = mEmptyRow
                    Exit Sub
                End If
            Else                ' (3)Amount>0
                 mEmptyRow = mLoopCount
            End If              ' (3)Amount>0
            End If              ' (2)Valid Column
            Else                ' (1)Account Head Code
               mEmptyRow = mLoopCount
            End If              ' (1)Account Head Code
            
            '------------------------To Give Advance to Sanchaya--------------------------' Sinoj
            If mTransactionType = gbTransactionTypePTax Then
                If vsGrid.Cell(flexcpText, mLoopCount, 0) = gbAcHeadCodeAdvancePTax And mvarDemandBasedFlag Then
                    If vsGrid.Cell(flexcpText, mLoopCount, 10) = "" Then                      ' Checking the Demand ID is Null By Sinoj
                        mBoolGiveAdvanceToSanchaya = True
                        mAdvAmtToSanchaya = mAdvAmtToSanchaya + val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                    End If
                End If
            ElseIf mTransactionType = gbTransactionTypeRentOnBuilding Then
                If vsGrid.Cell(flexcpText, mLoopCount, 0) = gbAcHeadCodeAdvanceBuilding And mvarDemandBasedFlag Then
                    If vsGrid.Cell(flexcpText, mLoopCount, 10) = "" And vsGrid.Cell(flexcpText, mLoopCount, 14) <> 1 Then                     ' Checking the Demand ID is Null By Sinoj
                        mBoolGiveAdvanceToSanchaya = True
                        mAdvAmtToSanchaya = mAdvAmtToSanchaya + val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                    End If
                End If
            ElseIf mTransactionType = gbTransactionTypeRentOnLand Then
                If vsGrid.Cell(flexcpText, mLoopCount, 0) = gbAcHeadCodeAdvanceLand And mvarDemandBasedFlag Then
                    If vsGrid.Cell(flexcpText, mLoopCount, 10) = "" And vsGrid.Cell(flexcpText, mLoopCount, 14) <> 1 Then                     ' Checking the Demand ID is Null By Sinoj
                        mBoolGiveAdvanceToSanchaya = True
                        mAdvAmtToSanchaya = mAdvAmtToSanchaya + val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                    End If
                End If
            End If
            '-----------------------------------------------------------------------------'
        Next mLoopCount
        
        '---------------------------------------------------------------------------------
        ' For D&F And PFA
        '---------------------------------------------------------------------------------
        If mTransactionType = gbTransactionTypeDandO Or mTransactionType = gbTransactionTypePFA Then
            If gbLinkWithDandOPFA = 1 Then
                If txtDemandPrefix.Text = "" Then
                    MsgBox "Please Enter Demand No", vbApplicationModal
                    Exit Sub
                End If
            End If
        End If
        '-------------------------------------------------------------------------------'
        ' END OF BLOCK ::                                                               '
        '-------------------------------------------------------------------------------'
        
        '===================================================='
        '       Added On 27/04/2009 By Cijith for KMBR       '
        '----------------------------------------------------'
        '   Checking Whether faConfig File Exists for KMBR   '
        '----------------------------------------------------'
        
        'Note:- Code Review Note by Aiby :: Date 31-Dec-2009
        '       This block of code can be removed from this part.
        '       Purpose this Block with surve is, It sets mKMBRflag
        '
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mTransactionType = gbTransactionTypePermitFeeFromKMBR Then
            objdb.SetConnection mCnn
            mSql = "Select tnyLinkWithKMBR from faConfig"
            Rec.Open mSql, mCnn
            If Rec!tnyLinkWithKMBR = 1 Then                                                              ' mKMBRAccess = Property Variable
                If mTransactionType = gbTransactionTypeApplicationForPermitKMBR And mKMBRAccess = 1 Then ' Set from KMBR from
                    mKMBRFlag = True
                ElseIf mTransactionType = gbTransactionTypePermitFeeFromKMBR Then
                    mKMBRFlag = True
                Else
                    mKMBRFlag = False
                End If
            Else
                mKMBRFlag = False
            End If
            Rec.Close
            Set mCnn = Nothing
        End If
        '----------------------------------------------------'
        
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
            If mKMBRFlag = True Then
            If cmbSeat.ListIndex = -1 Then
                MsgBox "Please Give Forwarded To Seat", vbInformation
                cmbSeat.SetFocus
                Exit Sub
            End If
            If txtMainPlace.Text = "" Then
                MsgBox "Please Give the Main Place", vbInformation
                txtMainPlace.SetFocus
                Exit Sub
            End If
            If txtPost.Text = "" Then
                MsgBox "Please Give Post Box", vbInformation
                txtPost.SetFocus
                Exit Sub
            End If
            If txtPin.Text = "" Then
                MsgBox "Please Give Pin Code", vbInformation
                txtPin.SetFocus
                Exit Sub
            End If
            If txtWardNo.Text = "" Then
                MsgBox "Please Ward No", vbInformation
                txtWardNo.SetFocus
                Exit Sub
            End If
            If txtDoorNo1.Text = "" Then
                MsgBox "Please Give Door Number", vbInformation
                txtDoorNo1.SetFocus
                Exit Sub
            End If
            If txtHouse.Text = "" Then
                MsgBox "Please Give House Name", vbInformation
                txtHouse.SetFocus
                Exit Sub
            End If
            
            Dim mCnnSoochika As New ADODB.Connection
            'If objDB.CreateNewConnection(mCnnSoochika, enuSourceString.SOOCHIKA) = True Then
            If ConnectSoochika(mCnnSoochika) = True Then
                mCnnSoochika.BeginTrans
                If gbLinkWithSoochika = 1 Then ' Urban
                    lSoochikaCurrentNo = SaveSoochika(mCnnSoochika)
                    
                    'changed by soumya V S oct 21
                  
                    If (frmUSevanaInward.Label22.Caption <> "") Then
                        ss = frmUSoochikaInward.updateseat(mCnnSoochika)
                    Else
                        frmUSoochikaInward.Label22.Caption = ""
                        frmUSoochikaInward.Label23.Caption = ""
                        frmUSoochikaInward.txtuserid.Text = ""
                        frmUSoochikaInward.txtseatid.Text = ""
                    End If
                Else '2 Unicode Soochika
                    lSoochikaCurrentNo = SaveSoochikaInwardDetails(mCnnSoochika)
                    SaveSoochikaInwardTrackDetails mCnnSoochika, lSoochikaCurrentNo 'paperless
                    
                'changed by soumya V S
                    
                    If (frmUSevanaInward.Label22.Caption <> "") Then
                     ss = frmUSoochikaInward.updateseat(mCnnSoochika)
                     
                    Else
                        frmUSoochikaInward.Label22.Caption = ""
                        frmUSoochikaInward.Label23.Caption = ""
                        frmUSoochikaInward.txtuserid.Text = ""
                        frmUSoochikaInward.txtseatid.Text = ""
                    End If
                End If
                If lSoochikaCurrentNo = -1 Or lSoochikaCurrentNo = 0 Then GoTo ErrorRollBackSoochika:
            End If
            
            End If
            'changed by soumya V S on oct21
            If ConnectSoochika(mCnnSoochika) = True Then
                    mSql = "SELECT tLBSettings.flgAttachment FROM tLBSettings"
                            Set Rec = mCnnSoochika.Execute(mSql)
                            If (Rec.Fields(0) = "1") Then
                            frmUSoochikaInward.SaveAttachment (lSoochikaCurrentNo)
                            End If
            End If
        End If
            'Attachment code 08-07-14
            'frmUSoochikaInward.SaveAttachment (lSoochikaCurrentNo)
        'End If
        '===================================================='
        
        '==================================================================================='
        ' Common Counter - mSoochikaConnected is Set from Sevana Inward
        '-----------------------------------------------------------------------------------'
        If mSoochikaConnected = True Then
            If gbSoochikaVer = 5 Then
                '--------------------------------------------------------
                '    Added By Akheel 09.03.11 For Unicode version
                '--------------------------------------------------------
                If objdb.CreateNewConnection(mCnnSoochika, enuSourceString.SoochikaUnicode) = True Then
                    mCnnSoochika.BeginTrans
                    If InwardMode = 0 Then
                        lSoochikaFileID = frmUSoochikaInward.SaveSoochika(mCnnSoochika)
                        'changed by soumya V S oct21
                    If (frmUSevanaInward.Label22.Caption <> "") Then
                    ss = frmUSoochikaInward.updateseat(mCnnSoochika)
                    Else
                        frmUSoochikaInward.Label22.Caption = ""
                        frmUSoochikaInward.Label23.Caption = ""
                        frmUSoochikaInward.txtuserid.Text = ""
                        frmUSoochikaInward.txtseatid.Text = ""
                        End If
                    Else
                        lSoochikaFileID = frmUSoochikaManualInward.SaveSoochika(mCnnSoochika)
                    End If
                    If lSoochikaFileID = -1 Or lSoochikaFileID = 0 Then GoTo ErrorRollBackSoochika:
                    lSoochikaCurrentNo = Right(lSoochikaFileID, 6) 'Added on 16-08-2011
                End If
                '--------------------------------------------------------
            Else
            '06-03-15
            'Dim mCnnSoochika As New ADODB.Connection
                If objdb.CreateNewConnection(mCnnSoochika, enuSourceString.SOOCHIKA) = True Then
                    mCnnSoochika.BeginTrans
                    lSoochikaCurrentNo = frmSoochikaInward.SaveSoochika(mCnnSoochika)
                    'changed by soumya V S
                   'changed for tvm corp
                   ' If (frmUSevanaInward.Label22.Caption <> "") Then
                    'ss = frmUSoochikaInward.updateseat(mCnnSoochika)
                    'Else
                      '  frmUSoochikaInward.Label22.Caption = ""
                        'frmUSoochikaInward.Label23.Caption = ""
                        'frmUSoochikaInward.txtuserid.Text = ""
                        'frmUSoochikaInward.txtseatid.Text = ""
                        'End If
                    If lSoochikaCurrentNo = -1 Or lSoochikaCurrentNo = 0 Then GoTo ErrorRollBackSoochika:
                End If
            End If
            
            'Attachment code 08-07-14
            'frmUSoochikaInward.SaveAttachment (lSoochikaCurrentNo)
        'End If
        'changed by soumya V S on oct21
        'changed for tvm corp
        'changed on Sep2015
           mSql = "SELECT tLBSettings.flgAttachment FROM tLBSettings"
            Set Rec = mCnnSoochika.Execute(mSql)
            If (Rec.Fields(0) = "1") Then
                frmUSoochikaInward.SaveAttachment (lSoochikaCurrentNo)
            End If
        End If
        '-----------------------------------------------------------------------------------'
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then ' CREATED NEW CONNECTION
            
            '-----------------------------------------------'Sinoj
            '           Validating Source of Fund           '
            
            '========================================='
            ' BEGIN TRANSACTION                       '
            '-----------------------------------------'
                'mCnn.BeginTrans
                'On Error GoTo ErrorRollBack:
            '========================================='
            Dim objTransactionType As New clsTransactionType
            Dim mCnnHO As New ADODB.Connection
            mSourceOfFundID = -1
            mFunctionId = -1
            If mDemandID > 0 Then   '   Getting Source of Fund From Demand Table
                mSql = "Select * From faIDEmandTBL Where numDemandID = " & mDemandID
                '-----------Added by Sunil--------------------
                If mZonal = 1 Then
                    If objdb.CreateNewConnection(mCnnHO, enuSourceString.SaankhyaHO) = True Then
                         Rec.Open mSql, mCnnHO
                    End If
                Else
                    Rec.Open mSql, mCnn
                End If
                '---------------------------------------------
             
                If Not (Rec.EOF And Rec.BOF) Then
                    If IsNull(Rec!intSourceFundID) Then
                        MsgBox "Source of Fund Not Saved in Demand", vbInformation
                        Exit Sub
                    End If
                    mSourceOfFundID = Rec!intSourceFundID
                    mFunctionId = Rec!intFunctionID
                Else
                    MsgBox "This Demand Number Not Identified", vbInformation
                    Exit Sub
                End If
                Rec.Close
            ElseIf mWebExtractMode Then
                mSourceOfFundID = frmWebExtracts.mWebSourceID
                mFunctionId = 2
                mFunctionaryID = 1
            Else                    '   Getting Source of Fund From TransactionType Table
                objTransactionType.SourceFundID = -1
                objTransactionType.FunctionID = -1
                objTransactionType.SetSourceOfFund CInt(mTransactionType)
                mSourceOfFundID = IIf(IsNull(objTransactionType.SourceFundID), -1, objTransactionType.SourceFundID)
                mFunctionId = IIf(IsNull(objTransactionType.FunctionID), -1, objTransactionType.FunctionID)
                'ADDED BY MINU ON 27/04/2013 FOR IR REGISTER
                If mInterruptedRegister = 0 Then
                   If mFunctionId <= 0 Then
                        MsgBox "Function not defined for the Transaction Type, Please make the Transaction Through Demand", vbInformation
                        Exit Sub
                    End If
                    If mSourceOfFundID <= 0 Then
                        MsgBox "Source of Fund not present in Transaction Type, Please make the Transaction Through Demand", vbInformation
                        Exit Sub
                    End If
                'ADDED BY MINU ON 27/04/2013 FOR IR REGISTER
                Else
                    If mFunctionId <= 0 Then
                        mFunctionId = 1
                        
                    End If
                    If mSourceOfFundID <= 0 Then
                        mSourceOfFundID = 4
                    End If
                End If
            End If
            '-----------------------------------------------'
'                objFunctionaries.SetFunctionary ("080000")
'                mFunctionaryID = objFunctionaries.FunctionaryID
'                Select Case mTransactionType
'                    Case 1 ' Property Tax
'                        objFunctions.SetFunction ("90910000")
'                        mFunctionID = objFunctions.FunctionID
'                    Case Else
'                        mFunctionID = Null
'                End Select

'                If mFunctionId <= 0 Then
'                    MsgBox "The Function not Defined, Please Make the Transaction through Demand", vbInformation
'                    Exit Sub
'                End If
                
                If val(txtAccountHead.Tag) > 0 Then
                    mDrAccountHeadID = val(txtAccountHead.Tag)
                Else
                    MsgBox "Error : Cash/Bank AccountHead is not set", vbInformation
                    Exit Sub
                End If
                
                '-------------------------------------------------------'
                ' Fill in Transaction Grid For Accounts Posting         '
                '-------------------------------------------------------'
                Call ListPostingHeadsInGridForGeneralReceipts
                If mGrandTotalValidityFlag Then
                    MsgBox "Difference in Grand Total and Item Total!", vbInformation
                    Exit Sub
                End If
                
'               '============For Card Transactions=========================Added By Sunil
             '  If val(txtInstrument.Tag) = 11 Then
'                    If txtInstNo = "" Then
'                        MsgBox "Enter the Instrument Number", vbInformation
'                        txtInstNo.SetFocus
'                        Exit Sub
'                    End If
'                    If txtDated.Text = "" Then
'                        MsgBox "Please Enter the Instrument Date", vbInformation
'                        txtDated.SetFocus
'                        Exit Sub
'                    End If
            '============For Card Transactions=============================Added By Sunil on 31-08-2011
                If val(txtInstrument.Tag) = 11 Then
                    If txtCardTransaction.Text = "" Then
                          If frameTransaction.Visible = True Then
                            MsgBox "Pleas Enter the Transaction Number"
                          Else
                            LockForm (False)
                            frameTransaction.Visible = True
                            Exit Sub
                          End If
                    Else
                          txtDated.Text = DdMmmYy(mTransactionDate)
                        '  mvchInstrumentNo_11 = val(txtCardTransaction.Text)
                          cmdPrint.Enabled = False
                    End If
                End If
                If cashBankValidateDr Then
                    MsgBox "Wrong Head in Debit ", vbApplicationModal
                    Exit Sub
                End If
                If cashBankValidateCr Then
                    MsgBox "Bank/ Cash Heads not Accepted as Credit ", vbApplicationModal
                    Exit Sub
                End If
              '==============================================================
                         
             '  End If
             
               '============================================================
                
                
                '-------------------------------------------------------'
                ' Exit Sub                                              '
                '-------------------------------------------------------'
                ' faVoucher                                             '
                '-------------------------------------------------------'
                Dim mintVoucherID_1                As Double
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                Dim mintTransactionTypeID_4        As Long
                Dim mtnyVoucherTypeID_5            As Integer
                Dim mintVoucherNo_6                As Variant
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
                Dim mnumSubLedgerID_21             As Variant
                Dim mintKeyID1_22                  As Variant
                Dim mintKeyID2_23                  As Variant
                Dim mintExternalApplicationID_24   As Long
                Dim mintExternalModuleID_25        As Long
                Dim mintFinancialYearID_26         As Long
                
                Dim mvchBank_33                    As String
                Dim mvchBankPlace_34               As String
                Dim mintFundID_35                  As Long
                Dim mRefNo                         As String
                Dim mRoundOff                      As Single
                Dim mAdvAmtAdj                     As Double
                Dim mtnyVoucherGroupID             As Variant
                Dim mnumLinkKeyID                  As Variant
                
                '@intVoucherID_1     [bigint],
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                
                mintTransactionTypeID_4 = val(txtTransactionType.Tag)
                mtnyVoucherTypeID_5 = 10
                If mInterruptedRegister <> 0 Then '= 3
                    mintVoucherNo_6 = mInterruptedRegisterReceiptNo
                Else
                    mintVoucherNo_6 = val(txtReceiptNo.Text)
                End If
                
                'mintVoucherNo_6 = val(txtReceiptNo.Text)
                mintBookNo_7 = val(txtBookNo.Text)
                mdtDate_8 = mTransactionDate
                mfltAmount_9 = val(txtTotal.Text)
                mintInstrumentTypeID_10 = val(txtInstrument.Tag)
                '=======For card Transaction
                If val(txtInstrument.Tag) = 11 Then
                    mvchInstrumentNo_11 = txtCardTransaction.Text
                Else
                    mvchInstrumentNo_11 = Trim(txtInstNo.Text)
                End If
                '====================
                mdtInstrumentDate_12 = IIf(Trim(txtDated) <> "", CheckDateInMMM(txtDated), Null)
                mvchDescription_13 = Trim(txtDescription.Text)
                If val(txtAdvance.Text) > 0 Then
                    If Len(mvchDescription_13) > 0 Then mvchDescription_13 = mvchDescription_13 + ", "
                    mvchDescription_13 = mvchDescription_13 + "Advance Adjusted Rs." + Trim(txtAdvance.Text)
                End If
                If cmbZone.ListIndex > 0 Then
                    mnumZoneID_14 = IIf(cmbZone.ItemData(cmbZone.ListIndex) > 0, cmbZone.ItemData(cmbZone.ListIndex), Null)
                End If
                If cmbZone.ListIndex > 0 Then
                    mnumZoneID_14 = cmbZone.ItemData(cmbZone.ListIndex)
                End If
                mnumWardID_15 = val(txtWardNo.Text)
                mintDoorNoP1_16 = val(txtDoorNo1.Text)
                mvchDoorNoP2_17 = Trim(txtDoorNo2.Text)
                mvchDoorNoP3_18 = ""
                mintUserID_19 = gbUserID
                mintCounterID_20 = gbCounterID
                mnumSubLedgerID_21 = mSubLedgerID ' mBuildingID ' Changed by Aiby on 10-Dec-2008 From Kollam Corp.
                mintKeyID1_22 = mDrAccountHeadID
                mintKeyID2_23 = mDemandID
                If mWebExtractMode = True Then
                    mintExternalApplicationID_24 = 118  ''saankhyaWeb project Receipt
                Else
                    mintExternalApplicationID_24 = AppID.Saankhya
                End If
                'Added by Sunil for Zonal Office Collection
                If mZonal = 1 Then
                    mintExternalModuleID_25 = 45
                Else
                    mintExternalModuleID_25 = 0
                End If
                mintFinancialYearID_26 = mYearID
                mvchBank_33 = Trim(txtBank)
                mvchBankPlace_34 = Trim(txtPlace)
                mintFundID_35 = 1
                mRefNo = Trim(txtRefNo.Text)
                mRoundOff = val(txtRoundOff)
                mAdvAmtAdj = val(txtAdvance.Text)
                If val(txtTransactionType.Tag) = gbTransactionTypeRentOnLand Or val(txtTransactionType.Tag) = gbTransactionTypeRentOnBuilding Then
                    If IsNull(mSubLedgerID) Then
                        If txtBuildingNo.Text <> "" Then
                            mnumSubLedgerID_21 = txtBuildingNo.Text
                        End If
                    End If
                ElseIf val(txtTransactionType.Tag) = gbTransactionTypeProfTaxTrade And gbLinkWithProfTradeWeb Then
                    If IsNull(mSubLedgerID) Then
                        If txtBuildingNo.Text <> "" Then
                            mnumSubLedgerID_21 = txtBuildingNo.Text
                        End If
                    Else
                        mnumSubLedgerID_21 = SubLedgerID
                    End If
                ElseIf val(txtTransactionType.Tag) = gbTransactionTypeProfTaxEmp And gbLinkWithProfEmpWeb Then
                    If IsNull(mSubLedgerID) Then
                        If txtBuildingNo.Text <> "" Then
                            mnumSubLedgerID_21 = txtBuildingNo.Text
                        End If
                    Else
                        mnumSubLedgerID_21 = SubLedgerID
                    End If
                End If
                If mInterruptedModeFlag Then
                    mtnyVoucherGroupID = 4
                    mintBookNo_7 = mIRBookID
                    If txtIntruptNoSuffix.Text <> "" Then
                        mvchDoorNoP3_18 = txtIntruptNoSuffix.Text
                    End If
                    
                    
                    
                    ''Dim mTotalBookPages As Integer
                    ''mtnyVoucherGroupID = 4
                    '''--------Interrupt Receipt Book---------
                    '''objDb.SetConnection mCnn
                    ''mSQL = "Select * From faInterruptedReceiptBooks Where tnyClosed<>1 And intCounterID=" & gbCounterID & "And intFinancialYearID=" & mYearID
                    ''Rec.Open mSQL, mCnn
                    ''If IsNull(Rec!intBookID) Then
                    ''    MsgBox "Book not Issued for this Counter"
                    ''    Exit Sub
                    ''Else
                    ''    mintBookNo_7 = Rec!intBookID
                    ''    mTotalBookPages = Rec!intCount
                    ''End If
                    ''Rec.Close
                    ''
                    ''If chkIntrNoSuffix.value = vbChecked Then
                    ''    If txtIntruptNoSuffix.Text = "" Then
                    ''        MsgBox "Please Enter Suffix for interrupReceipt.. Should be Alphabets .. "
                    ''        Exit Sub
                    ''    Else
                    ''        mvchDoorNoP3_18 = txtIntruptNoSuffix.Text
                    ''    End If
                    ''Else
                    ''    mvchDoorNoP3_18 = txtIntruptNoSuffix.Text
                    ''End If
                    '-------------------------------------------------------------------------'
                    ' INTERRUTED REGISTER MODE                                                '
                    '-------------------------------------------------------------------------'
                    ''If mInterruptedRegister = 0 Then  '  'ADDED BY MINU ON 29/11/2012 FOR IR REGISTER
                    '''--------Finding Total Receipts entered--------'
                    ''    Dim mIntTotalReceiptsEntered As Integer
                    ''    mSQL = "Select isNull(Count(*),0) intCount From" & vbNewLine
                    ''    mSQL = mSQL + "(Select intBookNo intBookID" & vbNewLine
                    ''    mSQL = mSQL + "From faVouchers" & vbNewLine
                    ''    mSQL = mSQL + "Where tnyVoucherGroupID = 4 And intBookNo = " & mintBookNo_7 & " And intFinancialYearID = " & mYearID & vbNewLine
                    ''    mSQL = mSQL + " Group By intVoucherNo,intBookNo"
                    ''    mSQL = mSQL + " Union All" & vbNewLine
                    ''    mSQL = mSQL + "Select intBookID" & vbNewLine
                    ''    mSQL = mSQL + "From faInterruptedCancelledReceipts" & vbNewLine
                    ''    mSQL = mSQL + "Where intBookID =  " & mintBookNo_7 & ") A"
                    ''    Rec.Open mSQL, mCnn
                    ''    mIntTotalReceiptsEntered = Rec!intCount
                    ''    Rec.Close
                    ''    If mTotalBookPages <= mIntTotalReceiptsEntered Then
                    ''        If gbLBType = 3 Or gbLBType = 4 Then
                    ''            MsgBox "The Book Completed, Please Request for another Book", vbInformation
                    ''            Set mCnn = Nothing
                    ''            mSQL = "Update faInterruptedReceiptBooks SET tnyClosed = 1 WHERE intBookID = " & mintBookNo_7
                    ''            objDB.SetConnection mCnn
                    ''            mCnn.Execute mSQL
                    ''            Set mCnn = Nothing
                    ''            Call FormInitialize
                    ''            Exit Sub
                    ''        End If
                    ''    End If
                    ''End If
                    '----------------------------------------------'
                    ' END BLOCK ::INTERRUTED REGISTER MODE
                    '----------------------------------------------'
                    
                ElseIf mZonal = 1 Then
                     mtnyVoucherGroupID = 5  'Added by Sunil
                Else
                     mtnyVoucherGroupID = Null
                End If
                
                If gbSectionID <> gbJSKSectionID Then
                    If mintInstrumentTypeID_10 = gbInstrumentLetterOfAuthority Then
                        mdtDate = mdtDate_8
                    End If
                    If mintTransactionTypeID_4 = gbTransactionTypeBFundSSSFund Then
                        mdtDate = mdtDate_8
                    End If
                    'If (gbLBPanchayat Or gbLBType = 4) And gbTransactionDate < DateValue(gbOnlinedate) Then     '***dtOnlineDate validation
                    
                    
                    If mDemandMode <> 0 Then
                        If mDemandMode <> 1 And IsNull(mDemandTrDate) = False Then
                            mdtDate = mDemandTrDate
                        Else
                            mdtDate = mTransactionDate
                        End If
                    Else
                        mdtDate = mdtDate_8
                    End If
                    'End If
        
                End If
                If mWebExtractMode = True Then
                    mdtDate = mTransactionDate
                    mSubLedgerID = txtOutDoorStaff.Tag  ''webExtractID
                    mnumSubLedgerID_21 = txtOutDoorStaff.Tag
                    mintFinancialYearID_26 = gbFinancialYearID
                End If
                If mZonal = 1 Then
                    mdtDate = Format(mZoneDate, "dd/mmm/yyyy")
                End If
                'mdtDate_8 = mCnn.Execute
                
                '========================================='
                ' BEGIN TRANSACTION                       '
                '-----------------------------------------'
                    mCnn.BeginTrans
                    On Error GoTo ErrorRollBack:
                '========================================='
                
                arrInput = Array( _
                -1, _
                gbLocalBodyID, _
                Null, _
                mintTransactionTypeID_4, _
                mtnyVoucherTypeID_5, _
                mintVoucherNo_6, _
                mintBookNo_7, _
                mdtDate, _
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
                mintExternalModuleID_25, mintFinancialYearID_26, gbShiftID, 1, 0, _
                mvchBank_33, mvchBankPlace_34, mintFundID_35, gbSeatID, gbSessionID, mRefNo, mRoundOff, mAdvAmtAdj, lSoochikaCurrentNo, 0, gbLocationID, mtnyVoucherGroupID, mnumLinkKeyID)

                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID_1 = arrOutPut(0, 0)
                    mVoucherID = mintVoucherID_1
                    mReceiptNo = arrOutPut(1, 0)
                    txtReceiptNo.Text = Left(mReceiptNo, 6) + "-" + mID(mReceiptNo, 7, Len(mReceiptNo))
                Else
                    GoTo ErrorRollBack:
                End If
                
                '-------------------------------------------------------'
                ' faVoucher Child
                '-------------------------------------------------------'
                'Dim mintVoucherID_1       As Double  '
                Dim mintLocalBodyID_2       As Long
                Dim mintSlNo_3              As Long
                Dim mintAccountHeadID_4     As Long
                Dim mtnyDebitOrCredit_5     As Byte
                Dim mintYearID_6            As Long
                Dim mtnyPeriodID_7          As Byte
                Dim mtnyArrearFlag_8        As Byte
'                Dim mnumDemandID_9          As Double
                Dim mfltAmount_10           As Double
                Dim mnumDemandID_9          As Variant   'Modified by Sunil
                
                For mLoopCount = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 Then
                            '----------------------------------------------------------------------------'
                            'NOTE=> vsGrid.Cell(flexcpText, mLoopCount, 14) :: Those Rows Which          '
                            '       Do not want to Save in Child Table eg. Advance Property Tax Adjusted '
                            '----------------------------------------------------------------------------'
                            mintLocalBodyID_2 = gbLocalBodyID
                            mintSlNo_3 = mLoopCount
                            mintAccountHeadID_4 = vsGrid.Cell(flexcpText, mLoopCount, 6)
                            mtnyDebitOrCredit_5 = 0
                            mintYearID_6 = val(vsGrid.Cell(flexcpText, mLoopCount, 7))
                            mtnyPeriodID_7 = val(vsGrid.Cell(flexcpText, mLoopCount, 8))
                            If mintYearID_6 < mYearID Then
                                mtnyArrearFlag_8 = 1
                            Else
                                mtnyArrearFlag_8 = 0
                            End If
                            
'                            mnumDemandID_9 = mDemandID 'vsGrid.Cell(flexcpText, mLoopCount, 10) 'Modified By sunil
                            mnumDemandID_9 = vsGrid.Cell(flexcpText, mLoopCount, 10) 'Modified By sunil
                            mfltAmount_10 = val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                            
                            Set arrInput = Nothing
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
                            objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                        End If
                    Else
                        Exit For
                    End If
                Next mLoopCount
                
                '-------------------------------------------------------'
                ' faVoucher Address
                '-------------------------------------------------------'
                '@intVoucherID   [bigint],
                '@intLocalBodyID    [int],
                '@vchName   [varchar](100),
                '@vchInit1  [varchar](2) = Null,
                '@vchInit2  [varchar](2) = Null ,
                '@vchInit3  [varchar](2) = Null,
                '@vchInit4  [varchar](2) = Null,
                '@vchHouseName  [varchar](100) = Null,
                '@vchStreetName [varchar](100) = Null,
                '@vchLocalPlace [char](10) = Null,
                '@vchMainPlace  [varchar](100)= Null,
                '@vchPostOffice [varchar](100) = Null,
                '@vchDistrict   [varchar](50)= Null,
                '@vchPinNumber  [varchar](6) = Null,
                '@vchPhone  [varchar](15)= Null),
                '@intWardNo  [int] = Null,
                '@intDoorNo [int] = Null,
                '@vchDoorNo2    [varChar](10) = Null
                
                '-----------------------'
                '       Added Newly     '
                '-----------------------'
                'vchName = Trim(txtPayee.Text)
                vchName = Trim(txtName.Text)
                vchHouseName = Trim(txtHouse.Text)
                vchInit1 = Trim(txtInit1.Text)
                vchInit2 = Trim(txtInit2.Text)
                vchInit3 = Trim(txtInit3.Text)
                vchInit4 = Trim(txtInit4.Text)
                vchStreetName = Trim(txtStreet.Text)
                vchLocalPlace = Trim(txtLocalPlace.Text)
                vchMainPlace = Trim(txtMainPlace.Text)
                vchPostOffice = Trim(txtPost.Text)
                vchPinNumber = txtPin.Text
                vchPhone = txtPhone.Text
                intWardNo = txtWardNo.Text
                intDoorNo = txtDoorNo1.Text
                vchDoorNo2 = txtDoorNo2.Text
                '-----------------------'
                arrInput = Array(mintVoucherID_1, _
                        gbLocalBodyID, _
                        vchName, _
                        vchInit1, _
                        vchInit2, _
                        vchInit3, _
                        vchInit4, _
                        vchHouseName, _
                        vchStreetName, _
                        vchLocalPlace, _
                        vchMainPlace, _
                        vchPostOffice, _
                        vchDistrict, _
                        vchPinNumber, _
                        vchPhone, _
                        intWardNo, _
                        intDoorNo, _
                        vchDoorNo2)
                objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn
                
                '-------------------------------------------------------' Sinoj
                '                     faVoucher Sub                     '
                '-------------------------------------------------------'
                Dim objVrSub As uVoucherSub
                With objVrSub
                    .intVoucherID = mintVoucherID_1
                    .decProjectID = Null
                    .intSourceOfFundID = mSourceOfFundID
                    .intCategoryID = Null
                    .intSectorID = Null
                    .intAllotmentID = Null
                    .intAgreementID = Null
                    .intCashBookID = Null
                    .intImplementingOfficerID = Null
                    .intCreditorTypeID = Null
                    .intCreditorsID = Null
                    .intTypeID = Null
                    .intLocalBodyID = gbLocalBodyID
                    
                    arrInput = Array(.intVoucherID, _
                                    .intLocalBodyID, _
                                    .decProjectID, _
                                    .intSourceOfFundID, _
                                    .intCategoryID, _
                                    .intSectorID, _
                                    .intAllotmentID, _
                                    .intAgreementID, _
                                    .intCashBookID, _
                                    .intImplementingOfficerID, _
                                    .intCreditorTypeID, _
                                    .intCreditorsID, _
                                    .intTypeID)
                    objdb.ExecuteSP "spSaveVoucherSub", arrInput, , , mCnn
                End With
                '-------------------------------------------------------'
                                
                '-------------------------------------------------------'
                ' Transactions                                          '
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
                Dim mnumtnyVoucherGroupID As Variant 'Added by Sunil on 29-07-2011
                intTransactionID_1 = -1
                mintLocalBodyID_2 = gbLocalBodyID
                mintFinancialYearID_3 = mYearID
                mdtTransactionDate_4 = mTransactionDate
                mintExternalApplicationID_5 = AppID.Saankhya
                If mZonal = 1 Then        'Added by Sunil on 29-07-2011
                    mnumtnyVoucherGroupID = 5
                    mintExternalApplicationModuleID_6 = 45
                Else
                    mintExternalApplicationModuleID_6 = 0
                End If
                
                mintFunctionID_7 = mFunctionId
                mintFunctionaryID_8 = mFunctionaryID
                mintFieldID_9 = IIf(val(txtWard) < 1, Null, val(txtWard))
                mintFundID_10 = Null
                mintBudgetCentreID_11 = Null
                mvchNarration_12 = Trim(txtDescription.Text)
                If val(txtAdvance.Text) > 0 Then
                    mvchNarration_12 = mvchNarration_12 + "Advance Amount Adjusted Rs." + Trim(txtAdvance.Text)
                End If
                mintTransactionTypeID_13 = mTransactionType
                mintVoucherNo_14 = mintVoucherID_1
                mintProcessID_15 = Null
                mvchGroup_16 = "R"
                mintGroupID_17 = 10
                mintKeyID_18 = Null     'mDemandID 'Added on 3-Sep-2008
                mnumSubLedgerID_19 = mBuildingID
                'mintUserID_20 = gbUserID
                
                arrInput = Array( _
                intTransactionID_1, _
                mintLocalBodyID_2, _
                mintFinancialYearID_3, _
                mdtDate, _
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
                mintVoucherNo_14, _
                mnumtnyVoucherGroupID)
                
                
                Set arrOutPut = Nothing
                objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    intTransactionID_1 = arrOutPut(0, 0)
                Else
                    GoTo ErrorRollBack:
                End If
                
                '-------------------------------------------------------'
                ' Transaction Child                                     '
                '-------------------------------------------------------'
                
                '=========================================================================================='
                '                                                                                          '
                ' BLOCK: I   :  Accounting Part of Advance Adjustment of Property Tax                      '
                '                                                                                          '
                ' a) Advance will be saved in Voucher Table as every normal voucher as it saves. Its Acco- '
                ' -uning part will handled by Transaction Tables. Order in which the advance settled off by'
                ' Penal Interest, PTax(Arrear)+LC, PTax(Current)+Lc                                        '
                ' b) This Block Only Handles Property Tax Advance                                          '
                ' c)                                                                                       '
                '------------------------------------------------------------------------------------------'
                If mAdvAmtAdj > 0 Then
                    Dim mTrChild As uTrChild
                    Dim mFineFlag As Boolean
                    Dim mFineAmt As Double
                    Dim mAmt As Double
                    Dim mPTax As Double
                    Dim mLC As Double
                    Dim mCess As Double
                    Dim mSL As Integer
                    Dim mExitLoopFlag As Boolean
                    Dim mByHeadID As Integer
                    Dim mNoticeFee As Integer
                    Dim mNoticeAmt As Double
                    'NOTE:- Check TransactionTypes
                    If mintTransactionTypeID_4 = gbTransactionTypePTax Then
                        'NOTE:- Posting of Advance Collection of Property Tax
                        mSL = 2
                        With mTrChild
                            .intTransactionID = intTransactionID_1
                            .intSerialNo = mSL
                            .intAccountHeadID = gbAcHeadIDAdvancePTax
                            .fltAmount = mAdvAmtAdj
                            .tinDebitOrCreditFlag = 1
                            .intByAccountHeadID = Null
                            .vchNarration = "Total Advance Collection Adjusted"
                            .intFundID = 1
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        End With
                        
                        'NOTE:- Checking for Penal Interest
                        For mLoopCount = 1 To vsGrid.Rows - 1
                            If val(vsGrid.TextMatrix(mLoopCount, 6)) = gbAcHeadIDPenalInterest Then
                                'Note:- Found Penal Interest in Grid
                                '       Which is expected Only once Penal Interest Appear in Grid
                                mFineFlag = True ' Fine Exists
                                mFineAmt = val(vsGrid.TextMatrix(mLoopCount, 11))
                                If mAdvAmtAdj >= mFineAmt Then ' Advance Amount is greater than total Penal Interest
                                    mAmt = mFineAmt
                                    mFineAmt = 0
                                Else   'Note:- Advance Amount will completely settled off by Penal interest
                                    mAmt = mAdvAmtAdj
                                    mFineAmt = mFineAmt - mAmt
                                    '(A)-->> Note:- Remaining Fine Should Set off With Cash/Bank Heads
                                End If
                                mAdvAmtAdj = mAdvAmtAdj - mAmt
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDPenalInterest
                                    .fltAmount = mAmt
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                    .vchNarration = "Advance Collection Adjusted With Penal Interest"
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                                Exit For
                            End If
                        Next mLoopCount
                        If mAdvAmtAdj > 0 Then
                            For mLoopCount = 1 To vsGrid.Rows - 1 ''' Added on 21 dec 2016 For setting off Demand Notice fee in Advace amount
                                If val(vsGrid.TextMatrix(mLoopCount, 6)) = gbAcHeadIDNoticeFee Then
    
                                    'Note:- Found Notice Fee in Grid
                                    '       Which is expected Only once Notice Fee Appear in Grid
                                    mNoticeFee = True ' Fine Exists
                                    mNoticeAmt = val(vsGrid.TextMatrix(mLoopCount, 11))
                                    If mAdvAmtAdj >= mNoticeAmt Then ' Advance Amount is greater than total Penal Interest
                                        mAmt = mNoticeAmt
                                        mNoticeAmt = 0
                                    Else   'Note:- Advance Amount will completely settled off by Penal interest
                                        mAmt = mAdvAmtAdj
                                        mNoticeAmt = mNoticeAmt - mAmt
                                        '(A)-->> Note:- Remaining Fine Should Set off With Cash/Bank Heads
                                    End If
                                    mAdvAmtAdj = mAdvAmtAdj - mAmt
                                    With mTrChild
                                        mSL = mSL + 1
                                        .intTransactionID = intTransactionID_1
                                        .intSerialNo = mSL
                                        .intAccountHeadID = gbAcHeadIDNoticeFee
                                        .fltAmount = mAmt
                                        .tinDebitOrCreditFlag = 0
                                        .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                        .vchNarration = "Advance Collection Adjusted With Penal Interest"
                                        .intFundID = 1
                                        
                                        arrInput = Array(.intTransactionID, _
                                        .intSerialNo, _
                                        .intAccountHeadID, _
                                        .fltAmount, _
                                        .tinDebitOrCreditFlag, _
                                        .intByAccountHeadID, _
                                        .vchNarration, _
                                        .intFundID)
                                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                    End With
                                    Exit For
                                End If
                            Next mLoopCount 'NOTE:- END of Setting Notice Fee if find in Grid
                        End If
                        mLoop = 1 'NOTE:- Using in ELSE Part
                        'NOTE:- Remaining Balance in Advance Collection
                        If mAdvAmtAdj > 0 Then
                            '----------------------------------------------------------------------'
                            'NOTE:- After Penal Interest Set Off                                   '
                            '       Checking for Property Tax Arrear or Current heads in Grid      '
                            '       in the same order as it appears in Grid                        '
                            '----------------------------------------------------------------------'
                            For mLoopCount = 1 To vsGrid.Rows - 1
                                If vsGrid.TextMatrix(mLoopCount, 0) <> gbAcHeadCodeAdvancePTax Then
                               'Note:- Sum of Property Tax + LC
                                mPTax = val(vsGrid.TextMatrix(mLoopCount, 11))
                                If mLoopCount < vsGrid.Rows Then
                                    mLoopCount = mLoopCount + 1 'Note:- Finding the Library Cess
                                    If vsGrid.TextMatrix(mLoopCount, 0) = gbAcHeadCodeLibraryCess Then
                                        mLC = val(vsGrid.TextMatrix(mLoopCount, 11))
                                    End If
                                    ' Added On 08.09.2010 For Finding Poor home Cess Amount
                                    If mPoorHomeCess Then
                                        mLoopCount = mLoopCount + 1 'Note:- Finding the PoorHomeCess Cess
                                        If vsGrid.TextMatrix(mLoopCount, 0) = gbAcHeadCodePoorHomeCess Then
                                            mCess = val(vsGrid.TextMatrix(mLoopCount, 11))
                                        End If
                                    End If
                                End If
                                mAmt = mPTax + mLC + mCess 'Sum PTax + LC+mCess
                                mSL = mSL + 1
                                'Note:- IF Sum of PTax+LC Greater than Advance Amount
                                If mAdvAmtAdj >= mAmt Then
                                    With mTrChild
                                        .intTransactionID = intTransactionID_1
                                        .intSerialNo = mSL
                                        'Note:- mLoopCount - 1 => Loop count is on LC, its to find PTax head One should
                                        '       check in the previous row
                                        If val(vsGrid.TextMatrix(mLoopCount - 1, 6)) = gbAcHeadIDPropertyTaxArrear Then
                                            .intAccountHeadID = gbAcHeadIDPropertyTaxArrear
                                        Else
                                            .intAccountHeadID = gbAcHeadIDPropertyTaxCurrent
                                        End If
                                        .fltAmount = mPTax
                                        .tinDebitOrCreditFlag = 0
                                        .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                        .vchNarration = "Adv.Adjusted With Property Tax " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                        .intFundID = 1
                                        
                                        arrInput = Array(.intTransactionID, _
                                        .intSerialNo, _
                                        .intAccountHeadID, _
                                        .fltAmount, _
                                        .tinDebitOrCreditFlag, _
                                        .intByAccountHeadID, _
                                        .vchNarration, _
                                        .intFundID)
                                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                        
                                        'NOTE:- Library Cess in the very next row in Grid (Thats what Expected!! ;) )
                                        mSL = mSL + 1
                                        .intTransactionID = intTransactionID_1
                                        .intSerialNo = mSL
                                        .intAccountHeadID = gbAcHeadIDLibraryCess
                                        .fltAmount = mLC
                                        .tinDebitOrCreditFlag = 0
                                        .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                        .vchNarration = "Adv. Collection Adjusted With Library Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                        .intFundID = 1
                                        
                                        arrInput = Array(.intTransactionID, _
                                        .intSerialNo, _
                                        .intAccountHeadID, _
                                        .fltAmount, _
                                        .tinDebitOrCreditFlag, _
                                        .intByAccountHeadID, _
                                        .vchNarration, _
                                        .intFundID)
                                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                        If mPoorHomeCess Then
                                            'Poor Home Cess Advance Posting
                                             If mCess > 0 Then
                                                 mSL = mSL + 1
                                                .intTransactionID = intTransactionID_1
                                                .intSerialNo = mSL
                                                .intAccountHeadID = gbAcHeadIDPoorHomeCess
                                                .fltAmount = mCess
                                                .tinDebitOrCreditFlag = 0
                                                .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                                .vchNarration = "Adv. Collection Adjusted With Poor home Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                                .intFundID = 1
                                                
                                                arrInput = Array(.intTransactionID, _
                                                .intSerialNo, _
                                                .intAccountHeadID, _
                                                .fltAmount, _
                                                .tinDebitOrCreditFlag, _
                                                .intByAccountHeadID, _
                                                .vchNarration, _
                                                .intFundID)
                                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                            End If
                                        End If
                                    End With
                                    'mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC)
                                    mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC + mCess)
                                Else
                                    '
                                    ' NOTE:- Advance Amount is Less than Total PTax+LC.
                                    '        Remaining Advance Amount will be split into PTax & LC by ratio.
                                    '        The rest part will again adjusted against Cash/Bank Account.
                                   
'                                    mLC = mAdvAmtAdj - mPTax
'                                    mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC)
                                    If mPoorHomeCess Then
                                        mPTax = Format(mAdvAmtAdj * 100 / 107, "0.0")
                                        mLC = Round(Format(mPTax * 5 / 100, "0.0"))
                                        mCess = mAdvAmtAdj - mPTax - mLC
                                    Else
                                        mPTax = Format(mAdvAmtAdj * 100 / 105, "0.0")
                                        mLC = mAdvAmtAdj - mPTax
                                    End If
                                    mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC + mCess)
                                    
                                    With mTrChild
                                            mByHeadID = gbAcHeadIDAdvancePTax
Step2:
                                        
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            ' Note:- mLoop - 1 => Loop count is on LC, its to find PTax head One should
                                            '        check in the previous row
                                            If val(vsGrid.TextMatrix(mLoopCount - 1, 6)) = gbAcHeadIDPropertyTaxArrear Then
                                                .intAccountHeadID = gbAcHeadIDPropertyTaxArrear
                                            Else
                                                .intAccountHeadID = gbAcHeadIDPropertyTaxCurrent
                                            End If
                                            .fltAmount = mPTax
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mByHeadID
                                            .vchNarration = "Adv.Adjusted With Property Tax " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                            
                                            'NOTE:- Library Cess in the very next row
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            .intAccountHeadID = gbAcHeadIDLibraryCess
                                            .fltAmount = mLC
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mByHeadID
                                            .vchNarration = "Adv. Collection Adjusted With Library Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                            If mPoorHomeCess Then
                                                'NOTE:- Poor Home Cess in the very next row
                                                mSL = mSL + 1
                                                .intTransactionID = intTransactionID_1
                                                .intSerialNo = mSL
                                                .intAccountHeadID = gbAcHeadIDPoorHomeCess
                                                .fltAmount = mCess
                                                .tinDebitOrCreditFlag = 0
                                                .intByAccountHeadID = mByHeadID
                                                .vchNarration = "Adv. Collection Adjusted With Library Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                                .intFundID = 1
                                                
                                                arrInput = Array(.intTransactionID, _
                                                .intSerialNo, _
                                                .intAccountHeadID, _
                                                .fltAmount, _
                                                .tinDebitOrCreditFlag, _
                                                .intByAccountHeadID, _
                                                .vchNarration, _
                                                .intFundID)
                                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                           End If
                                            'Rest part to adjust with Cash Or Bank AccountHead
                                        If Not mExitLoopFlag Then
'                                            mPTax = val(vsGrid.TextMatrix(mLoopCount - 1, 11)) - mPTax
'                                            mLC = val(vsGrid.TextMatrix(mLoopCount, 11)) - mLC
'                                            mExitLoopFlag = True
'                                            mByHeadID = mDrAccountHeadID
                                            If mPoorHomeCess Then
                                                mPTax = val(vsGrid.TextMatrix(mLoopCount - 2, 11)) - mPTax
                                                mLC = val(vsGrid.TextMatrix(mLoopCount - 1, 11)) - mLC
                                                mCess = val(vsGrid.TextMatrix(mLoopCount, 11)) - mCess
                                            
                                            Else
                                                mPTax = val(vsGrid.TextMatrix(mLoopCount - 1, 11)) - mPTax
                                                mLC = val(vsGrid.TextMatrix(mLoopCount, 11)) - mLC
                                            End If
                                            mExitLoopFlag = True
                                            mByHeadID = mDrAccountHeadID
            
                                            GoTo Step2:
                                        Else
                                            Exit For
                                        End If
                                    End With
                                End If
                                End If ' If vsGrid.TextMatrix(mLoopCount, 0) = gbAcHeadCodeAdvancePTax Then
                            Next mLoopCount
                            
                            If mLoopCount < vsGrid.Rows - 1 Then
                                mLoop = mLoopCount + 1
                            End If
                            GoTo Step3:
                        Else 'Note:- Else Part OF Condition [If mAdvAmtAdj > 0 Then] : After Penal Interest
                             'NOTE:- Advance Amount is settled.
                             '       Rest part of the accounting posting which is collected as Cash or Bank.
Step3:
                             'NOTE:- Cash Or Bank With SerialNo 1
                             With mTrChild
                                 .intTransactionID = intTransactionID_1
                                 .intSerialNo = 1
                                 .intAccountHeadID = mDrAccountHeadID
                                 .fltAmount = mfltAmount_9
                                 .tinDebitOrCreditFlag = 1
                                 .intByAccountHeadID = Null
                                 .vchNarration = Null
                                 .intFundID = 1
                                
                                 arrInput = Array(.intTransactionID, _
                                 .intSerialNo, _
                                 .intAccountHeadID, _
                                 .fltAmount, _
                                 .tinDebitOrCreditFlag, _
                                 .intByAccountHeadID, _
                                 .vchNarration, _
                                 .intFundID)
                                 objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                             End With
                             For mLoopCount = mLoop To vsGrid.Rows - 1
                                If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                                    If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 And val(vsGrid.Cell(flexcpText, mLoopCount, 6)) <> gbAcHeadIDPenalInterest Then
                                    
                                        'NOTE=> vsGrid.Cell(flexcpText, mLoopCount, 14) :: Those Rows Which Do not
                                        '       want to Save in Child Table eg. Advance Property Tax Adjusted
                                        With mTrChild
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            .intAccountHeadID = val(vsGrid.Cell(flexcpText, mLoopCount, 6))
                                            .fltAmount = val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mDrAccountHeadID
                                            .vchNarration = Null
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                        End With
                                    End If
                                End If
                             Next mLoopCount
                             'NOTE:- IF Penal Interest is not Set off completely then
                             '-->> (A)  Continuation
                             If mFineAmt > 0 Then
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDPenalInterest
                                    .fltAmount = mFineAmt
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = mDrAccountHeadID
                                    .vchNarration = Null
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                             End If
                             
                             'NOTE:-Round of Adjustment adjusted
                             If mRoundOff > 0 Then
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDRoundOff
                                    .fltAmount = mRoundOff
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = mDrAccountHeadID
                                    .vchNarration = "Round Of Adjustment"
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                             End If
                             GoTo GotoCommitTran:  'NOTE:- Complete the Transaction By CommitTrans
                        End If ' END OF Condition [If mAdvAmtAdj > 0 Then] : After Penal Interest Set off
                    ElseIf mintTransactionTypeID_4 = gbTransactionTypeRentOnBuilding Then
                        Call SaveRentAdv(intTransactionID_1, gbAcHeadCodeAdvanceBuilding, mCnn)
                        GoTo GotoCommitTran:
                    ElseIf mintTransactionTypeID_4 = gbTransactionTypeRentOnLand Then
                        Call SaveRentAdv(intTransactionID_1, gbAcHeadCodeAdvanceLand, mCnn)
                        GoTo GotoCommitTran:
                    End If     ' End of Checking Transaction Type : Property Tax
                End If         ' End of Advance Collection Posting Block 1
                '=========================================================================================='
                ' END OF BLOCK 1 : Advance Adjustment of Property Tax - Integrated Sanchaya Mode           '
                '=========================================================================================='
                For mLoop = 1 To vsGridTransactions.Rows - 1
                    '-------------------------------------------------------------'
                    'NOTE=> ALL TRANSACTIONS EXCEPT PROPERTY TAX - POSTING HERE   '
                    '-------------------------------------------------------------'
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
                                    gbFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                Next mLoop
                '-------------------------------------------------------'
                ' Round Off Adjustment to Transaction Child             '
                '-------------------------------------------------------'
                If val(txtRoundOff) > 0 Then
                    mintAccountHeadID_4 = gbAcHeadIDRoundOff
                    If mintAccountHeadID_4 = -1 Then
                        mintAccountHeadID_4 = Null
                    End If
                    arrInput = Array(intTransactionID_1, _
                                    mLoop, _
                                    mintAccountHeadID_4, _
                                    val(txtRoundOff), _
                                    0, _
                                    mDrAccountHeadID, _
                                    "Round Off Adj.", _
                                    gbFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                End If
                
                '-------------------------------------------------------'
                ' Update Demand Table                                   '
                '-------------------------------------------------------'
                'If mTransactionType = 1 Then
                
                Dim mStatusFlag As Integer
                If val(txtDemandNo.Tag) > 0 Then
                    mDemandID = txtDemandNo.Tag
                    mStatusFlag = 1
                    arrInput = Array(mDemandID, mStatusFlag, mintVoucherID_1)
                    objdb.ExecuteSP "spUpdateIDemandStatus", arrInput, , , mCnn
                Else
                    For mLoop = 1 To vsGrid.Rows - 1
                        If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked And mDemandID <> vsGrid.Cell(flexcpText, mLoop, 10) Then
                            mDemandID = val(vsGrid.TextMatrix(mLoop, 10))
                            mStatusFlag = 1
                            arrInput = Array(mDemandID, mStatusFlag, mintVoucherID_1)
                            objdb.ExecuteSP "spUpdateIDemandStatus", arrInput, , , mCnn
                        End If
                    Next mLoop
                End If
                '----------------------------------------------------------------------'
                'Checking Payment Order Cancellation TransactionType(Refund Of Payment)'
                '----------------------------------------------------------------------'
                If mTransactionType = gbTransactionTypeRefundOfPayment Then     ' Payment Order Cancellation
                    If val(txtDemandNo.Tag) > 0 Then
                        mDemandID = txtDemandNo.Tag
                        Rec.Open "Select intKeyID2 From faIDemandTBL Where numDemandID = " & mDemandID, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mPayOrderNo = Rec!intKeyID2
                        End If
                        Rec.Close
                        Rec.Open "Select intVoucherID,intVoucherNo From faVouchers Where tnyVoucherTypeID = 20 And intKeyID2 = " & mPayOrderNo, mCnn
                        ' Reconciliation Process
                        If Not (Rec.BOF And Rec.EOF) Then
                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(mTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mVoucherID) & "'Where intVoucherID = " & Rec!intVoucherID
                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,numTockenID = Null,dtRealisationDate = '" & Format(mTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mVoucherID
                            ' Reverse ID updation '
                            mCnn.Execute "Update faVouchers set tnysync=Null,intExternalModuleID = 70, numLinkKeyID = " & Rec!intVoucherID & " And intKeyID2=" & mPayOrderNo & "Where intVoucherID = " & mVoucherID
                            mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID = 70 Where intVoucherID = " & mVoucherID
                        End If
                        Rec.Close
                    End If
                Else
                    Dim mModule As Integer
                    If val(txtDemandNo.Tag) > 0 Then
                        mDemandID = txtDemandNo.Tag
                        Rec.Open "Select intKeyID2,tnyExtModuleID From faIDemandTBL Where numDemandID = " & mDemandID, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mPayOrderNo = Rec!intKeyID2
                            mModule = IIf(IsNull(Rec!tnyExtModuleID), 0, Rec!tnyExtModuleID)
                        End If
                        Rec.Close
                        If mModule = 70 Then
                            Rec.Open "Select intVoucherID,intVoucherNo From faVouchers Where tnyVoucherTypeID = 20 And intKeyID2 = " & mPayOrderNo, mCnn
                            If Not (Rec.BOF And Rec.EOF) Then
                                mCnn.Execute "Update faVouchers set tnysync=Null,intExternalModuleID = 70, numLinkKeyID = " & Rec!intVoucherID & " ,intKeyID2=" & mPayOrderNo & " Where intVoucherID = " & mVoucherID
                                mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID = 70 Where intVoucherID = " & mVoucherID
                            End If
                            Rec.Close
                        End If
                    End If
                End If
                '---------------------------------------------------------------------'
                
                '========================================='
                ' Sharing Data to KMBR and SOOCHIKA       '
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
                    If mKMBRFlag = True Then
                        Dim mCnnKMBR As New ADODB.Connection
                        If objdb.CreateNewConnection(mCnnKMBR, enuSourceString.KMBR) = True Then
                            mCnnKMBR.BeginTrans
                            If SaveSanketham(lSoochikaCurrentNo, mCnnKMBR) = True Then
                                
                                mCnnSoochika.CommitTrans
                                mCnnKMBR.CommitTrans
                                
                            Else
                                GoTo ErrorRollBack:
                            End If
                        End If
                    End If
                End If
                '========================================='
                
                
                '========================================='
                '             Saving to Sevana Reg        '
                '========================================='
                If mSoochikaConnected = True Then
                    Dim mCnnSevanaReg As New ADODB.Connection
                    If objdb.CreateNewConnection(mCnnSevanaReg, enuSourceString.SevanaRegn) = True Then
                        mCnnSevanaReg.BeginTrans
                        On Error GoTo ErrorRollBack:
                        
                        If gbSoochikaVer = 5 Then
                            '--------------------------------------------------------
                            '    Added By Akheel 09.03.11 For Unicode version
                            '--------------------------------------------------------
                            If (InwardMode = 0) Then
                                If frmUSoochikaInward.SaveSevana(lSoochikaFileID, frmUSevanaInward.SevanaTypeID, frmUSevanaInward.SevanaKioskID, mReceiptNo, mfltAmount_9, mCnnSevanaReg) = True Then
                                    mCnnSoochika.CommitTrans
                                    mCnnSevanaReg.CommitTrans
                                Else
                                    GoTo ErrorRollBack:
                                End If
                            Else
                                If frmUSoochikaManualInward.SaveSevana(lSoochikaFileID, frmUSevanaInward.SevanaTypeID, frmUSevanaInward.SevanaKioskID, mReceiptNo, mfltAmount_9, mCnnSevanaReg) = True Then
                                    mCnnSoochika.CommitTrans
                                    mCnnSevanaReg.CommitTrans
                                Else
                                    GoTo ErrorRollBack:
                                End If
                            End If
                            '--------------------------------------------------------
                        Else
                            If frmSoochikaInward.SaveSevana(lSoochikaCurrentNo, mReceiptNo, mfltAmount_9, mCnnSevanaReg) = True Then
                                mCnnSoochika.CommitTrans
                                mCnnSevanaReg.CommitTrans
                            Else
                                GoTo ErrorRollBack:
                            End If
                        End If
                    End If
                End If
                '========================================='
                
''''''                '-------------------------------------------------'
''''''                '             Saving to Sevana Pension        '
''''''                '-------------------------------------------------'
''''''                If gbLinkWithMOReturn = True Then
''''''                    Dim mCnnSevanaPension   As New ADODB.Connection
''''''                    Dim mArrInMO            As Variant
''''''                    Dim mRowCount           As Integer
''''''
''''''                    If objDB.CreateNewConnection(mCnnSevanaPension, enuSourceString.SevanaPension) Then
''''''                        mCnnSevanaPension.BeginTrans
''''''                        On Error GoTo ErrorRollBack:
''''''                        For mRowCount = 1 To vsGrid.Rows - 1
''''''                            mArrInMO = Array(gbLocalBodyID, _
''''''                                        vsGrid.TextMatrix(mReceiptDate, 17), _
''''''                                        vsGrid.TextMatrix(mReceiptDate, 18), _
''''''                                        vsGrid.TextMatrix(mReceiptDate, 19), _
''''''                                        txtReceiptNo.Text, _
''''''                                        txtDate.Text, _
''''''                                        txtGrandTotal.Text _
''''''                                        )
''''''                            objDB.ExecuteSP "KMAM_Remittance", mArrInMO, , , mCnnSevanaPension, adCmdStoredProc
''''''                        Next
''''''                    End If
''''''                        mCnnSevanaPension.CommitTrans
''''''                End If
''''''                '------------------------------------------------------------'
                
GotoCommitTran:
                
                '========================================='
                ' TRANSACTION COMMITTING                  '
                '-----------------------------------------'
                    mCnn.CommitTrans
                    Set mCnn = Nothing
                    On Error GoTo 0
                '========================================='
                Call LockForm(False)
                mdtDate = gbTransactionDate
                mPreviousYearMode = 0
                mGrandTotal = mGrandTotal + val(txtTotal)
                If mStartingReceiptNo = 0 Then
                    mStartingReceiptNo = txtReceiptNo.Text
                    lblFromReceiptNo.Caption = mStartingReceiptNo
                End If
                lblToReceiptNo.Caption = txtReceiptNo.Text
                lblGroupTotal.Caption = mGrandTotal
                If mWebExtractMode = True Then
                    If (objdb.SetConnection(mCnn)) Then
                        objdb.ExecuteSP "Update faWebExtracts set intExtractTypeID=1,numKeyID=" & mVoucherID & " Where intWebExtractID=" & mSubLedgerID, , , , mCnn, adCmdText
                        MsgBox "Saved Sucessfully"
                        Unload Me
                    End If
                End If
                If mReverseMode = True Then
                    If mintVoucherID_1 > 0 Then
                        frmListReverseEntryRequests.ReceiptID = mintVoucherID_1
                        frmListReverseEntryRequests.mReceiptMode = True
                    End If
                ElseIf mWebExtractMode = True Then
                
                Else
                    If mInterruptedModeFlag = False Then
                        If gbLBPanchayat = 1 Then               'ADDED BY MINU ON 26/09/2011
                            If gbLBID = 26666 Then
                                Call PrintReceipt(mintVoucherID_1)
                            Else
'                                Call PrintReceipt_ForNewFormat(mintVoucherID_1)
                                 'Call PrintReceipt_ForNewFormatRes(mintVoucherID_1) ''' Implemented By Reshma Replacing PrintReceipt_ForNewFormat
                                 'Kill "C:\Report.txt"
                                mPFlag = PrintReceipt_ForNewFormatRes(mintVoucherID_1)
                                Kill "C:\Report.txt"
                            End If
                        Else
                            mPFlag = PrintReceipt(mintVoucherID_1) 'PrintReceipt_ForNewFormatResForUrban(mintVoucherID_1)
                            'mPFlag = PrintReceipt_ForNewFormatResForUrban(mintVoucherID_1)
                            Kill "C:\Report.txt"
                            
                        End If 'Call PrintReceipt(mintVoucherID_1)
                    Else
                        If chkIntrNoSuffix.Value = vbChecked Then
                            ''------To Update DoorNoP3 of First Record------
                            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                            Dim mVrNo As String
                            mVrNo = Replace(txtReceiptNo.Text, "-", "")
                            mSql = "Update faVouchers set tnysync=Null,vchDoorNoP3='A' Where vchDoorNoP3 is Null  And intVoucherNo=" & mVrNo
                            mCnn.Execute mSql
                            ''
                        End If
                        Call UpdateIRRegister(mintVoucherID_1, mTransactionDate, mfltAmount_9)
                    End If
                End If
               
            '=================== For Zoanl Collection================= Added by sunil
            If mDemandMode = 9 Then
                    Dim nn As Variant
                    Dim rr As Variant
                    Dim mCount As Variant
                    Dim mSQLFin As String
                    Dim mCnnFin As New ADODB.Connection
                    Dim mCnnHODemand As New ADODB.Connection
                    Dim RecFin As New ADODB.Recordset
                    nn = frmTransactionTypeWiseDemandInbox.vsGrid.Rows
                    For rr = 1 To nn - 1
                        If frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(rr, 1) = "" Then
                             Exit For
                         End If
                    Next rr
                    mCount = rr - 1
                    mSQLFin = " Select count (distinct intVoucherNo) as count from faVouchers"
                    mSQLFin = mSQLFin + " Left Join faVoucherChild on faVouchers.intVoucherID=faVoucherChild.intVoucherID"
                    mSQLFin = mSQLFin + " left Join faIdemandTbl on faVoucherChild.numDemandID=faIdemandTbl.numDemandID"
                    mSQLFin = mSQLFin + " Where faVoucherChild.numDemandID =" & mnumDemandID_9

                    If (objdb.CreateNewConnection(mCnnFin, enuSourceString.Saankhya)) = True Then
                            RecFin.Open mSQLFin, mCnnFin
                    End If
                    If Not (RecFin.EOF And RecFin.BOF) Then
                        If mCount = RecFin!count Then

                            If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO)) Then
                                arrInput = Array(mDemandID, mVoucherID, mdtDate, 1) ''Changed gbTransactionDate with mdtDate
                                objdb.ExecuteSP "spUpdateDemandStatus", arrInput, , , mCnn, adCmdStoredProc
                            Else
                                MsgBox "SaankhyaHo Connection Does not exists"
                            End If
                       End If
                   End If
                   
                    '==========Updte Demand Status in Child Table=============== Added by sunil 22-08-2011
                    If frmTransactionTypeWiseDemandInbox.vsGrid.TextMatrix(frmTransactionTypeWiseDemandInbox.vsGrid.Row, 5) = 1 Then
                        Dim tnyStat As Variant
                        tnyStat = 11
                    Else
                        tnyStat = 1
                    End If
                    
                    'mCnn.Close
                   ' mCnn.Open
                      If (objdb.CreateNewConnection(mCnnHODemand, enuSourceString.SaankhyaHO)) Then
                                arrInput = Array(mDemandID, mVoucherID, mdtDate, tnyStat, val(txtTransactionType.Tag)) ''Changed gbTransactionDate with mdtDate
                                objdb.ExecuteSP "spUpdateDemandChildStatus", arrInput, , , mCnnHODemand, adCmdStoredProc
                            Else
                                MsgBox "SaankhyaHo Connection Does not exists"
                      End If
                      
                    '=============================================================
             Else
                '========================================='
                ' Soochika Inward Printing
                '========================================='
                If mSoochikaConnected = True Then
                    On Error GoTo 0
                    '---
                    'Added By Anisha On 28/Jun/2012
                     Dim mCnnUSoochika As New ADODB.Connection
                     
                    If (objdb.CreateNewConnection(mCnnUSoochika, enuSourceString.SoochikaUnicode)) Then
                        arrInput = Array(lSoochikaFileID, val(txtTotal.Text))
                        objdb.ExecuteSP "TblSMSStatus_U", arrInput, , , mCnnUSoochika, adCmdStoredProc
                    End If
                    '------
                    If gbSoochikaVer = 5 Then
                        If (InwardMode = 0) Then
                          '  frmUSoochikaInward.ShowAckReport (lSoochikaFileID)
                            Unload frmUSevanaInward
                            frmUSoochikaInward.DisableControls
                            frmUSoochikaInward.cmdNew.Enabled = True
                            frmUSoochikaInward.cmdSave.Enabled = False
                            frmUSoochikaInward.cmdNew.SetFocus
                        Else
                            Unload frmUSevanaInward
                            frmUSoochikaManualInward.DisableControls
                            frmUSoochikaManualInward.cmdNew.Enabled = True
                            frmUSoochikaManualInward.cmdSave.Enabled = False
                            frmUSoochikaManualInward.cmdNew.SetFocus
                        End If
                    Else
                      '  frmSoochikaInward.Ack (frmSoochikaInward.lSoochikaFeildID)
                        Unload frmSevanaInward
                        frmSoochikaInward.ClearDetails
                    End If
                    'Call FormInitialize    ' Changed by Aiby on 28-Sep-2012
                    'Unload Me              ' Changed by Aiby on 28-Sep-2012
                    cmdNew.Enabled = False  ' Changed by Aiby on 28-Sep-2012
                    
                End If
                
                  
                'Call FormInitialize
                
                
                '------------------Anju For Kochi Corp P Tax Integration----------
                '------------------Added On 29 Sept 2015--------------------------
                
                
'''''''                If gbLocalBodyID = 169 Then
'''''''                    Dim mToPerid        As String
'''''''                    Dim mFromPerid      As String
'''''''                    Dim mBuildingId_Web  As String
'''''''                    Dim mColAmnt         As String
'''''''                    Dim mCol_Date        As String
'''''''                    Dim mColReceipt_No   As String
'''''''                    Dim mFromHalf       As String
'''''''                    Dim mToHalf         As String
'''''''                    Dim mColFine        As String
'''''''                    Dim mUrll            As String
'''''''                    Dim mColPenalInterast            As Integer
'''''''                    Dim mColPost       As String
'''''''                    Dim xmlHttp         As Object
'''''''                    Dim mXmlString      As Variant
'''''''                    Dim oRs             As ADODB.Recordset
'''''''                    Dim oNode           As Object 'MSXML2.IXMLDOMNode
'''''''                    Dim oSubNodes       As Object 'MSXML2.IXMLDOMSelection
'''''''                    Dim oDoc            As Object
'''''''                    Dim params          As String
'''''''                    Dim xmlString       As String
'''''''
'''''''                    Set xmlHttp = CreateObject("MSXML2.xmlHttp")
'''''''
'''''''
'''''''                    If mTransactionType = gbTransactionTypePTax Then
'''''''                        Set Rec = GetRecordSet("spGetVoucherDetails_Tcs " & mintVoucherID_1, adOpenKeyset, adLockOptimistic)
'''''''                        If Not (Rec.EOF And Rec.BOF) Then
'''''''                            mBuildingId_Web = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)   'id
'''''''                            mCol_Date = IIf(IsNull(Format(Rec!dtDate, "yyyy-mm-dd")), "", Rec!dtDate)   'dateOfReceipt
'''''''                            mColReceipt_No = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)        'Receipt No
'''''''                            mColFine = IIf(IsNull(Rec!Fine), "", Rec!Fine)                             'Fine
'''''''                            mColAmnt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)                    'amountPaid
'''''''                            mFromPerid = IIf(IsNull(Rec!Fyear), "", Rec!Fyear)             'paymentPeriodFrom
'''''''                            mToPerid = IIf(IsNull(Rec!ToYear), "", Rec!ToYear)                   'paymentPeriodTo
'''''''                            mFromHalf = IIf(IsNull(Rec!FPeriod), "", Rec!FPeriod)          'halfYearFrom
'''''''                            mToHalf = IIf(IsNull(Rec!TPeriod), "", Rec!TPeriod)                'halfYearTo
'''''''                        End If
'''''''
'''''''                        mColPost = mColPost + CStr(mBuildingId_Web) + "~" + CStr(mintVoucherID_1) + "~" + Format(mCol_Date, "yyyy-mm-dd") + "~"
'''''''                        mColPost = mColPost + CStr(mColReceipt_No) + "~" + CStr(mColFine) + "~" + CStr(mColAmnt) + "~"
'''''''                        mColPost = mColPost + CStr(mFromPerid) + "~" + CStr(mToPerid) + "~" + CStr(mFromHalf) + "~"
'''''''                        mColPost = mColPost + CStr(mToHalf) + "~" + "NA"
'''''''                        mUrll = gbDefaultUrl + "/updatePaymentDtls?paymentUpdateParam=" + mColPost
'''''''                        xmlHttp.Open "POST", mUrll, False
'''''''                        xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-"
'''''''                        xmlHttp.send
'''''''                        'MsgBox xmlHttp.responseText
'''''''                    End If
'''''''                '========================================='
'''''''                ' Sharing Data to SanchayaWeb            '
'''''''                '-----------------------------------------'
'''''''                Else
                
                If gbFetchDemandFromWeb = 1 Then
                    If mTransactionType = gbTransactionTypePTax Then
                        If mDemandMode <= 1 Then
                            Dim mCollPost       As String
                            Dim mColZoneID      As String
                            Dim mBuildingIdWeb  As String
                            Dim mColAmt            As String
                            Dim mColDate        As String
                            Dim mColReceiptNo   As String
                            Dim mColBookNo      As String
                            Dim mColPeriodId     As String
                            Dim mColYearID       As String
                            Dim mHash           As String
                            Dim mCollOut        As String
        '                    Dim node            As IXMLDOMNode
        '                    Dim DataNodes       As IXMLDOMNodeList
                            Dim mUrl            As String
                            Dim objSOAP         As Variant
                            Dim mLen            As Integer
                            Dim mColAccID       As String
                            Dim mColKeyID       As String
                        
                    
                    
                        
                            mUrl = gbDefaultUrlSanchayaPost
                            Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                            objSOAP.MSSoapInit mUrl + "?WSDL"
                            Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                            If Not (Rec.EOF And Rec.BOF) Then
                                While Not Rec.EOF
                                
                                    mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                                    mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                                    mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                    mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
                                    mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                                    mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                                    mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                                    mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                                    mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                                    mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                                    If mColAccID <> gbAcHeadIDPenalInterest Then
                                        If mColAccID <> gbAcHeadIDNoticeFee Then
                                            mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
                                            mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mintVoucherID_1) + "#"
                                            mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
                                            mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
                                            mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
                                            mCollPost = mCollPost + "~"   ''Modified on 27/Dec/2016
                                        End If
                                    End If
                                    Rec.MoveNext
                                    'mCollPost = mCollPost + "~"
                                Wend
                                'mLen = Len(mCollPost) - 1
                                mCollPost = Left$(mCollPost, Len(mCollPost) - 1)
                                mHash = CStr(mintVoucherID_1) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
                                mCollOut = objSOAP.Saankhyaa_CollectionPosting(mCollPost, mHash)
                            End If
                        End If
                    End If
                '========================================='
                ' Sharing Data to SanchayaLite            '
                '-----------------------------------------'
                ElseIf gbLinkWithPropertyTax Then
                    If mTransactionType = 1 Then
                        Dim mchvReceiptNO As String
                        Dim mIsAdvance As Integer
                        
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If Not (Rec.EOF And Rec.BOF) Then
                        
                            Set mCnn = Nothing
                            If objdb.CreateNewConnection(mCnn, SanchayaLite) Then
                                
                                Dim intKeyID As Variant
                                Dim chvReceiptNo As Variant
                                Dim chvReceiptDate As Variant
                                Dim intCollectionYear As Variant
                                Dim tnySource As Variant
                                Dim tnyPaymentReceived As Variant
                                Dim numLocation As Variant
                                Dim fltAmt As Variant
                                
                                If frmPropertyTax.mvarDifferentZoneFlag = False Then
                                    While Not Rec.EOF
                                        '@intKeyID   Int,
                                        '@intVoucherID  BigInt,
                                        '@chvReceiptNo varChar(20),
                                        '@chvReceiptDate varChar(12),
                                        '@intCollectionYear Int,
                                        '@tnySource Tinyint ,
                                        '@tnyPaymentReceived TinyInt
                                        '@numLocation Numeric
                                        intKeyID = Rec!numDemandID
                                        mintVoucherID_1 = Rec!intVoucherID
                                        chvReceiptNo = Rec!intVoucherNo
                                        chvReceiptDate = Rec!dtDate
                                        intCollectionYear = mYearID
                                        tnySource = 2
                                        tnyPaymentReceived = 1
                                        
'                                        arrInput = Array(intKeyID, _
'                                                            mintVoucherID_1, _
'                                                            chvReceiptNo, _
'                                                            chvReceiptDate, _
'                                                            intCollectionYear, _
'                                                            tnySource, _
'                                                            tnyPaymentReceived)
'                                        objdb.ExecuteSP "spCloseDemandFromSaankhya", arrInput, , , mCnn
                                        arrInput = Array(gbLBID, _
                                            gbLocationID, _
                                            mvarSubLedgerID, _
                                            Rec!intYearID, _
                                            Rec!tnyPeriodID, _
                                            mintVoucherID_1, _
                                            chvReceiptNo, _
                                            chvReceiptDate, _
                                            mYearID, _
                                            Rec!intAccountHeadID, _
                                            Rec!fltAmount, _
                                            Null, _
                                            2, _
                                            0, _
                                            gbLBName, _
                                            115)
                                        
                                        objdb.ExecuteSP "spSanPropertyTaxCollectionPosting_IU", arrInput, , , mCnn
                                        Rec.MoveNext
                                    Wend
                                    
                                    arrInput = Array(gbLBID, gbLocationID, chvReceiptNo, chvReceiptDate)
                                    objdb.ExecuteSP "spSanCollectionAutherization", arrInput, , , mCnn
                                    
                                    '========================================================================================'
                                    ' ADVANCE CLOSING USING GRID VALUE                                                       '
                                    '========================================================================================'
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) = 1 Then
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            arrInput = Array(intKeyID, _
                                                            mintVoucherID_1, _
                                                            chvReceiptNo, _
                                                            chvReceiptDate, _
                                                            intCollectionYear, _
                                                            tnySource, _
                                                            tnyPaymentReceived)
                                            objdb.ExecuteSP "spCloseDemandFromSaankhya", arrInput, , , mCnn
                                        End If
                                    Next
                                Else
                                    '====================================================================='
                                    '   Modified On 12-aug-2009 by Cijith For Sanchaya Zonal Connectivity'
                                    '====================================================================='
                                    Dim intcnt As Integer
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If vsGrid.TextMatrix(mLoopCount, 6) = "" Then Exit For
                                        If vsGrid.TextMatrix(mLoopCount, 6) <> 113 Then
                                            intcnt = intcnt + 1
                                        End If
                                    Next
                                    
                                    arrInput = Array(gbLocationID, mintVoucherID_1, Rec!intVoucherNo, _
                                                    Rec!dtDate, Rec!numSubLedgerID, _
                                                    Rec!numZoneID, mAssesmentYearID, _
                                                    Rec!numWardId, Rec!intDoorNoP1, Rec!vchDoorNoP2, _
                                                    vchName, 2, mYearID, 0, Rec!fltTotalAmt, _
                                                    intcnt, 1)
                                    objdb.ExecuteSP "HOsnSaanOtherCollectionsI", arrInput, , , mCnn
                                    
                                    Dim numSanchayaHeadId As Integer
                                    Dim numSankhyaHeadID As Integer
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                                            numSankhyaHeadID = val(vsGrid.Cell(flexcpText, mLoopCount, 6))
                                            If numSankhyaHeadID = 1385 Or numSankhyaHeadID = 1386 Then
                                               numSanchayaHeadId = 1
                                            ElseIf numSankhyaHeadID = 1126 Then
                                                numSanchayaHeadId = 2
                                            ElseIf numSankhyaHeadID = 1157 Then
                                                numSanchayaHeadId = 4
                                            ElseIf numSankhyaHeadID = 113 Then ' Modified By Aiby  To Give Penal Interest
                                                numSanchayaHeadId = 90
                                            Else
                                                numSanchayaHeadId = 0
                                            End If
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            fltAmt = IIf(val(vsGrid.Cell(flexcpText, mLoopCount, 5)) = 0, val(vsGrid.Cell(flexcpText, mLoopCount, 4)), val(vsGrid.Cell(flexcpText, mLoopCount, 5)))
                                            arrInput = Array(gbLocationID, _
                                                        mintVoucherID_1, _
                                                        mLoopCount, _
                                                        val(vsGrid.Cell(flexcpText, mLoopCount, 7)), _
                                                        val(vsGrid.Cell(flexcpText, mLoopCount, 8)), _
                                                        numSanchayaHeadId, _
                                                        fltAmt, _
                                                        intKeyID)
                                            objdb.ExecuteSP "HOsnSaanOtherCollectionsSubI", arrInput, , , mCnn
                                        End If
                                    Next
                                End If
                                '========================================================================================'
                                
                                '----------------------------------------------------------------------------------------'
                                '----------------------------Give Advance to Sanchaya------------------------------------'
                                '---------------------------------------Sinoj Added--------------------------------------'
                                
                                If mBoolGiveAdvanceToSanchaya Then
                                    numLocation = cmbZone.ItemData(cmbZone.ListIndex)
                                    arrInput = Array(gbLocationID, _
                                                10, _
                                                gbLocalBodyID, _
                                                mYearID, _
                                                gbCurrentPeriodID, _
                                                mSubLedgerID, _
                                                "Advance Collection From Saankhya(JSK)" + vbNewLine + txtDescription.Text, _
                                                mAdvAmtToSanchaya, _
                                                mYearID, _
                                                0, _
                                                chvReceiptNo, _
                                                mTransactionDate, _
                                                gbLocationID, _
                                                2, _
                                                numLocation, _
                                                mintVoucherID_1)
                                    objdb.ExecuteSP "spSanSnPropertyTaxDemandAdvance", arrInput, , , mCnn

                                End If
                                '----------------------------------------------------------------------------------------'
                            Else
                                MsgBox "(Sanchaya)Connection Error:", vbInformation
                            End If
                        End If
                    End If
                End If

                '============== ==========================='
                ' Updating Demand Details For Rent on Land / Buildings (DB_Sanchaya)       '
                ' Codded By Anisha
                '-----------------------------------------'
                If gbLinkWithRentOnLand Then
                    If mTransactionType = gbTransactionTypeRentOnBuilding Or mTransactionType = gbTransactionTypeRentOnLand Then
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If Not (Rec.EOF And Rec.BOF) Then
                                If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                                    Dim numRLBDemand As Variant
                                    Dim numZonalOfficeID As Variant
                                    Dim numVoucherId As Variant
                                    Dim dtReceiptDate As String
                                    Dim tnyReceiptSource As Integer
                                    Dim numSubItemId   As Variant
                                    While Not Rec.EOF
                                        numRLBDemand = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID) 'intKeyID From Sanchaya
                                        numZonalOfficeID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                                        numVoucherId = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                                        chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                        dtReceiptDate = Rec!dtDate
                                        tnyReceiptSource = 2
                                        arrInput = Array(numRLBDemand, _
                                                        numVoucherId, _
                                                        numZonalOfficeID, _
                                                        gbLocation, _
                                                        chvReceiptNo, _
                                                        dtReceiptDate, _
                                                        mYearID, _
                                                        tnyReceiptSource)
                                        objdb.ExecuteSP "spSanSnRentDemandClose", arrInput, , , mCnn, adCmdStoredProc
                                        Rec.MoveNext
                                    Wend
                                    
                                    '-----------------------------------------------------------------------'
                                    '                       Addvance Clossing using Grid Value              '
                                    
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) = 1 Then
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            arrInput = Array(intKeyID, _
                                                        numVoucherId, _
                                                        numZonalOfficeID, _
                                                        gbLocation, _
                                                        chvReceiptNo, _
                                                        dtReceiptDate, _
                                                        mYearID, _
                                                        tnyReceiptSource)
                                            objdb.ExecuteSP "spSanSnRentDemandClose", arrInput, , , mCnn, adCmdStoredProc
                                        End If
                                    Next
                                    '-----------------------------------------------------------------------------'
                                    '-------To Give Advance Collection Amount To Sanchaya
                                    If mBoolGiveAdvanceToSanchaya Then
                                           numSubItemId = val(txtDoorNo2.Tag) 'Sets txtDoorNo2.tag Value as SubItemID from RentOnLand Form
                                          arrInput = Array(gbLocalBodyID, _
                                                      mnumSubLedgerID_21, _
                                                      numSubItemId, _
                                                      mYearID, _
                                                      Month(mTransactionDate) + 10, _
                                                      mAdvAmtToSanchaya, _
                                                      numVoucherId, _
                                                      numZonalOfficeID, _
                                                      "", _
                                                      chvReceiptNo, _
                                                      dtReceiptDate, _
                                                      mYearID, _
                                                      tnyReceiptSource)
                                        objdb.ExecuteSP "spSanSnRentDemandAdvance", arrInput, , , mCnn
                                    End If
                                End If
                        End If
                    End If
                End If
                
                
                '========================================='
                ' Updating Demand Details For ProfessionTax Traders (DB_Sanchaya)       '
                ' Codded By Poornima
                '-----------------------------------------'
                If gbLinkWithProfTaxEmp Then
                    If mTransactionType = gbTransactionTypeProfTaxTrade Then
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If Not (Rec.EOF And Rec.BOF) Then
                                If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                                    Dim numProfDemand As Variant
                                    Dim numZonalOficeID As Variant
                                    Dim mVID As Variant   ' VoucherID
                                    Dim mReceiptDate As String
                                    Dim mReceiptSource As Integer
                                    Dim mSubItemId   As Variant
                                    While Not Rec.EOF
                                    
                                                                       
                                        numProfDemand = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID) 'intKeyID From Sanchaya
                                        numProfDemand = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                                        numZonalOficeID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                                        mVID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                                        chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                        mReceiptDate = Rec!dtDate
                                        mReceiptSource = 2
                                        
                                        'arrInput = Array(numProfDemand, _
                                        '                mVID, _
                                        '                numZonalOficeID, _
                                        '                gbLocation, _
                                        '                chvReceiptNo, _
                                        '                mReceiptDate, _
                                        '                mYearID, _
                                        '                mReceiptSource)
                                        'objDB.ExecuteSP "spSanSnProfTaxDemandClose", arrInput, , , mCnn, adCmdStoredProc
                                        '
                                        
                                        
                                        
                                        arrInput = Array(gbLBID, _
                                            gbTransactionTypeProfTaxTrade, _
                                            gbLocationID, _
                                            numProfDemand, _
                                            Null, _
                                            Rec!intYearID, _
                                            Rec!tnyPeriodID, _
                                            mintVoucherID_1, _
                                            chvReceiptNo, _
                                            mReceiptDate, _
                                            mYearID, _
                                            Rec!intAccountHeadID, _
                                            Rec!fltAmount, _
                                            2, _
                                            0, _
                                            gbLBName, _
                                            115)
                                        objdb.ExecuteSP "snSaankhyaTransactionTBL_I", arrInput, , , mCnn, adCmdStoredProc
                                        Rec.MoveNext
                                    Wend
                                    
                                    
                                    'spSanProfTaxCollectionPosting
                                    '@intLBID int,
                                    '@intLocationID numeric,
                                    '@chvReceiptNo varchar(50),
                                    '@chvReceiptDate varchar(50),
                                    '@intTransType int
                                    arrInput = Array(gbLBID, gbLocationID, chvReceiptNo, mReceiptDate, gbTransactionTypeProfTaxTrade)
                                    objdb.ExecuteSP "spSanProfTaxCollectionPosting", arrInput, , , mCnn
                                    
                                 
                                End If
                                    '-----------------------------------------------------------------------'
                                    '                       Addvance Clossing using Grid Value              '
                                    
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) = 1 Then
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            arrInput = Array(intKeyID, _
                                                        mVID, _
                                                        numZonalOficeID, _
                                                        gbLocation, _
                                                        chvReceiptNo, _
                                                        mReceiptDate, _
                                                        mYearID, _
                                                        mReceiptSource)
                                            objdb.ExecuteSP "spSanSnProfTaxDemandClose", arrInput, , , mCnn, adCmdStoredProc
                                        End If
                                    Next
                            End If
                         End If
                      End If
                                    
                 ''Added On  11/Apr/2018     For Profession tax Trade/institution Web Integration in Tvm Corp
                 If gbLinkWithProfTradeWeb Or gbLinkWithProfEmpWeb Then
                    If mTransactionType = gbTransactionTypeProfTaxTrade Or mTransactionType = gbTransactionTypeProfTaxEmp Then
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        Dim mInstType As String
                        Dim mArrOutDemand       As Variant
                        If mTransactionType = gbTransactionTypeProfTaxTrade Then
                            mInstType = 1
                        ElseIf mTransactionType = gbTransactionTypeProfTaxEmp Then
                            mInstType = 2
                        End If
                        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                        
                         mUrl = gbDefaultUrl
                         On Error Resume Next
                         objSOAP.MSSoapInit (mUrl + "?WSDL")
                         If Not (Rec.EOF And Rec.BOF) Then
                                
                                    Dim numDemandTrade As Variant
                                    Dim numProfInst As Variant
                                    Dim numZonalID As Variant
                                  '  Dim mReceiptSource As Integer
                                    'Dim mSubItemId   As Variant
                                    Dim mTotAmt As Variant
                                    Dim mCollAmt    As Variant
                                    Dim intYearID As Integer
                                    Dim mPeriodID As Integer
                                    Dim mPenalAmt As Integer
                                    Dim mCredential As String
                                    
                                    mCredential = "ikm@revenue@sanchaya"
                                    While Not Rec.EOF
                                    
                                        mPenalAmt = 0
                                        numDemandTrade = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID) 'intKeyID From Sanchaya
                                        numProfInst = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                                        numZonalID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                                        mVID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                                        chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                        mReceiptDate = Rec!dtDate
                                        mTotAmt = IIf(IsNull(Rec!fltTotalAmt), "", Rec!fltTotalAmt)
                                        mCollAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                                        mYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                                        mPeriodID = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                                        mReceiptSource = 2
                                        
                                        If Rec!intAccountHeadID = gbAcHeadIDPenalInterest Then
                                           mPenalAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                                        End If
                                        
                                        mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(numDemandTrade) + "#" + CStr(numProfInst) + "#"
                                        mCollPost = mCollPost + CStr(numZonalID) + "#" + CStr(mintVoucherID_1) + "#" + CStr(chvReceiptNo) + "#"
                                        mCollPost = mCollPost + CStr(mReceiptDate) + "#" + CStr(mTotAmt) + "#" + CStr(mReceiptSource) + "#"
                                        mCollPost = mCollPost + CStr(mCollAmt) + "#" + CStr(mYearID) + "#" + CStr(mPeriodID) + "#"
                                        'mCollPost = mCollPost + CStr(mPenalAmt) + "#" + CStr(mInstType) + "#" + CStr(mCredential)
                                        mCollPost = mCollPost + CStr(mPenalAmt) + "#" + CStr(mCredential)
                                        mCollPost = mCollPost + "~"
                                     Rec.MoveNext
                                    Wend
                                    mCollPost = mCollPost
                                    mArrOutDemand = objSOAP.save_prof_receiptdetails(mCollPost)
'                               (  'voucherdetails'=>lb_prof'#demandno_prof#institution_id#zonal_id#VoucherID#collection#year#period#penal)
                      End If
                    End If
                End If
                
                
                '========================================='
                ' Insert Receipt Details On DB_SanchayaLite  For PFA,D&O  Licence Fee    '
                ' Created On ON 18/02/2010 By Anisha        '
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypeDandO Or mTransactionType = gbTransactionTypePFA Then
                    Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                    If gbLinkWithDandOPFA Then
                        'Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
                            Dim numReceiptLocationId As Double
                            Dim intReceiptYear      As Integer
                            Dim numZoneID           As Variant
                            Dim flagSankhya             As Integer
                            tnyReceiptSource = 2
                            tnyPaymentReceived = 1
                            flagSankhya = 0
                            If Not (Rec.EOF And Rec.BOF) Then
                                chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                chvReceiptDate = Format(IIf(IsNull(Rec!dtDate), "", Rec!dtDate), "dd/mm/yyyy")
                                mDemandID = Rec!numDemandID
                                numZoneID = cmbZone.ItemData(cmbZone.ListIndex) 'gbLocationID
                                mVoucherID = Rec!intVoucherID
                                arrInput = Array(numZoneID, _
                                            gbLocalBodyID, _
                                            mVoucherID, _
                                            mDemandID, _
                                            chvReceiptDate, _
                                            flagSankhya)
                                objdb.ExecuteSP "spsnLicSanCollection_I", arrInput, , , mCnn, adCmdStoredProc
                            End If
                        Else
                            MsgBox "Connection to Sanchaya Doesn't Exists"
                        End If
                    ElseIf gbLinkWithDandOWeb And mTransactionType = gbTransactionTypeDandO Then

                        'Dim mArrOutDemand       As Variant
                        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                        
                         mUrl = gbDefaultUrl
                         On Error Resume Next
                         objSOAP.MSSoapInit (mUrl + "?WSDL")
                        
                         If Not (Rec.EOF And Rec.BOF) Then
                             chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                             chvReceiptDate = IIf(IsNull(Rec!dtDate), "", Format(Rec!dtDate, "dd/MMM/yyyy"))
                             mDemandID = Rec!numDemandID
                             numZoneID = cmbZone.ItemData(cmbZone.ListIndex) 'gbLocationID
                             mVoucherID = Rec!intVoucherID
                             Dim mCredencial As String
                             Dim VoucherDetails As String
                             mCredencial = "ikm@revenue@sanchaya"

                             arrInput = gbLocalBodyID & "#" & mDemandID & "#" & chvReceiptNo & "#" _
                                            & chvReceiptDate & "#" & flagSankhya & "#" & mVoucherID
                             mArrOutDemand = (objSOAP.savereceiptdetails(arrInput))
                         Else
                             MsgBox "Connection to Sanchaya Doesn't Exists"
                         End If
                    End If
                End If
                '========================================='
                ' Updating Receipt Details IN KMBR        '
                ' Modified ON 04/05/2009 By Cijith        '
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypePermitFeeFromKMBR And mKMBRFlag = True Then
                    If objdb.CreateNewConnection(mCnn, enuSourceString.KMBR) Then
                        arrInput = Array(mReceiptNo, mdtDate, mDemandID, mVoucherID)
                        objdb.ExecuteSP "UGetReceiptNoDate", arrInput, , , mCnn, adCmdStoredProc
                    End If
                End If
                '=================Zonal Collection===================
                'Updating Status of Demand in SaankhyaHo
                'If mTransactionType = gbTransactionTypeZonalCollection And gbLinkWithFinanceHO = 1 Then
                '
                If (mDemandMode = 2 Or mTransactionType = gbTransactionTypeZonalCollection) And gbLinkWithFinanceHO = 1 Then
                    If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO)) Then
                        arrInput = Array(mDemandID, mVoucherID, mdtDate, 1) ''Changed gbTransactionDate with mdtDate
                        objdb.ExecuteSP "spUpdateDemandStatus", arrInput, , , mCnn, adCmdStoredProc
                    Else
                        MsgBox "SaankhyaHo Connection Does not exists"
                    End If
                End If
                '====================================================
                
                '========================================='
                ' Updating Status of Fine wave        '
                ' Added ON 04/05/2009 By Anisha       '
                '-----------------------------------------'
                If mFinewave Then
                    If (objdb.SetConnection(mCnn)) Then
                        objdb.ExecuteSP "Update faFineWaiver set tnyStatus=0  Where intVoucherNo=(Select intVoucherNo From faVouchers Where intVoucherID=" & mVoucherID & ")", , , , mCnn, adCmdText
                    End If
                End If
                '''''''''''WebExtractMode post project related receipts and Update intKeyId as VoucherId in faWebExtract Table
                
                If mReverseMode = 1 Then
                    frmReverseApproval.ReceiptVrID = mVoucherID
                End If
               End If
        Else
                Debug.Print "Error in establishing connection with Saankhya DB"
                Exit Sub
        End If

        '--------------------To Calculate The Group Total ----------------------'
        'Call GroupCalc
        '-----------------------------------------------------------------------'
        mSubLedgerID = Null
            '-------------------------------------------------------'
            'INTERRUPTED REGISTER-BY MINU
            '-------------------------------------------------------'
                'If mInterruptedRegister = 1 Then
                  '  Call UpdateIRRegister(mintVoucherID_1, mTransactionDate, mfltAmount_9)
                'End If
            '-------------------------------------------------------------------------'
            ' END BLOCK::INTERRUTED REGISTER MODE                                                '
            '-------------------------------------------------------------------------'
        Exit Sub
ErrorRollBack:
        MsgBox "Saankhya Error Handler: " & Error$
        mCnn.RollbackTrans
        Set mCnn = Nothing
        
        '---------------------------------------------------------------'
        ' KMBR Roll Back
        '---------------------------------------------------------------'
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
            If mKMBRFlag = True Then
                mCnnSoochika.RollbackTrans
                mCnnKMBR.RollbackTrans
            End If
        End If
        
'''''''        '---------------------------------------------------------------'
'''''''        ' Sevana Pension Roll Back
'''''''        '---------------------------------------------------------------'
'''''''        If mTransactionType = gbTransactionTypeMOReturnsSocialSecurityPension Then
'''''''            If gbLinkWithMOReturn Then
'''''''                mCnnSevanaPension.RollbackTrans
'''''''            End If
'''''''        End If
'''''''        '---------------------------------------------------------------'
        
ErrorRollBackSoochika:

        If mSoochikaConnected = True Then
            If mCnnSevanaReg.State Then
                mCnnSevanaReg.RollbackTrans
            End If
            If mCnnSoochika.State Then
                mCnnSoochika.RollbackTrans
            End If
        End If
    'End If

       End Sub
    Private Sub cmdSearchAccountHead_Click()
        Dim mSql As String
        
        If val(txtInstrument.Tag) > 0 Then
            Select Case val(txtInstrument.Tag)
                Case 6:

                    If val(txtAccountHead.Tag) > 0 Then
                        '*******************VALIDATIONS FOR RECEIPTS FROM OTHER LSGI's*****************************
                           
                        Dim RecDate As New ADODB.Recordset
                        Dim mCnn    As New ADODB.Connection
                        Dim objdb   As New clsDB
                        Dim mMsg    As String
                        objdb.SetConnection mCnn
                        
                        If val(txtTransactionType.Tag) = 119 Or val(txtTransactionType.Tag) = 120 _
                        Or val(txtTransactionType.Tag) = 121 Or val(txtTransactionType.Tag) = 122 _
                        Or val(txtTransactionType.Tag) = 123 Then
                            RecDate.Open "Select *,getDate() as CurDate From faLBSettings", mCnn, adOpenDynamic, adLockOptimistic
                            If Not (RecDate.EOF And RecDate.BOF) Then
                                If CDate(RecDate!CurDate) < gbBankChangePermitDate Then
                                   mMsg = "You are going to Edit the Default treasury For this Transaction Type"
                                   MsgBox mMsg, vbInformation
                                Else
                                    mMsg = "The Default treasury Can not be Edited For this Transaction Type"
                                    MsgBox mMsg, vbInformation
                                    cmdSearchAccountHead.Enabled = False
                                    Exit Sub
                                End If
                                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
                                mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
                        
                            End If
                            RecDate.Close
                        End If
                    End If
                    '**********************************************************************
                
                Case 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13 '[Cheque]
                    mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
                    mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faBank
   
                Case 7  '[Treasury Bills]
                    mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
                    mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID WHERE  tinHiddenFlag = 0 AND (faAccountHeads.vchAccountHeadCode Like '45045%' Or faAccountHeads.vchAccountHeadCode Like '45065%' Or faAccountHeads.vchAccountHeadCode Like '45025%' ) "
                Case Else
                    mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE  tinHiddenFlag = 0 AND faAccountHeads.intGroupID = " & faCash
            End Select
            
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.cmdSearch.Enabled = False
            frmSearchAccountHeads.VoucherMode = 101
            frmSearchAccountHeads.Show vbModal
            txtAccountHead.SetFocus
        End If
    End Sub
    Private Sub cmdSearchDemandNo_Click()
        Load frmSearchDemandNo
        If val(txtTransactionType.Tag) > 0 Then
            frmSearchDemandNo.txtTransactionType.Tag = val(txtTransactionType.Tag)
        End If
        frmSearchDemandNo.Show vbModal
        txtDemandNo_LostFocus
        'txtDemandNo.SetFocus
    End Sub
    Private Sub cmdSearchInstrument_Click()
        Call ListMasters(2)
    End Sub
    Private Sub cmdSearchTransactionType_Click()
        'Call ListMasters(1)
        If mWebExtractMode = True Then
            frmSearchTransactionType.ModeOfTransaction = 2
            frmSearchTransactionType.SQLQry = "Select  vchTransactionType, intTransactionTypeID From faTransactiontype  Where intGroupID=20 and intTransactionTypeId in (1141,1151,1161,1171,1181,1191)"
            frmSearchTransactionType.Show vbModal
        Else
            frmSearchTransactionType.ModeOfTransaction = 1
            frmSearchTransactionType.Show vbModal
            If txtTransactionType.Enabled = False Then
                txtTransactionType.Enabled = True
            End If
        End If
        txtTransactionType.SetFocus
    End Sub
    Private Sub Command1_Click()
        Call PrintReceipt_ForNewFormat(1011)
    End Sub

    Private Sub Command2_Click()
        Call cmdSave_Click
    End Sub
    
   Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        txtDate.Text = DdMmmYy(gbTransactionDate)
        Call Calculate
        
        'Note:- Checking for Interrupted Receipt Mode When getting Focus
        Call CheckInterruptReceiptRequestStatus
        If mInterruptedModeFlag Then
            If mInterruptEditMode = False Then
                If Not (mInterruptedModeSoochikaFlag) Then
                    Call FormInitialize
                End If
                Call GetNextIRNumber
            Else
                If IsDate(mIRVoucherDate) Then
                    'txtDate.Text = DdMmmYy(CDate(mIRVoucherDate))
                    mdtDate = CDate(mIRVoucherDate)
                    'MsgBox mIRVoucherDate
                End If
            End If
        ElseIf mZonal = 1 Then
            mdtDate = Format(mZoneDate, "dd/mmm/yyyy")
        ElseIf mWebExtractMode = True Then
            mdtDate = DdMmmYy(mWebExtractDate)
        Else
            mdtDate = gbTransactionDate
        End If
        txtDate.Text = DdMmmYy(mdtDate)
    End Sub
    
    Private Sub Form_GotFocus()
        txtDate.Text = DdMmmYy(gbTransactionDate)
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 And Shift = 2 Then
            Call MsgBox("Search!", vbInformation)
        End If
        If KeyCode = vbKeyF8 Then
            txtTransactionType.Enabled = False
            If gbLocalBodyID = 171 Or gbLocalBodyID = 222 Then
               'mPTaxFormLoadFlag = True 'Only required if PropertyTax Calculator is enabled
            End If
            txtTransactionType.Tag = gbTransactionTypePTax
            txtTransactionType.Text = "Property Tax"
            'Call txtTransactionType_LostFocus
            
            '------------------------------------------------------------'
            ' Property Tax Calculator                                    '
            '------------------------------------------------------------'
            'If gbLocalBodyID = 171 Or gbLocalBodyID = 222 Then ' Changed for Municipality Implementation
                If gbLinkWithPropertyTax Then
                    Call FormInitialize
                    txtTransactionType.Tag = gbTransactionTypePTax
                    txtTransactionType.Text = "Property Tax"
                    frmPropertyTax.Show vbModal
                    If cmdSave.Enabled Then
                        cmdSave.SetFocus
                    End If
                ''''' Added On 20 jul 2015  Demand from web
                ElseIf gbFetchDemandFromWeb = 1 Then
                
                    Call FormInitialize
                    txtTransactionType.Tag = gbTransactionTypePTax
                    txtTransactionType.Text = "Property Tax"
                    frmPropertyTax.mDemandWeb = True
                    frmPropertyTax.Show vbModal
                    If cmdSave.Enabled Then
                        cmdSave.SetFocus
                    End If
                Else
                    frmPropertyTaxCalculator.Show vbModal
                End If
            'End If
            '------------------------------------------------------------'
        End If

        If KeyCode = vbKeyF9 Then
            txtTransactionType.Enabled = False
            If gbLinkWithRentOnLand Then
                Call FormInitialize
                txtTransactionType.Tag = gbTransactionTypeRentOnBuilding
                txtTransactionType.Text = "Rent on Building / Stalls"
                frmRentOnLandBuildings.mCategory = 1
'                frmRentOnLandBuildings.cmbCategory.Tag = 1
'                frmRentOnLandBuildings.cmbCategory.Text = "Building"
'                frmRentOnLandBuildings.cmbCategory.Enabled = False
                frmRentOnLandBuildings.Show vbModal
            End If
        End If
        
        If KeyCode = vbKeyF10 Then
            txtTransactionType.Enabled = False
            If gbFetchDemandFromWeb Then
                Call FormInitialize
                txtTransactionType.Tag = gbTransactionTypePTaxGp
                txtTransactionType.Text = "Property tax (Group)"
                frmPTaxMultipleBuilding.Show vbModal
            End If
        End If
        
        If Shift = 2 And KeyCode = vbKeyG Then
            vsGrid.SetFocus
        End If
    End Sub

    Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        If Shift = 2 And KeyCode = vbKeyG Then
            vsGrid.SetFocus
            vsGrid.Row = 1
        End If
        If Shift = 4 And KeyCode = vbKeyT Then
            txtTransactionType.SetFocus
        End If
    End Sub

    Private Sub Form_Load()
        XPC.InitSubClassing
        Call FillZone
        Call FillTransactionTypes
        Call FormInitialize
        vsGrid.ColComboList(0) = "|..."
        If cmbDZone.ListCount = 0 Then
            cmbDZone.AddItem "Main Office"
            cmbDZone.ItemData(cmbDZone.NewIndex) = 1
            cmbDZone.ListIndex = 0
        End If
        lblFromReceiptNo.Caption = txtReceiptNo.Text
        Call FillGridYear
        '-----------------------------------------'
        'NOTE:-Log File For Debug System Errors
        '-----------------------------------------'
         LogFile " Receipt Module : " & gbUserName & "  " & gbComputerName
        '-----------------------------------------'
        Call FillSeats
    End Sub
    
    Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If lblGroupTotal.FontBold Then
            lblGroupTotal.FontBold = False
            lblGroupTotal.ForeColor = vbDefault
        End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
        mPreviousYearMode = 0
        mInterruptEditMode = False
        mInterruptedModeFlag = False
        mInterruptedRegister = 0
        mInterruptedRegisterID = -1
    End Sub

    Private Sub lblGroupTotal_DblClick()
        If MsgBox("Reset the Starting ReceiptNo for Group Totalling?", vbYesNo) = vbYes Then
            lblFromReceiptNo.Caption = txtReceiptNo.Text
            lblToReceiptNo.Caption = txtReceiptNo.Text
            lblGroupTotal.Caption = 0
            mGrandTotal = 0
            mStartingReceiptNo = txtReceiptNo.Text
        End If
    End Sub

    Private Sub lblGroupTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblGroupTotal.FontBold = True
        lblGroupTotal.ForeColor = vbBlue
    End Sub
    
    Private Sub lstMasters_DblClick()
        Select Case val(lstMasters.Tag)
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
                    txtAccountHead.Tag = objBank.BankAccountHeadID
                End If
        End Select
        
        If lstMasters.Text = "Property Tax" Then
            'txtDemandNo.Visible = False
            'txtDemandPrefix.Visible = False
            cmdSearchDemandNo.Visible = False
        Else
            txtDemandNo.Visible = True
            txtDemandPrefix.Visible = True
            cmdSearchDemandNo.Visible = True
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

    Private Sub Timer1_Timer()
        If mTimer = 0 Then
            lblInterruptStatus.Visible = True
            imgWarning.Visible = True
            mTimer = 1
            Exit Sub
        End If
        If mTimer = 1 Then
            lblInterruptStatus.Visible = False
            imgWarning.Visible = False
            mTimer = 0
            Exit Sub
        End If
        If mTimer = 2 Then
            lblInterruptedReceipt.Visible = True
            mTimer = 3
            Exit Sub
        End If
        If mTimer = 3 Then
            lblInterruptedReceipt.Visible = False
            mTimer = 2
            Exit Sub
        End If
    
    End Sub

    Private Sub txtAccountHead_GotFocus()
        Dim mStr As String
        If gbSearchID > 0 Then
            If gbSearchID = gbAcHeadIDCash Then
                Dim objAc   As New clsAccounts
                objAc.SetAccounts (gbSearchID)
                gbSearchID = -1
                gbSearchStr = ""
                If objAc.AccountHeadID > 0 Then
                    txtAccountHead.Text = objAc.AccountHead & " [ " & objAc.AccountCode & " ]"
                    txtAccountHead.Tag = objAc.AccountHeadID
                    txtInstNo.Text = ""
                    txtDated.Text = ""
                    txtBank.Text = ""
                    txtPlace.Text = ""
                    txtInstNo.Enabled = False
                    txtDated.Enabled = False
                    txtBank.Enabled = False
                    txtPlace.Enabled = False
                    
                    txtInstrument.Text = "Cash"
                    txtInstrument.Tag = gbInstrumentCash
                End If
            Else
                Dim objBank As New clsBank
                objBank.SetBankInfoByAccID (gbSearchID)
                gbSearchID = -1
                gbSearchStr = ""
                'objBank.SetBankInfo (lstMasters.ItemData(lstMasters.ListIndex))
                If objBank.BankID > 0 Then
                    txtAccountHead.Text = objBank.BankName & " [ " & objBank.BankAccountHeadCode & " ]"
                    txtAccountHead.Tag = objBank.BankAccountHeadID
                    
                    'Change By Aiby :11/Nov/2011
                    'txtInstNo.Text = ""
                    'txtDated.Text = ""
                    'txtBank.Text = ""
                    'txtPlace.Text = ""
                    
                    txtInstNo.Enabled = True
                    txtDated.Enabled = True
                    txtBank.Enabled = True
                    txtPlace.Enabled = True
                    
                    'Added By Anisha :2/Jul/2014
                    If CDate(txtDate.Text) <= GetLastReconDate(val(txtAccountHead.Tag)) Then
                        mStr = ""
                        mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                        mStr = mStr + " No new Transaction is allowed to Enter during the period."
                        MsgBox mStr, vbInformation
                        txtAccountHead.Tag = -1
                        txtAccountHead.Text = ""
                        Exit Sub
                    End If
                    
                End If
            End If
        End If
    End Sub
    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
    End Sub
    Private Sub txtAddress_DblClick()
        fraDemandDetails.Visible = True
        Call LoadAddressVariable
        fraDemandDetails.ZOrder (0)
        fraParty.Visible = False
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
        Dim mDt As Date
        Dim mIDate As Date
        txtDated.Text = CheckDateInMMM(txtDated.Text)
        '------------------------------------------'
        '   Added To Validate the Cheque Date      '
        '------------------------------------------'
        mDt = txtDated
        If mInterruptedModeFlag = False Then
            If val(txtInstrument.Tag) = 5 Or val(txtInstrument.Tag) = 4 Then
                If mDt > gbTransactionDate Then
                    MsgBox "Post dated cheques will not accepted!", vbInformation
                    txtDated.Text = DdMmmYy(gbTransactionDate)
                    txtDated.SetFocus
                End If
                If mDt < DateAdd("d", -180, gbTransactionDate) Then
                    MsgBox "Upto Six Months Validity Cheques Can Only Accept", vbInformation
                    txtDated.Text = DdMmmYy(gbTransactionDate)
                    txtDated.SetFocus
                End If
            End If
        Else
            '------------------------------------
            '----Interrupt receipt Cheque Date validation
            '----Modified On 21/3/2011 By Anisha
            '-------------------------------------
            mIDate = txtDate
            If val(txtInstrument.Tag) = 5 Or val(txtInstrument.Tag) = 4 Then
                If mDt > mIDate Then
                    MsgBox "Post dated cheques will not accepted!", vbInformation
                    txtDated.Text = DdMmmYy(mIDate)
                    txtDated.SetFocus
                End If
                If mDt < DateAdd("d", -180, mIDate) Then
                    MsgBox "Upto Six Months Validity Cheques Can Only Accept", vbInformation
                    txtDated.Text = DdMmmYy(mIDate)
                    txtDated.SetFocus
                End If
            End If
        End If
    End Sub
    Private Sub txtDistrict_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then Call PressTabKey
    End Sub

    Private Sub txtDemandNo_GotFocus()
    '        Dim mStr As String
    '        txtDemandNo.SelStart = 0
    '        txtDemandNo.SelLength = Len(txtDemandNo)
    '        If gbSearchID > 0 Then
    '            txtDemandPrefix = Token(gbSearchStr, "-")
    '            txtDemandNo = gbSearchStr
    '            gbSearchID = -1
    '            gbSearchStr = ""
    '            Call DisplayDemandDetails
    '        End If
    End Sub
    
    Private Sub txtDemandNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
            Exit Sub
        End If
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub
    
    Public Sub txtDemandNo_LostFocus()
        Dim mStr As String
        txtDemandNo.SelStart = 0
        txtDemandNo.SelLength = Len(txtDemandNo)
        If gbSearchID > 0 Then
            txtDemandPrefix = Token(gbSearchStr, "-")
            txtDemandNo = gbSearchStr
            gbSearchID = -1
            gbSearchStr = ""
            Call DisplayDemandDetails
        ElseIf Trim(txtDemandNo.Text) <> "" Then
            Call DisplayDemandDetails
        End If
    End Sub
    
    Private Sub txtDemandPrefix_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call PressTabKey
            Exit Sub
        End If
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

'''''    Private Sub txtDemandPrefix_LostFocus()
'''''        ''Integration with DO And PFA Using Service
'''''        ''Created On 16-Feb-10 By Anisha
'''''
'''''        Dim client          As New MSSOAPLib.SoapClient
'''''        Dim objDb           As New clsDB
'''''        Dim mcnn            As New ADODB.Connection
'''''        Dim Rec             As New Recordset
'''''        Dim mRec            As New Recordset
'''''        Dim objSOAP         As Variant
'''''        Dim mUrl            As String
'''''        Dim mArrOutChild    As String
'''''        Dim mArrOutDemand   As String
'''''        Dim mDemand         As Variant
'''''        Dim mArrIN          As Variant
'''''        Dim mSql            As String
'''''        If chkLinkDemand.value = vbChecked Then
'''''            Set objSOAP = CreateObject("MSSOAP.SOAPClient")
'''''            mUrl = gbDefaultUrl 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultUrl")
'''''            On Error Resume Next
'''''            objSOAP.MSSoapInit (mUrl + "?WSDL")
'''''            If Err.Number = -2147352567 Then
'''''                MsgBox "Please Uncheck And Enter The Demand No"
'''''                Exit Sub
'''''            End If
'''''            On Error GoTo 0
'''''            If txtDemandPrefix.Text = "" Then
'''''                MsgBox "Please Enter Demand Number"
''''''                txtDemandPrefix.SetFocus
'''''                Exit Sub
'''''            Else
'''''                mDemand = Trim(txtDemandPrefix.Text)
''''''                mSQL = "Select * From faVouchers Where tnyStatus<>4 And numSubLedgerID=" & mDemand
''''''                objDB.SetConnection mCnn
''''''                    Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
''''''                    If Not (Rec.EOF And Rec.BOF) Then
''''''                        MsgBox "Already taken Receipt For the Demand No: " & mDemand & ""
''''''                        Call FormInitialize
''''''                        Exit Sub
''''''                    End If
'''''            End If
'''''            mArrIN = Array(mDemand, gbLocalBodyID)
'''''            mArrOutDemand = (objSOAP.GetLicenceDemandTBL(mDemand, gbLocalBodyID))
'''''            If mArrOutDemand <> "" Then
'''''                Set Rec = xmlTORecordset(mArrOutDemand)
'''''                If Not Rec.EOF Then
'''''                    If Not (Rec.EOF And Rec.BOF) Then
'''''                        If Rec!Status <> 2 Then
'''''                            txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
'''''                            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
'''''                            txtDoorNo2.Text = IIf(IsNull(Rec!chvDoorNo), "", Rec!chvDoorNo)
'''''                            txtName.Text = IIf(IsNull(Rec!chvOwnerName), "", Rec!chvOwnerName)
'''''                            txtBuildingNo.Text = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
'''''                            mSubLedgerId = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
'''''                            txtHouse.Text = IIf(IsNull(Rec!chvInstName), "", Rec!chvInstName)
'''''                            txtStreet.Text = IIf(IsNull(Rec!chvBuildingName), "", Rec!chvBuildingName)
'''''                            txtLocalPlace.Text = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
'''''                            txtMainPlace.Text = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
'''''                            txtPost.Text = IIf(IsNull(Rec!chvPo), "", Rec!chvPo)
'''''                            txtPin.Text = IIf(IsNull(Rec!chvPincode), "", Rec!chvPincode)
'''''                            txtPhone.Text = IIf(IsNull(Rec!chvMobileNo), "", Rec!chvMobileNo)
'''''                            mArrOutChild = (objSOAP.GetLicenceDemandChild(mDemand, gbLocalBodyID))
'''''                            Set mRec = xmlTORecordset(mArrOutChild)
'''''                            Call FillGrid(mRec)
'''''                        Else
'''''                            MsgBox "This Demand is Cancelled/Expired"
'''''                            Exit Sub
'''''                        End If
'''''                    End If
'''''                Else
'''''                    MsgBox "No Such Demand Exists", vbApplicationModal
'''''                    Call FormInitialize
'''''                    Exit Sub
'''''                End If
'''''            Else
'''''                MsgBox "Demand Does not Exists", vbApplicationModal
'''''                Call FormInitialize
'''''                Exit Sub
'''''            End If
'''''        End If
'''''    End Sub

Private Sub txtDemandPrefix_LostFocus()
        ''Integration with DO And PFA Using Service
        ''Created On 16-Feb-10 By Anisha
        ''Modified by Syalima On Jan 2018 to implement web Integration
        
        Dim client              As New MSSOAPLib.SoapClient
        Dim objdb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New Recordset
        Dim mRec                As New Recordset
        Dim objSOAP             As Variant
        Dim mUrl                As String
        Dim mArrOutChild        As String
        Dim mArrOutDemand       As String
        Dim mDemand             As Variant
        Dim mArrIn              As Variant
        Dim mSql                As String
        Dim mxmlvalue           As Variant
        Dim mStatus             As String
        Dim mStatCh             As String
        Dim mLicenceDemandChild As Variant
        Dim mLoop               As Integer

        If chkLinkDemand.Value = vbChecked Then
            If gbLinkWithDandOWeb = 1 Then
                    Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
        
                    mUrl = gbDefaultUrl
                    On Error Resume Next
                    objSOAP.MSSoapInit (mUrl + "?WSDL")
                    If err.Number = -2147352567 Then
                        MsgBox "Please Uncheck And Enter The Demand No"
                        Exit Sub
                    End If
                    On Error GoTo 0
                    If txtDemandPrefix.Text = "" Then
                        MsgBox "Please Enter Demand Number"
        '                txtDemandPrefix.SetFocus
                        Exit Sub
                    Else
                        mDemand = Trim(txtDemandPrefix.Text)
                    End If
                    mArrIn = Array(mDemand, gbLocalBodyID)
                    Dim mCredencial As String
                    mCredencial = "ikm@revenue@sanchaya"
                    mArrOutDemand = (objSOAP.getlicencedemand(mDemand, gbLocalBodyID, "ikm@revenue@sanchaya"))
                    If mArrOutDemand = "" Then
                        mStatus = ""
                    ElseIf mArrOutDemand = "null" Then
                        mStatus = ""
                    Else
                        mStatus = 1
                    End If
                If mStatus <> "" Then
                    
                    mxmlvalue = convertJsonToVariantArray(mArrOutDemand)
                    mLoop = 0
                   ' For mLoop = 0 To UBound(mxmlvalue)
                   
                        txtWardNo.Text = IIf(IsNull(mxmlvalue(mLoop, 2)), "", mxmlvalue(mLoop, 2))
                        txtDoorNo1.Text = IIf(mxmlvalue(mLoop, 3) = "null", "", mxmlvalue(mLoop, 3))
                        txtDoorNo2.Text = IIf(mxmlvalue(mLoop, 4) = "null", "", mxmlvalue(mLoop, 4))
                        txtName.Text = IIf(mxmlvalue(mLoop, 5) = "null", "", mxmlvalue(mLoop, 5))
                        'txtBuildingNo.Text=mxmlvalue()
                        txtHouse.Text = IIf(mxmlvalue(mLoop, 7) = "null", "", mxmlvalue(mLoop, 7))
                        'txtStreet.Text=mxmlvalue()
                        txtLocalPlace.Text = IIf(mxmlvalue(mLoop, 9) = "null", "", mxmlvalue(mLoop, 9))
                        txtMainPlace.Text = IIf(mxmlvalue(mLoop, 10) = "null", "", mxmlvalue(mLoop, 10))
                        txtPost.Text = IIf(mxmlvalue(mLoop, 11) = "null", "", mxmlvalue(mLoop, 11))
                        txtPin.Text = IIf(mxmlvalue(mLoop, 12) = "null", "", mxmlvalue(mLoop, 12))
                        txtPhone.Text = IIf(mxmlvalue(mLoop, 13) = "null", "", mxmlvalue(mLoop, 13))
                        mSubLedgerID = IIf(mxmlvalue(mLoop, 6) = "null", "", mxmlvalue(mLoop, 6))
                  'Next
                        mArrOutChild = (objSOAP.getlicencedemandchild(mDemand, gbLocalBodyID, "ikm@revenue@sanchaya", gbLBType))
                        If mArrOutChild = "" Then
                            mStatCh = ""
                        ElseIf mArrOutChild = "null" Then
                            mStatCh = ""
                        Else
                            mStatCh = 1
                        End If
                        If mStatCh = 1 Then
                        
                            mLicenceDemandChild = convertJsonToVariantArray(mArrOutChild)
                            vsGrid.Rows = 2
                            'mLoopgrid = 1
                            For mLoop = 0 To UBound(mLicenceDemandChild)
                                vsGrid.TextMatrix(mLoop + 1, 0) = mLicenceDemandChild(mLoop, 3)
                    '           vsGrid.TextMatrix(mLoop, 1) = objAcc.AccountHead
                                vsGrid.TextMatrix(mLoop + 1, 2) = mLicenceDemandChild(mLoop, 4)
                                vsGrid.TextMatrix(mLoop + 1, 3) = mLicenceDemandChild(mLoop, 5)
                                vsGrid.TextMatrix(mLoop + 1, 4) = ""
                                vsGrid.TextMatrix(mLoop + 1, 5) = mLicenceDemandChild(mLoop, 6)
                                'vsGrid.TextMatrix(mLoop, 6) = objAcc.AccountHeadID
                                'vsGrid.TextMatrix(mLoop, 7) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                                'vsGrid.TextMatrix(mLoop, 8) = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                                vsGrid.TextMatrix(mLoop + 1, 9) = ""
                                vsGrid.TextMatrix(mLoop + 1, 10) = mLicenceDemandChild(mLoop, 7)
                                vsGrid.TextMatrix(mLoop + 1, 11) = mLicenceDemandChild(mLoop, 6)
                                vsGrid.TextMatrix(mLoop + 1, 12) = ""
                                vsGrid.TextMatrix(mLoop + 1, 13) = ""
                                vsGrid.TextMatrix(mLoop + 1, 14) = ""
                                vsGrid.TextMatrix(mLoop + 1, 15) = ""
                                vsGrid.TextMatrix(mLoop + 1, 16) = ""
                                vsGrid.Rows = vsGrid.Rows + 1
                                Call Calculate
                            'mLoop = mLoop + 1
                            Next
                            vsGrid.Editable = flexEDNone
                        Else
                            MsgBox "Demand Child details Does not Exists", vbApplicationModal
                            Exit Sub
                        End If
                Else
                    MsgBox "Demand Does not Exists", vbApplicationModal
                    
                    Call FormInitialize
                    Exit Sub
                End If
            Else
                MsgBox "Integration of D and O is not Activated"
                Exit Sub
            End If
        End If
    End Sub
    Private Function convertJsonToVariantArray(ByVal jsonString As String) As Variant()
        Dim cleanedUpArray() As Variant
        Dim brokenUpRows As Variant
        'Dim brokenUpRows  As String
        Dim mLoop1 As Integer
        If jsonString <> "" Then
            'Remove the first and last square bracket in the string
            jsonString = Right$(jsonString, Len(jsonString) - 2)
            jsonString = Left$(jsonString, Len(jsonString) - 2)
            
            'Break up the string in an array
            brokenUpRows = Split(jsonString, "},{")
            Dim Counter As Integer
            Counter = 0
            Dim counter2 As Long
            Dim brokenUpCols As Variant
            Dim strDomain As Variant
            Dim arr As Variant
            Dim str(10, 20) As Variant
            Dim counter3 As Integer
            'Dim Cnt As Integer
            'Dim counter As Integer
            
            Counter = 0
            counter3 = 0
            'Dim counter2 As Integer
            'Dim brokenUpCols As Variant
            
            ReDim linkArray(UBound(brokenUpRows)) As String
            
            For Counter = 0 To UBound(brokenUpRows)
                brokenUpCols = Split(brokenUpRows(Counter), ",")
                If Counter = 0 Then
                    ReDim cleanedUpArray(UBound(brokenUpRows), UBound(brokenUpCols)) As Variant
                End If
                For counter2 = 0 To UBound(brokenUpCols)
                    cleanedUpArray(Counter, counter2) = brokenUpCols(counter2)
                    'syalima
                    arr = Split(cleanedUpArray(Counter, counter2), ":")
    
                    If UBound(arr) > 0 Then
                        str(Counter, counter2) = Trim(Replace(arr(1), """", " "))
                    End If
                    'syalima
                Next
            Next
            convertJsonToVariantArray = str
        End If
'        If UBound(brokenUpRows) >= 0 Then
'            ReDim brokenUpCols(UBound(brokenUpRows)) As String
'            For counter = 0 To UBound(brokenUpRows)
'                brokenUpCols(counter) = Split(brokenUpRows(0), ",")
'                If counter >= 0 Then
'                    ReDim cleanedUpArray(UBound(brokenUpRows), UBound(brokenUpCols)) As Variant
'                End If
'                For counter2 = 0 To UBound(brokenUpCols)
'                    cleanedUpArray(counter, counter2) = brokenUpCols(counter, counter2)
'                    arr(counter2) = Split(cleanedUpArray(counter, counter2), ":")
'
'                    If UBound(arr) > 0 Then
'                        str(counter, counter2) = Trim(Replace(arr(1), """", " "))
'                    End If
'                Next
'            Next
'            convertJsonToVariantArray = str
'             Else
'            ' MsgBox "Demand Does not Exists", vbApplicationModal
'            Call FormInitialize
'        End If
    End Function

'    Private Sub txtDemandPrefix_LostFocus()
'            ''Integration with DO And PFA Using new Web Service
'            ''Modified On  3 feb 2017 16-Feb-10 By Anisha
'
'            Dim client          As New MSSOAPLib.SoapClient
'            Dim objDB           As New clsDB
'            Dim mCnn            As New ADODB.Connection
'            Dim Rec             As New Recordset
'            Dim mRec            As New Recordset
'            Dim objSOAP         As Variant
'            Dim mUrl            As String
'            Dim mArrOutChild    As String
'            Dim mArrOutDemand   As String
'            Dim mDemand         As Variant
'            Dim mArrIN          As Variant
'            Dim mSQL            As String
'            Dim mXmlStream      As New ADODB.Stream
'            If chkLinkDemand.value = vbChecked Then
'
'             If txtDemandPrefix.Text = "" Then
'                     MsgBox "Please Enter Demand Number"
'     '                txtDemandPrefix.SetFocus
'                     Exit Sub
'                 Else
'                 mDemand = Trim(txtDemandPrefix.Text)
'
'             End If
'
'            mUrl = gbDefaultUrl
'
'            Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
'            objSOAP.MSSoapInit mUrl + "?WSDL"
'
'
'            On Error GoTo WebConnectionERROR:
'            mDemand = CStr(IIf(IsNull(val(txtDemandPrefix.Text)), 0, val(txtDemandPrefix.Text)))
'            'mDemand = Trim(txtDemandPrefix.Text)
'                If mDemand <> "0" Then
'                    mArrOutDemand = objSOAP.GetLicenceDemand(mDemand, "208")
'                End If
'
'            On Error GoTo ERROR_AfterWEBService:
'                mXmlStream.Open
'                mXmlStream.WriteText mArrOutDemand
'                mXmlStream.Position = 0
'                Rec.Open mXmlStream
'                mXmlStream.Close
'                If mArrOutDemand <> "" Then
'                If Not (Rec.BOF And Rec.EOF) Then
'                    If Rec!Status <> 2 Then
'                        txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
'                        txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
'                        txtDoorNo2.Text = IIf(IsNull(Rec!chvdoorno), "", Rec!chvdoorno)
'                        txtName.Text = IIf(IsNull(Rec!chvOwnerName), "", Rec!chvOwnerName)
'                        txtBuildingNo.Text = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
'                        mSubLedgerID = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
'                        txtHouse.Text = IIf(IsNull(Rec!chvInstName), "", Rec!chvInstName)
'                        txtStreet.Text = IIf(IsNull(Rec!chvbuildingname), "", Rec!chvbuildingname)
'                        txtLocalPlace.Text = IIf(IsNull(Rec!chvLocalPlace), "", Rec!chvLocalPlace)
'                        txtMainPlace.Text = IIf(IsNull(Rec!chvMainPlace), "", Rec!chvMainPlace)
'                        txtPost.Text = IIf(IsNull(Rec!chvpo), "", Rec!chvpo)
'                        txtPin.Text = IIf(IsNull(Rec!chvpincode), "", Rec!chvpincode)
'                        txtPhone.Text = IIf(IsNull(Rec!chvmobileno), "", Rec!chvmobileno)
'                        mArrOutChild = objSOAP.GetLicenceDemandChild(mDemand, 208)
'                        On Error GoTo ERROR_AfterWEBService:
'                        mXmlStream.Open
'                        mXmlStream.WriteText mArrOutChild
'                        mXmlStream.Position = 0
'                        mRec.Open mXmlStream
'                        mXmlStream.Close
'
'                        Call FillGrid(mRec)
'                        Exit Sub
'                    Else
'                        MsgBox "This Demand is Cancelled/Expired"
'                        Exit Sub
'                    End If
'                Else
'                        MsgBox "No Such Demand Exists", vbApplicationModal
'                        Call FormInitialize
'                        Exit Sub
'                    End If
'                End If
'            Else
'                MsgBox "Demand Does not Exists", vbApplicationModal
'                Call FormInitialize
'                Exit Sub
'            End If
'
'WebConnectionERROR:
'        MsgBox "Connection to Web Service Failed :: " & Error, vbInformation
'        Exit Sub
'ERROR_AfterWEBService:
'        MsgBox Error
'    End Sub
    Private Sub FillGrid(Rec As ADODB.Recordset)
        Dim mLoop           As Integer
        Dim objAcc           As New clsAccounts
        vsGrid.Rows = 2
        mLoop = 1
        While Not (Rec.EOF And Rec.EOF)
            objAcc.SetAccountCode (Rec!vchAccountHeadCode)
            vsGrid.TextMatrix(mLoop, 0) = objAcc.AccountCode
            vsGrid.TextMatrix(mLoop, 1) = objAcc.AccountHead
            vsGrid.TextMatrix(mLoop, 2) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
            vsGrid.TextMatrix(mLoop, 3) = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
            vsGrid.TextMatrix(mLoop, 4) = ""
            vsGrid.TextMatrix(mLoop, 5) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
            vsGrid.TextMatrix(mLoop, 6) = objAcc.AccountHeadID
            vsGrid.TextMatrix(mLoop, 7) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
            vsGrid.TextMatrix(mLoop, 8) = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
            vsGrid.TextMatrix(mLoop, 9) = ""
            vsGrid.TextMatrix(mLoop, 10) = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
            vsGrid.TextMatrix(mLoop, 11) = IIf(IsNull(Rec!Amount), "", Rec!Amount)
            vsGrid.TextMatrix(mLoop, 12) = ""
            vsGrid.TextMatrix(mLoop, 13) = ""
            vsGrid.TextMatrix(mLoop, 14) = ""
            vsGrid.TextMatrix(mLoop, 15) = ""
            vsGrid.TextMatrix(mLoop, 16) = ""
            vsGrid.Rows = vsGrid.Rows + 1
            Call Calculate
            mLoop = mLoop + 1
            Rec.MoveNext
        Wend
        vsGrid.Editable = flexEDNone
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
        Dim objAc As New clsAccounts
        '------------------------------------------------------------------------------'
        ' When Instrument Cheque is selected then it Set's the Default Bank            '
        '------------------------------------------------------------------------------'
        'If Val(txtInstrument.Tag) = gbInstrumentCheque Or Val(txtInstrument.Tag) = 8 Then '8=Bank Pay_in_Slip
        If val(txtInstrument.Tag) <> gbInstrumentCash Then
            If cmdSearchAccountHead.Tag = 2 Then
                If val(txtAccountHead.Tag) > 0 Then
                    objBk.SetBankInfoByAccID (val(txtAccountHead.Tag))
                    If objBk.BankID > 0 Then
                        txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                    Else
                        GoTo DefaultBank:
                    End If
                Else
DefaultBank:
                    objBk.SetBankInfoByAccID (mDefaultBankID)
                    txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                    txtAccountHead.Tag = objBk.BankAccountHeadID
                End If
            Else
                objBk.SetBankInfoByAccID (mDefaultBankID)
                txtAccountHead.Text = objBk.BankName & " [ " & objBk.BankAccountHeadCode & " ]"
                txtAccountHead.Tag = objBk.BankAccountHeadID
            End If
            cmdSearchAccountHead.Tag = 2
        Else
            If val(cmdSearchAccountHead.Tag) <> 1 Then
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
                
                'Added by Aiby on 11-Nov-2011
                txtInstNo.Text = ""
                txtDated.Text = ""
                txtBank.Text = ""
                txtPlace.Text = ""
                txtInstNo.Enabled = False
                txtDated.Enabled = False
                txtBank.Enabled = False
                txtPlace.Enabled = False
                
            End If
        End If
        
                
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim mStr As String
        If mInterruptedModeFlag = False Then '' Added By Anisha
            If mPreviousYearMode = 1 Then
                arrInput = Array(gbCounterID, val(txtInstrument.Tag), gbFinancialYearID - 1)
            Else
                arrInput = Array(gbCounterID, val(txtInstrument.Tag), gbFinancialYearID)
            End If
            Set Rec = objdb.ExecuteSP("spGetNextReceiptNo", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
            If IsArray(arrOutPut) Then
                txtReceiptNo.Text = arrOutPut(0, 0)
            End If
            Rec.Close
            If Len(txtReceiptNo) > 6 Then
                mStr = Left(txtReceiptNo, 6)
                mStr = mStr + "-" + Right(txtReceiptNo, Len(txtReceiptNo) - 6)
                txtReceiptNo.Text = mStr
            End If
        End If
    End Sub
    Private Sub txtIntruptNoSuffix_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call PressTabKey
            Exit Sub
        End If
        If (KeyAscii > Asc("A") And KeyAscii <= Asc("Z")) Or KeyAscii = 8 Or (KeyAscii > Asc("a") And KeyAscii <= Asc("z")) Then
        ''''' Validating all aphabets and Back space Exclude A
            If CheckInterruptedNoSuffixExists(Chr(KeyAscii)) Then
                MsgBox "Alredy Exists", vbInformation
                KeyAscii = 0
            End If
        Else
            MsgBox "Please Enter Single Alphabets .. Exclude 'A'"
            txtIntruptNoSuffix.Text = ""
            KeyAscii = 0
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
    
    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPin_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
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
    
Private Sub txtTransactionType_Change()
    Dim mIndex As Long
    Dim mStr
    If Not mSkipFlag Then
        If mKeyCode = 8 Or mKeyCode = 46 Or txtTransactionType.Text = "" Then 'Tab or delete
            
        End If
        If Not (mBkSpaceFlag Or mKeyCode = 40 Or mKeyCode = 38) Then
            
            With lstTransactionType
                mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal txtTransactionType.Text)
                If mIndex >= 0 Then
                    .ListIndex = mIndex
                End If
            End With
            
            If mIndex >= 0 Then
                mStr = txtTransactionType.Text
                txtTransactionType.Text = lstTransactionType.List(mIndex)
                txtTransactionType.SelStart = Len(mStr)
                
                If Len(txtTransactionType.Text) - Len(mStr) > 0 Then
                    txtTransactionType.SelLength = Len(txtTransactionType.Text) - Len(mStr)
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
        If gbSearchStr <> "" Then
            If val(txtTransactionType.Tag) <> gbSearchID Then
                If mWebExtractMode = False Then
                    vsGrid.Rows = 1
                    vsGrid.Rows = 15
                    Call Calculate
                End If
            End If
            txtTransactionType.Text = gbSearchStr
            txtTransactionType.Tag = gbSearchID
            gbSearchCode = ""
            gbSearchID = -1
            gbSearchStr = ""
            Call txtTransactionType_KeyPress(13)
        End If
         
        txtTransactionType.SelStart = 0
        txtTransactionType.SelLength = Len(txtTransactionType)
        If Trim(txtTransactionType.Text) = "" Then
            'ListMasters (1)
            'lstMasters.Refresh
        End If
    End Sub
        
    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then ' vbKeyDelete Then
            txtTransactionType.Text = ""
            lstTransactionType.ListIndex = -1
        End If
        If KeyCode = 40 Then 'Down Arrow
            If lstTransactionType.ListIndex > -1 Then
                lstTransactionType.ListIndex = (lstTransactionType.ListIndex + 1) Mod lstTransactionType.ListCount
                txtTransactionType.Text = lstTransactionType.Text
            End If
        ElseIf KeyCode = 38 Then 'Uparrow
            If lstTransactionType.ListIndex > -1 Then
                If lstTransactionType.ListIndex = 0 Then
                    lstTransactionType.ListIndex = lstTransactionType.ListCount - 1
                    txtTransactionType.Text = lstTransactionType.Text
                Else
                    lstTransactionType.ListIndex = (lstTransactionType.ListIndex - 1) Mod lstTransactionType.ListCount
                    txtTransactionType.Text = lstTransactionType.Text
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
          Call PressTabKey
        End If
    End Sub
'
'    Private Sub txtTransactionType_LostFocus()
'        Dim mIndex As Long
'        Dim objAcc As New clsAccounts
'        '----------- Block -------------------- '
'        'Note:- Modified on 31-Mar-2010 By Aiby
'        If cmdSave.Enabled = False Then
'            Exit Sub
'        End If
'        '---- End of Block -------------------- '
'
'
'        chkLinkDemand.Visible = False
'        chkLinkDemand.value = 0
'        txtDemandPrefix.Width = 1080
'
'lblStart:
'        With lstTransactionType
'            mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal txtTransactionType.Text)
'            If mIndex >= 0 Then
'                txtTransactionType.Tag = lstTransactionType.ItemData(mIndex)
'            Else
'                txtTransactionType.Text = "Other Receipts"
'                txtTransactionType.Tag = ""
'                GoTo lblStart:
'            End If
'        End With
'
'        mvarDemandBasedFlag = False
'        txtDemandPrefix.Text = ""
'
'        Select Case val(txtTransactionType.Tag)
'            Case Is = gbTransactionTypePTax
'                'On Error Resume Next
'                If gbLinkWithPropertyTax Then
'                    On Error Resume Next
'                    If mPTaxFormLoadFlag = False Then
'                        If Not mInterruptEditMode Then
'                            Call FormInitialize
'                        End If
'                        txtTransactionType.Tag = gbTransactionTypePTax
'                        txtTransactionType.Text = "Property Tax"
'                        frmPropertyTax.Show vbModal
'                    End If
'                End If
'            Case Is = gbTransactionTypeProfTaxTrade
'                If gbLinkWithProfTaxEmp = 1 Then
'                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 1
'                    frmSearchProfTaxInstitutions.Show vbModal
'                End If
'            Case Is = gbTransactionTypeRentOnBuilding
'                If gbLinkWithRentOnLand Then
'                    If mRentSearchFormLoadedFlag = False Then
'                        Call FormInitialize
'                        mRentSearchFormLoadedFlag = True
'                        txtTransactionType.Tag = gbTransactionTypeRentOnBuilding
'                        txtTransactionType.Text = "Rent on Building / Stalls"
'                        frmRentOnLandBuildings.mCategory = 1
''                        frmRentOnLandBuildings.cmbCategory.Tag = 1
'
''                        frmRentOnLandBuildings.cmbCategory.Text = "Building"
''                        frmRentOnLandBuildings.cmbCategory.Enabled = False
'                        frmRentOnLandBuildings.Show vbModal
'                    End If
'                End If
'            Case Is = gbTransactionTypeRentOnLand
'                If gbLinkWithRentOnLand Then
'                    Call FormInitialize
'                    txtTransactionType.Tag = gbTransactionTypeRentOnLand
'                    txtTransactionType.Text = "Rent on Land/Bunks"
'                    frmRentOnLandBuildings.mCategory = 2
''                    frmRentOnLandBuildings.cmbCategory.Tag = 2
''                    frmRentOnLandBuildings.cmbCategory.Text = "Land"
'
'                    frmRentOnLandBuildings.Show vbModal
'                End If
'            Case Is = gbTransactionTypeBrith, gbTransactionTypeDeath
'                '-------------------------------------------------------'
'                '   Added for checking the Birth/Death Intagration      '
'                '-------------------------------------------------------'
'                Dim mSQL As String
'                Dim Rec As New ADODB.Recordset
'                Dim mCnn As New ADODB.Connection
'                Dim objDb As New clsDB
'                    objDb.SetConnection mCnn
'                mSQL = "Select tnyLinkWithBandDSchedules from faConfig"
'                Rec.Open mSQL, mCnn
'                If Rec!tnyLinkWithBandDSchedules = 1 Then
'                    frmScheduleRatesForBirthDeath.mReceiptOrDemandFlag = 1
'                    frmScheduleRatesForBirthDeath.Visible = True
'                    frmScheduleRatesForBirthDeath.ZOrder (0)
'                End If
'                If Rec.State = 1 Then Rec.Close
'                If mCnn.State = 1 Then mCnn.Close
'                '-------------------------------------------------------'
'                '                frmScheduleRatesForBirthDeath.mReceiptOrDemandFlag = 1
'                '                frmScheduleRatesForBirthDeath.Visible = True
'                '                frmScheduleRatesForBirthDeath.ZOrder (0)
'            Case Is = gbTransactionTypeZonalCollection
'                lblOutDoorStaff(8).Caption = "Zone"
'                lblOutDoorStaff(8).Visible = True
'                txtOutDoorStaff.Text = ""
'                txtOutDoorStaff.Tag = ""
'            Case Is = gbTransactionTypeOutDoor
'                lblOutDoorStaff(8).Caption = "Out Door Staff"
'                lblOutDoorStaff(8).Visible = True
'                txtOutDoorStaff.Text = ""
'                txtOutDoorStaff.Tag = ""
'            Case Is = gbTransactionTypeApplicationForPermitKMBR
'                '------------------------------------------------------------'
'                '       Added For KMBR Integration ON 24/04/2009             '
'                '------------------------------------------------------------'
'                objDb.SetConnection mCnn
'                mSQL = "Select tnyLinkWithKMBR from faConfig"
'                Rec.Open mSQL, mCnn
'                If Rec!tnyLinkWithKMBR = 1 Then
'                    frmKMBRIntegration.Show 1
'                    If mKMBRAccess = 1 Then
'                        objAcc.SetAccountCode ("140409900")
'                        vsGrid.TextMatrix(1, 0) = objAcc.AccountCode
'                        vsGrid.TextMatrix(1, 1) = objAcc.AccountHead
'                        vsGrid.TextMatrix(1, 6) = objAcc.AccountHeadID
'                        vsGrid.TextMatrix(1, 5) = 75
'                        vsGrid.Row = 1
'                        Call ValuesForHiddenColumns
'                        Call Calculate
'                    End If
'                    chkLinkDemand.Visible = True
'                    txtDemandPrefix.Width = 2280
'                End If
'            Case Is = gbTransactionTypePermitFeeFromKMBR
'                '------------------------------------------------------------'
'                '       Added For KMBR Integration ON 24/04/2009             '
'                '------------------------------------------------------------'
'                objDb.SetConnection mCnn
'                mSQL = "Select tnyLinkWithKMBR from faConfig"
'                Rec.Open mSQL, mCnn
'                If Rec!tnyLinkWithKMBR = 1 Then
'                    vsGrid.Enabled = False
'                End If
'            Case Is = gbTransactionTypeSaleOfTenderForm
'                '------------------------------------------------------------'
'                '       Connecting to Sugama                                 '
'                '------------------------------------------------------------'
'                If gbLinkWithSugama Then
'                    frmSugSaleofTender.Show vbModal
'                End If
'            Case Is = gbTransactionTypeDandO, gbTransactionTypePFA
'                '--------------------------------------------------------------------------------'
'                ' Interface Option to Link with External Demand Database Or Not                  '
'                '--------------------------------------------------------------------------------'
'                If gbLinkWithDandOPFA = 1 Then
'                    chkLinkDemand.Visible = True
'                    chkLinkDemand.value = 1
'                    txtDemandPrefix.Width = 2280
'                End If
'            Case Is = 9999
'                'Call FillAccountHeads
'                'Call FillGridYear
'            Case Else
'                txtDemandPrefix.Text = "1" & Format(gbFinancialYearID - 2000, "00") & Right(Format(val(txtTransactionType.Tag), "0000"), 3)
'        End Select
'
'        Call FillGridYear
'    End Sub
        
    Private Sub txtTransactionType_LostFocus()
     If mZonal <> 1 Then 'Added Sunil for Zonal Collection
        Dim mIndex As Long
        Dim objAcc As New clsAccounts
        '----------- Block -------------------- '
        'Note:- Modified on 31-Mar-2010 By Aiby
        If cmdSave.Enabled = False Then
            Exit Sub
        End If
        '---- End of Block -------------------- '
        
        
        chkLinkDemand.Visible = False
        chkLinkDemand.Value = 0
        txtDemandPrefix.Width = 1080

lblStart:
        With lstTransactionType
            mIndex = SendMessage(.hwnd, LB_FINDSTRING, -1, ByVal txtTransactionType.Text)
            If mIndex >= 0 Then
                txtTransactionType.Tag = lstTransactionType.ItemData(mIndex)
            Else
                If mWebExtractMode = True Then
                    txtTransactionType.Text = "Project Expenditure -General- Capital "

                    txtTransactionType.Tag = 1141
                Else
                    txtTransactionType.Text = "Other Receipts"
                    txtTransactionType.Tag = 9999
                End If
                GoTo lblStart:
                
            End If
        End With
        
        mvarDemandBasedFlag = False
        txtDemandPrefix.Text = ""
        
        Select Case val(txtTransactionType.Tag)
            Case Is = gbTransactionTypePTax
                'On Error Resume Next
                If gbLinkWithPropertyTax Then
                    On Error Resume Next
                    If mPTaxFormLoadFlag = False Then
                        If Not mInterruptEditMode Then
                            Call FormInitialize
                        End If
                        txtTransactionType.Tag = gbTransactionTypePTax
                        txtTransactionType.Text = "Property Tax"
                        frmPropertyTax.Show vbModal
                        If cmdSave.Enabled Then
                            cmdSave.SetFocus
                        End If
                    End If
                ElseIf gbFetchDemandFromWeb = 1 Then
                    On Error Resume Next
                    If mPTaxFormLoadFlag = False Then
                        If Not mInterruptEditMode Then
                            Call FormInitialize
                        End If
                        txtTransactionType.Tag = gbTransactionTypePTax
                        txtTransactionType.Text = "Property Tax"
                        frmPropertyTax.mDemandWeb = True
                        frmPropertyTax.Show vbModal
                        If cmdSave.Enabled Then
                            cmdSave.SetFocus
                        End If
                    End If
                End If
            Case Is = gbTransactionTypeProfTaxTrade
                If gbLinkWithProfTaxEmp = 1 Then
                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 1
                    frmSearchProfTaxInstitutions.Show vbModal
                    mSubLedgerID = txtHouse.Tag
                
                End If
                If gbLinkWithProfTradeWeb = 1 Then
                    frmPofessionTaxTrades.mPTType = 1
                    frmPofessionTaxTrades.Show vbModal
                    
                End If
            Case Is = gbTransactionTypeProfTaxEmp
                If gbLinkWithProfEmpWeb = 1 Then
                    frmPofessionTaxTrades.mPTType = 2
                    frmPofessionTaxTrades.Show vbModal
                End If
            Case Is = gbTransactionTypeRentOnBuilding
                If gbLinkWithRentOnLand Then
                    Call FormInitialize
                    If mRentSearchFormLoadedFlag = False Then
                        mRentSearchFormLoadedFlag = True
                        txtTransactionType.Tag = gbTransactionTypeRentOnBuilding
                        txtTransactionType.Text = "Rent on Building / Stalls"
                        frmRentOnLandBuildings.mCategory = 1
'                        frmRentOnLandBuildings.cmbCategory.Tag = 1
                        
'                        frmRentOnLandBuildings.cmbCategory.Text = "Building"
'                        frmRentOnLandBuildings.cmbCategory.Enabled = False
                        frmRentOnLandBuildings.Show vbModal
                    End If
                End If
            Case Is = gbTransactionTypeRentOnLand
                If gbLinkWithRentOnLand Then
                    If mRentSearchFormLoadedFlag = False Then
                        mRentSearchFormLoadedFlag = True
                        Call FormInitialize
                        txtTransactionType.Tag = gbTransactionTypeRentOnLand
                        txtTransactionType.Text = "Rent on Land/Bunks"
                        frmRentOnLandBuildings.mCategory = 2
    '                    frmRentOnLandBuildings.cmbCategory.Tag = 2
    '                    frmRentOnLandBuildings.cmbCategory.Text = "Land"
                        frmRentOnLandBuildings.Show vbModal
                    End If
                End If
            Case Is = gbTransactionTypeBrith, gbTransactionTypeDeath
                '-------------------------------------------------------'
                '   Added for checking the Birth/Death Intagration      '
                '-------------------------------------------------------'
                Dim mSql As String
                Dim Rec As New ADODB.Recordset
                Dim mCnn As New ADODB.Connection
                Dim objdb As New clsDB
                    objdb.SetConnection mCnn
                mSql = "Select tnyLinkWithBandDSchedules from faConfig"
                Rec.Open mSql, mCnn
                If Rec!tnyLinkWithBandDSchedules = 1 Then
                    frmScheduleRatesForBirthDeath.mReceiptOrDemandFlag = 1
                    frmScheduleRatesForBirthDeath.Visible = True
                    frmScheduleRatesForBirthDeath.ZOrder (0)
                End If
                If Rec.State = 1 Then Rec.Close
                If mCnn.State = 1 Then mCnn.Close
                '-------------------------------------------------------'
                '                frmScheduleRatesForBirthDeath.mReceiptOrDemandFlag = 1
                '                frmScheduleRatesForBirthDeath.Visible = True
                '                frmScheduleRatesForBirthDeath.ZOrder (0)
            Case Is = gbTransactionTypeZonalCollection
                lblOutDoorStaff(8).Caption = "Zone"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Text = ""
                txtOutDoorStaff.Tag = ""
            Case Is = gbTransactionTypeOutDoor
                lblOutDoorStaff(8).Caption = "Out Door Staff"
                lblOutDoorStaff(8).Visible = True
                txtOutDoorStaff.Text = ""
                txtOutDoorStaff.Tag = ""
            Case Is = gbTransactionTypeApplicationForPermitKMBR
                '------------------------------------------------------------'
                '       Added For KMBR Integration ON 24/04/2009             '
                '------------------------------------------------------------'
                objdb.SetConnection mCnn
                mSql = "Select tnyLinkWithKMBR from faConfig"
                Rec.Open mSql, mCnn
                If Rec!tnyLinkWithKMBR = 1 Then
                    frmKMBRIntegration.Show 1
                    If mKMBRAccess = 1 Then
                        objAcc.SetAccountCode (gbAcHeadCodeOtherFee) '("140409900")
                        vsGrid.TextMatrix(1, 0) = objAcc.AccountCode
                        vsGrid.TextMatrix(1, 1) = objAcc.AccountHead
                        vsGrid.TextMatrix(1, 6) = objAcc.AccountHeadID
                        vsGrid.TextMatrix(1, 5) = 75
                        vsGrid.Row = 1
                        Call ValuesForHiddenColumns(vsGrid.Row)
                        Call Calculate
                    End If
                    chkLinkDemand.Visible = True
                    txtDemandPrefix.Width = 2280
                End If
            Case Is = gbTransactionTypePermitFeeFromKMBR
                '------------------------------------------------------------'
                '       Added For KMBR Integration ON 24/04/2009             '
                '------------------------------------------------------------'
                objdb.SetConnection mCnn
                mSql = "Select tnyLinkWithKMBR from faConfig"
                Rec.Open mSql, mCnn
                If Rec!tnyLinkWithKMBR = 1 Then
                    vsGrid.Enabled = False
                End If
            Case Is = gbTransactionTypeSaleOfTenderForm
                '------------------------------------------------------------'
                '       Connecting to Sugama                                 '
                '------------------------------------------------------------'
                If gbLinkWithSugama Then
                    frmSugSaleofTender.Show vbModal
                End If
            Case Is = gbTransactionTypeDandO, gbTransactionTypePFA
                '--------------------------------------------------------------------------------'
                ' Interface Option to Link with External Demand Database Or Not                  '
                '--------------------------------------------------------------------------------'
                If gbLinkWithDandOPFA = 1 Or gbLinkWithDandOWeb = 1 Then
                    chkLinkDemand.Visible = True
                    chkLinkDemand.Value = 1
                    txtDemandPrefix.Width = 2280
                End If
            Case Is = gbTransactionTypeMOReturnsSocialSecurityPension
                Call FormInitialize
                txtTransactionType.Tag = gbTransactionTypeMOReturnsSocialSecurityPension
                txtTransactionType.Text = "Money Order Returns Social Security Pension"
                If gbLinkWithMOReturn Then
                    frmMOReturned.Show vbModal
                End If
            Case Is = 9999
                'Call FillAccountHeads
                'Call FillGridYear
            Case Else
                txtDemandPrefix.Text = "1" & Format(gbFinancialYearID - 2000, "00") & Right(Format(val(txtTransactionType.Tag), "0000"), 3)
        End Select
        
        Call FillGridYear
    End If
    End Sub

    Private Sub txtWard_LostFocus()
        If val(txtWard) > 0 Then
            txtWard.Tag = val(txtWard)
        Else
            txtWard.Tag = ""
        End If
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
                If vsGrid.TextMatrix(vsGrid.Row - 1, 0) = "" Or _
                   (val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 And _
                   val(vsGrid.TextMatrix(vsGrid.Row - 1, 5)) <= 0) Then
                   Cancel = True
                   Exit Sub
                End If
            End If
            
            If Col = 4 Or Col = 5 Then
                If Trim(vsGrid.TextMatrix(Row, 0)) = "" Then
                    Cancel = True
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
                vsGrid.Col = vsGrid.Col + 2
                vsGrid.Redraw = flexRDDirect
                gbSearchStr = ""
            ElseIf Col = 1 Then
                Cancel = True
            End If
            
    End Sub
    
    Private Sub vsGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
        If NewRow >= vsGrid.Rows - 1 Then ' Edited By Sinoj as old row as newRow
            vsGrid.Rows = vsGrid.Rows + 5
        End If
    End Sub
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If mInterruptedModeFlag = False Then
            If vsGrid.Row > 9 And val(txtTransactionType.Tag) <> gbTransactionTypePTax Then
                MsgBox "Can't print more than 9 rows in this Demand", vbInformation
                Exit Sub
            End If
        End If
        Dim mSql As String
        If txtTransactionType.Text = "Other Receipts" And val(txtTransactionType.Tag) = 0 Then
            txtTransactionType.Tag = 9999
            txtTransactionType.Text = "Other Receipts"
        End If
        If val(txtTransactionType.Tag) > 0 And val(txtTransactionType.Tag) < 9999 Then
            mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
            mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
            mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & " And faAccountHeads.tinHiddenFlag = 0 And faAccountHeads.intGroupID is Null Order By faTransactionTypeChild.intOrder"
            frmSearchAccountHeads.SQLString = mSql '"Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        Else
            If gbLBPanchayat = 1 Then
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null and intMinorAccountHeadID<>220 Order By faAccountHeads.vchAccountHeadCode"
            
            Else
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null and intMinorAccountHeadID<>248 Order By faAccountHeads.vchAccountHeadCode"
            End If
        End If
        
        frmSearchAccountHeads.VoucherMode = 100
        frmSearchAccountHeads.Show vbModal
        
    End Sub
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        Dim mAmt As Double
        'If vsGrid.Row > 0 Then
        If Row > 0 Then
            If Col = 0 Then
                Set objAccHead = New clsAccounts
                objAccHead.SetAccountCode (Trim(vsGrid.TextMatrix(Row, 0)))
                If objAccHead.AccountHeadID > 0 Then
                    vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                    vsGrid.TextMatrix(Row, 1) = objAccHead.AccountHead
                    vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                Else
                    '------------------------------------------'''''Added and edited By Sinoj'''''''
                    If vsGrid.TextMatrix(Row, 1) <> "" Then
                        vsGrid.RemoveItem (Row)
                    End If
'                    vsGrid.TextMatrix(Row, 0) = ""
'                    vsGrid.TextMatrix(Row, 1) = ""
'                    vsGrid.TextMatrix(Row, 6) = ""
'                    vsGrid.TextMatrix(Row, 4) = ""
'                    vsGrid.TextMatrix(Row, 5) = ""
                    Call Calculate
                    '------------------------------------------'''''Added and editted By Sinoj'''''''
                End If
            ElseIf Col = 1 And vsGrid.ComboIndex > -1 Then
                Set objAccHead = New clsAccounts
                If objAccHead.FindAccountByHead(Trim(vsGrid.ComboItem)) Then
                vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
            ElseIf Col = 4 Then
                mAmt = val(vsGrid.TextMatrix(Row, 4))
                If mRoundOffDecimalPlace Then
                    mAmt = Format(mAmt, "#0")
                    vsGrid.TextMatrix(Row, 4) = Format(mAmt, "0.00")
                Else
                    vsGrid.TextMatrix(Row, 4) = Format(val(vsGrid.TextMatrix(Row, 4)), "0.00")
                End If
                
                If (mAmt - Int(mAmt)) > 0 Then
                    mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
                End If
                vsGrid.TextMatrix(Row, 4) = Format(mAmt, "0.00")
                
                If val(vsGrid.TextMatrix(Row, 4)) > 0 Then
                    vsGrid.TextMatrix(Row, 5) = ""
                End If
                Call Calculate
            ElseIf Col = 5 Then
                mAmt = val(vsGrid.TextMatrix(Row, 5))
                If mRoundOffDecimalPlace Then
                    mAmt = Format(mAmt, "#0")
                    vsGrid.TextMatrix(Row, 5) = Format(mAmt, "0.00")
                Else
                    vsGrid.TextMatrix(Row, 5) = Format(val(vsGrid.TextMatrix(Row, 5)), "0.00")
                End If
                
                If (mAmt - Int(mAmt)) > 0 Then
                    mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
                End If
                vsGrid.TextMatrix(Row, 5) = Format(mAmt, "0.00")
                
                
                
                If val(vsGrid.TextMatrix(Row, 5)) > 0 Then
                    vsGrid.TextMatrix(Row, 4) = ""
                End If
                Call Calculate
            End If
            Call ValuesForHiddenColumns(Row)
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
    
    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If mInterruptedModeFlag = False Then
            If vsGrid.Row > 9 And val(txtTransactionType.Tag) <> gbTransactionTypePTax Then
                KeyAscii = 0
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
    Private Sub funUpdate()
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim objFunctions As New clsFunction
        Dim objFunctionaries As New clsFunctionary
        Dim mFunctionaryID  As Variant
        Dim mFunctionId As Variant
        Dim mLoopCount As Long
        Dim mLoop As Long
        Dim Rec As New ADODB.Recordset
        Dim mDemandID As Variant
        Dim objAc As New clsAccounts
        Dim lSoochikaCurrentNo As Variant
        Dim mSql As String
        
        Dim mTransactionDate As Date
        Dim mYearID As Integer
        
        If mPreviousYearMode = 1 Then
            If IsDate(txtDate) Then
                mTransactionDate = CDate(txtDate.Text)
                If Not (mTransactionDate >= DateAdd("yyyy", -1, gbStartingDate) And mTransactionDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
                    MsgBox "Transaction Date found mismatch!", vbInformation
                    Exit Sub
                End If
                mYearID = gbFinancialYearID - 1
            Else
                MsgBox "Transaction Date found mismatch!", vbInformation
                Exit Sub
            End If
        ElseIf mInterruptedModeFlag = True Then
            mYearID = gbFinancialYearID
            mTransactionDate = DdMmmYy(txtDate.Text)
        ElseIf mWebExtractMode = True Then
            mTransactionDate = mWebExtractDate
        Else
            mYearID = gbFinancialYearID
            mTransactionDate = gbTransactionDate
        End If
        
        '''Added Anisha
        If CDate(mTransactionDate) <= GetLastReconDate(val(txtAccountHead.Tag)) Then
            MsgBox "Selected Bank/Treasury of the month is Reconciled, not allowed to do transactions"
            txtAccountHead.Tag = -1
            txtAccountHead.Text = ""
            Exit Sub
        End If
        
        If val(txtReceiptNo.Tag) < 0 Then
            MsgBox "Didn't able to locate this Voucher details for Updation!", vbInformation
            Exit Sub
        End If
        If (val(txtTotal)) <= 0 Then
            MsgBox "Please Check the Amount...!", vbInformation
            Exit Sub
        End If
        
        If Trim(txtName.Text) = "" Then
            MsgBox "Please Enter the name of Person who is remitting the Amount", vbInformation
            txtName.SetFocus
            Exit Sub
        End If
        
        If vsGrid.Rows > 1 Then
            If vsGrid.TextMatrix(1, 0) = "" Or _
               (val(vsGrid.TextMatrix(1, 4)) <= 0 And _
                val(vsGrid.TextMatrix(1, 5)) <= 0) Then
                MsgBox "No Item has been entered in the Grid..!", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "No Item has been entered in the Grid..!", vbInformation
        End If
        
        If val(txtInstrument.Tag) = 5 Then
            If Trim(txtInstNo) = "" Then
                MsgBox "Please Enter the Cheque No.", vbInformation
                txtInstNo.SetFocus
                Exit Sub
            End If
            If Not IsDate(txtDated) Then
                MsgBox "Please Enter the Cheque Date", vbInformation
                txtDated.SetFocus
                Exit Sub
            End If
            If Trim(txtBank) = "" Then
                MsgBox "Please Enter the Name of Bank/Branch, Who issued the cheque...", vbInformation
                txtBank.SetFocus
                Exit Sub
            End If
            If Trim(txtPlace) = "" Then
                MsgBox "Please Enter the place of Bank issued the Cheque..", vbInformation
                txtPlace.SetFocus
                Exit Sub
            End If
        End If
        If txtTransactionType.Tag = -1 Then
            MsgBox "Please Select TransactionType..", vbInformation
            txtTransactionType.SetFocus
            Exit Sub
        Else
            mTransactionType = val(txtTransactionType.Tag)
        End If
        If mTransactionType = gbTransactionTypePTax Then
            If val(txtWardNo) < 1 Then
                MsgBox "Enter the Ward No ", vbInformation
                txtWardNo.SetFocus
                Exit Sub
            End If
            
            If val(txtDoorNo1) < 1 Then
                MsgBox "Enter the valid Door No ", vbInformation
                txtDoorNo1.SetFocus
                Exit Sub
            End If
        ElseIf mTransactionType = gbTransactionTypeOutDoor Then
            mSubLedgerID = val(txtOutDoorStaff.Tag)
        End If
        
        If val(txtDemandNo.Tag) > 0 Then
            mDemandID = txtDemandNo.Tag
        End If
    
        '-------------------------------------------------------------------------------'
        ' S E R V E R   D A T E   V A R I F I C A T I O N                               '
        '-------------------------------------------------------------------------------'
        If mInterruptedModeFlag = False Then
            objdb.SetConnection mCnn
            Set Rec = mCnn.Execute("Select GetDate()")
            If IsDate(Rec.Fields(0)) Then
                mdtDate = DdMmmYy(Rec.Fields(0))
            Else
                MsgBox "Didn't able to Access Server Date", vbInformation
                Exit Sub
            End If
            Rec.Close
            mCnn.Close
        End If
        
        '-------------------------------------------------------------------------------'
        ' C h e c k    I t e m s   i n    G r i d    C o r r e c t l y     F i l l e d  '
        '-------------------------------------------------------------------------------'
        Dim mEmptyRow As Integer
        mEmptyRow = 9999
        For mLoopCount = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then      ' (1)Account Head Code
                If mEmptyRow < mLoopCount Then                        ' (4)Checks any Previos Row is incomplete
                    MsgBox "Row is not completed!", vbInformation
                    vsGrid.Row = mEmptyRow
                    Exit Sub
                End If
            If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 Then ' (2)Valid Column
            If val(vsGrid.Cell(flexcpText, mLoopCount, 11)) > 0 Then  ' (3)Amount>0
                If mEmptyRow < mLoopCount Then                        ' (4)Checks any Previos Row is incomplete
                    MsgBox "Row is not completed!", vbInformation
                    vsGrid.Row = mEmptyRow
                    Exit Sub
                End If
            Else                ' (3)Amount>0
                 mEmptyRow = mLoopCount
            End If              ' (3)Amount>0
            End If              ' (2)Valid Column
            Else                ' (1)Account Head Code
               mEmptyRow = mLoopCount
            End If              ' (1)Account Head Code
        Next mLoopCount
        
        '---------------------------------------------------------------------------------
        'For D&F And PFA
        '---------------------------------------------------------------------------------
        If mTransactionType = gbTransactionTypeDandO Or mTransactionType = gbTransactionTypePFA Then
            If gbLinkWithDandOPFA = 1 Then
                If txtDemandPrefix.Text = "" Then
                    MsgBox "Please Enter Demand No", vbApplicationModal
                End If
            End If
        End If
        '-------------------------------------------------------------------------------'
        ' END OF BLOCK ::                                                               '
        '-------------------------------------------------------------------------------'
        
        '===================================================='
        '       Added On 27/04/2009 By Cijith for KMBR       '
        '----------------------------------------------------'
        '   Checking Whether faConfig File Exists for KMBR   '
        '----------------------------------------------------'
        
        'Note:- Code Review Note by Aiby :: Date 31-Dec-2009
        '       This block of code can be removed from this part.
        '       Purpose this Block with surve is, It sets mKMBRflag
        '
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Or mTransactionType = gbTransactionTypePermitFeeFromKMBR Then
            'objDb.SetConnection mCnn
            'mSql = "Select tnyLinkWithKMBR from faConfig"
            'Rec.Open mSql, mCnn
            'If Rec!tnyLinkWithKMBR = 1 Then                                                              ' mKMBRAccess = Property Variable
            '    If mTransactionType = gbTransactionTypeApplicationForPermitKMBR And mKMBRAccess = 1 Then ' Set from KMBR from
            '        mKMBRFlag = True
            '    ElseIf mTransactionType = gbTransactionTypePermitFeeFromKMBR Then
            '        mKMBRFlag = True
            '    Else
            '        mKMBRFlag = False
            '    End If
            'Else
            '    mKMBRFlag = False
            'End If
            'Rec.Close
            'Set mCnn = Nothing
        End If
        '----------------------------------------------------'
        
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
            If mKMBRFlag = True Then
            If cmbSeat.ListIndex = -1 Then
                MsgBox "Please Give Forwarded To Seat", vbInformation
                cmbSeat.SetFocus
                Exit Sub
            End If
            If txtMainPlace.Text = "" Then
                MsgBox "Please Give the Main Place", vbInformation
                txtMainPlace.SetFocus
                Exit Sub
            End If
            If txtPost.Text = "" Then
                MsgBox "Please Give Post Box", vbInformation
                txtPost.SetFocus
                Exit Sub
            End If
            If txtPin.Text = "" Then
                MsgBox "Please Give Pin Code", vbInformation
                txtPin.SetFocus
                Exit Sub
            End If
            If txtWardNo.Text = "" Then
                MsgBox "Please Ward No", vbInformation
                txtWardNo.SetFocus
                Exit Sub
            End If
            If txtDoorNo1.Text = "" Then
                MsgBox "Please Give Door Number", vbInformation
                txtDoorNo1.SetFocus
                Exit Sub
            End If
            If txtHouse.Text = "" Then
                MsgBox "Please Give House Name", vbInformation
                txtHouse.SetFocus
                Exit Sub
            End If
            
            Dim mCnnSoochika As New ADODB.Connection
            If objdb.CreateNewConnection(mCnnSoochika, enuSourceString.SOOCHIKA) = True Then
                mCnnSoochika.BeginTrans
                lSoochikaCurrentNo = SaveSoochika(mCnnSoochika)
                If lSoochikaCurrentNo = -1 Or lSoochikaCurrentNo = 0 Then GoTo ErrorRollBackSoochika:
            End If
            End If
            'changed by soumya V S on 14.05.14
                    frmUSoochikaInward.SaveAttachment (lSoochikaCurrentNo)
        End If
        '===================================================='
        
        '==================================================================================='
        ' Common Counter - mSoochikaConnected is Set from Sevana Inward
        '-----------------------------------------------------------------------------------'
        If mSoochikaConnected = True Then
            'Dim mCnnSoochika As New ADODB.Connection
            If objdb.CreateNewConnection(mCnnSoochika, enuSourceString.SOOCHIKA) = True Then
                mCnnSoochika.BeginTrans
                lSoochikaCurrentNo = frmSoochikaInward.SaveSoochika(mCnnSoochika)
                If lSoochikaCurrentNo = -1 Or lSoochikaCurrentNo = 0 Then GoTo ErrorRollBackSoochika:
            End If
        End If
        '-----------------------------------------------------------------------------------'
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then ' CREATED NEW CONNECTION
                objFunctionaries.SetFunctionary ("080000")
                mFunctionaryID = objFunctionaries.FunctionaryID
                Select Case mTransactionType
                    Case 2 ' Property Tax
                        objFunctions.SetFunction ("90910000")
                        mFunctionId = objFunctions.FunctionID
                    Case Else
                        mFunctionId = Null
                End Select
                If val(txtAccountHead.Tag) > 0 Then
                    mDrAccountHeadID = val(txtAccountHead.Tag)
                Else
                    MsgBox "Error : Cash/Bank AccountHead is not set", vbInformation
                    Exit Sub
                End If
                
                '-------------------------------------------------------'
                ' Fill in Transaction Grid For Accounts Posting         '
                '-------------------------------------------------------'
                Call ListPostingHeadsInGridForGeneralReceipts
                If mGrandTotalValidityFlag Then
                    MsgBox "Difference in Grand Total and Item Total!", vbInformation
                    Exit Sub
                End If
                
                
                '-------------------------------------------------------'
                ' Exit Sub                                              '
                '-------------------------------------------------------'
                ' faVoucher                                             '
                '-------------------------------------------------------'
                Dim mintVoucherID_1                As Long
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                Dim mintTransactionTypeID_4        As Long
                Dim mtnyVoucherTypeID_5            As Integer
                Dim mintVoucherNo_6                As Variant
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
                Dim mnumSubLedgerID_21             As Variant
                Dim mintKeyID1_22                  As Variant
                Dim mintKeyID2_23                  As Variant
                Dim mintExternalApplicationID_24   As Long
                Dim mintExternalModuleID_25        As Long
                Dim mintFinancialYearID_26         As Long
                
                Dim mvchBank_33                    As String
                Dim mvchBankPlace_34               As String
                Dim mintFundID_35                  As Long
                Dim mRefNo                         As String
                Dim mRoundOff                      As Single
                Dim mAdvAmtAdj                     As Double
                Dim mtnyVoucherGroupID             As Variant
                Dim mnumLinkKeyID                  As Variant
                
                '@intVoucherID_1     [bigint],
                '@intLocalBodyID_2  [int],
                '@intTransactionID_3    [bigint],
                mintVoucherID_1 = val(txtReceiptNo.Tag)
                mintTransactionTypeID_4 = val(txtTransactionType.Tag)
                mtnyVoucherTypeID_5 = 10
                mintVoucherNo_6 = val(txtReceiptNo.Text)
                mintBookNo_7 = val(txtBookNo.Text)
                mdtDate_8 = gbTransactionDate
                mfltAmount_9 = val(txtTotal.Text)
                mintInstrumentTypeID_10 = val(txtInstrument.Tag)
                mvchInstrumentNo_11 = Trim(txtInstNo.Text)
                mdtInstrumentDate_12 = IIf(Trim(txtDated) <> "", CheckDateInMMM(txtDated), Null)
                mvchDescription_13 = Trim(txtDescription.Text)
                If val(txtAdvance.Text) > 0 Then
                    If Len(mvchDescription_13) > 0 Then mvchDescription_13 = mvchDescription_13 + ", "
                    mvchDescription_13 = mvchDescription_13 + "Advance Adjusted Rs." + Trim(txtAdvance.Text)
                End If
                If cmbZone.ListIndex > 0 Then
                    mnumZoneID_14 = IIf(cmbZone.ItemData(cmbZone.ListIndex) > 0, cmbZone.ItemData(cmbZone.ListIndex), Null)
                End If
                If cmbZone.ListIndex > 0 Then
                    mnumZoneID_14 = cmbZone.ItemData(cmbZone.ListIndex)
                End If
                mnumWardID_15 = val(txtWardNo.Text)
                mintDoorNoP1_16 = val(txtDoorNo1.Text)
                mvchDoorNoP2_17 = Trim(txtDoorNo2.Text)
                If txtIntruptNoSuffix.Text <> "" Then
                    mvchDoorNoP3_18 = Trim(txtIntruptNoSuffix.Text)
                End If
                mintUserID_19 = gbUserID
                mintCounterID_20 = gbCounterID
                mnumSubLedgerID_21 = mSubLedgerID ' mBuildingID ' Changed by Aiby on 10-Dec-2008 From Kollam Corp.
                mintKeyID1_22 = mDrAccountHeadID
                mintKeyID2_23 = mDemandID
                mintExternalApplicationID_24 = AppID.Saankhya
                mintExternalModuleID_25 = 0
                mintFinancialYearID_26 = mYearID
                mvchBank_33 = Trim(txtBank)
                mvchBankPlace_34 = Trim(txtPlace)
                mintFundID_35 = 1
                mRefNo = Trim(txtRefNo.Text)
                mRoundOff = val(txtRoundOff)
                mAdvAmtAdj = val(txtAdvance.Text)
                If mInterruptedModeFlag Then
                    mtnyVoucherGroupID = 4
                    '--------Interrupt Receipt Book---------
                    'objDb.SetConnection mCnn
                    mSql = "Select * From faInterruptedReceiptBooks Where tnyClosed<>1 And intCounterID=" & gbCounterID & "And intFinancialYearID=" & mYearID
                    Rec.Open mSql, mCnn
                
                    If Not (Rec.BOF And Rec.EOF) Then
                        If IsNull(Rec!intBookID) Then
                            MsgBox "Book not Issued for this Counter", vbApplicationModal
                            Exit Sub
                        Else
                            mintBookNo_7 = Rec!intBookID
                        End If
                    Else
                        MsgBox "This Book is Issued to Another Counter/Book is Closed", vbApplicationModal
                        Exit Sub
                    End If
                    Rec.Close
                Else
                    mtnyVoucherGroupID = Null
                End If
                
                'mdtDate_8 = mCnn.Execute
                
                '========================================='
                ' BEGIN TRANSACTION                       '
                '-----------------------------------------'
                    mCnn.BeginTrans
                    On Error GoTo ErrorRollBack:
                '========================================='
                If gbLinkWithPropertyTax Then
                    Call funCancelPropertyTax(mintVoucherNo_6, mintVoucherID_1)
'                ElseIf gbFetchDemandFromWeb = 1 Then
'                     PTaxWebDemand (val(mintVoucherID_1))
                End If
                
                arrInput = Array( _
                mintVoucherID_1, _
                gbLocalBodyID, _
                Null, _
                mintTransactionTypeID_4, _
                mtnyVoucherTypeID_5, _
                mintVoucherNo_6, _
                mintBookNo_7, _
                mdtDate, _
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
                mintExternalModuleID_25, mintFinancialYearID_26, gbShiftID, 1, 0, _
                mvchBank_33, mvchBankPlace_34, mintFundID_35, gbSeatID, gbSessionID, mRefNo, mRoundOff, mAdvAmtAdj, lSoochikaCurrentNo, 0, gbLocationID, mtnyVoucherGroupID, mnumLinkKeyID)

                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID_1 = arrOutPut(0, 0)
                    If mintVoucherID_1 = -1 Then
                        mVoucherID = mintVoucherID_1
                        mReceiptNo = arrOutPut(1, 0)
                    Else
                        mVoucherID = mintVoucherID_1
                    End If
                Else
                    GoTo ErrorRollBack:
                End If
                
                '-------------------------------------------------------'
                ' faVoucher Child
                '-------------------------------------------------------'
                'Dim mintVoucherID_1       As Double  '
                Dim mintLocalBodyID_2       As Long
                Dim mintSlNo_3              As Long
                Dim mintAccountHeadID_4     As Long
                Dim mtnyDebitOrCredit_5     As Byte
                Dim mintYearID_6            As Long
                Dim mtnyPeriodID_7          As Byte
                Dim mtnyArrearFlag_8        As Byte
                Dim mnumDemandID_9          As Double
                Dim mfltAmount_10           As Double
                
                mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & mintVoucherID_1
                
                For mLoopCount = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 Then
                            '----------------------------------------------------------------------------'
                            'NOTE=> vsGrid.Cell(flexcpText, mLoopCount, 14) :: Those Rows Which          '
                            '       Do not want to Save in Child Table eg. Advance Property Tax Adjusted '
                            '----------------------------------------------------------------------------'
                            mintLocalBodyID_2 = gbLocalBodyID
                            mintSlNo_3 = mLoopCount
                            mintAccountHeadID_4 = vsGrid.Cell(flexcpText, mLoopCount, 6)
                            mtnyDebitOrCredit_5 = 0
                            mintYearID_6 = val(vsGrid.Cell(flexcpText, mLoopCount, 7))
                            mtnyPeriodID_7 = val(vsGrid.Cell(flexcpText, mLoopCount, 8))
                            If mintYearID_6 < mYearID Then
                                mtnyArrearFlag_8 = 1
                            Else
                                mtnyArrearFlag_8 = 0
                            End If
                            mnumDemandID_9 = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                            mfltAmount_10 = val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                            
                            Set arrInput = Nothing
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
                            objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                        End If
                    Else
                        Exit For
                    End If
                Next mLoopCount
                
                '-------------------------------------------------------'
                ' faVoucher Address
                '-------------------------------------------------------'
                '@intVoucherID   [bigint],
                '@intLocalBodyID    [int],
                '@vchName   [varchar](100),
                '@vchInit1  [varchar](2) = Null,
                '@vchInit2  [varchar](2) = Null ,
                '@vchInit3  [varchar](2) = Null,
                '@vchInit4  [varchar](2) = Null,
                '@vchHouseName  [varchar](100) = Null,
                '@vchStreetName [varchar](100) = Null,
                '@vchLocalPlace [char](10) = Null,
                '@vchMainPlace  [varchar](100)= Null,
                '@vchPostOffice [varchar](100) = Null,
                '@vchDistrict   [varchar](50)= Null,
                '@vchPinNumber  [varchar](6) = Null,
                '@vchPhone  [varchar](15)= Null),
                '@intWardNo  [int] = Null,
                '@intDoorNo [int] = Null,
                '@vchDoorNo2    [varChar](10) = Null
                
                '-----------------------'
                '       Added Newly     '
                '-----------------------'
                'vchName = Trim(txtPayee.Text)
                vchName = Trim(txtName.Text)
                vchHouseName = Trim(txtHouse.Text)
                vchInit1 = Trim(txtInit1.Text)
                vchInit2 = Trim(txtInit2.Text)
                vchInit3 = Trim(txtInit3.Text)
                vchInit4 = Trim(txtInit4.Text)
                vchStreetName = Trim(txtStreet.Text)
                vchLocalPlace = Trim(txtLocalPlace.Text)
                vchMainPlace = Trim(txtMainPlace.Text)
                vchPostOffice = Trim(txtPost.Text)
                vchPinNumber = txtPin.Text
                vchPhone = txtPhone.Text
                intWardNo = txtWardNo.Text
                intDoorNo = txtDoorNo1.Text
                vchDoorNo2 = txtDoorNo2.Text
                '-----------------------'
                
                mCnn.Execute "Delete From faVoucherAddress Where intVoucherID = " & mintVoucherID_1
                
                arrInput = Array(mintVoucherID_1, _
                        gbLocalBodyID, _
                        vchName, _
                        vchInit1, _
                        vchInit2, _
                        vchInit3, _
                        vchInit4, _
                        vchHouseName, _
                        vchStreetName, _
                        vchLocalPlace, _
                        vchMainPlace, _
                        vchPostOffice, _
                        vchDistrict, _
                        vchPinNumber, _
                        vchPhone, _
                        intWardNo, _
                        intDoorNo, _
                        vchDoorNo2)
                objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn
                
                '-------------------------------------------------------'
                ' Transactions                                          '
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
                
                intTransactionID_1 = val(txtDate.Tag)
                mintLocalBodyID_2 = gbLocalBodyID
                mintFinancialYearID_3 = gbFinancialYearID
                mdtTransactionDate_4 = gbTransactionDate
                mintExternalApplicationID_5 = AppID.Saankhya
                mintExternalApplicationModuleID_6 = 0
                mintFunctionID_7 = mFunctionId
                mintFunctionaryID_8 = mFunctionaryID
                mintFieldID_9 = IIf(val(txtWard) < 1, Null, val(txtWard))
                mintFundID_10 = Null
                mintBudgetCentreID_11 = Null
                mvchNarration_12 = Trim(txtDescription.Text)
                If val(txtAdvance.Text) > 0 Then
                    mvchNarration_12 = mvchNarration_12 + "Advance Amount Adjusted Rs." + Trim(txtAdvance.Text)
                End If
                mintTransactionTypeID_13 = mTransactionType
                mintVoucherNo_14 = mintVoucherID_1
                mintProcessID_15 = Null
                mvchGroup_16 = "R"
                mintGroupID_17 = 10
                mintKeyID_18 = Null     'mDemandID 'Added on 3-Sep-2008
                mnumSubLedgerID_19 = mBuildingID
                'mintUserID_20 = gbUserID
                
                arrInput = Array( _
                intTransactionID_1, _
                mintLocalBodyID_2, _
                mintFinancialYearID_3, _
                mdtDate, _
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
                objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    intTransactionID_1 = arrOutPut(0, 0)
                Else
                    GoTo ErrorRollBack:
                End If
                
                '-------------------------------------------------------'
                ' Transaction Child                                     '
                '-------------------------------------------------------'
                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & intTransactionID_1
                '=========================================================================================='
                '                                                                                          '
                ' BLOCK: I   :  Accounting Part of Advance Adjustment of Property Tax                      '
                '                                                                                          '
                ' a) Advance will be saved in Voucher Table as every normal voucher as it saves. Its Acco- '
                ' -uning part will handled by Transaction Tables. Order in which the advance settled off by'
                ' Penal Interest, PTax(Arrear)+LC, PTax(Current)+Lc                                        '
                ' b) This Block Only Handles Property Tax Advance                                          '
                ' c)                                                                                       '
                '------------------------------------------------------------------------------------------'
                If mAdvAmtAdj > 0 Then
                    Dim mTrChild As uTrChild
                    Dim mFineFlag As Boolean
                    Dim mFineAmt As Double
                    Dim mAmt As Double
                    Dim mPTax As Double
                    Dim mLC As Double
                    Dim mSL As Integer
                    Dim mExitLoopFlag As Boolean
                    Dim mByHeadID As Integer
                    
                    'NOTE:- Check TransactionTypes
                    If mintTransactionTypeID_4 = gbTransactionTypePTax Then
                        'NOTE:- Posting of Advance Collection of Property Tax
                        mSL = 2
                        With mTrChild
                            .intTransactionID = intTransactionID_1
                            .intSerialNo = mSL
                            .intAccountHeadID = gbAcHeadIDAdvancePTax
                            .fltAmount = mAdvAmtAdj
                            .tinDebitOrCreditFlag = 1
                            .intByAccountHeadID = Null
                            .vchNarration = "Total Advance Collection Adjusted"
                            .intFundID = 1
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        End With
                        
                        'NOTE:- Checking for Penal Interest
                        For mLoopCount = 1 To vsGrid.Rows - 1
                            If val(vsGrid.TextMatrix(mLoopCount, 6)) = gbAcHeadIDPenalInterest Then
                                'Note:- Found Penal Interest in Grid
                                '       Which is expected Only once Penal Interest Appear in Grid
                                mFineFlag = True ' Fine Exists
                                mFineAmt = val(vsGrid.TextMatrix(mLoopCount, 11))
                                If mAdvAmtAdj >= mFineAmt Then ' Advance Amount is greater than total Penal Interest
                                    mAmt = mFineAmt
                                    mFineAmt = 0
                                Else   'Note:- Advance Amount will completely settled off by Penal interest
                                    mAmt = mAdvAmtAdj
                                    mFineAmt = mFineAmt - mAmt
                                    '(A)-->> Note:- Remaining Fine Should Set off With Cash/Bank Heads
                                End If
                                mAdvAmtAdj = mAdvAmtAdj - mAmt
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDPenalInterest
                                    .fltAmount = mAmt
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                    .vchNarration = "Advance Collection Adjusted With Penal Interest"
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                                Exit For
                            End If
                        Next mLoopCount 'NOTE:- END of Setting Penal Interest if find in Grid
                        mLoop = 1 'NOTE:- Using in ELSE Part
                        'NOTE:- Remaining Balance in Advance Collection
                        If mAdvAmtAdj > 0 Then
                            '----------------------------------------------------------------------'
                            'NOTE:- After Penal Interest Set Off                                   '
                            '       Checking for Property Tax Arrear or Current heads in Grid      '
                            '       in the same order as it appears in Grid                        '
                            '----------------------------------------------------------------------'
                            For mLoopCount = 1 To vsGrid.Rows - 1
                                If vsGrid.TextMatrix(mLoopCount, 0) <> gbAcHeadCodeAdvancePTax Then
                               'Note:- Sum of Property Tax + LC
                                mPTax = val(vsGrid.TextMatrix(mLoopCount, 11))
                                mLoopCount = mLoopCount + 1 'Note:- Finding the Library Cess
                                If vsGrid.TextMatrix(mLoopCount, 0) = gbAcHeadCodeLibraryCess Then
                                    mLC = val(vsGrid.TextMatrix(mLoopCount, 11))
                                End If
                                mAmt = mPTax + mLC 'Sum PTax + LC
                                mSL = mSL + 1
                                'Note:- IF Sum of PTax+LC Greater than Advance Amount
                                If mAdvAmtAdj >= mAmt Then
                                    With mTrChild
                                        .intTransactionID = intTransactionID_1
                                        .intSerialNo = mSL
                                        'Note:- mLoopCount - 1 => Loop count is on LC, its to find PTax head One should
                                        '       check in the previous row
                                        If val(vsGrid.TextMatrix(mLoopCount - 1, 6)) = gbAcHeadIDPropertyTaxArrear Then
                                            .intAccountHeadID = gbAcHeadIDPropertyTaxArrear
                                        Else
                                            .intAccountHeadID = gbAcHeadIDPropertyTaxCurrent
                                        End If
                                        .fltAmount = mPTax
                                        .tinDebitOrCreditFlag = 0
                                        .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                        .vchNarration = "Adv.Adjusted With Property Tax " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                        .intFundID = 1
                                        
                                        arrInput = Array(.intTransactionID, _
                                        .intSerialNo, _
                                        .intAccountHeadID, _
                                        .fltAmount, _
                                        .tinDebitOrCreditFlag, _
                                        .intByAccountHeadID, _
                                        .vchNarration, _
                                        .intFundID)
                                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                        
                                        'NOTE:- Library Cess in the very next row in Grid (Thats what Expected!! ;) )
                                        mSL = mSL + 1
                                        .intTransactionID = intTransactionID_1
                                        .intSerialNo = mSL
                                        .intAccountHeadID = gbAcHeadIDLibraryCess
                                        .fltAmount = mLC
                                        .tinDebitOrCreditFlag = 0
                                        .intByAccountHeadID = gbAcHeadIDAdvancePTax
                                        .vchNarration = "Adv. Collection Adjusted With Library Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                        .intFundID = 1
                                        
                                        arrInput = Array(.intTransactionID, _
                                        .intSerialNo, _
                                        .intAccountHeadID, _
                                        .fltAmount, _
                                        .tinDebitOrCreditFlag, _
                                        .intByAccountHeadID, _
                                        .vchNarration, _
                                        .intFundID)
                                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                    End With
                                    mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC)
                                Else
                                    '
                                    ' NOTE:- Advance Amount is Less than Total PTax+LC.
                                    '        Remaining Advance Amount will be split into PTax & LC by ratio.
                                    '        The rest part will again adjusted against Cash/Bank Account.
                                    mPTax = Format(mAdvAmtAdj * 100 / 105, "0.0")
                                    mLC = mAdvAmtAdj - mPTax
                                    mAdvAmtAdj = mAdvAmtAdj - (mPTax + mLC)
                                    With mTrChild
                                            mByHeadID = gbAcHeadIDAdvancePTax
Step2:
                                        
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            ' Note:- mLoop - 1 => Loop count is on LC, its to find PTax head One should
                                            '        check in the previous row
                                            If val(vsGrid.TextMatrix(mLoopCount - 1, 6)) = gbAcHeadIDPropertyTaxArrear Then
                                                .intAccountHeadID = gbAcHeadIDPropertyTaxArrear
                                            Else
                                                .intAccountHeadID = gbAcHeadIDPropertyTaxCurrent
                                            End If
                                            .fltAmount = mPTax
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mByHeadID
                                            .vchNarration = "Adv.Adjusted With Property Tax " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                            
                                            'NOTE:- Library Cess in the very next row
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            .intAccountHeadID = gbAcHeadIDLibraryCess
                                            .fltAmount = mLC
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mByHeadID
                                            .vchNarration = "Adv. Collection Adjusted With Library Cess " & vsGrid.TextMatrix(mLoopCount, 2) & "-" & vsGrid.TextMatrix(mLoopCount, 3)
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                            
                                            'Rest part to adjust with Cash Or Bank AccountHead
                                        If Not mExitLoopFlag Then
                                            mPTax = val(vsGrid.TextMatrix(mLoopCount - 1, 11)) - mPTax
                                            mLC = val(vsGrid.TextMatrix(mLoopCount, 11)) - mLC
                                            mExitLoopFlag = True
                                            mByHeadID = mDrAccountHeadID
                                            GoTo Step2:
                                        Else
                                            Exit For
                                        End If
                                    End With
                                End If
                                End If ' If vsGrid.TextMatrix(mLoopCount, 0) = gbAcHeadCodeAdvancePTax Then
                            Next mLoopCount
                            
                            If mLoopCount < vsGrid.Rows - 1 Then
                                mLoop = mLoopCount + 1
                            End If
                            GoTo Step3:
                        Else 'Note:- Else Part OF Condition [If mAdvAmtAdj > 0 Then] : After Penal Interest
                             'NOTE:- Advance Amount is settled.
                             '       Rest part of the accounting posting which is collected as Cash or Bank.
Step3:
                             'NOTE:- Cash Or Bank With SerialNo 1
                             With mTrChild
                                 .intTransactionID = intTransactionID_1
                                 .intSerialNo = 1
                                 .intAccountHeadID = mDrAccountHeadID
                                 .fltAmount = mfltAmount_9
                                 .tinDebitOrCreditFlag = 1
                                 .intByAccountHeadID = Null
                                 .vchNarration = Null
                                 .intFundID = 1
                                
                                 arrInput = Array(.intTransactionID, _
                                 .intSerialNo, _
                                 .intAccountHeadID, _
                                 .fltAmount, _
                                 .tinDebitOrCreditFlag, _
                                 .intByAccountHeadID, _
                                 .vchNarration, _
                                 .intFundID)
                                 objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                             End With
                             For mLoopCount = mLoop To vsGrid.Rows - 1
                                If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                                    If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) <> 1 And val(vsGrid.Cell(flexcpText, mLoopCount, 6)) <> gbAcHeadIDPenalInterest Then
                                    
                                        'NOTE=> vsGrid.Cell(flexcpText, mLoopCount, 14) :: Those Rows Which Do not
                                        '       want to Save in Child Table eg. Advance Property Tax Adjusted
                                        With mTrChild
                                            mSL = mSL + 1
                                            .intTransactionID = intTransactionID_1
                                            .intSerialNo = mSL
                                            .intAccountHeadID = val(vsGrid.Cell(flexcpText, mLoopCount, 6))
                                            .fltAmount = val(vsGrid.Cell(flexcpText, mLoopCount, 11))
                                            .tinDebitOrCreditFlag = 0
                                            .intByAccountHeadID = mDrAccountHeadID
                                            .vchNarration = Null
                                            .intFundID = 1
                                            
                                            arrInput = Array(.intTransactionID, _
                                            .intSerialNo, _
                                            .intAccountHeadID, _
                                            .fltAmount, _
                                            .tinDebitOrCreditFlag, _
                                            .intByAccountHeadID, _
                                            .vchNarration, _
                                            .intFundID)
                                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                        End With
                                    End If
                                End If
                             Next mLoopCount
                             'NOTE:- IF Penal Interest is not Set off completely then
                             '-->> (A)  Continuation
                             If mFineAmt > 0 Then
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDPenalInterest
                                    .fltAmount = mFineAmt
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = mDrAccountHeadID
                                    .vchNarration = Null
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                             End If
                             
                             'NOTE:-Round of Adjustment adjusted
                             If mRoundOff > 0 Then
                                With mTrChild
                                    mSL = mSL + 1
                                    .intTransactionID = intTransactionID_1
                                    .intSerialNo = mSL
                                    .intAccountHeadID = gbAcHeadIDRoundOff
                                    .fltAmount = mRoundOff
                                    .tinDebitOrCreditFlag = 0
                                    .intByAccountHeadID = mDrAccountHeadID
                                    .vchNarration = "Round Of Adjustment"
                                    .intFundID = 1
                                    
                                    arrInput = Array(.intTransactionID, _
                                    .intSerialNo, _
                                    .intAccountHeadID, _
                                    .fltAmount, _
                                    .tinDebitOrCreditFlag, _
                                    .intByAccountHeadID, _
                                    .vchNarration, _
                                    .intFundID)
                                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                End With
                             End If
                             GoTo GotoCommitTran:  'NOTE:- Complete the Transaction By CommitTrans
                        End If ' END OF Condition [If mAdvAmtAdj > 0 Then] : After Penal Interest Set off
                    ElseIf mintTransactionTypeID_4 = gbTransactionTypeRentOnBuilding Then
                        Call SaveRentAdv(intTransactionID_1, gbAcHeadCodeAdvanceBuilding, mCnn)
                    ElseIf mintTransactionTypeID_4 = gbTransactionTypeRentOnLand Then
                       Call SaveRentAdv(intTransactionID_1, gbAcHeadCodeAdvanceLand, mCnn)
                    End If     ' End of Checking Transaction Type : Property Tax
                End If         ' End of Advance Collection Posting Block 1
                '=========================================================================================='
                ' END OF BLOCK 1 : Advance Adjustment of Property Tax - Integrated Sanchaya Mode           '
                '=========================================================================================='
                For mLoop = 1 To vsGridTransactions.Rows - 1
                    '-------------------------------------------------------------'
                    'NOTE=> ALL TRANSACTIONS EXCEPT PROPERTY TAX - POSTING HERE   '
                    '-------------------------------------------------------------'
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
                                    gbFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                Next mLoop
                '-------------------------------------------------------'
                ' Round Off Adjustment to Transaction Child             '
                '-------------------------------------------------------'
                If val(txtRoundOff) > 0 Then
                    mintAccountHeadID_4 = gbAcHeadIDRoundOff
                    If mintAccountHeadID_4 = -1 Then
                        mintAccountHeadID_4 = Null
                    End If
                    arrInput = Array(intTransactionID_1, _
                                    mLoop, _
                                    mintAccountHeadID_4, _
                                    val(txtRoundOff), _
                                    0, _
                                    mDrAccountHeadID, _
                                    "Round Off Adj.", _
                                    gbFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                End If
                
                '-------------------------------------------------------'
                ' Update Demand Table                                   '
                '-------------------------------------------------------'
                'If mTransactionType = 1 Then
                
                Dim mStatusFlag As Integer
                If val(txtDemandNo.Tag) > 0 Then
                    mDemandID = txtDemandNo.Tag
                    mStatusFlag = 1
                    arrInput = Array(mDemandID, mStatusFlag, mintVoucherID_1)
                    objdb.ExecuteSP "spUpdateIDemandStatus", arrInput, , , mCnn
                Else
                    For mLoop = 1 To vsGrid.Rows - 1
                        If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked And mDemandID <> vsGrid.Cell(flexcpText, mLoop, 10) Then
                            mDemandID = val(vsGrid.TextMatrix(mLoop, 10))
                            mStatusFlag = 1
                            arrInput = Array(mDemandID, mStatusFlag, mintVoucherID_1)
                            objdb.ExecuteSP "spUpdateIDemandStatus", arrInput, , , mCnn
                        End If
                    Next mLoop
                End If
                
                '========================================='
                ' Sharing Data to KMBR and SOOCHIKA       '
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
                    'If mKMBRFlag = True Then
                    '    Dim mCnnKMBR As New ADODB.Connection
                    '    If objDb.CreateNewConnection(mCnnKMBR, enuSourceString.KMBR) = True Then
                    '        mCnnKMBR.BeginTrans
                    '        If SaveSanketham(lSoochikaCurrentNo, mCnnKMBR) = True Then
                    '
                    '            mCnnSoochika.CommitTrans
                    '            mCnnKMBR.CommitTrans
                    '
                    '        Else
                    '            GoTo ErrorRollBack:
                    '        End If
                    '    End If
                    'End If
                End If
                '========================================='
                
                
                '========================================='
                '             Saving to Sevana Reg        '
                '========================================='
                If mSoochikaConnected = True Then
                    'Dim mCnnSevanaReg As New ADODB.Connection
                    'If objDb.CreateNewConnection(mCnnSevanaReg, enuSourceString.SevanaRegn) = True Then
                    '    mCnnSevanaReg.BeginTrans
                    '    On Error GoTo ErrorRollBack:
                    '    If frmSoochikaInward.SaveSevana(lSoochikaCurrentNo, mReceiptNo, mfltAmount_9, mCnnSevanaReg) = True Then
                    '        mCnnSoochika.CommitTrans
                    '        mCnnSevanaReg.CommitTrans
                    '    Else
                    '        GoTo ErrorRollBack:
                    '    End If
                    'End If
                End If
                '========================================='
                
GotoCommitTran:
                
                '========================================='
                ' TRANSACTION COMMITTING                  '
                '-----------------------------------------'
                    mCnn.CommitTrans
                    Set mCnn = Nothing
                    On Error GoTo 0
                '========================================='
                Call LockForm(False)
                
                mGrandTotal = mGrandTotal + val(txtTotal)
                If mStartingReceiptNo = 0 Then
                    mStartingReceiptNo = txtReceiptNo.Text
                    lblFromReceiptNo.Caption = mStartingReceiptNo
                End If
                lblToReceiptNo.Caption = txtReceiptNo.Text
                lblGroupTotal.Caption = mGrandTotal
                If mInterruptedModeFlag = False Then
'                     If gbLBPanchayat = 1 Then               'ADDED BY MINU ON 26/09/2011
'                            Call PrintReceipt_ForNewFormat_ModifiedByMinu(mintVoucherID_1)
'                     Else
'                            Call PrintReceipt(mintVoucherID_1)
'                     End If 'Call PrintReceipt(mintVoucherID_1)
                End If
                '========================================='
                ' Soochika Inward Printing
                '========================================='
                If mSoochikaConnected = True Then
                    'On Error GoTo 0
                    'frmSoochikaInward.Ack (frmSoochikaInward.lSoochikaFeildID)
                    'Unload frmSevanaInward
                    'frmSoochikaInward.ClearDetails
                    'Call FormInitialize
                    'Unload Me
                End If
                
                'Call FormInitialize
                
                
                
'''                If gbFetchDemandFromWeb = 1 Then
'''                    If mTransactionType = gbTransactionTypePTax Then
'''                        If mDemandMode <= 1 Then
'''                            Dim mCollPost       As String
'''                            Dim mColZoneID      As String
'''                            Dim mBuildingIdWeb  As String
'''                            Dim mColAmt            As String
'''                            Dim mColDate        As String
'''                            Dim mColReceiptNo   As String
'''                            Dim mColBookNo      As String
'''                            Dim mColPeriodId     As String
'''                            Dim mColYearID       As String
'''                            Dim mHash           As String
'''                            Dim mCollOut        As String
'''        '                    Dim node            As IXMLDOMNode
'''        '                    Dim DataNodes       As IXMLDOMNodeList
'''                            Dim mUrl            As String
'''                            Dim objSOAP         As Variant
'''                            Dim mLen            As Integer
'''                            Dim mColAccID       As String
'''                            Dim mColKeyID       As String
'''
'''
'''                            mUrl = gbDefaultUrlSanchayaPost
'''                            Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
'''                            objSOAP.MSSoapInit mUrl + "?WSDL"
'''                            Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
'''                            If Not (Rec.EOF And Rec.BOF) Then
'''                                While Not Rec.EOF
'''
'''                                    mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'''                                    mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'''                                    mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'''                                    mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
'''                                    mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
'''                                    mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
'''                                    mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
'''                                    mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
'''                                    mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'''                                    mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
'''                                    If mColAccID <> gbAcHeadIDPenalInterest Then
'''                                        If mColAccID <> gbAcHeadIDNoticeFee Then
'''                                            mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
'''                                            mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mintVoucherID_1) + "#"
'''                                            mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
'''                                            mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
'''                                            mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
'''                                            mCollPost = mCollPost + "~"   ''Modified on 27/Dec/2016
'''                                        End If
'''                                    End If
'''                                    Rec.MoveNext
'''                                    'mCollPost = mCollPost + "~"
'''                                Wend
'''                                'mLen = Len(mCollPost) - 1
'''                                mCollPost = Left$(mCollPost, Len(mCollPost) - 1)
'''                                mHash = CStr(mintVoucherID_1) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
'''                                mCollOut = objSOAP.Saankhyaa_CollectionPosting(mCollPost, mHash)
'''                            End If
'''                        End If
'''                    End If
                '========================================='
                ' Sharing Data to SanchayaLite            '
                '-----------------------------------------'
                If gbLinkWithPropertyTax Then
                    If mTransactionType = 1 Then
                        Dim mchvReceiptNO As String
                        Dim mIsAdvance As Integer
                        
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If Not (Rec.EOF And Rec.BOF) Then
                            Set mCnn = Nothing
                            If objdb.CreateNewConnection(mCnn, SanchayaLite) Then
                                
                                Dim intKeyID As Variant
                                Dim chvReceiptNo As Variant
                                Dim chvReceiptDate As Variant
                                Dim intCollectionYear As Variant
                                Dim tnySource As Variant
                                Dim tnyPaymentReceived As Variant
                                Dim numLocation As Variant
                                Dim fltAmt As Variant
                                
                                If frmPropertyTax.mvarDifferentZoneFlag = False Then
                                
                                    While Not Rec.EOF
                                        '@intKeyID   Int,
                                        '@intVoucherID  BigInt,
                                        '@chvReceiptNo varChar(20),
                                        '@chvReceiptDate varChar(12),
                                        '@intCollectionYear Int,
                                        '@tnySource Tinyint ,
                                        '@tnyPaymentReceived TinyInt
                                        '@numLocation Numeric
                                        intKeyID = Rec!numDemandID
                                        mintVoucherID_1 = Rec!intVoucherID
                                        chvReceiptNo = Rec!intVoucherNo
                                        chvReceiptDate = Rec!dtDate
                                        intCollectionYear = gbFinancialYearID
                                        tnySource = 2
                                        tnyPaymentReceived = 1
                                        
                                        arrInput = Array(intKeyID, _
                                                            mintVoucherID_1, _
                                                            chvReceiptNo, _
                                                            chvReceiptDate, _
                                                            intCollectionYear, _
                                                            tnySource, _
                                                            tnyPaymentReceived)
                                        objdb.ExecuteSP "spCloseDemandFromSaankhya", arrInput, , , mCnn
                                        
                                        Rec.MoveNext
                                    Wend
                                    '========================================================================================'
                                    ' ADVANCE CLOSING USING GRID VALUE                                                       '
                                    '========================================================================================'
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If val(vsGrid.Cell(flexcpText, mLoopCount, 14)) = 1 Then
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            arrInput = Array(intKeyID, _
                                                            mintVoucherID_1, _
                                                            chvReceiptNo, _
                                                            chvReceiptDate, _
                                                            intCollectionYear, _
                                                            tnySource, _
                                                            tnyPaymentReceived)
                                            objdb.ExecuteSP "spCloseDemandFromSaankhya", arrInput, , , mCnn
                                        End If
                                    Next
                                Else
                                    '====================================================================='
                                    '   Modified On 12-aug-2009 by Cijith For Sanchaya Zonal Connectivity'
                                    '====================================================================='
                                    Dim intcnt As Integer
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If vsGrid.TextMatrix(mLoopCount, 6) = "" Then Exit For
                                        If vsGrid.TextMatrix(mLoopCount, 6) <> 113 Then
                                            intcnt = intcnt + 1
                                        End If
                                    Next
                                    arrInput = Array(gbLocationID, mintVoucherID_1, Rec!intVoucherNo, _
                                                    Rec!dtDate, Rec!numSubLedgerID, _
                                                    Rec!numZoneID, mAssesmentYearID, _
                                                    Rec!numWardId, Rec!intDoorNoP1, Rec!vchDoorNoP2, _
                                                    vchName, 2, gbFinancialYearID, 0, Rec!fltTotalAmt, _
                                                    intcnt, 1)
                                    objdb.ExecuteSP "HOsnSaanOtherCollectionsI", arrInput, , , mCnn
                                        
                                    Dim numSanchayaHeadId As Integer
                                    Dim numSankhyaHeadID As Integer
                                    For mLoopCount = 1 To vsGrid.Rows - 1
                                        If vsGrid.Cell(flexcpText, mLoopCount, 0) <> "" Then
                                            numSankhyaHeadID = val(vsGrid.Cell(flexcpText, mLoopCount, 6))
                                            If numSankhyaHeadID = 1385 Or numSankhyaHeadID = 1386 Then
                                               numSanchayaHeadId = 1
                                            ElseIf numSankhyaHeadID = 1126 Then
                                                numSanchayaHeadId = 2
                                            ElseIf numSankhyaHeadID = 1157 Then
                                                numSanchayaHeadId = 4
                                            ElseIf numSankhyaHeadID = 113 Then ' Modified By Aiby  To Give Penal Interest
                                                numSanchayaHeadId = 90
                                            Else
                                                numSanchayaHeadId = 0
                                            End If
                                            intKeyID = val(vsGrid.Cell(flexcpText, mLoopCount, 10))
                                            fltAmt = IIf(val(vsGrid.Cell(flexcpText, mLoopCount, 5)) = 0, val(vsGrid.Cell(flexcpText, mLoopCount, 4)), val(vsGrid.Cell(flexcpText, mLoopCount, 5)))
                                            arrInput = Array(gbLocationID, _
                                                        mintVoucherID_1, _
                                                        mLoopCount, _
                                                        val(vsGrid.Cell(flexcpText, mLoopCount, 7)), _
                                                        val(vsGrid.Cell(flexcpText, mLoopCount, 8)), _
                                                        numSanchayaHeadId, _
                                                        fltAmt, _
                                                        intKeyID)
                                            objdb.ExecuteSP "HOsnSaanOtherCollectionsSubI", arrInput, , , mCnn
                                        End If
                                    Next
                                End If
                                '========================================================================================'
                            Else
                                MsgBox "(Sanchaya)Connection Error:", vbInformation
                            End If
                        End If
                    End If
                End If
                
                
                '========================================='
                ' Updating Demand Details For Rent on Land'
                ' and Buildings (DB_Sanchaya)       '
                ' Codded By Anisha
                '-----------------------------------------'
                If gbLinkWithRentOnLand Then
                    If mTransactionType = gbTransactionTypeRentOnBuilding Then
                        Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                        If Not (Rec.EOF And Rec.BOF) Then
                                If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                                    Dim IsAdvance As Integer
                                    Dim numRLBDemand As Variant
                                    Dim numZonalOfficeID As Variant
                                    Dim numVoucherId As Variant
                                    Dim dtReceiptDate As String
                                    Dim tnyReceiptSource As Integer
                                    Dim fltAdvance As Double
                                    Dim intYearID As Integer
                                    Dim chvPeriodID As String
                                    While Not Rec.EOF
                                        IsAdvance = 0
                                        numRLBDemand = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                                        numZonalOfficeID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                                        numVoucherId = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                                        chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                        dtReceiptDate = Rec!dtDate
                                        tnyReceiptSource = 2
                                        fltAdvance = 0
                                        intYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                                        chvPeriodID = Rec!tnyPeriodID
                                        arrInput = Array(IsAdvance, _
                                                        numRLBDemand, _
                                                        numZonalOfficeID, _
                                                        numVoucherId, _
                                                        chvReceiptNo, _
                                                        dtReceiptDate, _
                                                        tnyReceiptSource, _
                                                        gbLocationID, _
                                                        fltAdvance, _
                                                        intYearID, _
                                                        chvPeriodID, _
                                                        mSubLedgerID)
                                        objdb.ExecuteSP "spSanRentDemandClose", arrInput, , , mCnn, adCmdStoredProc
                                        Rec.MoveNext
                                    Wend
                                    
                                End If
                        End If
                    End If
                End If
                '========================================='
                ' Insert Receipt Details On DB_SanchayaLite  For PFA,D&O  Licence Fee    '
                ' Created On ON 18/02/2010 By Anisha        '
                'Modified On Jan 2018 for D&O Web Integration
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypeDandO Or mTransactionType = gbTransactionTypePFA Then
                    Set Rec = GetRecordSet("spGetVoucherDetails " & mintVoucherID_1 & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                    Dim numReceiptLocationId As Double
                    Dim intReceiptYear      As Integer
                    Dim numZoneID           As Variant
                    Dim flagSankhya             As Integer
                    tnyReceiptSource = 2
                    tnyPaymentReceived = 1
                    flagSankhya = 0
                    If gbLinkWithDandOPFA Then
                        If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
                            If Not (Rec.EOF And Rec.BOF) Then
                                chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                                chvReceiptDate = Format(IIf(IsNull(Rec!dtDate), "", Rec!dtDate), "dd/mm/yyyy")
                                mDemandID = Rec!numDemandID
                                numZoneID = cmbZone.ItemData(cmbZone.ListIndex) 'gbLocationID
                                mVoucherID = Rec!intVoucherID
                                arrInput = Array(numZoneID, _
                                            gbLocalBodyID, _
                                            mVoucherID, _
                                            mDemandID, _
                                            chvReceiptDate, _
                                            flagSankhya)
                                objdb.ExecuteSP "spsnLicSanCollection_I", arrInput, , , mCnn, adCmdStoredProc
                            End If
                        Else
                            MsgBox "Connection to Sanchaya Doesn't Exists"
                        End If
'                        for D&O Web Integration
                    ElseIf gbLinkWithDandOWeb And mTransactionType = gbTransactionTypeDandO Then
                        Dim objSOAP             As Variant
                        Dim mArrOutDemand       As Variant
                        Dim mUrl                As String
                        Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                        
                         mUrl = gbDefaultUrl 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultUrl")
                         On Error Resume Next
                         objSOAP.MSSoapInit (mUrl + "?WSDL")
                        
                         If Not (Rec.EOF And Rec.BOF) Then
                             chvReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                             chvReceiptDate = IIf(IsNull(Rec!dtDate), "", Format(Rec!dtDate, "dd/MMM/yyyy"))
                            ' mRcptDt = Format$(chvReceiptDate, "dd/MMM/yyyy") '"yyyy/mm/dd")
                             mDemandID = Rec!numDemandID
                             numZoneID = cmbZone.ItemData(cmbZone.ListIndex) 'gbLocationID
                             mVoucherID = Rec!intVoucherID
                             Dim mCredencial As String
                             Dim VoucherDetails As String
                             mCredencial = "ikm@revenue@sanchaya"
                             arrInput = Array(gbLocalBodyID, _
                                           mDemandID, _
                                           chvReceiptNo, _
                                           chvReceiptDate, _
                                           flagSankhya, _
                                           mVoucherID)
                                           
                              arrInput = gbLocalBodyID & "#" & mDemandID & "#" & chvReceiptNo & "#" _
                                            & chvReceiptDate & "#" & flagSankhya & "#" & mVoucherID
                             mArrOutDemand = (objSOAP.savereceiptdetails(arrInput))
                         Else
                             MsgBox "Connection to Sanchaya Doesn't Exists"
                         End If
                        
                    End If
                End If
                '========================================='
                ' Updating Receipt Details IN KMBR        '
                ' Modified ON 04/05/2009 By Cijith        '
                '-----------------------------------------'
                If mTransactionType = gbTransactionTypePermitFeeFromKMBR And mKMBRFlag = True Then
                    If objdb.CreateNewConnection(mCnn, enuSourceString.KMBR) Then
                        arrInput = Array(mReceiptNo, mdtDate, mDemandID, mVoucherID)
                        objdb.ExecuteSP "UGetReceiptNoDate", arrInput, , , mCnn, adCmdStoredProc
                    End If
                End If
                '=================Zonal Collection===================
                'Updating Status of Demand in SaankhyaHo
                
                'If mTransactionType = gbTransactionTypeZonalCollection And gbLinkWithFinanceHO = 1 Then
                If mDemandMode = 2 And gbLinkWithFinanceHO = 1 Then
                    If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaHO)) Then
                        arrInput = Array(mDemandID, mVoucherID, mdtDate, 1) ''Changed gbTransactionDate with mdtDate
                        objdb.ExecuteSP "spUpdateDemandStatus", arrInput, , , mCnn, adCmdStoredProc
                    Else
                        MsgBox "SaankhyaHo Connection Does not exists"
                    End If
                End If
                '====================================================
                
                '========================================='
                ' Updating Status of Fine wave        '
                ' Added ON 04/05/2009 By Anisha       '
                '-----------------------------------------'
                If mFinewave Then
                    If (objdb.SetConnection(mCnn)) Then
                        objdb.ExecuteSP "Update faFineWaiver set tnyStatus=0  Where intVoucherNo=(Select intVoucherNo From faVouchers Where intVoucherID=" & mVoucherID & ")", , , , mCnn, adCmdText
                    End If
                End If
                
        Else
                Debug.Print "Error in establishing connection with Saankhya DB"
                Exit Sub
        End If

        '--------------------To Calculate The Group Total ----------------------'
        'Call GroupCalc
        '-----------------------------------------------------------------------'
        Call UpdateIRRegister(mVoucherID, mdtDate, mfltAmount_9)  'ADDED BY MINU FOR IR REGISTER
        
        Exit Sub
ErrorRollBack:
        MsgBox "Saankhya Error Handler: " & Error$
        mCnn.RollbackTrans
        Set mCnn = Nothing
        
        '---------------------------------------------------------------'
        ' KMBR Roll Back
        '---------------------------------------------------------------'
        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
            If mKMBRFlag = True Then
                'mCnnSoochika.RollbackTrans
                'mCnnKMBR.RollbackTrans
            End If
        End If
        
ErrorRollBackSoochika:

        If mSoochikaConnected = True Then
            'If mCnnSevanaReg.State Then
            '    mCnnSevanaReg.RollbackTrans
            'End If
            'If mCnnSoochika.State Then
            '    mCnnSoochika.RollbackTrans
            'End If
        End If
    End Sub
        Private Function SaveRentAdv(ByVal mTID As Double, ByVal mAdvHeadCode As String, mCnn As ADODB.Connection) As Boolean
        Dim objdb   As New clsDB
        Dim objAcc  As New clsAccounts
        Dim Rec     As New ADODB.Recordset
       ' Dim mCnn    As New ADODB.Connection
        Dim mAdvAmt    As Double
        Dim mDrAdvAmt    As Double
        Dim mfltAmount  As Double
        Dim mRoundOff   As Double
        Dim mTrChild As uTrChild
        Dim mSL     As Integer
        Dim arrInput As Variant
        Dim mCount As Integer
        Dim mFineAmt    As Double
        Dim mCrAmt      As Double
        Dim mCrPAmt      As Double
        Dim mPAmt       As Double
        Dim mAdvHeadID  As Integer
        
        mfltAmount = val(txtTotal.Text)
        mRoundOff = val(txtRoundOff)
        mAdvAmt = val(txtAdvance.Text)
        mDrAdvAmt = mAdvAmt
        objAcc.SetAccountCode (mAdvHeadCode)
        mAdvHeadID = objAcc.AccountHeadID
        If mAdvAmt > 0 Then
            mSL = 2
            For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mCount, 0) <> "" And vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodePenalInterest Then
                    mFineAmt = vsGrid.TextMatrix(mCount, 11)
                    If mAdvAmt >= mFineAmt Then
                        mAdvAmt = mAdvAmt - mFineAmt
                        mCrPAmt = mFineAmt   'To Adj againt AdvanceHead
                        mFineAmt = 0                               'To Adj againt Cash/Bank
                    Else
                        mFineAmt = mFineAmt - mAdvAmt   'To Adj againt Cash/Bank
                        mCrPAmt = mAdvAmt    'To Adj againt AdvanceHead
                        mAdvAmt = 0
                    End If
                End If
            Next
            'Penal Amount Adjust with Advance ,Cr against Advance Head
            If mCrPAmt > 0 Then
                mCrAmt = mCrAmt + mCrPAmt  'Calculation Credit Amount Against Advance Head
                mSL = 3
                With mTrChild
                    .intTransactionID = mTID
                    .intSerialNo = mSL
                    .intAccountHeadID = gbAcHeadIDPenalInterest
                    .fltAmount = mCrPAmt
                    .tinDebitOrCreditFlag = 0
                    .intByAccountHeadID = mAdvHeadID
                    .vchNarration = "Total Advance Collection Adjusted For Penal"
                    .intFundID = 1
                    
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                End With
             End If
             For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mCount, 0) <> "" And _
                    (vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Or _
                    vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent Or _
                    vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear Or _
                    vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent) Then
                    If mCount + 1 < vsGrid.Rows Then
                        If vsGrid.TextMatrix(mCount + 1, 0) <> "" And vsGrid.TextMatrix(mCount + 1, 0) = gbAcHeadCodeServiceTax Then
                            mPAmt = val(vsGrid.TextMatrix(mCount, 11)) + val(vsGrid.TextMatrix(mCount + 1, 11))
                        End If
'                        If vsGrid.TextMatrix(mCount + 1, 0) <> "" And vsGrid.TextMatrix(mCount + 1, 0) = gbAcHeadCodeCGST Then
'                            mPAmt = val(vsGrid.TextMatrix(mCount, 11)) + val(vsGrid.TextMatrix(mCount + 1, 11))
'                        End If
                    End If
                    If mCount + 2 < vsGrid.Rows Then
'                        If vsGrid.TextMatrix(mCount + 2, 0) <> "" And vsGrid.TextMatrix(mCount + 2, 0) = gbAcHeadCodeSGST Then
'                            mPAmt = val(vsGrid.TextMatrix(mCount, 11)) + val(vsGrid.TextMatrix(mCount + 2, 11))
'                        End If
                    End If
                    ' Checking Advance Amount with Rent+Service Tax
                    ' Advance>=Rent+Service
                    If mAdvAmt >= mPAmt Then
                        mAdvAmt = mAdvAmt - mPAmt
                        mSL = mSL + 1
                        With mTrChild
                            'Adv Adjust with Rent
                            .intTransactionID = mTID
                            .intSerialNo = mSL
                            If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Then
                                .intAccountHeadID = gbAcHeadIDCivicAmenitiesArrear
                            ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent Then
                                .intAccountHeadID = gbAcHeadIDCivicAmenitiesCurrent
                            ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear Then
                                .intAccountHeadID = gbAcHeadIDRentLandArrear
                            ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent Then
                                .intAccountHeadID = gbAcHeadIDRentLandCurrent
                            End If
                            .fltAmount = vsGrid.TextMatrix(mCount, 11)
                            .tinDebitOrCreditFlag = 0
                            .intByAccountHeadID = mAdvHeadID
                            .vchNarration = "Advance Collection Adjusted For Rent"
                            .intFundID = 1
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                            
                            'AdvAdj for Service tax
                            mSL = mSL + 1
                            
                            .intSerialNo = mSL
                            .intAccountHeadID = gbAcHeadIDServiceTax
                            .fltAmount = vsGrid.TextMatrix(mCount + 1, 11)
                            .vchNarration = "Advance Collection Adjusted For Service tax"
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                        End With
                    Else
                        Dim mRent           As Double
                        Dim mSTax           As Double
                        Dim mAdvAdjRent     As Double
                        Dim mAdvAdjSTax     As Double
                        If mAdvAmt > 0 Then
                            mPAmt = mPAmt - mAdvAmt
                            mRent = Format(mAdvAmt * 100 / 110.3, "0.00")
                            mAdvAdjRent = mAdvAmt    'For Credit posting againt Cash(Remaining amount After substraction From Advance)
                            mSTax = mAdvAmt - mRent
                            mAdvAmt = 0
                            mSL = mSL + 1
                            With mTrChild
                                .intTransactionID = mTID
                                .intSerialNo = mSL
                                If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesCurrent
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear Then
                                    .intAccountHeadID = gbAcHeadIDRentLandArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent Then
                                    .intAccountHeadID = gbAcHeadIDRentLandCurrent
                                End If
                                .fltAmount = mRent
                                .tinDebitOrCreditFlag = 0
                                .intByAccountHeadID = mAdvHeadID
                                .vchNarration = "Total Advance Collection Adjusted For Rent"
                                .intFundID = 1
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                
                                'Service tax Advance Adjust
                                mSL = mSL + 1
                                .intSerialNo = mSL
                                .intAccountHeadID = gbAcHeadIDServiceTax
                                .fltAmount = mSTax
                                .vchNarration = "Advance Collection Adjusted For Service tax"
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                
                                
                                '----- Set Off Remaining Rent+S tax With Cash/Bank
                                mAdvAdjSTax = Format(mAdvAdjRent * 10.3 / 100, "0.00")
                                mAdvAdjRent = mAdvAdjRent - mAdvAdjSTax
                                mSL = mSL + 1
                                .intTransactionID = mTID
                                .intSerialNo = mSL
                                If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesCurrent
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear Then
                                    .intAccountHeadID = gbAcHeadIDRentLandArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent Then
                                    .intAccountHeadID = gbAcHeadIDRentLandCurrent
                                End If
                                .fltAmount = val(vsGrid.TextMatrix(mCount, 11)) - mRent
                                .tinDebitOrCreditFlag = 0
                                .intByAccountHeadID = mDrAccountHeadID
                                .vchNarration = ""
                                .intFundID = 1
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                
                                'Service tax with cash/bank
                                mSL = mSL + 1
                                .intSerialNo = mSL
                                .intAccountHeadID = gbAcHeadIDServiceTax
                                .fltAmount = val(vsGrid.TextMatrix(mCount + 1, 11)) - mSTax
                                .vchNarration = ""
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                            End With
                        Else
                            'Accounting Against Cash Head
                            mSL = mSL + 1
                            With mTrChild
                                .intTransactionID = mTID
                                .intSerialNo = mSL
                                If vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesArrear Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeCivicAmenitiesCurrent Then
                                    .intAccountHeadID = gbAcHeadIDCivicAmenitiesCurrent
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandArrear Then
                                    .intAccountHeadID = gbAcHeadIDRentLandArrear
                                ElseIf vsGrid.TextMatrix(mCount, 0) = gbAcHeadCodeRentLandCurrent Then
                                    .intAccountHeadID = gbAcHeadIDRentLandCurrent
                                End If
                                .fltAmount = vsGrid.TextMatrix(mCount, 11)
                                .tinDebitOrCreditFlag = 0
                                .intByAccountHeadID = mDrAccountHeadID
                                .vchNarration = ""
                                .intFundID = 1
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                                
                                'Service tax
                                mSL = mSL + 1
                                .intSerialNo = mSL
                                .intAccountHeadID = gbAcHeadIDServiceTax
                                .fltAmount = vsGrid.TextMatrix(mCount + 1, 11)
                                .vchNarration = ""
                                
                                arrInput = Array(.intTransactionID, _
                                .intSerialNo, _
                                .intAccountHeadID, _
                                .fltAmount, _
                                .tinDebitOrCreditFlag, _
                                .intByAccountHeadID, _
                                .vchNarration, _
                                .intFundID)
                                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                            End With
                        End If
                    End If
                End If
             Next
             '  Remaining Penal Interest After adv Adjusted Set off
             If mFineAmt > 0 Then
                   With mTrChild
                       mSL = mSL + 1
                       .intTransactionID = mTID
                       .intSerialNo = mSL
                       .intAccountHeadID = gbAcHeadIDPenalInterest
                       .fltAmount = mFineAmt
                       .tinDebitOrCreditFlag = 0
                       .intByAccountHeadID = mDrAccountHeadID
                       .vchNarration = Null
                       .intFundID = 1
                       
                       arrInput = Array(.intTransactionID, _
                       .intSerialNo, _
                       .intAccountHeadID, _
                       .fltAmount, _
                       .tinDebitOrCreditFlag, _
                       .intByAccountHeadID, _
                       .vchNarration, _
                       .intFundID)
                       objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                   End With
                End If
                'NOTE:-Round of Adjustment adjusted Cr with cash/Bank
                If mRoundOff > 0 Then
                   With mTrChild
                       mSL = mSL + 1
                       .intTransactionID = mTID
                       .intSerialNo = mSL
                       .intAccountHeadID = gbAcHeadIDRoundOff
                       .fltAmount = mRoundOff
                       .tinDebitOrCreditFlag = 0
                       .intByAccountHeadID = mDrAccountHeadID
                       .vchNarration = "Round Of Adjustment"
                       .intFundID = 1
                       
                       arrInput = Array(.intTransactionID, _
                       .intSerialNo, _
                       .intAccountHeadID, _
                       .fltAmount, _
                       .tinDebitOrCreditFlag, _
                       .intByAccountHeadID, _
                       .vchNarration, _
                       .intFundID)
                       objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                   End With
                End If
'             Advance Head Debit Post
                 With mTrChild
                    mSL = 2
                    .intTransactionID = mTID
                    .intSerialNo = mSL
                    .intAccountHeadID = mAdvHeadID
                    .fltAmount = mDrAdvAmt
                    .tinDebitOrCreditFlag = 1
                    .intByAccountHeadID = Null
                    .vchNarration = ""
                    .intFundID = 1
                    
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                    
    '                Debit Cash/bank posting
                    mSL = 1
                    .intTransactionID = mTID
                    .intSerialNo = mSL
                    .intAccountHeadID = mDrAccountHeadID
                    .fltAmount = mfltAmount
                    .tinDebitOrCreditFlag = 1
                    .intByAccountHeadID = Null
                    .vchNarration = ""
                    .intFundID = 1
                    
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                End With
        End If
    End Function
    Private Function cashBankValidateDr() As Boolean
    ''' Codded ON 4.7.12 By Anisha
    ''' Validation of Blocking Cash/Bank Heads For Jv
        Dim mSql As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mAccID      As Integer
        Dim mGAcc       As Integer
        Dim mCnt        As Integer
        mAccID = val(txtAccountHead.Tag)
        mSql = "Select * From faAccountHeads Where intGroupId not in(1,2) And intAccountHeadID= " & mAccID
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                cashBankValidateDr = True
                Exit Function
'                While Not (Rec.EOF)
'                    If mAccID = Rec!intAccountHeadID Then
'                        cashBankValidate = True
'                        Exit Function
'                    End If
''                    If txtAccountHeadCode.Text = Rec!vchAccountHeadCode Then
''                        cashBankValidate = True
''                        Exit Function
''                    End If
''                    If vsGrid.FindRow(Rec!vchAccountHeadCode, 1, 1, 1, 1) > 0 Then
''                         cashBankValidate = True
''                         Exit Function
''                    End If
'                    Rec.MoveNext
'                Wend
            End If
        End If
    End Function
    Private Function cashBankValidateCr() As Boolean
    ''' Codded ON 4.7.12 By Anisha
    ''' Validation of Blocking Cash/Bank Heads For Jv
        Dim mSql As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mAccID      As Integer
        Dim mGAcc       As Integer
        Dim mCnt        As Integer
        mAccID = val(txtAccountHead.Tag)
        mSql = "Select * From faAccountHeads Where intGroupId in(1,2) Order By intAccountHeadID Asc"
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                ' cashBankValidateCr = True
                While Not (Rec.EOF)
                    If vsGrid.FindRow(Rec!vchAccountHeadCode, 1, 0, 1, 1) > 0 Then
                         cashBankValidateCr = True
                         Exit Function
                    End If
                    Rec.MoveNext
                Wend
            End If
        End If
    End Function
    Private Function CheckInterruptedNoSuffixExists(mChar) As Boolean
        ''-----Added On 28/Jun/2011 By Anisha-------
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mSql As String
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select * From faVouchers Where intVoucherNo=" & txtReceiptNo.Text & "And vchDoorNoP3='" & mChar & "'"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            CheckInterruptedNoSuffixExists = True
        Else
            CheckInterruptedNoSuffixExists = False
        End If
    End Function
    Private Function UpdateIRRegister(mintVoucherID_1 As Variant, mdtDate As Date, mfltAmount_9 As Double)
        '-------------------------------'
        ' ADDED BY MINU FOR IR REGISTER '
        '-------------------------------'
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mVoucherNo As String
        Dim mSuffix As String
        Dim mVRNum As String
        
        mVRNum = CStr(mInterruptedRegisterReceiptNo)
        mVoucherNo = Token(mVRNum, " ")
        mSuffix = mVRNum

        If InterruptedRegister = 1 Then 'Cancel
            mSql = "Update faInterruptedRegister set intVoucherID= " & mintVoucherID_1 & ",dtVoucherDate=' " & DdMmmYy(mdtDate) & "',fltAmount=" & mfltAmount_9 & ",tnyStatus=4 "
            'mSQL = mSQL + " Where intReceiptNo=" & mVoucherNo & "         "
            mSql = mSql + " Where intID=" & mInterruptedRegisterID
            
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            'Call GetNextIRNumber
            'Call FormInitialize
            
            Call FormInitialize
    Call GetNextIRNumber
           

        ElseIf InterruptedRegister = 2 Then 'Edit
            mSql = "Update faInterruptedRegister set intVoucherID= " & mintVoucherID_1 & ",fltAmount=" & mfltAmount_9 & ",tnyStatus=4,tnyFlag=0 "
'            mSQL = mSQL + " Where intReceiptNo=" & val(mVoucherNo) & ""
'            If mSuffix <> "" Then
'            mSQL = mSQL + " And vchSuffix= '" & mSuffix & "' """
'            End If
            mSql = mSql + " Where intVoucherID=" & mVoucherID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            Unload Me
        ElseIf InterruptedRegister = 3 Then 'Insert Suffix
            mSql = "Update faInterruptedRegister set intVoucherID= " & mintVoucherID_1 & ",dtVoucherDate=' " & DdMmmYy(mdtDate) & "',fltAmount=" & mfltAmount_9 & ",tnyStatus=4 "
            'mSQL = mSQL + " Where intReceiptNo=" & mVoucherNo & " And vchSuffix= '" & txtIntruptNoSuffix.Text & "' "   'vchSuffix= ' " & txtIntruptNoSuffix.Text & "' ,
            mSql = mSql + " Where intID=" & mInterruptedRegisterID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            Unload Me
        End If
        
    End Function
    
    Private Sub GetNextIRNumber()
        Dim mSql As String
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        mSql = " SELECT Top 1 * FROM faInterruptedRegister "
        mSql = mSql + " Where intBookID = " & mIRBookID & " And IsNull(intVoucherID, 0) = 0 And IsNull(tnyCancelled, 0) = 0 Order by intReceiptNo"
        objdb.SetConnection mCnn
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            mInterruptedRegisterReceiptNo = Rec!intReceiptNo
            txtReceiptNo.Text = Rec!intReceiptNo
            mInterruptedRegisterID = Rec!intID
            If Rec!vchSuffix <> "" Then
                txtIntruptNoSuffix.Text = Rec!vchSuffix
            Else
                txtIntruptNoSuffix.Text = ""
            End If
        Else
            mSql = " There is no Interrupted Register left open"
            MsgBox mSql, vbInformation
            cmdNew.Enabled = False
            cmdSave.Enabled = False
            'Unload Me
        End If
        Rec.Close
        
        If IsDate(mInterruptedRegisterReceiptDate) Then
            If CDate(mInterruptedRegisterReceiptDate) < CDate(gbStartingDate) And CDate(mInterruptedRegisterReceiptDate) < CDate(gbEndingDate) Then
                mPreviousYearMode = 1
                'mYearID = gbFinancialYearID - 1
            Else
                mPreviousYearMode = 0
                'mYearID = gbFinancialYearID
            End If
        End If
        
    End Sub
    
    Public Property Get InterruptedMode() As Boolean
        InterruptedMode = mInterruptedModeFlag
    End Property
    
    Public Property Get InterruptedModeSoochika() As Boolean
        InterruptedModeSoochika = mInterruptedModeSoochikaFlag
    End Property

    Public Property Let InterruptedModeSoochika(mData As Boolean)
        mInterruptedModeSoochikaFlag = mData
    End Property

    
    Public Property Get InterruptEditMode() As Boolean
        InterruptEditMode = mInterruptEditMode
    End Property
    
    Public Property Let InterruptEditMode(mData As Boolean)
        mInterruptEditMode = mData
    End Property
    
    Public Property Let SubLedgerID(mSubLedgerID As Variant)
        mvarSubLedgerID = mSubLedgerID
    End Property
    
    Public Property Get SubLedgerID() As Variant
        SubLedgerID = mvarSubLedgerID
    End Property
    
    Public Property Let PRPReprintFlag(RePrint As Integer)
        mRePrintFlag = RePrint
    End Property
    Public Property Let DemandBasedFlag(mFlag As Boolean)
        mvarDemandBasedFlag = mFlag
    End Property
    
    '------------------------------------------------------------'
    '                   Added On 24/04/2009                      '
    '           By Cijith Sreedharan For KMBR Integration        '
    '------------------------------------------------------------'
    Public Property Let PermitType(mData As Integer)
        mPermitType = mData
    End Property
    
    Public Property Let BuildingType(mData As Double)
        mBuildingType = mData
    End Property
    
    Public Property Let KMBRAccess(mData As Integer)
        mKMBRAccess = mData
    End Property
    
    Public Property Let BuildingWard(mData As Double)
        mBuildingWard = mData
    End Property
    
    Public Property Let AssessmentYear(mData As Integer)
        mAssesmentYearID = mData
    End Property
    
    Public Property Let SoochikaConnected(mData As Boolean)
        mSoochikaConnected = mData
    End Property
    Public Property Let PoorHomeCess(mData As Boolean)
        mPoorHomeCess = mData
    End Property
    Public Property Let ZonalCollection(mData As Integer)  'Added by Sunil Babu
        mZonal = mData
    End Property
     Public Property Get InterruptedRegister() As Variant   'Added by Minu For IR
        InterruptedRegister = mInterruptedRegister
    End Property
      Public Property Let InterruptedRegister(mData As Variant) 'Added by Minu For IR
        mInterruptedRegister = mData
    End Property
    Public Property Get InterruptedRegisterReceiptNo() As Variant 'Added by Minu For IR
        InterruptedRegisterReceiptNo = mInterruptedRegisterReceiptNo
    End Property
      Public Property Let InterruptedRegisterReceiptNo(mData As Variant) 'Added by Minu For IR
        mInterruptedRegisterReceiptNo = mData
    End Property
    Public Property Get InterruptedRegisterReceiptDate() As Variant 'Added by Minu For IR
        InterruptedRegisterReceiptDate = mInterruptedRegisterReceiptDate
    End Property
      Public Property Let InterruptedRegisterReceiptDate(mData As Variant) 'Added by Minu For IR
        mInterruptedRegisterReceiptDate = mData
    End Property

    Public Property Get IRBookID() As Variant          'Added by Minu For IR
        IRBookID = mIRBookID
    End Property
      Public Property Let IRBookID(mData As Variant)   'Added by Minu For IR
        mIRBookID = mData
    End Property
    
    'paperless
Private Sub SaveSoochikaInwardTrackDetails(mCnn As ADODB.Connection, FID As Variant)
    Dim arrIn As Variant
    Dim ForwardTo As Double
    Dim Rec As New ADODB.Recordset
    ReDim arrIn(9)
    Dim objdb As New clsDB
    
'     Dim mCnnSoochikaMas As New ADODB.Connection
'            If objdb.CreateNewConnection(mCnnSoochikaMas, enuSourceString.DBMaster) = True Then
'            mCnnSoochikaMas.BeginTrans
'            End If
               
    
    arrIn(0) = FID
    ForwardTo = CDbl(mSeatPrefix + CStr(cmbSeat.ItemData(cmbSeat.ListIndex)))
    arrIn(1) = ForwardTo
    arrIn(2) = gbSeatID
    Set Rec = mCnn.Execute("SpSelectSeatDetails " & ForwardTo)
    If Not (Rec.EOF Or Rec.BOF) Then
        arrIn(3) = IIf(IsNull(Rec!numCurrentUserID), 0, Rec!numCurrentUserID)
    Else
        arrIn(3) = Null
    End If
    Rec.Close
    arrIn(4) = gbUserID
    arrIn(5) = "Processing"
    arrIn(6) = Null
    arrIn(7) = Null
    arrIn(8) = 0  'Changed by Renjitha on 29.02.2012 Form 1 to 0
    arrIn(9) = Null
    
    objdb.ExecuteSP "SpSaveInwardTrackDetails", arrIn, , , mCnn, adCmdStoredProc
End Sub
  Private Function GetLastReconDate(intBankID As Integer) As Variant
  ''Added Anisha On 2/Jul/2014
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mMonthID As Integer
        Dim mFinYear As Integer
               
        mSql = "Select * From faBanks Where intAccountHeadID=" & intBankID
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                GetLastReconDate = IIf(IsNull(Rec!dtReconEndDate), Null, Rec!dtReconEndDate)
            End If
        Rec.Close
 End Function
Private Function PTaxWebDemand(mVoucherID As Long)

'''For Interrupted Receipt Edit . P tax cancellation

                    Dim Rec             As New ADODB.Recordset
                    Dim mCollPost       As String
                    Dim mColZoneID      As String
                    Dim mBuildingIdWeb  As String
                    Dim mColAmt            As String
                    Dim mColDate        As String
                    Dim mColReceiptNo   As String
                    Dim mColBookNo      As String
                    Dim mColPeriodId     As String
                    Dim mColYearID       As String
                    Dim mHash           As String
                    Dim mCollOut        As String
'                    Dim node            As IXMLDOMNode
'                    Dim DataNodes       As IXMLDOMNodeList
                    Dim mUrl            As String
                    Dim objSOAP         As Variant
                    Dim mLen            As Integer
                    Dim mColAccID       As String
                    Dim mColKeyID       As String
                    'Dim Rec             As New ADODB.Recordset
                mUrl = gbDefaultUrlSanchayaPost
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
          
                    Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                    If Not (Rec.EOF And Rec.BOF) Then
                        While Not Rec.EOF
                            
                            mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                            mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                            mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
                            mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                            mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                            mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                            mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                            mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                            If mColAccID <> gbAcHeadIDPenalInterest Then
                                mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
                                mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mVoucherID) + "#"
                                mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
                                mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
                                mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
                            End If
                            Rec.MoveNext
                            mCollPost = mCollPost + "~"
                        Wend
                        mLen = Len(mCollPost) - 1
                        mCollPost = Left$(mCollPost, mLen - 1)
                        mHash = CStr(mVoucherID) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
                        mCollOut = objSOAP.Saankhyaa_CollectionPostingCancel(mCollPost, mHash)
                    End If
  
    End Function

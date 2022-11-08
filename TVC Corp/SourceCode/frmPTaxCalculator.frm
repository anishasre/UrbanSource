VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmPTaxCalculator 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tax Calculator"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFeeOnSplservice 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Fee On Special Service"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4050
      TabIndex        =   74
      Top             =   3915
      Width           =   4170
   End
   Begin VB.CheckBox chkCentralGovBulding 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Central Gov Building Service Charge"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   73
      Top             =   3870
      Width           =   4035
   End
   Begin VB.CheckBox chkServiceCess 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Service Cess "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   72
      Top             =   855
      Width           =   3885
   End
   Begin VB.CheckBox chkSurcharge 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Surcharge"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   71
      Top             =   855
      Width           =   4245
   End
   Begin VB.Frame frmCentralgov 
      BackColor       =   &H00DAF2F2&
      Height          =   1410
      Left            =   45
      TabIndex        =   56
      Top             =   4320
      Width           =   4020
      Begin VB.TextBox txtToYrGov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         MaxLength       =   4
         TabIndex        =   17
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtToPeriodGov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         MaxLength       =   1
         TabIndex        =   18
         Top             =   675
         Width           =   450
      End
      Begin VB.TextBox txtGovPenal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   495
         TabIndex        =   58
         Top             =   1080
         Width           =   1440
      End
      Begin VB.TextBox txtGovTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   57
         Top             =   1080
         Width           =   1425
      End
      Begin VB.TextBox txtGov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         MaxLength       =   15
         TabIndex        =   14
         Top             =   225
         Width           =   1335
      End
      Begin VB.OptionButton optAmtGov 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   225
         Width           =   855
      End
      Begin VB.OptionButton optRateGov 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   12
         Top             =   225
         Width           =   810
      End
      Begin VB.TextBox txtFromPeriodGov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         MaxLength       =   1
         TabIndex        =   16
         Top             =   675
         Width           =   450
      End
      Begin VB.TextBox txtFromYrGov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   495
         MaxLength       =   4
         TabIndex        =   15
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label5 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   3330
         TabIndex        =   76
         Top             =   675
         Width           =   120
      End
      Begin VB.Label Label19 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Penal"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   64
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label18 
         BackColor       =   &H00DAF2F2&
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
         Height          =   225
         Left            =   2070
         TabIndex        =   63
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label17 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   2115
         TabIndex        =   62
         Top             =   1260
         Width           =   15
      End
      Begin VB.Label Label16 
         BackColor       =   &H00DAF2F2&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   61
         Top             =   675
         Width           =   195
      End
      Begin VB.Label Label15 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   1395
         TabIndex        =   60
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label14 
         BackColor       =   &H00DAF2F2&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   59
         Top             =   675
         Width           =   435
      End
   End
   Begin VB.Frame frmFeeOnSpclservice 
      BackColor       =   &H00DAF2F2&
      Height          =   1455
      Left            =   4095
      TabIndex        =   38
      Top             =   4275
      Width           =   4110
      Begin VB.TextBox txtFromYrSpl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   495
         MaxLength       =   4
         TabIndex        =   8
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtFromPeriodSpl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         MaxLength       =   1
         TabIndex        =   9
         Top             =   675
         Width           =   450
      End
      Begin VB.TextBox txtToYrSpl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2475
         MaxLength       =   4
         TabIndex        =   10
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtToPeriodSpl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3510
         MaxLength       =   1
         TabIndex        =   11
         Top             =   675
         Width           =   450
      End
      Begin VB.OptionButton optRateSpl 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   225
         TabIndex        =   5
         Top             =   225
         Width           =   810
      End
      Begin VB.OptionButton optAmtspl 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1035
         TabIndex        =   6
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox txtAmntspl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2115
         MaxLength       =   15
         TabIndex        =   7
         Top             =   225
         Width           =   1335
      End
      Begin VB.TextBox txtTotalSpl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2475
         TabIndex        =   40
         Top             =   1035
         Width           =   1425
      End
      Begin VB.TextBox txPenalSpl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   495
         TabIndex        =   39
         Top             =   1035
         Width           =   1440
      End
      Begin VB.Label Label27 
         BackColor       =   &H00DAF2F2&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   55
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label26 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   1395
         TabIndex        =   54
         Top             =   675
         Width           =   120
      End
      Begin VB.Label Label25 
         BackColor       =   &H00DAF2F2&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   53
         Top             =   675
         Width           =   195
      End
      Begin VB.Label Label11 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   3375
         TabIndex        =   52
         Top             =   675
         Width           =   120
      End
      Begin VB.Label Label20 
         BackColor       =   &H00DAF2F2&
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
         Height          =   225
         Left            =   2070
         TabIndex        =   42
         Top             =   1035
         Width           =   510
      End
      Begin VB.Label Label21 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Penal"
         Height          =   225
         Left            =   45
         TabIndex        =   41
         Top             =   1035
         Width           =   420
      End
   End
   Begin VB.Frame frmSurcharge 
      BackColor       =   &H00DAF2F2&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   4050
      TabIndex        =   35
      Top             =   1215
      Width           =   4155
      Begin VB.TextBox txtToYrsur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2430
         MaxLength       =   4
         TabIndex        =   22
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtFromYrSur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   495
         MaxLength       =   4
         TabIndex        =   20
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtSurRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2700
         MaxLength       =   15
         TabIndex        =   19
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox txtToPeriodSur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3555
         MaxLength       =   1
         TabIndex        =   23
         Top             =   540
         Width           =   480
      End
      Begin VB.TextBox txtFromPeriodSur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   21
         Top             =   540
         Width           =   480
      End
      Begin VB.TextBox txtSurTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2475
         TabIndex        =   44
         Top             =   2340
         Width           =   1470
      End
      Begin VB.TextBox txtPenalSur 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   43
         Top             =   2340
         Width           =   1470
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00DAF2F2&
         Height          =   1320
         Left            =   315
         TabIndex        =   36
         Top             =   765
         Width           =   3705
         Begin VSFlex8LCtl.VSFlexGrid vsGridSurcharge 
            Height          =   1140
            Left            =   90
            TabIndex        =   37
            Top             =   135
            Width           =   3435
            _cx             =   6059
            _cy             =   2011
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
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPTaxCalculator.frx":0000
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
      End
      Begin VB.Label Label28 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Rate"
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
         Left            =   2205
         TabIndex        =   69
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label24 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   3465
         TabIndex        =   50
         Top             =   540
         Width           =   120
      End
      Begin VB.Label Label8 
         BackColor       =   &H00DAF2F2&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   49
         Top             =   540
         Width           =   480
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1530
         TabIndex        =   48
         Top             =   540
         Width           =   120
      End
      Begin VB.Label Label10 
         BackColor       =   &H00DAF2F2&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   47
         Top             =   540
         Width           =   195
      End
      Begin VB.Label Label23 
         BackColor       =   &H00DAF2F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   225
         Left            =   2070
         TabIndex        =   46
         Top             =   2340
         Width           =   375
      End
      Begin VB.Label Label22 
         BackColor       =   &H00DAF2F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Penal"
         Height          =   225
         Left            =   90
         TabIndex        =   45
         Top             =   2340
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdGenarateDemnd 
      Caption         =   "Generate Demand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3105
      TabIndex        =   34
      Top             =   5760
      Width           =   2070
   End
   Begin VB.Frame frmServiceCess 
      BackColor       =   &H00DAF2F2&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   45
      TabIndex        =   29
      Top             =   1215
      Width           =   3975
      Begin VB.Frame Frame2 
         BackColor       =   &H00DAF2F2&
         Height          =   1410
         Left            =   180
         TabIndex        =   75
         Top             =   585
         Width           =   3705
         Begin VSFlex8LCtl.VSFlexGrid vsGridSCess 
            Height          =   1005
            Left            =   225
            TabIndex        =   77
            Top             =   225
            Width           =   3390
            _cx             =   5980
            _cy             =   1773
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
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPTaxCalculator.frx":0065
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
      End
      Begin VB.TextBox txtPenalServiceCess 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   495
         TabIndex        =   66
         Top             =   2250
         Width           =   1470
      End
      Begin VB.TextBox txtTotalserviceCess 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2430
         TabIndex        =   65
         Top             =   2250
         Width           =   1470
      End
      Begin VB.CheckBox chkWaive 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Waive"
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
         Left            =   2880
         TabIndex        =   51
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox txtpenalSurcharge 
         Height          =   330
         Left            =   7965
         TabIndex        =   33
         Top             =   3240
         Width           =   1740
      End
      Begin VB.TextBox txtTotalsurcharge 
         Height          =   330
         Left            =   7965
         TabIndex        =   31
         Top             =   2880
         Width           =   1740
      End
      Begin VB.Label Label7 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Penal"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   68
         Top             =   2250
         Width           =   420
      End
      Begin VB.Label Label6 
         BackColor       =   &H00DAF2F2&
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
         Height          =   225
         Left            =   1980
         TabIndex        =   67
         Top             =   2250
         Width           =   420
      End
      Begin VB.Label Label13 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Penal"
         Height          =   225
         Left            =   7380
         TabIndex        =   32
         Top             =   3285
         Width           =   510
      End
      Begin VB.Label Label12 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Total"
         Height          =   225
         Left            =   7380
         TabIndex        =   30
         Top             =   2925
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAF2F2&
      Height          =   705
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   8175
      Begin VB.TextBox txtToYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtToPeriodID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7245
         MaxLength       =   1
         TabIndex        =   24
         Top             =   180
         Width           =   480
      End
      Begin VB.TextBox txtFromPeriodID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4860
         MaxLength       =   1
         TabIndex        =   3
         Top             =   180
         Width           =   480
      End
      Begin VB.TextBox txtFromYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3735
         MaxLength       =   4
         TabIndex        =   2
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtHalfYrTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label29 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   7155
         TabIndex        =   70
         Top             =   225
         Width           =   120
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DAF2F2&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5715
         TabIndex        =   28
         Top             =   225
         Width           =   195
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DAF2F2&
         Caption         =   "-"
         Height          =   435
         Left            =   4770
         TabIndex        =   27
         Top             =   225
         Width           =   120
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DAF2F2&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3105
         TabIndex        =   26
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Half Year Tax Rate"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   180
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmPTaxCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mHalfYears As String
    Dim mNoOfHalfYears As Integer
    Dim mNoOfHalfYearsSur As Integer
    Dim mMode As Variant   '5 = PTax Calculator
    Dim ToatalPenal As Double
    
    Dim dSTotal As Double
    Dim dSurTotal As Double
    Dim dCentralGovTotal As Double
    Dim dSplFeeTotal As Double
    
    
    Private bIsRightClk As Boolean
Private Sub chkCentralGovBulding_Click()
    txtGovPenal.Enabled = False
    txtGovTotal.Enabled = False
    txtGov.Enabled = False
    txtFromYrGov.Text = txtFromYear.Text
    txtFromPeriodGov.Text = txtFromPeriodID.Text
    txtToYrGov.Text = txtToYear.Text
    txtToPeriodGov.Text = txtToPeriodID.Text
         If chkCentralGovBulding.Value = 1 Then
            chkServiceCess.Enabled = False
            chkSurcharge.Enabled = False
            Call AmntValidation
              frmCentralgov.Enabled = True
         Else
            chkServiceCess.Enabled = True
            chkSurcharge.Enabled = True
             frmCentralgov.Enabled = True
             optAmtGov.Value = False
             optRateGov.Value = False
             txtGov.Text = ""
             txtFromYrGov.Text = ""
             txtFromPeriodGov.Text = ""
             txtToYrGov.Text = ""
             txtToPeriodGov.Text = ""
             txtGovPenal.Text = ""
             txtGovTotal.Text = ""
        End If
    
End Sub

Private Sub chkFeeOnSplservice_Click()
    Call AmntValidation
    txPenalSpl.Enabled = False
    txtTotalSpl.Enabled = False
    txtFromYrSpl.Text = txtFromYear.Text
    txtFromPeriodSpl.Text = txtFromPeriodID.Text
    txtToYrSpl.Text = txtToYear.Text
    txtToPeriodSpl.Text = txtToPeriodID.Text
    If chkFeeOnSplservice.Value = 1 Then
        frmFeeOnSpclservice.Enabled = True
    Else
        frmFeeOnSpclservice.Enabled = False
        optAmtspl.Value = False
        optRateSpl.Value = False
        txtAmntspl.Text = ""
        txtFromYrSpl.Text = ""
        txtFromPeriodSpl.Text = ""
        txtToYrSpl.Text = ""
        txtToPeriodSpl.Text = ""
        txPenalSpl.Text = ""
        txtTotalSpl.Text = ""
    End If
End Sub
Private Sub chkNonResidential_Click()
' If chkNonResidential.value = 1 Then
'    chkResidential.value = vbUnchecked
'        If chkServiceCess.value = 1 Then
'            chkServiceCess.value = vbUnchecked
'        End If
'        If chkSurcharge.value = 1 Then
'            chkSurcharge.value = vbUnchecked
'        End If
'        If chkCentralGovBulding.value = 1 Then
'            chkCentralGovBulding.value = vbUnchecked
'        End If
'        If chkFeeOnSplservice.value = 1 Then
'            chkFeeOnSplservice.value = vbUnchecked
'        End If
' Else
'       If chkServiceCess.value = 1 Then
'            chkServiceCess.value = vbUnchecked
'        End If
'        If chkSurcharge.value = 1 Then
'            chkSurcharge.value = vbUnchecked
'        End If
'        If chkCentralGovBulding.value = 1 Then
'            chkCentralGovBulding.value = vbUnchecked
'        End If
'        If chkFeeOnSplservice.value = 1 Then
'            chkFeeOnSplservice.value = vbUnchecked
'        End If
'
' End If
End Sub

'Private Sub chkResidential_Click()
'    If chkResidential.value = 1 Then
'        chkNonResidential.value = vbUnchecked
'        If chkServiceCess.value = 1 Then
'            chkServiceCess.value = vbUnchecked
'        End If
'        If chkSurcharge.value = 1 Then
'            chkSurcharge.value = vbUnchecked
'        End If
'        If chkCentralGovBulding.value = 1 Then
'            chkCentralGovBulding.value = vbUnchecked
'        End If
'        If chkFeeOnSplservice.value = 1 Then
'            chkFeeOnSplservice.value = vbUnchecked
'        End If
'    Else
'
'        If chkServiceCess.value = 1 Then
'            chkServiceCess.value = vbUnchecked
'        End If
'        If chkSurcharge.value = 1 Then
'            chkSurcharge.value = vbUnchecked
'        End If
'        If chkCentralGovBulding.value = 1 Then
'            chkCentralGovBulding.value = vbUnchecked
'        End If
'        If chkFeeOnSplservice.value = 1 Then
'            chkFeeOnSplservice.value = vbUnchecked
'        End If
'    End If
'
'End Sub
Private Sub chkServiceCess_Click()
         Dim dSani As Double
        Dim dWatr As Double
        Dim dStreet As Double
        Dim dDraing As Double
       ' Dim dSTotal As Double
        txtPenalServiceCess.Enabled = False
        txtTotalserviceCess.Enabled = False
        If txtHalfYrTax.Text = "" Then
            MsgBox "Enter the Half Year Tax RAte", vbInformation
            txtHalfYrTax.SetFocus
        Else
          If chkServiceCess.Value = 1 Then
                     
                    chkCentralGovBulding.Enabled = False
                      Call VsgridChecked
                      vsGridSCess.Editable = flexEDNone
                      Dim mHalfYears As String
                      frmServiceCess.Enabled = True
                      mHalfYears = txtHalfYrTax.Text
                      ' flexgrid'
                          dSani = mHalfYears * 4 / 100
                          dWatr = mHalfYears * 3 / 100
                          dStreet = mHalfYears * 2 / 100
                          dDraing = mHalfYears * 1 / 100
                          dSTotal = dSani + dWatr + dStreet + dDraing
                          
                          vsGridSCess.TextMatrix(0, 0) = mHalfYears * 4 / 100
                          vsGridSCess.TextMatrix(1, 0) = mHalfYears * 3 / 100
                          vsGridSCess.TextMatrix(2, 0) = mHalfYears * 2 / 100
                          vsGridSCess.TextMatrix(3, 0) = mHalfYears * 1 / 100
                          
                          
                            Dim Yr1 As Integer
                            Dim Yr2 As Integer
                            Dim mPeriod As Integer
                            Dim mNoOfDemands As Integer
                            Yr1 = val(txtFromYear)
                            Yr2 = val(txtToYear)
                            mPeriod = val(txtFromPeriodID)
                                If mPeriod > 1 Then
                                    mPeriod = 2
                                Else
                                    mPeriod = 1
                                End If
                            mNoOfDemands = Yr2 - Yr1 + 1
                            mNoOfDemands = mNoOfDemands * 2
                                If mPeriod = 2 Then
                                    mNoOfDemands = mNoOfDemands - 1
                                End If
                                If val(txtToPeriodID) = 1 Then
                                    mNoOfDemands = mNoOfDemands - 1
                                End If
                            mNoOfHalfYears = mNoOfDemands
                          
                          
                          txtTotalserviceCess.Text = dSTotal * mNoOfHalfYears
                          vsGridSCess.TextMatrix(0, 3) = "4%"
                          vsGridSCess.TextMatrix(1, 3) = "3%"
                          vsGridSCess.TextMatrix(2, 3) = "2%"
                          vsGridSCess.TextMatrix(3, 3) = "1%"
                      'Call CalculateTotal
                      txtPenalServiceCess.Text = CalculatePenal(val(txtFromYear.Text), val(txtToYear.Text), val(txtFromPeriodID.Text), val(txtToPeriodID.Text), val(dSTotal))
            
        
           Else
              If chkSurcharge.Value = 1 Then
                chkCentralGovBulding.Enabled = False
                Else
                chkCentralGovBulding.Enabled = True
              End If
              txtPenalServiceCess.Text = ""
              txtTotalserviceCess.Text = ""
              Call VsgridChecked
              frmServiceCess.Enabled = False
              txtPenalServiceCess.Text = ""
              txtTotalserviceCess.Text = ""
              chkWaive.Value = vbUnchecked
              vsGridSCess.TextMatrix(0, 1) = vbUnchecked
              vsGridSCess.TextMatrix(1, 1) = vbUnchecked
              vsGridSCess.TextMatrix(2, 1) = vbUnchecked
              vsGridSCess.TextMatrix(3, 1) = vbUnchecked
              vsGridSCess.TextMatrix(0, 3) = ""
              vsGridSCess.TextMatrix(1, 3) = ""
              vsGridSCess.TextMatrix(2, 3) = ""
              vsGridSCess.TextMatrix(3, 3) = ""
          End If
    End If
End Sub
Private Sub chkSurcharge_Click()
    txtPenalSur.Enabled = False
    txtSurTotal.Enabled = False
    txtFromYrSur.Text = txtFromYear.Text
    txtFromPeriodSur.Text = txtFromPeriodID.Text
    txtToYrsur.Text = txtToYear.Text
    txtToPeriodSur.Text = txtToPeriodID.Text
    If chkSurcharge.Value = 1 Then
        chkCentralGovBulding.Enabled = False
        frmSurcharge.Enabled = True
    Else
        If chkServiceCess.Value = 1 Then
           chkCentralGovBulding.Enabled = False
           Else
           chkCentralGovBulding.Enabled = True
        End If
        
        txtSurRate.Text = ""
        txtFromPeriodSur.Text = ""
        txtFromYrSur.Text = ""
        txtToYrsur.Text = ""
        txtToPeriodSur.Text = ""
        vsGridSurcharge.Clear
        vsGridSurcharge.TextMatrix(0, 0) = "Year"
        vsGridSurcharge.TextMatrix(0, 1) = "Period"
        vsGridSurcharge.TextMatrix(0, 2) = "Amount"
        frmSurcharge.Enabled = False
    End If
End Sub

Private Sub chkWaive_Click()
    If chkWaive.Value = vbChecked Then
        vsGridSCess.Editable = flexEDKbdMouse
    Else
        vsGridSCess.Editable = flexEDNone
    End If
    
End Sub
Private Sub cmdGenarateDemnd_Click()
If gbLBPanchayat = 1 Then
    If chkServiceCess.Value = 1 Then
        If frmPropertyTaxCalculator.NonResi = 0 Then
           Call GenarateDemand(val(txtFromYear.Text), val(txtFromPeriodID.Text), val(txtToYear.Text), val(txtToPeriodID.Text), val(gbAcHeadCodeServicceCessCurrent), val(gbAcHeadCodeServicceCessArrear), val(dSTotal), val(txtPenalServiceCess.Text))
           Else
           Call GenarateDemand(val(txtFromYear.Text), val(txtFromPeriodID.Text), val(txtToYear.Text), val(txtToPeriodID.Text), val(gbAcHeadCodeServicceCessCurrentNonR), val(gbAcHeadCodeServicceCessArrearNonR), val(dSTotal), val(txtPenalServiceCess.Text))
        End If
    End If
    If chkSurcharge.Value = 1 Then
        If frmPropertyTaxCalculator.NonResi = 0 Then
            Call GenarateDemand(val(txtFromYrSur.Text), val(txtFromPeriodSur.Text), val(txtToYrsur.Text), val(txtToPeriodSur.Text), val(gbAcHeadCodeSurPTCurrent), val(gbAcHeadCodeSurPTArrear), val(dSurTotal), val(txtPenalSur.Text))
            Else
            Call GenarateDemand(val(txtFromYrSur.Text), val(txtFromPeriodSur.Text), val(txtToYrsur.Text), val(txtToPeriodSur.Text), val(gbAcHeadCodeSurPTCurrentNonR), val(gbAcHeadCodeSurPTArrearNonR), val(dSurTotal), val(txtPenalSur.Text))
        End If
    End If
    If chkCentralGovBulding.Value = 1 Then
        Call GenarateDemand(val(txtFromYrGov.Text), val(txtFromPeriodGov.Text), val(txtToYrGov.Text), val(txtToPeriodGov.Text), val(gbAcHeadCodeSurCentralGovtBuildCurrent), val(gbAcHeadCodeSurCentralGovtBuildArrear), val(dCentralGovTotal), val(txtGovPenal.Text))
    End If
     If chkFeeOnSplservice.Value = 1 Then
        Call GenarateDemand(val(txtFromYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToYrSpl.Text), val(txtToPeriodSpl.Text), val(gbAcHeadCodeSplServicesCurrent), val(gbAcHeadCodeSplServicesArrear), val(dSplFeeTotal), val(txPenalSpl.Text))
    End If
    
Else
    If chkServiceCess.Value = 1 Then
        Call GenarateDemand(val(txtFromYear.Text), val(txtFromPeriodID.Text), val(txtToYear.Text), val(txtToPeriodID.Text), val(gbAcHeadCodeServicceCessCurrent), val(gbAcHeadCodeServicceCessArrear), val(dSTotal), val(txtPenalServiceCess.Text))
    End If
     If chkSurcharge.Value = 1 Then
        Call GenarateDemand(val(txtFromYrSur.Text), val(txtFromPeriodSur.Text), val(txtToYear.Text), val(txtToPeriodID.Text), val(gbAcHeadCodeSurPTCurrent), val(gbAcHeadCodeSurPTArrear), val(dSurTotal), val(txtPenalSur.Text))
    End If
    If chkCentralGovBulding.Value = 1 Then
        Call GenarateDemand(val(txtFromYrGov.Text), val(txtFromPeriodGov.Text), val(txtToYrGov.Text), val(txtToPeriodGov.Text), val(gbAcHeadCodeSurCentralGovtBuildCurrent), val(gbAcHeadCodeSurCentralGovtBuildArrear), val(dCentralGovTotal), val(txtGovPenal.Text))
    End If
    If chkFeeOnSplservice.Value = 1 Then
        Call GenarateDemand(val(txtFromYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToYrSpl.Text), val(txtToPeriodSpl.Text), val(gbAcHeadCodeSplServicesCurrent), val(gbAcHeadCodeSplServicesArrear), val(dSplFeeTotal), val(txPenalSpl.Text))
    End If
    
End If
      Unload Me
End Sub
Private Sub Form_Load()
        Call FormEnableF
        Call FormInitialization
        Call LockControl
        optRateSpl.Value = False
        optAmtspl.Value = False
        txtHalfYrTax.Enabled = True
        bIsRightClk = False
End Sub

Private Sub optAmtGov_Click()
txtGov.Enabled = True
txtGovTotal.Text = ""
txtGovPenal.Text = ""

End Sub
Private Sub optAmtspl_Click()
    txtAmntspl.Enabled = True
    txtTotalSpl.Text = ""
    txPenalSpl.Text = ""

End Sub

Private Sub optRateGov_Click()
    txtGov.Enabled = True
    txtGovTotal.Text = ""
    txtGovPenal.Text = ""
End Sub

Private Sub optRateSpl_Click()
    txtAmntspl.Enabled = True
    txtTotalSpl.Text = ""
    txPenalSpl.Text = ""

End Sub
Private Sub txtAmntSpl_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtAmntspl_LostFocus()
   If txtAmntspl.Text <> "" Then
    If optRateSpl.Value = True Then
     If txtAmntspl.Text > 100 Then
          MsgBox ("Rate Should Not be Grater Than 100")
          txtAmntspl.Text = 0
          Else
          Call SplAmount
     End If
    End If
          
     Call SplAmount
    Else
   MsgBox ("Enter the Amount/Rate")
 End If
   
   
End Sub
Private Sub txtFromPeriodGov_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtFromPeriodGov_LostFocus()
   If Trim(txtFromPeriodGov.Text) = "" Then
            txtFromPeriodGov.Text = 1
    End If
End Sub
Private Sub txtFromPeriodID_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtFromPeriodID_LostFocus()
   If Trim(txtFromPeriodID.Text) = "" Then
            txtFromPeriodID.Text = 1
    End If
End Sub

Private Sub txtFromPeriodSpl_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If

End Sub
Private Sub txtFromPeriodSpl_LostFocus()
   If Trim(txtFromPeriodSpl.Text) = "" Then
            txtFromPeriodSpl.Text = 1
    End If
End Sub

Private Sub txtFromPeriodSur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtFromPeriodSur_LostFocus()
        If Trim(txtFromPeriodSur) = "" Then
            txtFromPeriodSur.Text = "1"
        End If
End Sub
Private Sub txtFromYear_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Sub

Private Sub txtFromYear_LostFocus()
    Dim mYear As Integer
    mYear = val(txtFromYear)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtFromYear = mYear
   ' Call LockControl
End Sub
Private Sub txtFromYrGov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PressTabKey
        Exit Sub
    End If
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
    Else
         KeyAscii = 0
    End If
End Sub
Private Sub txtFromYrGov_LostFocus()
    Dim mYear As Integer
    mYear = val(txtFromYrGov)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtFromYrGov = mYear
End Sub
Private Sub txtFromYrSpl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

End Sub
Private Sub txtFromYrSpl_LostFocus()
    Dim mYear As Integer
    mYear = val(txtFromYrSpl.Text)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtFromYrSpl = mYear
End Sub
Private Sub txtFromYrSur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
End Sub
Private Sub txtFromYrSur_LostFocus()
    Dim mYear As Integer
    mYear = val(txtFromYrSur)
        If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
        If mYear < 1901 Then mYear = gbFinancialYearID
        txtFromYrSur = mYear
End Sub
Private Sub txtGov_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Sub

Private Sub txtGov_LostFocus()
 If txtGov.Text <> "" Then
    If optRateGov.Value = True Then
     If txtGov.Text > 100 Then
          MsgBox ("Rate Should Not be Grater Than 100")
          txtGov.Text = 0
          Else
          Call BuldingAmount
     End If
    End If
          
     Call BuldingAmount
    Else
   MsgBox ("Enter the Amount/Rate")
 End If

End Sub
Private Sub txtHalfYrTax_Keypress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHalfYrTax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        bIsRightClk = False
    End If
End Sub
Private Sub txtSurRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtSurRate_LostFocus()
         If txtSurRate.Text <> "" Then
                If txtSurRate.Text < 100 Then
                Call SurchargeAmount
                Else
                MsgBox ("Rate should not be grater than 100")
                txtSurRate.Text = 0
                vsGridSurcharge.Clear 1, 0
                End If
             Else
             MsgBox ("Enter Srcharge Rate")
        End If
 
End Sub

Private Sub txtToPeriodGov_LostFocus()
   If txtGov.Text <> "" Then
        Call BuldingAmount
        Else
        MsgBox ("Enter Rate/Amount")
   End If
End Sub
Private Sub txtToPeriodGov_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub

Private Sub txtToPeriodID_LostFocus()
        If Trim(txtToPeriodID) = "" Then
            txtToPeriodID.Text = "2"
        End If
End Sub
Private Sub txtToPeriodID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtToPeriodSpl_LostFocus()
   If txtAmntspl.Text <> "" Then
    SplAmount
    Else
     MsgBox ("Enter Rate/Amount")
   End If
    



'    Dim Yr1 As Integer
'    Dim Yr2 As Integer
'    Dim mPeriod As Integer
'    Dim mNoOfDemands As Integer
'    Yr1 = val(txtFromYrSpl)
'    Yr2 = val(txtToYrSpl)
'    mPeriod = val(txtFromPeriodSpl)
'    If mPeriod > 1 Then
'        mPeriod = 2
'        Else
'        mPeriod = 1
'    End If
'    mNoOfDemands = Yr2 - Yr1 + 1
'    mNoOfDemands = mNoOfDemands * 2
'    If mPeriod = 2 Then
'        mNoOfDemands = mNoOfDemands - 1
'    End If
'    If val(txtToPeriodID) = 1 Then
'        mNoOfDemands = mNoOfDemands - 1
'    End If
'    mNoOfHalfYears = mNoOfDemands
'
'    If optAmtspl.value = True Then
'        txtTotalSpl.Text = txtAmntspl.Text
'        txPenalSpl.Text = CalculatePenal(val(txtFromYrSpl.Text), val(txtToYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToPeriodSpl.Text), val(txtAmntspl.Text))
'    End If
'    If optRateSpl.value = True Then
'       txtTotalSpl.Text = txtHalfYrTax * mNoOfHalfYears * val(txtAmntspl.Text) / 100
'        txPenalSpl.Text = CalculatePenal(val(txtFromYrSpl.Text), val(txtToYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToPeriodSpl.Text), val(txtTotalSpl.Text))
'    End If
'   If Trim(txtToPeriodSpl.Text) = "" Then
'            txtToPeriodSpl.Text = 1
'    End If
End Sub

Private Sub KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtToPeriodSur_LostFocus()
    If txtSurRate.Text <> "" Then
         Call SurchargeAmount
         Else
         MsgBox ("Enter Srcharge Rate")
    End If
End Sub
Private Sub txtToPeriodSur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
            If KeyAscii <> Asc("2") And KeyAscii <> 8 Then KeyAscii = Asc("1")
        Else
            KeyAscii = 0
        End If
End Sub


Private Sub txtTotalserviceCess_LostFocus()
    Dim mTotal As Single
    mTotal = val(txtTotalserviceCess.Text)
       If (mTotal - Int(mTotal)) > 0 Then
            mTotal = Int(mTotal) + 1
       End If
  
    txtTotalserviceCess.Text = Format(mTotal, "0.00")
    txtTotalserviceCess.Text = Format(val(txtTotalserviceCess), "0.00")
     
End Sub

Private Sub txtToYear_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Sub
Private Sub txtToYear_LostFocus()
    Dim mYear As Integer
    mYear = val(txtToYear)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtToYear = mYear
End Sub
Private Sub txtToYrGov_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

End Sub
Private Sub txtToYrGov_LostFocus()
    Dim mYear As Integer
    mYear = val(txtToYrGov.Text)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtToYrGov = mYear
End Sub

Private Sub txtToYrSpl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

End Sub
Private Sub txtToYrSpl_LostFocus()
    Dim mYear As Integer
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtToYrSpl = mYear
End Sub
Private Sub txtToYrsur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            PressTabKey
            Exit Sub
        End If
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

End Sub
Private Sub txtToYrSur_LostFocus()
    Dim mYear As Integer
    mYear = val(txtToYrsur)
    If mYear > gbFinancialYearID Then mYear = gbFinancialYearID
    If mYear < 1901 Then mYear = gbFinancialYearID
    txtToYrsur = mYear
End Sub
Private Sub vsGridSCess_AfterEdit(ByVal Row As Long, ByVal Col As Long)
      Call CalculateTotal
      txtPenalServiceCess.Text = CalculatePenal(val(txtFromYear.Text), val(txtToYear.Text), val(txtFromPeriodID.Text), val(txtToPeriodID.Text), val(dSTotal))
'      Call SurChange
If Col = 1 Then

    If vsGridSCess.Cell(flexcpChecked, Row, Col) = 2 Then
        vsGridSCess.TextMatrix(Row, 3) = ""
    Else
        If Row = 0 Then
            vsGridSCess.TextMatrix(Row, 3) = "4%"
        ElseIf Row = 1 Then
            vsGridSCess.TextMatrix(Row, 3) = "3%"
        ElseIf Row = 2 Then
            vsGridSCess.TextMatrix(Row, 3) = "2%"
        ElseIf Row = 3 Then
            vsGridSCess.TextMatrix(Row, 3) = "1%"
        End If
    End If
End If
End Sub


Private Sub CalculateTotal()
    Dim mLoop As Integer
    Dim mTotalAmt As Double
    
    Dim Yr1 As Integer
    Dim Yr2 As Integer
    Dim mPeriod As Integer
    Dim mNoOfDemands As Integer
    Yr1 = val(txtFromYear)
    Yr2 = val(txtToYear)
    mPeriod = val(txtFromPeriodID)
        If mPeriod > 1 Then
           mPeriod = 2
           Else
           mPeriod = 1
        End If
           mNoOfDemands = Yr2 - Yr1 + 1
           mNoOfDemands = mNoOfDemands * 2
       If mPeriod = 2 Then
           mNoOfDemands = mNoOfDemands - 1
       End If
       If val(txtToPeriodID) = 1 Then
           mNoOfDemands = mNoOfDemands - 1
       End If
          mNoOfHalfYears = mNoOfDemands
    
        For mLoop = 0 To 3
            If vsGridSCess.Cell(flexcpChecked, mLoop, 1) = 1 Then
                mTotalAmt = mTotalAmt + val(vsGridSCess.TextMatrix(mLoop, 0))
            End If
        Next
    txtTotalserviceCess.Text = mTotalAmt * mNoOfHalfYears
    dSTotal = mTotalAmt
End Sub

 Private Function Fine(ByVal mYearID As Integer, ByVal mPeriodId As Integer, ByVal mUptoDate As Date, ByVal mPTax As Double) As Double
          ' Modified By : Aiby                                                            '
        '             : For                                        '
        '==============================================================================='
       
        Dim dtFromDt As Variant
        Dim mNoOfMonths As Long
        Dim mAmount     As Double
        Dim dtFromDate  As Date
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation Mode 1= Act and 2 = Circular                          '
        '-------------------------------------------------------------------------------'
        If gbFineCalculationMode = 1 Then
            If mPeriodId = 1 Then
                dtFromDt = DateSerial(mYearID, 10, 1)
            Else
                dtFromDt = DateSerial(mYearID + 1, 4, 1)
            End If
            
            If mYearID = gbFinancialYearID And mPeriodId = 2 Then
                Fine = 0
                Exit Function
            End If
            
            If mYearID < 2006 Then
                If mYearID = 2005 And mPeriodId = 2 Then
                    GoTo Skip
                End If
                If mUptoDate > DateSerial(2005, 9, 1) Then
                    'mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 4, 1), dtFromDt)) * 2 + 10
                    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 9, 1), dtFromDt)) * 2
                    mNoOfMonths = mNoOfMonths + 1
                    dtFromDt = DateSerial(2005, 10, 1)
                    mYearID = 2005
                    mPeriodId = 2
                Else
                    mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2
                    dtFromDt = mUptoDate
                    mYearID = Year(dtFromDt)
                    If Month(dtFromDt) > 9 And Month(dtFromDt) < 4 Then
                        mPeriodId = 2
                    Else
                        mPeriodId = 1
                    End If
                End If
                
            
                'If Year(mUptoDate) = 2005 Then
'                If mYearID = 2005 Then
'                    If mPeriodID = 1 Then
'                        mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
'                        If Month(mUptoDate) > 5 Then
'                            mNoOfMonths = mNoOfMonths - ((Month(mUptoDate) - 5) * 12)
'                        End If
'                    Else
'                        GoTo Skip:
'                    End If
                'End If
                'If Year(mUptoDate) < 2005 Then 'New Change For UptoDate
                'Else
                'If mYearID < 2005 Then 'New Change For UptoDate
                '    mNoOfMonths = Abs(DateDiff("M", DateSerial(2005, 5, 1), dtFromDt)) * 2 + 10
                '    dtFromDt = DateSerial(2005, 11, 1)
                'End If
                'Else
                 '   mNoOfMonths = Abs(DateDiff("M", mUptoDate, dtFromDt)) * 2 + 10
                'End If
                
                
                
            End If
Skip:
            If mUptoDate >= dtFromDt Then
                'mNoOfMonths = mNoOfMonths + (gbFinancialYearID - mYearID) * 12 'New Change For UptoDate
                mNoOfMonths = mNoOfMonths + 1 + Abs(DateDiff("M", mUptoDate, dtFromDt))  'New Change For UptoDate
            End If
            If mYearID = gbFinancialYearID And mPeriodId = 1 Then
                'mNoOfMonths = mNoOfMonths - 1
            End If
            'mNoOfMonths = mNoOfMonths + 1
            dtFromDate = DateAdd("m", 1, mUptoDate)
            'Debug.Print "No of Months (Fine) " & mNoOfMonths
            Fine = mPTax * mNoOfMonths / 100
            'If mNoOfMonths = 60 Then Stop
            Debug.Print "No of Months (Fine) " & mNoOfMonths & "    " & Fine
            Exit Function
        ElseIf gbFineCalculationMode = 2 Then
        '-------------------------------------------------------------------------------'
        ' NOTE:- Fine Calculation As Per Circular                                       '
        '-------------------------------------------------------------------------------'
           'mPTax = Format(mPTax * 2, "0.00")
            dtFromDt = DateSerial(mYearID, 11, 1)
            If mYearID = gbFinancialYearID Then
                Fine = 0
                Exit Function
            End If
            If mYearID < 2005 Then
                mNoOfMonths = Abs(DateDiff("m", DateSerial(2005, 8, 1), dtFromDt))
                dtFromDt = DateSerial(2005, 9, 1)
                mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt))
            End If
            mNoOfMonths = mNoOfMonths + Abs(DateDiff("m", gbTransactionDate, dtFromDt)) + 1
            Fine = mPTax * mNoOfMonths / 100
            Exit Function
        End If
    End Function
Private Sub CheckEnableF()
        chkServiceCess.Enabled = False
        chkSurcharge.Enabled = False
        chkCentralGovBulding.Enabled = False
        chkFeeOnSplservice.Enabled = False
End Sub
 Private Sub CheckEnableT()
        chkServiceCess.Enabled = True
        chkSurcharge.Enabled = True
        chkCentralGovBulding.Enabled = True
        chkFeeOnSplservice.Enabled = True
End Sub

Private Sub FormEnableF()
        frmServiceCess.Enabled = False
        frmSurcharge.Enabled = False
        frmCentralgov.Enabled = False
        frmFeeOnSpclservice.Enabled = False
End Sub
Private Sub FormEnableT()
    frmServiceCess.Enabled = True
    frmSurcharge.Enabled = True
    frmCentralgov.Enabled = True
    frmFeeOnSpclservice.Enabled = True
End Sub
Private Sub VsgridChecked()
    vsGridSCess.TextMatrix(0, 1) = vbChecked
    vsGridSCess.TextMatrix(1, 1) = vbChecked
    vsGridSCess.TextMatrix(2, 1) = vbChecked
    vsGridSCess.TextMatrix(3, 1) = vbChecked
End Sub
Private Sub VsgridUnChecked()
    vsGridSCess.TextMatrix(0, 1) = vbUnchecked
    vsGridSCess.TextMatrix(1, 1) = vbUnchecked
    vsGridSCess.TextMatrix(2, 1) = vbUnchecked
    vsGridSCess.TextMatrix(3, 1) = vbUnchecked
 End Sub
Private Sub SurchargeGrid()
    Dim sAmount As Double
    Dim mPeriodId As Integer
    Dim mNoOfDemands As Integer
    Dim mYearID As Integer
    Dim mRow As Integer
    Dim Yr1 As Integer
    Dim Yr2 As Integer
    Dim mPeriod As Integer
    Dim mNoOfHalfYearsSur As Integer
    Dim mPeriodCount As Integer
    Dim dAmount As Double
    Yr1 = val(txtFromYrSur)
    Yr2 = val(txtToYrsur)
     'Call NoOfHalfYr(val(txtFromYrSur.Text), val(txtToYrsur.Text), val(txtFromPeriodSur.Text), val(txtToPeriodSur.Text))
      mPeriod = val(txtFromPeriodSur)
        If mPeriod > 1 Then
            mPeriod = 2
        Else
            mPeriod = 1
        End If
    mNoOfDemands = Yr2 - Yr1 + 1
    mNoOfDemands = mNoOfDemands * 2
        If mPeriod = 2 Then
            mNoOfDemands = mNoOfDemands - 1
        End If
        
        If val(txtToPeriodSur) = 1 Then
            mNoOfDemands = mNoOfDemands - 1
        End If
    mNoOfHalfYearsSur = mNoOfDemands
    Dim mLoop As Integer
        mYearID = val(txtFromPeriodSur)
        mPeriodId = val(txtFromPeriodSur)
        mPeriodCount = val(txtFromPeriodSur)
        mRow = 1
            vsGridSurcharge.TextMatrix(mRow, 0) = val(txtFromYrSur)
            If val(txtFromPeriodSur) = 1 Then
                vsGridSurcharge.TextMatrix(mRow, 1) = "FirstHalf"
                Else
                vsGridSurcharge.TextMatrix(mRow, 1) = "SecondHalf"
            End If
            vsGridSurcharge.TextMatrix(mRow, 2) = val(txtHalfYrTax.Text) * val(txtSurRate.Text) / 100
            dAmount = vsGridSurcharge.TextMatrix(mRow, 2)
            dSurTotal = vsGridSurcharge.TextMatrix(mRow, 2)
         '  vsGridSurcharge.Rows = 2
         mLoop = 2
        For mLoop = 2 To mNoOfHalfYearsSur
        vsGridSurcharge.Rows = vsGridSurcharge.Rows + 1
       ' mRow = mRow + 1
           If mPeriodCount = 1 Then
                vsGridSurcharge.TextMatrix(mLoop, 0) = Yr1
                If val(mPeriodCount + 1) = 2 Then
                    vsGridSurcharge.TextMatrix(mLoop, 1) = "Second Half"
                    Else
                     vsGridSurcharge.TextMatrix(mLoop, 1) = "First Half"
                End If
              ' vsGridSurcharge.TextMatrix(mLoop, 1) = mPeriodCount + 1
                dAmount = dAmount + val(txtHalfYrTax.Text) * val(txtSurRate.Text) / 100
                vsGridSurcharge.TextMatrix(mLoop, 2) = val(txtHalfYrTax.Text) * val(txtSurRate.Text) / 100
                mPeriodCount = mPeriodCount + 1
                
            Else
                vsGridSurcharge.TextMatrix(mLoop, 0) = Yr1 + 1
                If val(mPeriodCount - 1) = 1 Then
                    vsGridSurcharge.TextMatrix(mLoop, 1) = "First Half"
                    Else
                    vsGridSurcharge.TextMatrix(mLoop, 1) = "Second Half"
                End If
               ' vsGridSurcharge.TextMatrix(mLoop, 1) = mPeriodCount - 1
                dAmount = dAmount + val(txtHalfYrTax.Text) * val(txtSurRate.Text) / 100
                vsGridSurcharge.TextMatrix(mLoop, 2) = val(txtHalfYrTax.Text) * val(txtSurRate.Text) / 100
                mPeriodCount = mPeriodCount - 1
                Yr1 = Yr1 + 1
            End If
       
        Next
        txtSurTotal.Text = dAmount
End Sub
Private Sub DateValidation(ByVal FromYear As Integer, ByVal ToYear As Integer, ByVal FromPeriod As Integer, ByVal ToPeriod As Integer)
    If FromYear = 0 Then
        MsgBox "Enter the Year", vbInformation
        txtFromYear.SetFocus
        Exit Sub
    End If
    If FromPeriod = 0 Then
        MsgBox "Enter the Period", vbInformation
        txtFromPeriodID.SetFocus
        Exit Sub
    End If
    If ToYear = 0 Then
        MsgBox "Enter the Year", vbInformation
        txtToYear.SetFocus
        Exit Sub
    End If
        If ToPeriod = 0 Then
        MsgBox "Enter the Period", vbInformation
        txtToPeriodID.SetFocus
        Exit Sub
    End If
 
End Sub
Private Function CalculatePenal(ByVal FromYear As Integer, ByVal ToYear As Integer, ByVal FromPeriod As Integer, ByVal ToPeriod As Integer, ByVal TaxRate As Double) As Double
    Dim mYearID As Integer
    Dim mLoopFlag As Boolean
    Dim mPeriodId As Integer
    Dim mFine As Double
    Call DateValidation(FromYear, ToYear, FromPeriod, ToPeriod)
            
        mFine = 0
    mLoopFlag = True
    mYearID = val(FromYear)
    mPeriodId = val(FromPeriod)
   
        mLoopFlag = True
        While mLoopFlag 'And (mYearID <= val(txtToYear) And mPeriodID <= val(txtToPeriodID))
            mFine = mFine + Fine(mYearID, mPeriodId, gbTransactionDate, val(TaxRate))
            'If mYearID = val(txtToYear) And mPeriodID = val(txtToPeriodID) Then ' Changed on 26-07-10 to check 2000/2
            If mYearID >= val(txtToYear) And mPeriodId = val(txtToPeriodID) Then
                mLoopFlag = False
            Else
                If mPeriodId < 2 Then
                    mPeriodId = 2
                Else
                    mYearID = mYearID + 1
                    mPeriodId = 1
                End If
            End If
            If mYearID > val(txtToYear) Then
                'If mPeriodID > val(txtToPeriodID) Then ' Changed by Aiby 26-07-10 Check Runtime Error 2000/2
                If mPeriodId >= val(txtToPeriodID) Then
                    mLoopFlag = False
                End If
            End If
        Wend
        'For mLoop = 1 To mNoOfHalfYears
            'mFine = mFine + CalculateFineforPTax(mYearID, mPeriodID, Val(lblPT.Caption))
        'Next mLoop
        
        If mFine - Int(mFine) > 0 Then
            mFine = mFine + (1 - (mFine - Int(mFine)))
        End If
       CalculatePenal = Format(mFine, "0.00")
 End Function
 Public Sub GenarateDemand(ByVal FromYear As Integer, ByVal FromPeriodID As Integer, ByVal ToYear As Integer, ByVal ToperiodId As Integer, ByVal gbAcHeadCodeCurrent As String, ByVal gbAcHeadCodeArrear As String, ByVal Total As Double, ByVal Penal As Double)
        Dim mLoop As Integer
        Dim mYearID As Integer
        Dim mPeriodId As Integer
        Dim objAcc As New clsAccounts
        Dim mRow As Integer
        Dim mFineFlag As Boolean
        Dim mTotal As Single
        Dim mPenal As Single
            Call NoOfHalfYr(val(FromYear), val(ToYear), val(FromPeriodID), val(ToperiodId))
            mYearID = val(FromYear)
            mPeriodId = val(FromPeriodID)
            'frmReceiptsCounter.vsGrid.Clear 1, 0

             mRow = frmPropertyTaxCalculator.PTRowCount     '1
           

            For mLoop = 1 To mNoOfHalfYears
                If mYearID < gbFinancialYearID Then
                    mFineFlag = True
                    objAcc.SetAccountCode gbAcHeadCodeArrear
                Else
                    objAcc.SetAccountCode gbAcHeadCodeCurrent
                End If
                If frmReceiptsCounter.vsGrid.Rows = mRow Then
                    frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 10
                End If
                
                If objAcc.AccountHeadID > 0 Then
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = mYearID & "-" & mYearID + 1
                    
                    If mPeriodId = 1 Then
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 1 '"First Half"
                    Else
                        frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = 2 '"Second Half"
                    End If
                    If mYearID < gbFinancialYearID Then
                            mTotal = val(Total)
                            If (mTotal - Int(mTotal)) > 0 Then
                               mTotal = Int(mTotal) + 1
                            End If
                         frmReceiptsCounter.vsGrid.TextMatrix(mRow, 4) = Format(val(mTotal), "#0")
                    Else
                         mTotal = val(Total)
                            If (mTotal - Int(mTotal)) > 0 Then
                               mTotal = Int(mTotal) + 1
                            End If
                         frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = Format(val(mTotal), "#0")
                    End If
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = mYearID
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodId
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = val(mTotal)
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                End If
                mRow = mRow + 1
                frmPropertyTaxCalculator.PTRowCount = mRow
                If frmReceiptsCounter.vsGrid.Rows = mRow Then
                    frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 2
                End If
    
                If mPeriodId = 1 Then
                    mPeriodId = 2
                Else
                    mPeriodId = 1
                    mYearID = mYearID + 1
                End If
            Next mLoop
            
            'mRowCount = mRow
             If val(Penal) > 0 Then
            '  mRowCount = mRow
                objAcc.SetAccountCode gbAcHeadCodePenalInterest
                If objAcc.AccountHeadID > 0 Then
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 0) = objAcc.AccountCode
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 1) = objAcc.AccountHead
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 2) = gbFinancialYearID & "-" & gbFinancialYearID + 1
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 3) = ""
                      mPenal = val(Penal)
                            If (mPenal - Int(mPenal)) > 0 Then
                               mPenal = Int(mPenal) + 1
                            End If
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 5) = Format(val(mPenal), "#0")
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 6) = objAcc.AccountHeadID
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 7) = gbFinancialYearID
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 8) = mPeriodId
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 9) = 1
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 10) = ""
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 11) = Format(val(mPenal), "#0")
                    frmReceiptsCounter.vsGrid.TextMatrix(mRow, 12) = ""
                    frmPropertyTaxCalculator.PTRowCount = frmPropertyTaxCalculator.PTRowCount + 1
                End If
            Else
              '  mRowCount = 1
            End If
            frmReceiptsCounter.Calculate
            frmReceiptsCounter.txtTransactionType.Tag = gbTransactionTypePTax
            frmReceiptsCounter.txtTransactionType.Text = "Property Tax"
    

 End Sub

 Private Sub AmntValidation()
    If txtHalfYrTax = "" Then
        MsgBox "Enter the HalfYear TaxRate", vbInformation
        txtHalfYrTax.SetFocus
        Exit Sub
    End If
End Sub
 Private Sub NoOfHalfYr(ByVal FromYear As Integer, ByVal ToYear As Integer, ByVal FromPeriodID As Integer, ByVal ToperiodId As Integer)
    Dim Yr1 As Integer
    Dim Yr2 As Integer
    Dim mPeriod As Integer
    Dim mNoOfDemands As Integer
    Yr1 = FromYear
    Yr2 = ToYear
    mPeriod = FromPeriodID
        If mPeriod > 1 Then
            mPeriod = 2
        Else
            mPeriod = 1
        End If
    mNoOfDemands = Yr2 - Yr1 + 1
    mNoOfDemands = mNoOfDemands * 2
        If mPeriod = 2 Then
            mNoOfDemands = mNoOfDemands - 1
        End If
        If ToperiodId = 1 Then
            mNoOfDemands = mNoOfDemands - 1
        End If
    mNoOfHalfYears = mNoOfDemands
End Sub
Property Let Mode(mData As Variant)
    mMode = mData
End Property
Property Get Mode() As Variant
    mMode = mMode
End Property
Public Sub FormInitialization()
    txtHalfYrTax.Text = frmPropertyTaxCalculator.txtTaxRate
    txtFromYear.Text = frmPropertyTaxCalculator.txtFromYear.Text
    txtFromPeriodID.Text = frmPropertyTaxCalculator.txtFromPeriodID
    txtToYear.Text = frmPropertyTaxCalculator.txtToYear
    txtToPeriodID.Text = frmPropertyTaxCalculator.txtToPeriodID
End Sub
Private Sub SurchargeAmount()
   If Trim(txtToPeriodSur.Text) = "" Then
            txtToPeriodSur.Text = 2
    End If
    If chkSurcharge.Value = 1 Then
        vsGridSurcharge.Clear
        vsGridSurcharge.TextMatrix(0, 0) = "Year"
        vsGridSurcharge.TextMatrix(0, 1) = "Period"
        vsGridSurcharge.TextMatrix(0, 2) = "Amount"
       Call SurchargeGrid
       txtPenalSur.Text = CalculatePenal(val(txtFromYrSur.Text), val(txtToYrsur.Text), val(txtFromPeriodSur.Text), val(txtToPeriodSur.Text), val(dSurTotal))
    End If
End Sub
Private Sub SplAmount()
    Dim Yr1 As Integer
    Dim Yr2 As Integer
    Dim mPeriod As Integer
    Dim mNoOfDemands As Integer
    Yr1 = val(txtFromYrSpl)
    Yr2 = val(txtToYrSpl)
    mPeriod = val(txtFromPeriodSpl)
    If mPeriod > 1 Then
        mPeriod = 2
        Else
        mPeriod = 1
    End If
    mNoOfDemands = Yr2 - Yr1 + 1
    mNoOfDemands = mNoOfDemands * 2
    If mPeriod = 2 Then
        mNoOfDemands = mNoOfDemands - 1
    End If
    If val(txtToPeriodID) = 1 Then
        mNoOfDemands = mNoOfDemands - 1
    End If
    mNoOfHalfYears = mNoOfDemands
    
    If optAmtspl.Value = True Then
        txtTotalSpl.Text = txtAmntspl.Text
        dSplFeeTotal = txtTotalSpl.Text
        txPenalSpl.Text = CalculatePenal(val(txtFromYrSpl.Text), val(txtToYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToPeriodSpl.Text), val(txtAmntspl.Text))
    End If
    If optRateSpl.Value = True Then
        txtTotalSpl.Text = txtHalfYrTax * mNoOfHalfYears * val(txtAmntspl.Text) / 100
        dSplFeeTotal = txtHalfYrTax * val(txtAmntspl.Text) / 100
        txPenalSpl.Text = CalculatePenal(val(txtFromYrSpl.Text), val(txtToYrSpl.Text), val(txtFromPeriodSpl.Text), val(txtToPeriodSpl.Text), val(dSplFeeTotal))
    End If
   If Trim(txtToPeriodSpl.Text) = "" Then
            txtToPeriodSpl.Text = 1
    End If
End Sub
Private Sub BuldingAmount()
     Dim Yr1 As Integer
    Dim Yr2 As Integer
    Dim mPeriod As Integer
    Dim mNoOfDemands As Integer
    Yr1 = val(txtFromYrGov)
    Yr2 = val(txtToYrGov)
    mPeriod = val(txtFromPeriodGov)
    If mPeriod > 1 Then
        mPeriod = 2
        Else
        mPeriod = 1
    End If
    mNoOfDemands = Yr2 - Yr1 + 1
    mNoOfDemands = mNoOfDemands * 2
    If mPeriod = 2 Then
        mNoOfDemands = mNoOfDemands - 1
    End If
    If val(txtToPeriodID) = 1 Then
        mNoOfDemands = mNoOfDemands - 1
    End If
    mNoOfHalfYears = mNoOfDemands

    If optAmtGov.Value = True Then
        txtGovTotal.Text = txtGov.Text
        dCentralGovTotal = txtGovTotal.Text
        txtGovPenal.Text = CalculatePenal(val(txtFromYrGov.Text), val(txtToYrGov.Text), val(txtFromPeriodGov.Text), val(txtToPeriodGov.Text), val(txtGov.Text))
        
    End If
    If optRateGov.Value = True Then
       txtGovTotal.Text = txtHalfYrTax * mNoOfHalfYears * val(txtGov.Text) / 100
       dCentralGovTotal = txtHalfYrTax * val(txtGov.Text) / 100
       txtGovPenal.Text = CalculatePenal(val(txtFromYrGov.Text), val(txtToYrGov.Text), val(txtFromPeriodGov.Text), val(txtToPeriodGov.Text), val(dCentralGovTotal))
    End If
 
    If Trim(txtToPeriodGov.Text) = "" Then
            txtToPeriodGov.Text = 2
    End If
    
End Sub


Private Sub vsGridSCess_Click()
 If chkWaive.Value = 1 Then
    If vsGridSCess.Col = 1 Then
        vsGridSCess.Editable = flexEDKbdMouse
    Else
         vsGridSCess.Editable = flexEDNone
    End If
 Else
     vsGridSCess.Editable = flexEDNone
 End If

End Sub

Private Sub LockControl()
    If val(frmPropertyTaxCalculator.txtTaxRate) = 0 Then
        txtHalfYrTax.Locked = False
     Else
        txtHalfYrTax.Locked = True
    End If
    
    If frmPropertyTaxCalculator.txtFromYear = "" Then
        txtFromYear.Locked = False
     Else
        txtFromYear.Locked = True
    End If
    
     If frmPropertyTaxCalculator.txtFromPeriodID = "" Then
        txtFromPeriodID.Locked = False
     Else
        txtFromPeriodID.Locked = True
    End If
    
     If frmPropertyTaxCalculator.txtToYear = "" Then
        txtToYear.Locked = False
     Else
        txtToYear.Locked = True
    End If
    
     If frmPropertyTaxCalculator.txtToPeriodID = "" Then
        txtToPeriodID.Locked = False
     Else
        txtToPeriodID.Locked = True
    End If

End Sub

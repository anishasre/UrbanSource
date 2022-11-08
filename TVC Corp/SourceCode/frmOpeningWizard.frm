VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmOpeningWizard 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Entry Wizard"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16890
   Icon            =   "frmOpeningWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   16890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContainer 
      Height          =   8070
      Left            =   2760
      TabIndex        =   7
      Top             =   315
      Width           =   14235
      Begin VB.Frame fraVerify 
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   660
         TabIndex        =   14
         Top             =   315
         Width           =   12750
         Begin VB.CommandButton cmdUdoVrGenerate 
            Caption         =   "Undo Voucher Generation"
            Enabled         =   0   'False
            Height          =   330
            Left            =   10485
            TabIndex        =   38
            Top             =   5985
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VSFlex8LCtl.VSFlexGrid vsVerifyContra 
            Height          =   1770
            Left            =   1290
            TabIndex        =   16
            Top             =   3120
            Visible         =   0   'False
            Width           =   8790
            _cx             =   15505
            _cy             =   3122
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOpeningWizard.frx":1CCA
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
         Begin VSFlex8LCtl.VSFlexGrid vsVerifyCB 
            Height          =   1770
            Left            =   1620
            TabIndex        =   36
            Top             =   4710
            Visible         =   0   'False
            Width           =   7485
            _cx             =   13203
            _cy             =   3122
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Rows            =   7
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOpeningWizard.frx":1DA2
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
         Begin VSFlex8LCtl.VSFlexGrid vsVerify 
            Height          =   2895
            Left            =   2430
            TabIndex        =   15
            Top             =   225
            Width           =   5325
            _cx             =   9393
            _cy             =   5106
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9.75
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
            Rows            =   10
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOpeningWizard.frx":1E7B
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
            OutlineBar      =   2
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
         Begin VB.Label lblVerifyMes 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   465
            Left            =   360
            TabIndex        =   35
            Top             =   5445
            Visible         =   0   'False
            Width           =   11355
         End
      End
      Begin VB.Frame fraOnline 
         Caption         =   "Online Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4290
         Left            =   900
         TabIndex        =   8
         Top             =   1350
         Visible         =   0   'False
         Width           =   12210
         Begin VB.TextBox txtOnlineDate 
            Height          =   285
            Left            =   3105
            TabIndex        =   9
            Top             =   1080
            Width           =   1995
         End
         Begin MSComCtl2.DTPicker dtpOnlinedate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-mmm-yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   5085
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1080
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   8421376
            CalendarTitleBackColor=   -2147483638
            CustomFormat    =   "dd/mm/yyyy"
            Format          =   64749571
            CurrentDate     =   39612
         End
         Begin VB.Label lblDateMessage 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   330
            Left            =   225
            TabIndex        =   37
            Top             =   495
            Visible         =   0   'False
            Width           =   10680
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SELECT SAANKHYA ONLINE DATE"
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
            Left            =   375
            TabIndex        =   13
            Top             =   1080
            Width           =   2610
         End
         Begin VB.Label lblActual 
            BackColor       =   &H80000009&
            Caption         =   "act"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1860
            Left            =   6615
            TabIndex        =   12
            Top             =   765
            Width           =   4335
         End
         Begin VB.Label lblObTrans 
            Caption         =   "ob"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6660
            TabIndex        =   11
            Top             =   1080
            Visible         =   0   'False
            Width           =   4290
         End
      End
      Begin VB.Frame fraGenerateVr 
         Height          =   1140
         Left            =   1215
         TabIndex        =   17
         Top             =   6705
         Visible         =   0   'False
         Width           =   10905
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   315
            TabIndex        =   18
            Top             =   360
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblMessage 
            Caption         =   "lblMessage"
            Height          =   285
            Left            =   1755
            TabIndex        =   19
            Top             =   765
            Visible         =   0   'False
            Width           =   6135
         End
      End
   End
   Begin VB.Frame fraLabel 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   16800
      Begin VB.Label lblHeading 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OPENING CASH BOOK ENTRY WIZARD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   34
         Top             =   45
         Width           =   16665
      End
      Begin VB.Image Image2 
         Height          =   330
         Left            =   0
         Picture         =   "frmOpeningWizard.frx":1F62
         Stretch         =   -1  'True
         Top             =   45
         Width           =   16740
      End
   End
   Begin VB.Frame fraButton 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2565
      TabIndex        =   1
      Top             =   8460
      Width           =   14280
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
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
         Left            =   6345
         TabIndex        =   6
         Top             =   45
         Width           =   1230
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   7560
         TabIndex        =   4
         Top             =   45
         Width           =   1230
      End
      Begin VB.CommandButton cmdPre 
         Caption         =   "Previous"
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
         Left            =   3870
         TabIndex        =   3
         Top             =   45
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   5130
         TabIndex        =   2
         Top             =   45
         Width           =   1230
      End
   End
   Begin VB.Frame frmMain 
      BorderStyle     =   0  'None
      Caption         =   "Process Flow"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8475
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   2535
      Begin VB.CheckBox chkOnline 
         Caption         =   "Online Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   45
         TabIndex        =   26
         Top             =   1305
         Width           =   195
      End
      Begin VB.CheckBox chkVr 
         Caption         =   "Generate Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   25
         Top             =   3690
         Width           =   195
      End
      Begin VB.CheckBox chkVerify 
         Caption         =   "Verify"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   24
         Top             =   3285
         Width           =   195
      End
      Begin VB.CheckBox chkOB 
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   23
         Top             =   1665
         Width           =   195
      End
      Begin VB.CheckBox chkR 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   22
         Top             =   2040
         Width           =   195
      End
      Begin VB.CheckBox chkP 
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   45
         TabIndex        =   21
         Top             =   2475
         Width           =   195
      End
      Begin VB.CheckBox chkCB 
         Caption         =   "Closing Balance"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   20
         Top             =   2865
         Width           =   195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   33
         Top             =   2475
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   270
         TabIndex        =   32
         Top             =   2880
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Verify RP Statement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   31
         Top             =   3285
         Width           =   2085
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Voucher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   30
         Top             =   3690
         Width           =   1890
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   29
         Top             =   1665
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   28
         Top             =   2070
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Online Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   27
         Top             =   1305
         Width           =   1230
      End
      Begin VB.Image Image1 
         Height          =   8430
         Left            =   0
         Picture         =   "frmOpeningWizard.frx":6460
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2520
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3330
      Top             =   8595
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
End
Attribute VB_Name = "frmOpeningWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private Const FRM_NUMBEROFPAGES = 6 ' = 7 pages total. Note: DO NOT CHANGE!!!.
    Private Const FRM_NUMBERTOUSE = FRM_NUMBEROFPAGES
    'Declare Page Array
    Dim FRM_ARRAY_PAGES(FRM_NUMBEROFPAGES) As Integer
    'Page number contstants.
    Private Const FRM_ONLINEDATE = 0
    Private Const FRM_OPENING = 1
    Private Const FRM_RECEIPT = 2
    Private Const FRM_PAYMENT = 3
    Private Const FRM_CLOSING = 4
    Private Const FRM_VERIFY = 5
    Private Const FRM_GENERATEVr = FRM_NUMBEROFPAGES
'    Private Const FRM_FINISH = FRM_NUMBEROFPAGES
    'Var to store active page #
    Private FRM_CURPAGE As Integer
    'Declare Page Title Array
    Dim FRM_ARRAY_TITLES(FRM_NUMBEROFPAGES) As String
    'Page Title Contants, ONLY CHANGE YOUR TITLES HERE!
    Private Const FRM_ONLINEDATE_TITLE = " ONLINE DATE "
    Private Const FRM_OPENING_TITLE = " OPENING CASH BOOK "
    Private Const FRM_RECEIPT_TITLE = " RECEIPT "
    Private Const FRM_PAYMENT_TITLE = " PAYMENT"
    Private Const FRM_CLOSING_TITLE = " CLOSING"
    Private Const FRM_VERIFY_TITLE = " VERIFY R&P STATEMENT"
    Private Const FRM_GENERATEVr_TITLE = " GENERATE VOUCHER "
'    Private Const FRM_FINISH_TITLE = " Finish "
    Private mvarFrameNo As Integer 'FRM_ONLINEDATE =0 ...
    
    Public mFreeze As Integer

    Public Sub cmdCancel_Click()
        Unload Me
    End Sub
    Public Sub cmdNext_Click()
       mvarFrameNo = 0
       Dim i As Integer
       For i = FRM_ONLINEDATE To FRM_NUMBERTOUSE - 1
          If FRM_ARRAY_PAGES(i) And (i < FRM_NUMBERTOUSE) Then
             FRM_CURPAGE = i + 1
             FRM_ARRAY_PAGES(i) = False '
             Exit For
          End If
       Next i
       If FRM_CURPAGE > FRM_ONLINEDATE Then
          cmdPre.Visible = True
          cmdPre.Enabled = True
       Else
          cmdPre.Enabled = False
       End If
       
       If FRM_CURPAGE < FRM_NUMBERTOUSE Then
          cmdNext.Enabled = True
          cmdCancel.Caption = "&Cancel"
       Else
          cmdNext.Enabled = False
          cmdCancel.Caption = "&Finish"
       End If
       
     If FRM_CURPAGE < FRM_NUMBERTOUSE Then
       FRM_ARRAY_PAGES(FRM_CURPAGE) = True
       fraContainer.Caption = FRM_ARRAY_TITLES(FRM_CURPAGE)
     Else 'Your using less pages, so if equal, show finish page.
       FRM_ARRAY_PAGES(FRM_NUMBEROFPAGES) = True ' Update jump to last page page
       fraContainer.Caption = FRM_ARRAY_TITLES(FRM_NUMBEROFPAGES) ' Set page title
     End If
     
     Select Case FRM_CURPAGE
           Case FRM_OPENING
                 Call OPENINGFRAME
           Case FRM_RECEIPT
                Call RECEIPTFRAME
           Case FRM_PAYMENT
                Call PAYMENTFRAME
           Case FRM_CLOSING
                Call CLOSINGFRAME
           Case FRM_VERIFY
                Call VERIFYFRAME
           Case FRM_GENERATEVr
                 Call GENERATEVrFRAME
    '       Case FRM_FINISH
        End Select
    End Sub
    Public Sub cmdPre_Click()
        Dim i As Integer
        For i = FRM_ONLINEDATE To FRM_NUMBERTOUSE
           If FRM_ARRAY_PAGES(i) And (i > FRM_ONLINEDATE) Then 'found it!
              FRM_CURPAGE = i - 1 'Make Prev page active.
              FRM_ARRAY_PAGES(i) = False ' Disable current page
              Exit For
'           Else 'You're using less pages, so catch it.
'              FRM_CURPAGE = i - 1 'Make Prev page active.
'              FRM_ARRAY_PAGES(FRM_NUMBEROFPAGES) = False ' Disable Finish page
'             ' Frame1(FRM_NUMBEROFPAGES).Visible = False ' and hide it.
           End If
        Next i
        If FRM_CURPAGE > FRM_ONLINEDATE Then
           cmdPre.Enabled = True
        Else                             'Set button states
           cmdPre.Enabled = False
        End If
        
        If FRM_CURPAGE < FRM_NUMBERTOUSE Then
           cmdNext.Enabled = True
           cmdCancel.Caption = "&Cancel"
        Else                             'And set Cancel button state Caption
           cmdNext.Enabled = False
           cmdCancel.Caption = "&Finish" 'Last page, so change button caption
        End If
        
        FRM_ARRAY_PAGES(FRM_CURPAGE) = True ' Update Array with Current page
        fraContainer.Caption = FRM_ARRAY_TITLES(FRM_CURPAGE) ' Set page title
        Select Case FRM_CURPAGE
           Case FRM_OPENING
                 Call OPENINGFRAME
           Case FRM_RECEIPT
                Call RECEIPTFRAME
           Case FRM_PAYMENT
                Call PAYMENTFRAME
           Case FRM_CLOSING
                Call CLOSINGFRAME
           Case FRM_VERIFY
                Call VERIFYFRAME
           Case FRM_GENERATEVr
                 Call GENERATEVrFRAME
    '       Case FRM_FINISH
        End Select

    End Sub
    Private Sub OPENINGFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = True
        End If
        lblHeading.Caption = "OPENING AMOUNT VERIFICATION"
        fraVerify.Visible = False
        fraButton.Visible = False
        fraGenerateVr.Visible = False
        'chkOnline.value = vbChecked
        fraOnline.Visible = False
        Call CheckBoxstatus(1)
        frmOBCashBook.BorderStyle = 0
        frmOBCashBook.Top = frmMenu.Top + 3500
        frmOBCashBook.Left = frmMenu.Left + Me.Left + 2635
        frmOBCashBook.Width = fraContainer.Width
        frmOBCashBook.Show vbModal
        fraButton.Visible = True
    End Sub
                
    Private Sub RECEIPTFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = True
        End If
        Call CheckBoxstatus(2)
        lblHeading.Caption = "RECEIPT"
        fraGenerateVr.Visible = False
        fraButton.Visible = False
        fraOnline.Visible = False
        fraVerify.Visible = False
        frmOBReceiptTransactions.BorderStyle = 0
        frmOBReceiptTransactions.Top = 2400
        frmOBReceiptTransactions.Left = frmMenu.Left + Me.Left + 2655
        frmOBReceiptTransactions.Width = fraContainer.Width
        frmOBReceiptTransactions.Show vbModal
    End Sub
    Private Sub PAYMENTFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
        Call CheckBoxstatus(3)
        lblHeading.Caption = "PAYMENT"
        fraGenerateVr.Visible = False
        fraButton.Visible = False
        fraOnline.Visible = False
        fraVerify.Visible = False
        frmOBPaymentTransactions.BorderStyle = 0
        frmOBPaymentTransactions.Top = 2500
        frmOBPaymentTransactions.Left = frmMenu.Left + Me.Left + 2655
        frmOBPaymentTransactions.Width = fraContainer.Width
        frmOBPaymentTransactions.Show vbModal
    End Sub
    Private Sub CLOSINGFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
        Call CheckBoxstatus(4)
        lblHeading.Caption = "CASH BOOK CLOSING BALANCE"
        fraGenerateVr.Visible = False
        fraButton.Visible = False
        fraOnline.Visible = False
        fraVerify.Visible = False
        frmOBClosingCashBook.BorderStyle = 0
        frmOBClosingCashBook.Top = 3500
        frmOBClosingCashBook.Left = frmMenu.Left + Me.Left + 2655
        frmOBClosingCashBook.Width = fraContainer.Width
        frmOBClosingCashBook.Show vbModal
    End Sub
    Private Sub VERIFYFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = False
        End If
        Call CheckBoxstatus(5)
        lblHeading.Caption = "VERIFY Receipt & Payment STATEMENT"
        'chkOB.value = vbChecked
        fraOnline.Visible = False
        fraGenerateVr.Visible = False
        fraButton.Visible = True
        fraVerify.Visible = True
        If frmOpeningWizard.mFreeze = 1 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
        fraVerify.Top = fraContainer.Top + (fraContainer.Height / 8)
        cmdSave.Caption = "Verify"
        Call FillVerifyData
    End Sub
    
    Private Sub GENERATEVrFRAME()
        If mFreeze = 1 Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
        Call CheckBoxstatus(6)
        lblHeading.Caption = "GENERATE VOUCHERS OF OPENING ENTRY IN SAANKHYA "
        fraOnline.Visible = False
        fraVerify.Visible = False
        fraGenerateVr.Visible = True
        fraGenerateVr.Top = fraContainer.Top + (fraContainer.Height / 4)
        cmdSave.Caption = "Generate Vr"
      End Sub
    Private Sub cmdSave_Click()
        Dim mSql        As String
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mDate       As String
        
        Dim mOnlineDate As Date
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If FRM_CURPAGE = FRM_ONLINEDATE Then
            If txtOnlineDate.Text <> "" Then
                mOnlineDate = Format(CDate(txtOnlineDate.Text), "dd/mmm/yyyy")
                mDate = "1/" + CStr(Month(mOnlineDate)) + "/" + CStr(Year(mOnlineDate))
                mDate = Format(mDate, "dd/mmm/yyyy")
                mSql = "Update faConfig set dtRPOpeningDate='" & Format(CDate(mDate), "dd/mmm/yyyy") & "'"
                mCnn.Execute (mSql)
                cmdSave.Enabled = False
                cmdNext.Enabled = True
                chkOnline.value = vbChecked
                Call cmdNext_Click
            Else
                MsgBox "Please Select Saankhya Online Date.....", vbApplicationModal
                Exit Sub
            End If
        End If
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If FRM_CURPAGE = FRM_VERIFY Then
            ''------Verify Difference Amount AdjustMent
            Call SaveContraAdj
        End If
        mCnn.Close
        If FRM_CURPAGE = FRM_GENERATEVr Then
            Call GenerateVoucher
            ''------Verify Difference Amount AdjustMent
        End If
    End Sub
    Private Sub SaveContraAdj()
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objCL       As New clsAccounts
        Dim objFun      As New clsFunction
        Dim mSql        As String
        Dim mCnt        As Integer
        Dim mintOBRPTransactionsID  As Double
        Dim AccID       As Integer
        Dim mAccCode    As String
        Dim mArrIn      As Variant
        Dim mFunID      As Integer
        Dim mFunCode    As String
        Dim mAmount     As Double
        If val(vsVerify.TextMatrix(4, 1)) = val(vsVerify.TextMatrix(8, 1)) Then
         
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Delete FROM faOBRPTransactions Where intVoucherTypeID=30 "
        mCnn.Execute (mSql)
        For mCnt = 1 To vsVerifyContra.Rows - 1
            If val(vsVerifyContra.TextMatrix(mCnt, 3)) <> gbAcHeadIDCash Then
                If vsVerifyContra.TextMatrix(mCnt, 2) <> "" And vsVerifyContra.TextMatrix(mCnt, 0) <> "" Then
                    If vsVerifyContra.TextMatrix(mCnt, 4) = "" Or val(vsVerifyContra.TextMatrix(mCnt, 4)) = 0 Then
                        mintOBRPTransactionsID = -1
                    Else
                         mintOBRPTransactionsID = val(vsVerifyContra.TextMatrix(mCnt, 4))
                    End If
                   mAmount = vsVerifyContra.TextMatrix(mCnt, 5)
                   AccID = val(vsVerifyContra.TextMatrix(mCnt, 3))
                   objCL.SetAccounts (AccID)
                   mAccCode = objCL.AccountCode
                   mFunID = 4
                   objFun.SetFunctionByID (mFunID)
                   mFunCode = objFun.FunctionCode
                   mArrIn = Array(mintOBRPTransactionsID, Null, Null, AccID, mAccCode, mFunID, mFunCode, mAmount, _
                   30, 0, 0, Null, Null, "Panchat Head Amount Adjusted" & gbTransactionDate)
                   objdb.ExecuteSP "spSaveOBRPTransactions", mArrIn, , , mCnn, adCmdStoredProc
                End If
            End If
        Next
            MsgBox "Verified Successfully.."
            cmdSave.Enabled = False
            cmdNext.Enabled = True
        Else
            lblVerifyMes.Visible = True
            lblVerifyMes.Caption = "Amount MisMatch .. Total of Receipt Side Does not Match Total of Payment Side"
            cmdSave.Enabled = False
            cmdNext.Enabled = False
            Exit Sub
        End If
    End Sub
    
      Private Sub GenerateVoucher()
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim Rect             As New ADODB.Recordset
        Dim arrInput        As Variant
        Dim arrOutPut       As Variant
        Dim arrOut          As Variant
        Dim mAmount         As Double
        Dim mintKeyID       As Long
        Dim mumVoucherNo    As Variant
        Dim mVoucherID      As Double
        Dim mTransactionID  As Double
        Dim mDrCr           As Integer
        Dim mSql            As String
        Dim mCrDrAmt        As Double
        Dim voucher         As uVoucher
        Dim mAccID          As Integer
        Dim mVrType         As Integer
        Dim mtnyDebitOrCredit     As Integer
        Dim mFundID         As Integer
        Dim mFunctionId     As Integer
        Dim mOBTrnID        As Double
        Dim mClDate         As Date
        Dim tnyVrgpID       As Variant
        Dim mYearID As Integer
        
        'On Error GoTo ErrRollBack:
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        Dim mRPDate As Date
        mSql = "Select dtRPOpeningDate From faConfig "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If IsDate(Rec!dtRPOpeningDate) Then
                mRPDate = Rec!dtRPOpeningDate
                
                If (mRPDate >= "01-Apr-2011" And mRPDate <= "31-Mar-2012") Then
                    mYearID = 2011
                ElseIf (mRPDate >= "01-Apr-2012" And mRPDate <= "31-Mar-2013") Then
                    mYearID = 2012
                Else
                    mYearID = gbFinancialYearID
                End If
            Else
                MsgBox "Opening Receipts & Payments Date is not specified", vbInformation
                Exit Sub
            End If
        End If
        Rec.Close
        
        
        mSql = "Select * From faOBRPTransactions Where ISNull(tnyRecovery,0) <> 1 And intVoucherID is Null"
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn
        
        
        
        ProgressBar1.Min = 0
        ProgressBar1.Max = Rec.RecordCount + 1
        ProgressBar1.value = 0
        ProgressBar1.value = 0
        mFundID = 1
        tnyVrgpID = Null
        mClDate = Format(DateAdd("d", -1, Format(txtOnlineDate.Text, "DD-MMM-YYYY")), "DD-MMM-YYYY")
        
        
        
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                    mOBTrnID = IIf(IsNull(Rec!intOBRPTransactionsID), -1, Rec!intOBRPTransactionsID)
                    mAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    If IsNull(Rec!intVoucherID) Then
                        mVoucherID = -1
                    Else
                        mVoucherID = Rec!intVoucherID
                    End If
                    
                    mAmount = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                    mVrType = IIf(IsNull(Rec!intVoucherTypeID), "", Rec!intVoucherTypeID)
                    mFunctionId = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    If Rec!intVoucherTypeID = 20 Then
                         mtnyDebitOrCredit = 1
                     ElseIf Rec!intVoucherTypeID = 10 Then
                         mtnyDebitOrCredit = 0
                     ElseIf Rec!intVoucherTypeID = 30 Then
                        If mAmount > 0 Then
                         mtnyDebitOrCredit = 0
                         Else
                         mtnyDebitOrCredit = 1
                         End If
                         tnyVrgpID = 2
                     End If
                    '--------------------*********-------------------'
                    '                    faVouchers                  '
                    '--------------------*********-------------------'
                    With voucher
                        .intVoucherID_1 = mVoucherID
                        .intLocalBodyID_2 = gbLocalBodyID
                        .intTransactionID_3 = Null
                        .intTransactionTypeID_4 = 3007
                        .tnyVoucherTypeID_5 = mVrType
                        .intVoucherNo_6 = Null
                        .intBookNo_7 = Null
                        .dtDate_8 = mClDate
                        .fltAmount_9 = Abs(mAmount)
                        .intInstrumentTypeID_10 = 1
                        .vchInstrumentNo_11 = Null
                        .dtInstrumentDate_12 = Null
                        .vchDescription_13 = "Vrs For Panchayat Cash Book "
                        .numZoneID_14 = gbnumZonalID
                        .numWardID_15 = Null
                        .intDoorNoP1_16 = Null
                        .vchDoorNoP2_17 = Null
                        .vchDoorNoP3_18 = Null
                        .intUserID_19 = gbUserID
                        .intCounterID_20 = gbCounterID
                        .numSubLedgerID_21 = Null
                        .intKeyID1_22 = gbAcHeadIDCash
                        .intKeyID2_23 = Null
                        .intExternalApplicationID_24 = 115
                        .intExternalModuleID_25 = 41
                        .intFinancialYearID_26 = mYearID
                        .tnyShiftID_27 = Null
                        .tnyPrintFlag_28 = Null
                        .tnyCancelFlag_29 = Null
                        .vchBank_33 = Null
                        .vchBankPlace_34 = Null
                        .intFundID_35 = mFundID
                        .numSeatID = gbSeatID
                        .intSessionID = Null
                        .vchRefNo = Null
                        .fltRoundOff = Null
                        .fltAdvAmtAdj = Null
                        .numInwardNo = Null
                        .tnyStatus_32 = Null
                        .numLocationID = Null
                        
                        arrInput = Array(.intVoucherID_1, _
                                                .intLocalBodyID_2, _
                                                .intTransactionID_3, _
                                                .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
                                                .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
                                                .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
                                                .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
                                                .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
                                                .intKeyID2_23, .intExternalApplicationID_24, _
                                                .intExternalModuleID_25, .intFinancialYearID_26, _
                                                .tnyShiftID_27, .tnyPrintFlag_28, _
                                                .tnyCancelFlag_29, .vchBank_33, _
                                                .vchBankPlace_34, .intFundID_35, _
                                                .numSeatID, .intSessionID, _
                                                .vchRefNo, .fltRoundOff, _
                                                .fltAdvAmtAdj, .numInwardNo, _
                                                .tnyStatus_32, .numLocationID)
                                
                        objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn, adCmdStoredProc
                        If IsNumeric(arrOutPut(0, 0)) Then
                                mVoucherID = arrOutPut(0, 0)
                                mumVoucherNo = arrOutPut(1, 0)
                        End If
                    End With
                    
                    mCnn.Execute "Update faOBRPTransactions set intVoucherID= " & mVoucherID & ",vchVoucherNo=" & mumVoucherNo & " Where intOBRPTransactionsID =" & mOBTrnID
                    '--------------------*****----------------------------'
                             'VoucherChild
                    '--------------------*****----------------------------'
                    Dim mSlNo                 As Long
                    Dim mintYearID            As Long
                    Dim mtnyPeriodID          As Integer
                    Dim mtnyArrearFlag        As Integer
                    Dim vChild                As uVChild
                    Dim mtnyCrFlag As Integer
                    mCnn.Execute "Delete From faVoucherChild Where intVoucherID =" & mVoucherID
                    With vChild
                        If Rec!intVoucherTypeID = 20 Then
                            mtnyCrFlag = 0
                        ElseIf Rec!intVoucherTypeID = 10 Then
                            mtnyCrFlag = 1
                        ElseIf Rec!intVoucherTypeID = 30 Then
                            If mAmount > 0 Then
                                mtnyCrFlag = 1
                            Else
                                mtnyCrFlag = 0
                            End If
                        End If
                        .intVoucherID_1 = mVoucherID
                        .intLocalBodyID_2 = gbLocalBodyID
                        .intSlNo_3 = 2
                        .tnyDebitOrCredit_5 = mtnyCrFlag
                        .intYearID_6 = Null
                        .tnyPeriodID_7 = Null
                        .tnyArrearFlag_8 = Null
                        .numDemandID_9 = Null
                        .fltAmount_10 = Abs(mAmount)
                        arrInput = Array( _
                                        .intVoucherID_1, _
                                        .intLocalBodyID_2, _
                                         1, _
                                        mAccID, _
                                        .tnyDebitOrCredit_5, _
                                        .intYearID_6, _
                                        .tnyPeriodID_7, _
                                        .tnyArrearFlag_8, _
                                        .numDemandID_9, _
                                        .fltAmount_10 _
                                 )
                        objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                    End With
                    '-------------------------------------'
                    ' Data for Transaction Table          '
                    '-------------------------------------'
                    Dim Trans As uTr
                    Dim mSqlt As String
                    mSqlt = "Select intTransactionID from faTransactions where intVoucherID =" & mVoucherID
                    Rect.Open mSqlt, mCnn
                    If Not (Rect.EOF And Rect.BOF) Then
                        mTransactionID = Rec!intTransactionID
                    Else
                        mTransactionID = -1
                    End If
                    Rect.Close
                    With Trans
                        .intTransactionID = mTransactionID
                        .intLocalBodyID = gbLocalBodyID
                        .intFinancialYearID = mYearID
                        .dtTransactionDate = mClDate
                        .intExternalApplicationID = Null
                        .intExternalApplicationModuleID = 41
                        .intFunctionID = mFunctionId
                        .intFunctionaryID = Null
                        .intFieldID = Null
                        .intFundID = mFundID
                        .intBudgetCentreID = Null
                        .vchNarration = "Vrs For Panchayat Cash Book "
                        .intTransactionTypeID = 3007
                        .intProcessID = Null
                        ''''''
                        If mVrType = 10 Then
                            .vchGroup = "R"
                        ElseIf mVrType = 20 Then
                            .vchGroup = "P"
                        ElseIf mVrType = 30 Then
                            .vchGroup = "C"
                        End If
                        .intGroupID = mVrType
                        ''''''
                        .intKeyID = Null
                        .numSubLedgerID = Null
                        .numUserID = gbUserID
                        .intVoucherID = mVoucherID
                         arrInput = Array(.intTransactionID, _
                        .intLocalBodyID, _
                        .intFinancialYearID, _
                        .dtTransactionDate, _
                        .intExternalApplicationID, _
                        .intExternalApplicationModuleID, _
                        .intFunctionID, _
                        .intFunctionaryID, _
                        .intFieldID, _
                        .intFundID, _
                        .intBudgetCentreID, _
                        .vchNarration, _
                        .intTransactionTypeID, _
                        .intProcessID, _
                        .vchGroup, _
                        .intGroupID, _
                        .intKeyID, _
                        .numSubLedgerID, _
                        .numUserID, _
                        .intVoucherID)
                                                                                                      
                        objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
                             If IsNumeric(arrOutPut(0, 0)) Then
                                mTransactionID = arrOutPut(0, 0)
                             Else
                                'GoTo ErrRollBack:
                             End If
                End With
            
                '----------------------------------------'
                ' Data for TransactionChild    '
                '----------------------------------------'
                  Dim transChild As uTrChild
                  
                  mCnn.Execute "Delete From faTransactionChild Where intTransactionID =" & mTransactionID
                  With transChild
                        If Rec!intVoucherTypeID = 20 Then
                            mtnyCrFlag = 0
                        ElseIf Rec!intVoucherTypeID = 10 Then
                            mtnyCrFlag = 1
                        ElseIf Rec!intVoucherTypeID = 30 Then
                            If mAmount > 0 Then
                                mtnyCrFlag = 0
                            Else
                                mtnyCrFlag = 1
                            End If
                        End If
                        .intTransactionID = mTransactionID
                        .intSerialNo = 1
                        .intAccountHeadID = gbAcHeadIDCash
                        .fltAmount = Abs(mAmount)
                        .tinDebitOrCreditFlag = mtnyCrFlag
                        .intByAccountHeadID = Null
                        .vchNarration = "Cash Book Opening"
                        .intFundID = mFundID
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
                    '----------------------------------------'
                    ' Data for TransactionChild From Grid   '
                    '----------------------------------------'
                   With transChild
                        .intTransactionID = mTransactionID
                        .intSerialNo = 2
                        .intAccountHeadID = mAccID
                        .fltAmount = Abs(mAmount)
                        .tinDebitOrCreditFlag = IIf(mtnyCrFlag = 0, 1, 0)
                        .intByAccountHeadID = gbAcHeadIDCash
                        .vchNarration = "Opening Balance OBRP"
                        .intFundID = mFundID
                                    
                        arrInput = Array(.intTransactionID, _
                                      .intSerialNo, _
                                      .intAccountHeadID, _
                                      .fltAmount, _
                                      .tinDebitOrCreditFlag, _
                                      .intByAccountHeadID, _
                                      .vchNarration, _
                                      .intFundID _
                                      )
                         objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
                  End With
                  ProgressBar1.value = Int(ProgressBar1.value * 100 / ProgressBar1.Max)
            Rec.MoveNext
            Wend
            ProgressBar1.value = ProgressBar1.Max
            lblMessage.Caption = Int(ProgressBar1.value * 100 / ProgressBar1.Max)
            lblMessage.Refresh
            MsgBox "Vouchers Generated Successfully", vbApplicationModal
            cmdSave.Enabled = False
            cmdCancel.Enabled = True
            Call DateRange
          End If
            
            
        Exit Sub
ErrRollBack:
        MsgBox "Saankhya Error" & err.Description
    End Sub
    
    Private Sub FillVerifyData()
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim objAcc      As New clsAccounts
        Dim mCashBalance As Double
        Dim mBankBalance As Double
        Dim mCnt        As Integer
        Dim mCVCnt      As Integer
        Dim mTotOB      As Double
        Dim mTotCl      As Double
        Dim mTotDiff    As Double
        Dim mRecovery   As Double
        Dim mArrClosing As Double
        Dim mActualClosing As Double
        Dim mDiffCash   As Double
        Dim mDiffAmtSum As Double
        vsVerify.MergeCol(0) = True
        vsVerify.MergeCells = flexMergeRestrictRows
        vsVerify.RowHidden(0) = True
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "SELECT faAccountHeads.intGroupID Gr,Sum(fltOpening) OP,Sum(fltClosing) CL  "
        mSql = mSql + " FROM faOBCashBook INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faOBCashBook.intAccountHeadID"
        mSql = mSql + " Group By faAccountHeads.intGroupID"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF Or Rec.BOF)
                If Rec!Gr = 1 Then
                    vsVerify.TextMatrix(1, 1) = IIf(IsNull(Rec!OP), "", Rec!OP) 'Cash Opening
                    vsVerify.TextMatrix(6, 1) = IIf(IsNull(Rec!CL), "", Rec!CL) 'Cash Closing vsVerify.TextMatrix(5, 1)
                Else
                    vsVerify.TextMatrix(2, 1) = IIf(IsNull(Rec!OP), "", Rec!OP) 'Bank/Treasury Opening
                    vsVerify.TextMatrix(7, 1) = IIf(IsNull(Rec!CL), "", Rec!CL) 'Bank/Treasury Closing vsVerify.TextMatrix(6, 1)
                End If
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        mSql = ""
        mSql = " Select sum(fltAmount) Amt,intVouchertypeID Gr From faOBRPTransactions Where intVouchertypeID not in (30)  Group By intVouchertypeID"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF Or Rec.BOF)
                 If Rec!Gr = 10 Then
                      vsVerify.TextMatrix(3, 1) = IIf(IsNull(Rec!Amt), "", Rec!Amt) 'Receipt  ' txtRSubTot.Text
                 ElseIf Rec!Gr = 20 Then
                     vsVerify.TextMatrix(5, 1) = IIf(IsNull(Rec!Amt), "", Rec!Amt) 'Payment vsVerify.TextMatrix(4, 1) txtPSubTot.Text
                 ElseIf Rec!Gr = 0 Then
                     mRecovery = IIf(IsNull(Rec!Amt), "", Rec!Amt) 'Payment vsVerify.TextMatrix(4, 1) txtPSubTot.Text
                 End If
            Rec.MoveNext
            Wend
        End If
        vsVerify.TextMatrix(3, 1) = val(vsVerify.TextMatrix(3, 1)) + mRecovery 'Receipt
        vsVerify.TextMatrix(5, 1) = val(vsVerify.TextMatrix(5, 1)) + mRecovery 'Payment
        Rec.Close
        vsVerify.TextMatrix(4, 1) = (val(vsVerify.TextMatrix(1, 1)) + val(vsVerify.TextMatrix(2, 1)) + val(vsVerify.TextMatrix(3, 1))) 'SubTot 1
        vsVerify.Cell(flexcpBackColor, 4, 1, 4, 1) = &HB0B00F
        vsVerify.Cell(flexcpFontBold, 4, 1, 4, 1) = True
        mCashBalance = (val(vsVerify.TextMatrix(1, 1)) + val(vsVerify.TextMatrix(3, 1))) - val(vsVerify.TextMatrix(5, 1))
        mBankBalance = val(vsVerify.TextMatrix(2, 1))

        vsVerify.TextMatrix(8, 1) = val(vsVerify.TextMatrix(5, 1)) + val(vsVerify.TextMatrix(7, 1)) + val(vsVerify.TextMatrix(6, 1)) 'mCashBalance + mBankBalance  'total
        vsVerify.Cell(flexcpBackColor, 8, 1, 8, 1) = &HB0B00F
        vsVerify.Cell(flexcpFontBold, 8, 1, 8, 1) = True
        mTotOB = mCashBalance + mBankBalance
        mTotCl = val(vsVerify.TextMatrix(1, 2)) + val(vsVerify.TextMatrix(2, 2))
        mTotDiff = mTotOB - mTotCl
        If val(vsVerify.TextMatrix(4, 1)) <> val(vsVerify.TextMatrix(8, 1)) Then
            lblVerifyMes.Visible = True
            lblVerifyMes.Caption = "Amount MisMatch .. Total of Receipt Side Does not Match Total of Payment Side"
            cmdSave.Enabled = False
            cmdNext.Enabled = False
            Exit Sub
        End If
        '''''
        
        '''''   To Find  Transactions that Occured in before Saankhya Online Date
        Dim mClDate As String
        mClDate = Format(txtOnlineDate.Text, "DD-MMM-YYYY")
            mSql = ""
            mSql = "set dateformat dmy Select * From faVouchers Where tnyVoucherTypeID in (10,20,30) And isNull(tnyStatus,0)<>4 And intTransactionTypeID<>3007 And dtdate<'" & mClDate & "'"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF) Then
                lblVerifyMes.Visible = True
                lblVerifyMes.Caption = "Some Transactions Done Before Saankhya Online Date..Please Check And Correct it.."
                cmdSave.Enabled = False
                cmdNext.Enabled = False
                Exit Sub
            End If
            Rec.Close
        '''''
        
        '--------Fill Hidden grid To Find The Difference Amount
        mSql = ""
        mSql = "Select faOBCashBook.intAccountHeadID,faOBCashBook.vchAccountHeadCode,fltOpening,fltClosing, fltClosing ArCl From faOBCashBook "
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        If Rec.RecordCount > 0 Then
            vsVerifyCB.Rows = 1
            If Not (Rec.EOF And Rec.BOF) Then
                vsVerifyCB.Rows = Rec.RecordCount + 1
                vsVerifyCB.Col = 0
                vsVerifyCB.Row = 1
                vsVerifyCB.ColSel = 5
                vsVerifyCB.RowSel = vsVerifyCB.Rows - 1
                mSql = Rec.GetString(, , vbTab, Chr(13))
                vsVerifyCB.Clip = mSql
            End If
        End If
        
        '------------------------------------------------------
        Rec.Close
        If vsVerifyCB.FindRow(gbAcHeadIDCash, , 0, 1, 1) > 0 Then
            vsVerifyCB.TextMatrix(vsVerifyCB.FindRow(gbAcHeadIDCash, , 0, 1, 1), 4) = mCashBalance
        End If
      
      
        For mCnt = 1 To vsVerifyCB.Rows - 1
            'vsVerifyCB.TextMatrix(mCnt, 5) = CLng(val(vsVerifyCB.TextMatrix(mCnt, 4)) - val(vsVerifyCB.TextMatrix(mCnt, 2))) 'mDiff Amount
            vsVerifyCB.TextMatrix(mCnt, 5) = CDbl(val(vsVerifyCB.TextMatrix(mCnt, 4)) - val(vsVerifyCB.TextMatrix(mCnt, 2))) 'mDiff Amount ''' CHANGED BY AIBY on 19-OCT-2013
            If val(vsVerifyCB.TextMatrix(mCnt, 0)) <> gbAcHeadIDCash Then
                mDiffAmtSum = mDiffAmtSum + CDbl(val(vsVerifyCB.TextMatrix(mCnt, 5))) 'Total Bank Amount to Adjust with cash
            End If
        Next
        
        
        
        
        '''---- Cash Closing----------------
        Dim mContraAdjTot   As Double
        Dim mContraAdj      As Double
        Dim mTotalCash      As Double
        Dim mTotalBank      As Double
        Dim mBankCl         As Double
        Dim mClashCl        As Double
        Dim mBankDiff       As Double
        Dim mCashDiff       As Double
'        Dim mContraAdj      As Double
        mArrClosing = mCashBalance + mBankBalance
        mActualClosing = val(vsVerify.TextMatrix(6, 1)) + val(vsVerify.TextMatrix(7, 1))  'Closing of cash And Bank
        vsVerifyContra.Rows = 1
        mClashCl = val(vsVerify.TextMatrix(6, 1))
        mBankCl = val(vsVerify.TextMatrix(7, 1))
        mCashDiff = Abs(mCashBalance - mClashCl)
        If val(mArrClosing) = val(mActualClosing) And Abs(mCashDiff) = Abs(val(mDiffAmtSum)) Then
        'If val(mArrClosing) = val(mActualClosing) Then 'And Abs(mCashDiff) = Abs(val(mDiffAmtSum))
            mCVCnt = 1
            mCashDiff = mCashBalance '- mClashCl
            If Abs(mCashDiff) > 0 Then
                For mCnt = 1 To vsVerifyCB.Rows - 1
                If val(vsVerifyCB.TextMatrix(mCnt, 0)) <> gbAcHeadIDCash Then
                    If val(vsVerifyCB.TextMatrix(mCnt, 5)) <> 0 Then
                        vsVerifyContra.AddItem ("")
                        objAcc.SetAccountID (val(vsVerifyCB.TextMatrix(mCnt, 0)))
                        vsVerifyContra.TextMatrix(mCVCnt, 0) = val(vsVerifyCB.TextMatrix(mCnt, 1)) 'AccHeadCode
                        vsVerifyContra.TextMatrix(mCVCnt, 1) = objAcc.AccountHead                  'AccHead
                        vsVerifyContra.TextMatrix(mCVCnt, 2) = val(vsVerifyCB.TextMatrix(mCnt, 5)) 'Amount
                        vsVerifyContra.TextMatrix(mCVCnt, 5) = val(vsVerifyCB.TextMatrix(mCnt, 5))  'Amount
                        vsVerifyContra.TextMatrix(mCVCnt, 3) = val(vsVerifyCB.TextMatrix(mCnt, 0)) 'AccHeadID
                        mCVCnt = mCVCnt + 1
                    End If
                End If
                Next
            End If
           If mFreeze = 1 Then
                cmdSave.Enabled = False
                'cmdNext.Enabled = True
            Else
             lblVerifyMes.Visible = False
             cmdSave.Enabled = True
             cmdNext.Enabled = False
            End If
        Else
            lblVerifyMes.Visible = True
            lblVerifyMes.Caption = " Amount Mismatch.. Please Check .."
            cmdSave.Enabled = False
            cmdNext.Enabled = False
            Exit Sub
        End If
        
        '''---------------------------------
        mCVCnt = 1
       ' vsVerifyContra.Rows = 1
        Dim mTotGrdDiff As Double
        mTotGrdDiff = 0
    End Sub
'    Private Sub IsVoucherGe()
'
'    End Sub
    Private Sub cmdUdoVrGenerate_Click()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        
        cmdUdoVrGenerate.Enabled = False
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Delete From faVoucherChild Where intVoucherID in (Select intVoucherID From faVouchers Where intTransactionTypeID=3007)" '(Select intVoucherID From faOBRPTransactions Where intVoucherID is not Null)"
        mCnn.Execute mSql
        mSql = ""
        mSql = "Delete From faTransactionChild Where intTransactionID in("
        mSql = mSql + " Select intTransactionID From faTransactions Where intVoucherID in (Select intVoucherID From faVouchers Where intTransactionTypeID=3007))" '(Select intVoucherID From faOBRPTransactions Where intVoucherID is not Null))"
        mCnn.Execute mSql
        mSql = ""
        mSql = "Delete From faTransactions Where intVoucherID in (Select intVoucherID From faVouchers Where intTransactionTypeID=3007)" '(Select intVoucherID From faOBRPTransactions Where intVoucherID is not Null)"
        mCnn.Execute mSql
        mSql = ""
        mSql = "Delete From faVouchers Where intVoucherID in (Select intVoucherID From faVouchers Where intTransactionTypeID=3007)" '(Select intVoucherID From faOBRPTransactions Where intVoucherID is not Null)"
        mCnn.Execute mSql
        mSql = ""
        mSql = "Update faOBRPTransactions set intVoucherID=null,vchVoucherNo=Null"
        mCnn.Execute mSql
        MsgBox "All R&P Vouchers cancelled ...", vbInformation
        Unload Me
    End Sub

    Private Sub dtpOnlinedate_CloseUp()
        Dim mDate As Date
        mDate = Format(dtpOnlinedate.value, "DD-MMM-YYYY")
        If Day(mDate) = 1 Then
            mDate = "1/" + CStr(Month(mDate)) + "/" + CStr(Year(mDate))
            txtOnlineDate.Text = Format(mDate, "DD-MMM-YYYY")
        Else
            txtOnlineDate.Text = ""
            MsgBox "You are not Selected First Date of the Month..", vbInformation
            mDate = "1/" + CStr(Month(mDate)) + "/" + CStr(Year(mDate))
            txtOnlineDate.Text = Format(mDate, "DD-MMM-YYYY")
            Exit Sub
        End If
    End Sub
    Private Sub DateRange()
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim TrMinDate   As Date
        Dim OpDate      As Date
        Dim mLastDay    As Date
        Dim mFirstDay   As Date
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mSql = " Select Convert(varchar(12),Min(dtdate),103) as MinDate "
'        mSql = mSql + " Convert(varchar(12),'1/'+Convert(varchar(5),(Month(Min(dtdate))))+'/'+Convert(varchar(5),Year(Min(dtdate))),103) as StartDate"
'        mSql = mSql + " From faVouchers Where Convert(varChar(15),dtdate,103)= Convert(varChar(15),dtTimeStamp,103)"
        mSql = " Select Convert(varchar(12),Min(dtdate),103) as MinDate,"
        mSql = mSql + " Convert(varchar(12),'1/'+Convert(varchar(5),(Month(Min(dtdate))))+'/'+Convert(varchar(5),Year(Min(dtdate))),103) as StartDate"
        mSql = mSql + " From faVouchers Where Convert(varChar(15),dtdate,103)= Convert(varChar(15),dtTimeStamp,103)"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            TrMinDate = Format(IIf(IsNull(Rec!MinDate), gbStartingDate, Rec!MinDate), "dd/mmm/yyyy")
            mFirstDay = Format(IIf(IsNull(Rec!StartDate), gbStartingDate, Rec!StartDate), "dd/mmm/yyyy")
        Else
            TrMinDate = gbStartingDate
            mFirstDay = gbStartingDate
        End If
        Rec.Close
        mSql = ""
        mSql = "Select top 1 dateadd(d,1,dtTransactionDate) dtTransactionDate From faTransactions Where intTransactionTypeID = 3000"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            OpDate = IIf(IsNull(Rec!dtTransactionDate), "", Rec!dtTransactionDate)
        End If
        Rec.Close
        
        dtpOnlinedate.MinDate = mFirstDay 'Format("1/" + CStr(Month(TrMinDate)) + "/" + CStr(Year(TrMinDate)), "dd/mmm/yyyy") 'OpDate"
        'dtpOnlinedate.MaxDate = TrMinDate
        mLastDay = DateAdd("m", 1, mFirstDay) 'Format("1/" + CStr(Month(TrMinDate) + 1) + "/" + CStr(Year(TrMinDate)), "dd/mmm/yyyy")
        dtpOnlinedate.MaxDate = DateAdd("d", -1, mLastDay)
        dtpOnlinedate.CustomFormat = "dd/mmm/yyyy"
        dtpOnlinedate.value = Format(TrMinDate, "dd/mmm/yyyy")

        lblObTrans.Caption = "Double Entry Opening Balance Sheet Entry Date :-" & Format(OpDate, "DD-MMM-YYYY")
        lblActual.Caption = "First Online Transaction done through Saankhya Double Entry is On :-" & Format(TrMinDate, "DD-MMM-YYYY") _
        & vbNewLine & " Select First Date of that month"

        mSql = ""
        mSql = "Select count(*) mCnt From faOBRPTransactions Where isNull(tnyRecovery,0)<>1 And isNull(intVoucherID,0)<>0"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!mCnt > 0 Then
                mFreeze = 1
                cmdSave.Enabled = False
    
            ''''Comented on 27 aug 2014 To block OBRP regeneration
'                cmdUdoVrGenerate.Visible = True
            Else
                mFreeze = 0
                cmdUdoVrGenerate.Visible = False
            End If
        End If
        Rec.Close
    End Sub
    Private Sub CheckBoxstatus(mFlag As Integer)
        Select Case mFlag
            Case 0 'OnLine date
                    chkOnline.value = vbUnchecked
                    chkOB.value = vbUnchecked
                    chkR.value = vbUnchecked
                    chkP.value = vbUnchecked
                    chkCB.value = vbUnchecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 1 'FRM_Opening
                    chkOnline.value = vbChecked
                    chkOB.value = vbUnchecked
                    chkR.value = vbUnchecked
                    chkP.value = vbUnchecked
                    chkCB.value = vbUnchecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 2 'Receipt
                    chkOnline.value = vbChecked
                    chkOB.value = vbChecked
                    chkR.value = vbUnchecked
                    chkP.value = vbUnchecked
                    chkCB.value = vbUnchecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 3 'Payment
                    chkOnline.value = vbChecked
                    chkOB.value = vbChecked
                    chkR.value = vbChecked
                    chkP.value = vbUnchecked
                    chkCB.value = vbUnchecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 4 'Closing
                    chkOnline.value = vbChecked
                    chkOB.value = vbChecked
                    chkR.value = vbChecked
                    chkP.value = vbChecked
                    chkCB.value = vbUnchecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 5 'Verify
                    chkOnline.value = vbChecked
                    chkOB.value = vbChecked
                    chkR.value = vbChecked
                    chkP.value = vbChecked
                    chkCB.value = vbChecked
                    chkVerify.value = vbUnchecked
                    chkVr.value = vbUnchecked
            Case 6 'FRM_GENERATEVr
                    chkOnline.value = vbChecked
                    chkOB.value = vbChecked
                    chkR.value = vbChecked
                    chkP.value = vbChecked
                    chkCB.value = vbChecked
                    chkVerify.value = vbChecked
                    chkVr.value = vbUnchecked
        End Select
    End Sub

    Private Sub Form_Activate()
'        Me.Top = 0
'        Me.Left = (Screen.Width - Me.Width) / 2
    End Sub
    
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        FRM_ARRAY_TITLES(FRM_ONLINEDATE) = FRM_ONLINEDATE_TITLE
        FRM_ARRAY_TITLES(FRM_OPENING) = FRM_OPENING_TITLE
        FRM_ARRAY_TITLES(FRM_RECEIPT) = FRM_RECEIPT_TITLE
        FRM_ARRAY_TITLES(FRM_PAYMENT) = FRM_PAYMENT_TITLE
        FRM_ARRAY_TITLES(FRM_CLOSING) = FRM_CLOSING_TITLE
        FRM_ARRAY_TITLES(FRM_VERIFY) = FRM_VERIFY_TITLE
        FRM_ARRAY_TITLES(FRM_GENERATEVr) = FRM_GENERATEVr_TITLE
'        FRM_ARRAY_TITLES(FRM_FINISH) = FRM_FINISH_TITLE
        FRM_ARRAY_TITLES(FRM_NUMBEROFPAGES) = FRM_GENERATEVr_TITLE
        FRM_ARRAY_PAGES(FRM_ONLINEDATE) = True 'Tell the Array it is.
        FRM_CURPAGE = FRM_ONLINEDATE
        Me.Top = 0
        Me.Left = (Screen.Width - Me.Width) / 2
        vsVerify.MergeRow(0) = True
        vsVerify.MergeCells = flexMergeRestrictRows
        cmdNext.Enabled = False
        Call DateRange
        Call FillDate
        fraOnline.Visible = True
        fraOnline.Top = fraContainer.Top + (fraContainer.Height / 4)
        fraOnline.Left = fraContainer.Left
        'fraOnline.Width = fraContainer.Width
        fraVerify.Visible = False
        Call GetLastPostingYear
    End Sub
    
    Private Sub FillDate()
        Dim mSql        As String
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim objdb       As New clsDB
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If FRM_CURPAGE = FRM_ONLINEDATE Then
            cmdPre.Visible = False
            mSql = "Select dtRPOpeningDate From faConfig"
            'Set Rec = objDB.ExecuteSP(mSql, , , , mCnn, adCmdText)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If IsNull(Rec!dtRPOpeningDate) = False Then
                    txtOnlineDate.Text = Format(IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate), "DD-MMM-YYYY")
'                    dtpOnlinedate.Enabled = False
'                    txtOnlineDate.Enabled = False
'                    cmdSave.Enabled = False
                    chkOnline.value = vbChecked
                    cmdNext.Enabled = True
                    lblDateMessage.Visible = True
                    lblDateMessage.Caption = "Already Saved Online Date...  " + CStr(Format(IIf(IsNull(Rec!dtRPOpeningDate), "", Rec!dtRPOpeningDate), "DD-MMM-YYYY"))
                End If
            End If
            Rec.Close
        End If
    End Sub
    Public Property Let FrameNo(mData As Integer)
        mvarFrameNo = mData
    End Property
    Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If UnloadMode = 0 Then
            If cmdCancel.Caption = "Cancel" Then
                If MsgBox("You haven't Finished the Opening cash Book Wizard, are you sure.. you want to quit? ", vbQuestion + vbYesNo, "Close Wizard") = vbYes Then
 '                    If FormLoad("frmOBCashBook") Then
'                        Unload frmOBCashBook
''                    ElseIf IsLoaded("frmOBReceiptTransactions") Then
''                        Unload frmOBReceiptTransactions
''                    ElseIf IsLoaded("frmOBReceiptTransactions") Then
''                        Unload frmOBReceiptTransactions
''                    ElseIf IsLoaded("frmOBPaymentTransactions") Then
''                       Unload frmOBPaymentTransactions
''                    ElseIf IsLoaded("frmOBClosingCashBook") Then
''                        Unload frmOBClosingCashBook
'                    End If
                Else
                   Cancel = vbCancel 'pressed No
                End If
            End If
        Else
            Unload Me
            Unload frmOBCashBook
            Unload frmOBReceiptTransactions
            Unload frmOBPaymentTransactions
            Unload frmOBClosingCashBook
        End If
    End Sub
    Private Sub txtActualCL_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    Private Sub txtArriverCL_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub

    Private Sub txtPSubTot_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub

    Private Sub txtRSubTot_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub

    Private Sub GetLastPostingYear()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim mPostingYearID As Integer
        Dim mCount As Integer

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "SELECT  intFinYearID FROM faPostingIndex WHERE tnyStage=2 AND tnyVerifyCash=1 AND tnyVerifyBS=1"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            mPostingYearID = Rec!intFinYearID
            mCount = 1
        Else
            mCount = 0
        End If
        
        If mCount = 1 Then
             cmdUdoVrGenerate.Enabled = False
        Else
            cmdUdoVrGenerate.Enabled = True
        End If
        Rec.Close
        mCnn.Close
    End Sub

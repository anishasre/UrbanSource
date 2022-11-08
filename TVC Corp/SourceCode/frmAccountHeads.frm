VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAccountHeadsNew 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AccountHeads"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Type of Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   210
      TabIndex        =   3
      Top             =   750
      Width           =   1965
      Begin VB.OptionButton optAsset 
         Appearance      =   0  'Flat
         Caption         =   "Asset"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   7
         Top             =   1125
         Width           =   1470
      End
      Begin VB.OptionButton optLiability 
         Appearance      =   0  'Flat
         Caption         =   "Liability"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   6
         Top             =   870
         Width           =   1470
      End
      Begin VB.OptionButton optExpenditure 
         Appearance      =   0  'Flat
         Caption         =   "Expenditure"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   5
         Top             =   600
         Width           =   1470
      End
      Begin VB.OptionButton optIncome 
         Appearance      =   0  'Flat
         Caption         =   "Income"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5700
      TabIndex        =   2
      Top             =   5970
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   5970
      Width           =   1155
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   3300
      TabIndex        =   0
      Top             =   5970
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5775
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10186
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   -2147483624
      TabCaption(0)   =   "MajorHeads"
      TabPicture(0)   =   "AccountHeadsNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "VSGridMajor"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "MinorHeads"
      TabPicture(1)   =   "AccountHeadsNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "VSGridMinor"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "DetailedHeads"
      TabPicture(2)   =   "AccountHeadsNew.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblProgress"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "VSGridDetail"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "pbAccHead"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame3 
         Height          =   2295
         Left            =   -74940
         TabIndex        =   40
         Top             =   360
         Width           =   10575
         Begin VB.TextBox txtTrimFirstDigit 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3330
            MaxLength       =   9
            TabIndex        =   49
            Top             =   810
            Width           =   225
         End
         Begin VB.TextBox txtMajorHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            TabIndex        =   42
            Top             =   810
            Width           =   4695
         End
         Begin VB.TextBox txtMajorCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3570
            MaxLength       =   8
            TabIndex        =   41
            Top             =   810
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2280
            TabIndex        =   43
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2235
         Left            =   -74940
         TabIndex        =   31
         Top             =   330
         Width           =   10575
         Begin VB.TextBox txtMinorCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3990
            MaxLength       =   6
            TabIndex        =   37
            Top             =   1050
            Width           =   990
         End
         Begin VB.TextBox txtMinorHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4995
            TabIndex        =   36
            Top             =   1050
            Width           =   4695
         End
         Begin VB.TextBox txtMajorCodeFromMinor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3420
            TabIndex        =   35
            Top             =   750
            Width           =   1545
         End
         Begin VB.TextBox txtMajorHeadFromMinor 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4980
            TabIndex        =   34
            Top             =   750
            Width           =   4695
         End
         Begin VB.CommandButton cmdMajorSearchFromMinor 
            Caption         =   "..."
            Height          =   300
            Left            =   9720
            TabIndex        =   33
            Top             =   750
            Width           =   375
         End
         Begin VB.TextBox txtTrimMajor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3420
            TabIndex        =   32
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minor Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2370
            TabIndex        =   39
            Top             =   1110
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2370
            TabIndex        =   38
            Top             =   810
            Width           =   1005
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2235
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   10545
         Begin VB.CommandButton cmdDetailSearch 
            Caption         =   "..."
            Height          =   300
            Left            =   9900
            TabIndex        =   50
            Top             =   870
            Width           =   375
         End
         Begin VB.TextBox txtDetailedCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   24
            Top             =   870
            Width           =   750
         End
         Begin VB.CommandButton cmdSearchMajorByDetail 
            Caption         =   "..."
            Height          =   300
            Left            =   9900
            TabIndex        =   23
            Top             =   270
            Width           =   375
         End
         Begin VB.TextBox txtDetailedHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5340
            TabIndex        =   22
            Top             =   870
            Width           =   4515
         End
         Begin VB.TextBox txtSchedule 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3750
            TabIndex        =   21
            Top             =   1170
            Width           =   2400
         End
         Begin VB.TextBox txtOpeningBalance 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3750
            TabIndex        =   20
            Top             =   1470
            Width           =   2400
         End
         Begin VB.OptionButton optDebit 
            Appearance      =   0  'Flat
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6210
            TabIndex        =   19
            Top             =   1530
            Width           =   810
         End
         Begin VB.OptionButton optCredit 
            Appearance      =   0  'Flat
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   7050
            TabIndex        =   18
            Top             =   1530
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox txtAlias 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3750
            TabIndex        =   17
            Top             =   1770
            Width           =   2400
         End
         Begin VB.TextBox txtMinorHeadByDetail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5340
            TabIndex        =   16
            Top             =   570
            Width           =   4515
         End
         Begin VB.TextBox txtMinorCodeByDetail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3750
            TabIndex        =   15
            Top             =   570
            Width           =   1560
         End
         Begin VB.TextBox txtMajorCodeByDetail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3750
            TabIndex        =   14
            Top             =   270
            Width           =   1560
         End
         Begin VB.TextBox txtMajorHeadByDetail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5340
            TabIndex        =   13
            Top             =   270
            Width           =   4515
         End
         Begin VB.CommandButton cmdSearchMinorByDetail 
            Caption         =   "..."
            Height          =   300
            Left            =   9900
            TabIndex        =   12
            Top             =   570
            Width           =   375
         End
         Begin VB.TextBox txtTrimMinor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3750
            MaxLength       =   5
            TabIndex        =   11
            Top             =   870
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Detailed Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2490
            TabIndex        =   30
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2925
            TabIndex        =   29
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Balance"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2235
            TabIndex        =   28
            Top             =   1470
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alias"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3270
            TabIndex        =   27
            Top             =   1800
            Width           =   435
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minor Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2685
            TabIndex        =   26
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major Head"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2700
            TabIndex        =   25
            Top             =   300
            Width           =   1005
         End
      End
      Begin MSComctlLib.ProgressBar pbAccHead 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   4830
         Visible         =   0   'False
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VSFlex8LCtl.VSFlexGrid VSGridMajor 
         Height          =   2655
         Left            =   -73830
         TabIndex        =   44
         Top             =   2760
         Width           =   8565
         _cx             =   15108
         _cy             =   4683
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"AccountHeadsNew.frx":0054
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
      Begin VSFlex8LCtl.VSFlexGrid VSGridMinor 
         Height          =   2715
         Left            =   -73800
         TabIndex        =   45
         Top             =   2700
         Width           =   8565
         _cx             =   15108
         _cy             =   4789
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"AccountHeadsNew.frx":0106
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
      Begin VSFlex8LCtl.VSFlexGrid VSGridDetail 
         Height          =   1995
         Left            =   1230
         TabIndex        =   46
         Top             =   2730
         Width           =   8565
         _cx             =   15108
         _cy             =   3519
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"AccountHeadsNew.frx":01D0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minor Head"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74040
         TabIndex        =   48
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   1200
         TabIndex        =   47
         Top             =   5340
         Width           =   8625
      End
   End
End
Attribute VB_Name = "frmAccountHeadsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
       Dim mSearchID As Variant
       Dim mGroupID As Variant
       Dim mEditFlag As Boolean
       Dim mEditFlg As Boolean
       Dim mFlag As Variant
    Sub OnLostFocusQuery(Query As String, txtHead As TextBox)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        mCnn.ConnectionString = objDB.GetConnectionString(enuSourceString.Saankhya)
        mCnn.Open
        Rec.Open Query, mCnn
        If Not Rec.EOF Then
            mFlag = Rec.Fields(0)
            txtHead.Text = Rec.Fields(2)
        Else
            mFlag = Null
        End If
    End Sub
    
   
   'For Clearing all variables

    Private Sub FormInitialize()
        txtMajorCode.Text = ""
        txtMajorCode.Tag = -1
        txtMajorHead.Text = ""
        txtMajorHead.Tag = -1
        txtMinorCode.Text = ""
        txtMinorCode.Tag = -1
        txtMinorHead.Text = ""
        txtMinorHead.Tag = -1
        txtDetailedCode.Text = ""
        txtDetailedCode.Tag = -1
        txtDetailedHead.Text = ""
        txtMinorCodeByDetail.Text = ""
        txtMinorHeadByDetail.Text = ""
        txtMajorCodeByDetail.Text = ""
        txtMajorHeadByDetail.Text = ""
        txtAlias.Text = ""
        txtSchedule.Text = ""
        txtOpeningBalance.Text = ""
        txtMajorCodeFromMinor.Text = ""
        txtMajorHeadFromMinor.Text = ""
        txtTrimMajor.Text = ""
        txtTrimMinor.Text = ""
        optDebit.Value = True
        mEditFlag = False
        lblProgress.Caption = ""
        txtMinorCodeByDetail.Tag = -1
        txtTrimFirstDigit.Text = ""
        VSGridDetail.Rows = 1
        VSGridMajor.Rows = 1
        VSGridMinor.Rows = 1
    End Sub

    'For Searching Major AccountHeads

    Private Sub ShowSearchAccountHead()
        Dim mSQL As String
        Dim mTypeID As Variant
        
        Call FormInitialize
        If optIncome.Value Then mTypeID = 1
        If optExpenditure.Value Then mTypeID = 2
        If optLiability.Value Then mTypeID = 3
        If optAsset.Value Then mTypeID = 4
        If Not IsNumeric(mTypeID) Then mTypeID = Null
                    Select Case mTypeID
                Case 1
                    mSQL = "Select (vchMajorAccountHeadCode + '  ' + vchMajorAccountHead) as AccHead, intMajorAccountHeadID From faMajorAccountHeads Where tinType=1"
                Case 2
                    mSQL = "Select (vchMajorAccountHeadCode + '  ' + vchMajorAccountHead) as AccHead, intMajorAccountHeadID From faMajorAccountHeads Where tinType=2"
                Case 3
                    mSQL = "Select (vchMajorAccountHeadCode + '  ' + vchMajorAccountHead) as AccHead, intMajorAccountHeadID From faMajorAccountHeads Where tinType=3"
                Case 4
                    mSQL = "Select (vchMajorAccountHeadCode + '  ' + vchMajorAccountHead) as AccHead, intMajorAccountHeadID From faMajorAccountHeads Where tinType=4"
             End Select
            frmSearchAccountHeads.SQLString = mSQL
            frmSearchAccountHeads.Show vbModal
    End Sub

    'For canceling the present action and Clearing the form
    
    Private Sub cmdCancel_Click()
        Call FormInitialize
    End Sub

    Private Sub cmdDetailSearch_Click()
        Call SearchDetailedAccountHead
        Call txtDetailedCode_GotFocus
    End Sub

    'For Searching MajorAccountHead From Minor's tab and Show Corresponding
    'Major Head and Code in MinorHead's Tab
    
    Private Sub cmdMajorSearchFromMinor_Click()
        Call ShowSearchAccountHead
        Call txtMajorCodeFromMinor_GotFocus
    End Sub
    
'Defines What will happen in this event at each tab's Window

    Private Sub cmdNew_Click()
        Dim mTypeID As Long
        If SSTab.Tab = 0 Then
            Call FormInitialize
            If optIncome.Value Then mTypeID = 1
            If optExpenditure.Value Then mTypeID = 2
            If optLiability.Value Then mTypeID = 3
            If optAsset.Value Then mTypeID = 4
            txtTrimFirstDigit.Text = mTypeID
        ElseIf SSTab.Tab = 1 Then
            txtMinorCode.Text = ""
            txtMinorCode.Tag = -1
            txtMinorHead.Text = ""
            txtMinorHead.Tag = -1
            mEditFlag = False
            If Val(txtMajorCodeFromMinor.Tag) > -1 Then
                txtTrimMajor.Text = mID(Trim(txtMajorCodeFromMinor.Text), 1, 3)
            End If
        ElseIf SSTab.Tab = 2 Then
            txtDetailedCode.Text = ""
            txtDetailedHead.Text = ""
            txtDetailedCode.Tag = -1
            txtDetailedHead.Tag = -1
            txtOpeningBalance.Text = ""
            txtAlias.Text = ""
            txtSchedule.Text = ""
            optDebit.Value = False
            optCredit.Value = False
            mEditFlag = False
            If Val(txtMinorCodeByDetail.Tag) > -1 Then
                txtTrimMinor.Text = mID(Trim(txtMinorCodeByDetail.Text), 1, 5)
            End If
        End If
    End Sub

'Save Procedure has 3 different functions according to the tab selection
'If Tab 0 is the working pane, then user can, simply edit an exsting MajorHead or create a New one.
'In the case of Tab 1, Creating a new MinorHead Or editing an existing one, under corresponding majorHead can be done-
    '- User is not expected to modify corresponding majorHead at present.
'If the selected working pane is tab 2 then, this procedure does the following things:
    '- Either edit/ save a detailedHead Under Corresponding major and minor heads.
    '- Create a zero'th Transaction ID if an AccountHead's OPening Balance is Fixed initially (One time execution).
    '- Save the Corresponding Accountheads Details(For Opening Balance) in Transaction Child's Table.
    '- At each time when a transaction amount is posted to a particular account head then it's opening balance in transaction child's table and current balance in account heads table will be updated.
 
    Private Sub cmdSave_Click()
            Dim objDB                   As New clsDB
            Dim objAcc                  As New clsAccounts
            Dim objAc                   As New clsAccounts
            Dim mCnn                    As New ADODB.Connection
            Dim mDetailedAccountHeadID  As Long
            Dim mMajorAccountHeadID     As Long
            Dim mMajorCode              As String
            Dim mMajorHead              As String
            Dim mMinorAccountHeadID     As Long
            Dim mDetailedCode           As String
            Dim mSecondaryCode          As String
            Dim mTypeID                 As Long
            Dim mAmt                    As Double
            Dim arrInput                As Variant
            Dim Rec                     As New ADODB.Recordset
            Dim mSQL                    As String
            Dim mMinorCode              As String
            Dim mMinorHead              As String
            Dim mScheduleID             As Long
            Dim mDetailedHead           As String
            Dim recAccountHeadID        As New ADODB.Recordset
            
            If SSTab.Tab = 0 Then
                '---------------------------------------------------'
                '  Validations For Major                            '
                '---------------------------------------------------'
                If Val(txtMajorCode.Tag) = -1 Then
                    mEditFlag = False
                    mMajorCode = Trim(txtTrimFirstDigit.Text) & Trim(txtMajorCode.Text)
                    If Trim(txtMajorHead.Text) = "" Then MsgBox "Description Cannot be left Blank", vbInformation, "Saankhya": Exit Sub
                    If Len(txtMajorCode.Text) < 8 Then
                            MsgBox "Enter a code with 9 digits"
                            txtMajorCode.Text = ""
                            txtMajorHead.Text = ""
                            txtMajorCode.SetFocus
                            Exit Sub
                    End If
                    If txtMajorHead.Text = "" Then MsgBox "Account Head Cannot be left blank", vbInformation, "Saankhya"
                    mMajorHead = Trim(txtMajorHead.Text)
                    If optIncome.Value Then mTypeID = 1
                    If optExpenditure.Value Then mTypeID = 2
                    If optLiability.Value Then mTypeID = 3
                    If optAsset.Value Then mTypeID = 4
                ElseIf Val(txtMajorCode.Tag) > -1 And mEditFlag Then
                    mMajorAccountHeadID = Val(txtMajorCode.Tag)
                    mMajorCode = Trim(txtTrimFirstDigit.Text) & Trim(txtMajorCode.Text)
                    If Len(txtMajorCode.Text) < 8 Then
                            MsgBox "Enter a code with 9 digits"
                            txtMajorCode.Text = ""
                            txtMajorHead.Text = ""
                            txtMajorCode.SetFocus
                            Exit Sub
                    End If
                    mMajorHead = Trim(txtMajorHead.Text)
                    If optIncome.Value Then mTypeID = 1
                    If optExpenditure.Value Then mTypeID = 2
                    If optLiability.Value Then mTypeID = 3
                    If optAsset.Value Then mTypeID = 4
                Else
                    mEditFlag = False
                    mMajorCode = Trim(txtMajorCode.Text) & Trim(txtTrimFirstDigit.Text)
                    If Len(txtMajorCode.Text) < 8 Then
                            MsgBox "Enter a code with 9 digits"
                            txtMajorCode.Text = ""
                            txtMajorHead.Text = ""
                            txtMajorCode.SetFocus
                            Exit Sub
                    End If
                    mMajorHead = Trim(txtMajorHead.Text)
                    If optIncome.Value Then mTypeID = 1
                    If optExpenditure.Value Then mTypeID = 2
                    If optLiability.Value Then mTypeID = 3
                    If optAsset.Value Then mTypeID = 4
                End If
                '---------------------------------------------------'
                '  Updating                                         '
                '---------------------------------------------------'
                objDB.SetConnection mCnn
                arrInput = Array(IIf(mEditFlag, mMajorAccountHeadID, Null), _
                                    mMajorCode, _
                                    mMajorHead, _
                                    mTypeID _
                                    )
                                    arrInput(0) = mFlag
                objDB.ExecuteSP "spSaveMajorAccountHeads", arrInput, , , mCnn
                Call FormInitialize
                Call FillVSGridMajor(mTypeID)
                
            ElseIf SSTab.Tab = 1 Then
                '---------------------------------------------------'
                '  Validations For Minor                            '
                '---------------------------------------------------'
                If txtMajorCodeFromMinor.Tag = -1 Then
                    MsgBox "Select a Major Head", vbInformation
                    Exit Sub
                ElseIf txtMajorCodeFromMinor.Tag > -1 Then
                    mMinorAccountHeadID = Val(txtMinorCode.Tag)
                    If mMinorAccountHeadID = -1 Then
                        mEditFlag = False
                        mMajorAccountHeadID = Val(txtMajorCodeFromMinor.Tag)
                        If txtMinorHead.Text = "" Then MsgBox "Account head cannot be left blank", vbInformation, "Saankhya": Exit Sub
                        If Len(Trim(txtMinorCode.Text)) < 6 Then
                            MsgBox "Enter a code with 6 digits"
                            txtMinorCode.Text = ""
                            txtMinorHead.Text = ""
                            txtMinorCode.SetFocus
                            Exit Sub
                        End If
                        txtTrimMajor.Text = mID(Trim(txtMajorCodeFromMinor.Text), 1, 3)
                        mMinorCode = Trim(txtTrimMajor.Text) & Trim(txtMinorCode.Text)
                        mMinorHead = Trim(txtMinorHead.Text)
                        mMinorAccountHeadID = Val(txtMinorCode.Tag)
                        If optIncome.Value Then mTypeID = 1
                        If optExpenditure.Value Then mTypeID = 2
                        If optLiability.Value Then mTypeID = 3
                        If optAsset.Value Then mTypeID = 4
                    Else
                        mEditFlag = True
                        mMajorAccountHeadID = Val(txtMajorCodeFromMinor.Tag)
                        If Len(Trim(txtMinorCode.Text)) < 6 Then
                            MsgBox "Enter a code with 6 digits"
                            txtMinorCode.Text = ""
                            txtMinorHead.Text = ""
                            txtMinorCode.SetFocus
                            Exit Sub
                        End If
                        
                        mMinorCode = Trim(txtTrimMajor.Text) & Trim(txtMinorCode.Text)
                        mMinorHead = Trim(txtMinorHead.Text)
                        mMinorAccountHeadID = Val(txtMinorCode.Tag)
                        If optIncome.Value Then mTypeID = 1
                        If optExpenditure.Value Then mTypeID = 2
                        If optLiability.Value Then mTypeID = 3
                        If optAsset.Value Then mTypeID = 4
                    End If
                End If
                '---------------------------------------------------'
                '  Updating                                         '
                '---------------------------------------------------'
                objDB.SetConnection mCnn
                arrInput = Array(IIf(mEditFlag, mMinorAccountHeadID, Null), _
                                    mMinorCode, _
                                    mMinorHead, _
                                    mMajorAccountHeadID, _
                                    mTypeID _
                                    )
                arrInput(0) = mFlag
                objDB.ExecuteSP "spSaveMinorAccountHeads", arrInput, , , mCnn
                Call FormInitialize
                Call FillVSGridMinorByMajor(mTypeID, Val(txtMajorCodeFromMinor.Tag))
                
            ElseIf SSTab.Tab = 2 Then
                '---------------------------------------------------'
                '  Validations For Detail                           '
                '---------------------------------------------------'
                If Val(txtMajorCodeByDetail.Tag) = -1 Then
                    MsgBox "Select A major Head", vbInformation
                    txtMajorCodeByDetail.SetFocus
                    mEditFlg = False
                    Exit Sub
                End If
                If Val(txtMinorCodeByDetail.Tag) = -1 Then
                    MsgBox "Select A minor Head"
                    txtMinorCodeByDetail.SetFocus
                    mEditFlg = False
                    Exit Sub
                End If
                If Trim(txtDetailedHead.Text) = "" Then MsgBox "Detailed Description Cannot be Left Blank", vbInformation, "Saankhya": mEditFlg = False: Exit Sub
                objAcc.SetAccountID (Val(txtDetailedCode.Tag))
                mDetailedAccountHeadID = objAcc.AccountHeadID
                If mDetailedAccountHeadID = -1 Then
                    mEditFlag = False
                    mMajorAccountHeadID = Val(txtMajorCodeByDetail.Tag)
                    mMinorAccountHeadID = Val(txtMinorCodeByDetail.Tag)
                    If optIncome.Value Then mTypeID = 1
                    If optExpenditure.Value Then mTypeID = 2
                    If optLiability.Value Then mTypeID = 3
                    If optAsset.Value Then mTypeID = 4
                    txtTrimMinor.Text = mID(Trim(txtMinorCodeByDetail.Text), 1, 5)
                            If Len(Trim(txtDetailedCode.Text)) < 4 Then
                            MsgBox "Enter a code with 4 digits"
                            txtDetailedCode.Text = ""
                            txtDetailedHead.Text = ""
                            txtDetailedCode.Tag = -1
                            txtDetailedHead.Tag = -1
                            txtOpeningBalance.Text = ""
                            txtAlias.Text = ""
                            txtSchedule.Text = ""
                            optDebit.Value = False
                            optCredit.Value = False
                            If Val(txtMinorCodeByDetail.Tag) > -1 Then
                                txtTrimMinor.Text = mID(Trim(txtMinorCodeByDetail.Text), 1, 5)
                            End If
                            txtDetailedCode.SetFocus
                            mEditFlg = False
                            Exit Sub
                    End If
                    mDetailedCode = Trim(txtTrimMinor.Text) & Trim(txtDetailedCode.Text)
                    mDetailedHead = Trim(txtDetailedHead.Text)
'                    If optCredit.Value Then
'                        mAmt = Abs(Val(txtOpeningBalance.Text)) * -1
'                    Else
                        mAmt = Abs(Val(txtOpeningBalance.Text))
'                    End If
                    txtAlias.Text = Trim(txtAlias.Text)
                    mScheduleID = Val(txtSchedule.Tag)
                Else
                    mEditFlag = True
                    mDetailedAccountHeadID = objAcc.AccountHeadID
                    mMajorAccountHeadID = objAcc.MajorAccountHeadID
                    mMinorAccountHeadID = objAcc.MinorAccountHeadID
                    If optIncome.Value Then mTypeID = 1
                    If optExpenditure.Value Then mTypeID = 2
                    If optLiability.Value Then mTypeID = 3
                    If optAsset.Value Then mTypeID = 4
                    If Len(Trim(txtDetailedCode.Text)) < 4 Then
                            MsgBox "Enter a code with 4 digits"
                            txtDetailedCode.Text = ""
                            txtDetailedHead.Text = ""
                            txtDetailedCode.Tag = -1
                            txtDetailedHead.Tag = -1
                            txtOpeningBalance.Text = ""
                            txtAlias.Text = ""
                            txtSchedule.Text = ""
                            optDebit.Value = False
                            optCredit.Value = False
                            If Val(txtMinorCodeByDetail.Tag) > -1 Then
                                txtTrimMinor.Text = mID(Trim(txtMinorCodeByDetail.Text), 1, 5)
                            End If
                            txtDetailedCode.SetFocus
                            mEditFlag = False
                            mEditFlg = False
                            Exit Sub
                    End If
                    mDetailedCode = Trim(mID(txtMinorCodeByDetail.Text, 1, 5)) & Trim(txtDetailedCode.Text)
                    mDetailedHead = Trim(txtDetailedHead.Text)
                    mTypeID = objAcc.mType
'                    If optCredit.Value Then
'                        mAmt = Abs(Val(txtOpeningBalance.Text)) * -1
'                    Else
                        mAmt = Abs(Val(txtOpeningBalance.Text))
'                    End If
                    
                End If
                '
'            '---------------------------------------------------'
'            '  Updating                                         '
'            '---------------------------------------------------'
            objDB.SetConnection mCnn
                arrInput = Array(IIf(mEditFlag, Val(txtDetailedCode.Tag), Null), _
                            mDetailedCode, _
                            mDetailedHead, _
                            Format(mAmt, "0.00"), _
                            Trim(txtAlias.Text), _
                            mMinorAccountHeadID, _
                            mMajorAccountHeadID, _
                            Val(txtSchedule.Tag), _
                            mTypeID, _
                            gbLocalBodyID, _
                            gbFinancialYearID, _
                            IIf(optDebit, 1, 0), _
                            Format(mAmt, "0.00") _
                            )
                            
                                arrInput(0) = mFlag
           Set recAccountHeadID = objDB.ExecuteSP("spSaveDetailedHead", arrInput, , , mCnn)
                mSQL = "Select Count(*) From faTransactions Where intTransactionID = 0"
                Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
                If Rec.Fields(0).Value = 0 Then
                    Dim intTransactionID_1   As Double
                    Dim mintLocalBodyID_2  As Long
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
                    Dim mintTransactionTypeID_13   As Variant
                    Dim mintVoucherNo_14   As Variant
                    Dim mintProcessID_15    As Variant
                    Dim mintGroupID_17    As Variant
                    Dim mvchGroup_16   As String
                    Dim mintKeyID_18   As Variant
                    Dim mnumSubLedgerID_19    As Variant
                    Dim mintUserID_20  As Variant

                    intTransactionID_1 = 0
                    mintLocalBodyID_2 = gbLocalBodyID
                    mintFinancialYearID_3 = gbFinancialYearID
                    mdtTransactionDate_4 = DdMmmYy(gbStartingDate)
                    mintExternalApplicationID_5 = AppID.Saankhya
                    mintExternalApplicationModuleID_6 = 0
                    mintFunctionID_7 = Null
                    mintFunctionaryID_8 = Null
                    mintFieldID_9 = Null
                    mintFundID_10 = Null
                    mintBudgetCentreID_11 = Null
                    mvchNarration_12 = "Opening Balance"
                    mintTransactionTypeID_13 = Null
                    mintVoucherNo_14 = Null
                    mintProcessID_15 = Null
                    mvchGroup_16 = "JV"
                    mintGroupID_17 = 40
                    mintKeyID_18 = Null
                    mnumSubLedgerID_19 = Null
                    mintUserID_20 = 0

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

                    objDB.ExecuteSP "spSaveTransactions", arrInput, , , mCnn
                End If
                If mAmt > 0 Then
                    If Not recAccountHeadID.EOF Then
                            arrInput = Array(0, _
                            2, _
                            recAccountHeadID.Fields(0), _
                            Format(mAmt, "0.00"), _
                            IIf(optDebit, 1, 0), _
                            Null, _
                            "Opening Balance", _
                            Null _
                            )
                        objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                    End If
                End If
            End If
            If mEditFlag Then
                Call UpdateOPeningBalnce(mDetailedAccountHeadID)
            End If
            Call FormInitialize
            VSGridDetail.Rows = 1
            
    End Sub
    
'For searching MajorHeads from detailed head's tab and to show it.

    Private Sub cmdSearchMajorByDetail_Click()
        Call ShowSearchAccountHead
        Call txtMajorCodeByDetail_GotFocus
    End Sub

'For searching minor heads from detailed heads's tab.

    Private Sub SearchMinorFromDetail()
        Dim mSQL As String
        Dim mTypeID As Variant
        Call FormInitialize
        If optIncome.Value Then mTypeID = 1
        If optExpenditure.Value Then mTypeID = 2
        If optLiability.Value Then mTypeID = 3
        If optAsset.Value Then mTypeID = 4
        If Not IsNumeric(mTypeID) Then mTypeID = Null

                    Select Case mTypeID
                Case 1
                    mSQL = "Select (vchMinorAccountHeadCode + '  ' + vchMinorAccountHead) as AccHead, intMinorAccountHeadID From faMinorAccountHeads Where tinType=1 and intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag)
                Case 2
                    mSQL = "Select (vchMinorAccountHeadCode + '  ' + vchMinorAccountHead) as AccHead, intMinorAccountHeadID From faMinorAccountHeads Where tinType=2 and intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag)
                Case 3
                    mSQL = "Select (vchMinorAccountHeadCode + '  ' + vchMinorAccountHead) as AccHead, intMinorAccountHeadID From faMinorAccountHeads Where tinType=3 and intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag)
                Case 4
                    mSQL = "Select (vchMinorAccountHeadCode + '  ' + vchMinorAccountHead) as AccHead, intMinorAccountHeadID From faMinorAccountHeads Where tinType=4 and intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag)
             End Select
            frmSearchAccountHeads.SQLString = mSQL
            frmSearchAccountHeads.Show vbModal
    End Sub

'call the search function for minor heads from detaild head's tab and show it.

    Private Sub cmdSearchMinorByDetail_Click()
        Call SearchMinorFromDetail
        Call txtMinorCodeByDetail_GotFocus
    End Sub



'Defines the position of the form.

    Private Sub Form_Activate()
        Me.Top = 0
        frmAccountHeadsNew.Left = (frmMenu.Width - Me.Width) / 2
    End Sub

'When the form is loaded initially, then the default enabled tab is majorHeads with it's grid filled with major heads corresponding to the typeID.

    Private Sub Form_Load()
        FormInitialize
        If SSTab.Tab = 0 Then
            Call FormLoadMajor
      End If
    End Sub

'when Asset option is selected from major head's tab, MajorHead's grid will be filled with type=Asset.

    Private Sub optAsset_Click()
        If SSTab.Tab = 0 Then
            If optAsset.Value Then
                Call FillVSGridMajor(4)
                txtTrimFirstDigit.Text = 4
            End If
        End If
    End Sub

'when Expenditure option is selected from major head's tab, MajorHead's grid will be filled with type=expenditure.

    Private Sub optExpenditure_Click()
        If SSTab.Tab = 0 Then
            If optExpenditure.Value Then
                Call FillVSGridMajor(2)
                txtTrimFirstDigit.Text = 2
            End If
        End If
    End Sub

'when Income option is selected from major head's tab, MajorHead's grid will be filled with type=Income.

    Private Sub optIncome_Click()
        If SSTab.Tab = 0 Then
            If optIncome.Value Then
                Call FillVSGridMajor(1)
                txtTrimFirstDigit.Text = 1
            End If
        End If
    End Sub

'when Liablility option is selected from major head's tab, MajorHead's grid will be filled with type=Liablility.

    Private Sub optLiability_Click()
        If SSTab.Tab = 0 Then
            If optLiability.Value Then
                Call FillVSGridMajor(3)
                txtTrimFirstDigit.Text = 3
            End If
        End If
    End Sub

'Defines what should happen when sstab's Zeroth or first tab is clicked.

    Private Sub SSTab_Click(PreviousTab As Integer)
        Dim mSQL As String
        If SSTab.Tab = 0 Then
            Call FormLoadMajor
       ElseIf SSTab.Tab = 1 Then
          VSGridMinor.Clear 0, 1
          cmdMajorSearchFromMinor.SetFocus
        End If
    End Sub
    
    'this event is intented to display the detailed head details when detailed head code is entered in corresponding text box
    '(But at present it is not functioning properly, since, detailed code is conacatinted with first 5 digits of the 9 digit code which is expected to appear in non-editable textbox-
        '- "txtTrimMinor" and remaining 4 digits from the the detailed code's text box.
        'if the textbox "txtTrimMinor" is made editable, then there by it can be possible to display the detailed head details by entering the corresponding 5 digits + the 4 digits of 9 digit code.
        'But, if it is made editable, then the user will get the freedom to modify the first 5 digits of the detailed code also in addition to the last 4 digits of detailed head.)
        
    Private Sub txtDetailedCode_GotFocus()
        Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long
           
           If gbSearchStr <> "" Then
               mEditFlag = True
               Dim mStr As String
               txtDetailedCode.Text = mID(Token(gbSearchStr, " "), 6, 9)
               txtDetailedHead.Text = Trim(gbSearchStr)
               txtDetailedCode.Tag = gbSearchID
               objAc.SetAccounts (gbSearchID)
                            
                    
                    txtTrimMinor.Text = mID(objAc.AccountCode, 1, 5)
                    txtMinorCodeByDetail.Tag = objAc.MinorAccountHeadID
                    txtMinorCodeByDetail.Text = objAc.MinorAccountHeadCode
                    txtMinorHeadByDetail.Text = objAc.MinorAccountHead
                    txtMajorCodeByDetail.Tag = objAc.MajorAccountHeadID
                    txtMajorCodeByDetail.Text = objAc.MajorAccountHeadCode
                    txtMajorHeadByDetail.Text = objAc.MajorAccountHead
                    txtOpeningBalance.Text = objAc.OpeningBalance
                    txtAlias.Text = objAc.Alias
                    mTypeID = objAc.mType
                    If objAc.DebitOrCredit = 1 Then
                        optDebit = True
                    Else
                        optCredit = True
                    End If
           
               gbSearchStr = ""
               gbSearchID = -1
            Else
                Call dispDetailedHead
            End If
            
            
            
'            Dim objAc As New clsAccounts
'           Dim mSQL As String
'           If gbSearchStr <> "" Then
'               Dim mStr As String
'               txtDetailedCode.Text = Trim(Token(gbSearchStr, " "))
'               txtDetailedHead.Text = Trim(gbSearchStr)
'               txtDetailedCode.Tag = gbSearchID
'               objAc.SetAccounts (gbSearchID)
'               txtDetailedHead.Tag = objAc.AccountType
'               txtPrimaryCode.Text = objAc.AccountCode
'               txtMinorByDetailHide.Tag = objAc.MinorAccountHeadID
'               mGroupID = objAc.GroupID
'               txtGroup.Tag = mGroupID
'               gbSearchStr = ""
'               gbSearchID = -1
'               Call DisplayHeadsByDetailedHead
'           End If
'           txtDetailedHead.SelStart = 0
'           txtDetailedHead.SelLength = Len(txtDetailedHead)
            
                
    txtMinorCodeByDetail_GotFocus
    End Sub
    
    ' to allow the user to enter only numbers
    
    Private Sub txtDetailedCode_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

'this event is intented to display the detailed head details when detailed head code is entered in corresponding text box
    '(But at present it is not functioning properly, since, detailed code is conacatinted with first 5 digits of the 9 digit code which is expected to appear in non-editable textbox-
        '- "txtTrimMinor" and remaining 4 digits from the the detailed code's text box.
        'if the textbox "txtTrimMinor" is made editable, then there by it can be possible to display the detailed head details by entering the corresponding 5 digits + the 4 digits of 9 digit code.
        'But, if it is made editable, then the user will get the freedom to modify the first 5 digits of the detailed code also in addition to the last 4 digits of detailed head.)
        
    Private Sub txtDetailedCode_LostFocus()
        Call dispDetailedHead
        Call OnLostFocusQuery("Select * from faAccountHeads where vchAccountHeadCode='" & Trim(mID(txtMinorCodeByDetail.Text, 1, 5)) + Trim(txtDetailedCode.Text) & "'", txtDetailedHead)
    End Sub

'To display the major head detials when correspoding major code is enterd.

    Private Sub txtMajorCode_GotFocus()
           Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long
           
           If gbSearchStr <> "" Then
               Dim mStr As String
               mEditFlag = True
               txtMajorCode.Text = Trim(Token(gbSearchStr, " "))
               txtMajorHead.Text = Trim(gbSearchStr)
               txtMajorCode.Tag = gbSearchID
               objAc.SetAccounts (gbSearchID)
               txtMajorHead.Tag = objAc.AccountType
               txtAlias.Text = objAc.Alias
               gbSearchStr = ""
               gbSearchID = -1
           Else
            Call dispMajorHead
           End If
           txtMajorHead.SelStart = 0
           txtMajorHead.SelLength = Len(txtMajorHead)
    End Sub
    
'To restrict the user from entering alphabets

    Private Sub txtMajorCode_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

'To display the major head detials when correspoding major code is enterd in Major Head's tab.

    Private Sub txtMajorCode_LostFocus()
        If Len(Trim(txtMajorCode.Text)) > 1 And Val(txtMajorCode.Text) > 0 Then
            txtMajorCode.Text = mID(txtMajorCode.Text, 1, 2) + "000000"
        End If
        Call dispMajorHead
        Call OnLostFocusQuery("Select * from faMajorAccountHeads where vchMajorAccountHeadCode='" & txtTrimFirstDigit.Text + txtMajorCode.Text & "'", txtMajorHead)
    End Sub

'To display the major head detials from the detailed head tab.

    Private Sub txtMajorCodeByDetail_GotFocus()
           Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim RecMajorFromMinor As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long
           
           If gbSearchStr <> "" Then
               Dim mStr As String
               txtMajorCodeByDetail.Text = Trim(Token(gbSearchStr, " "))
               txtMajorHeadByDetail.Text = Trim(gbSearchStr)
               txtMajorCodeByDetail.Tag = gbSearchID
               mSQL = "Select * From faMajorAccountHeads Where faMajorAccountHeads.intMajorAccountHeadID= " & gbSearchID
               Set Rec = GetRecordSet(mSQL)
               If Rec.RecordCount > 0 Then
                txtMajorCodeByDetail.Text = Rec!vchMajorAccountHeadCode
                txtMajorCodeByDetail.Tag = Rec!intMajorAccountHeadID
                txtMajorHeadByDetail.Text = Rec!vchMajorAccountHead
                mTypeID = Rec!tinType
               End If
               gbSearchStr = ""
               gbSearchID = -1
            
            ElseIf Val(txtMajorCodeByDetail.Tag) > -1 Then
                mSQL = "Select * From faMajorAccountHeads Where intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag)
                Set Rec = GetRecordSet(mSQL)
                If Rec.RecordCount > 0 Then
                    mEditFlag = True
                         txtMajorCodeByDetail.Text = Rec!vchMajorAccountHeadCode
                         txtMajorCodeByDetail.Tag = Rec!intMajorAccountHeadID
                         txtMajorHeadByDetail.Text = Rec!vchMajorAccountHead
                         mTypeID = Rec!tinType
                End If
            Else
                txtMajorCodeByDetail.Tag = -1
                txtMajorCodeByDetail.Text = ""
                txtMajorHeadByDetail.Text = ""
                mTypeID = 0
           End If
           txtMajorHeadByDetail.SelStart = 0
           txtMajorHeadByDetail.SelLength = Len(txtMajorHeadByDetail)
    End Sub

'To display the major head detials from the Minor head tab.

    Private Sub txtMajorCodeFromMinor_GotFocus()
           Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim RecMajorFromMinor As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long
           
           If gbSearchStr <> "" Then
               Dim mStr As String
               txtMajorCodeFromMinor.Text = Trim(Token(gbSearchStr, " "))
               txtMajorHeadFromMinor.Text = Trim(gbSearchStr)
               txtMajorCodeFromMinor.Tag = gbSearchID
               mSQL = "Select * From faMajorAccountHeads Where faMajorAccountHeads.intMajorAccountHeadID= " & gbSearchID
               Set Rec = GetRecordSet(mSQL)
               If Rec.RecordCount > 0 Then
                txtMajorCodeFromMinor.Text = Rec!vchMajorAccountHeadCode
                txtMajorCodeFromMinor.Tag = Rec!intMajorAccountHeadID
                txtMajorHeadFromMinor.Text = Rec!vchMajorAccountHead
                txtTrimMajor.Text = mID(Rec!vchMajorAccountHeadCode, 1, 3)
                mTypeID = Rec!tinType
               End If
               gbSearchStr = ""
               gbSearchID = -1
            ElseIf Val(txtMajorCodeFromMinor.Tag) > -1 Then
                mSQL = "Select * From faMajorAccountHeads Where intMajorAccountHeadID=" & Val(txtMajorCodeFromMinor.Tag)
                Set Rec = GetRecordSet(mSQL)
                If Rec.RecordCount > 0 Then
                         txtMajorCodeFromMinor.Text = Rec!vchMajorAccountHeadCode
                         txtMajorHeadFromMinor.Text = Rec!vchMajorAccountHead
                         mTypeID = Rec!tinType
                End If
            Else
                txtMajorCodeFromMinor.Tag = -1
                txtMajorCodeFromMinor.Text = ""
                txtMajorHeadFromMinor.Text = ""
                mTypeID = 0
           End If
           txtMajorHeadFromMinor.SelStart = 0
           txtMajorHeadFromMinor.SelLength = Len(txtMajorHead)
           Call FillVSGridMinorByMajor(mTypeID, Val(txtMajorCodeFromMinor.Tag))
    End Sub

'To allow the user to enter only numbers

    Private Sub txtMinorCode_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

'To display the Minor head detials when Minor code is enterd in Minor Head's tab.

    Private Sub txtMinorCode_LostFocus()
        Call dispMinorHead
        If Len(txtMinorCode.Text) > 1 And Val(txtMinorCode.Text) > 0 Then
            txtMinorCode.Text = mID(txtMinorCode.Text, 1, 2) + "0000"
        End If
        Call OnLostFocusQuery("Select * from faMinorAccountHeads where vchMinorAccountHeadCode='" & mID(txtMajorCodeFromMinor.Text, 1, 3) + txtMinorCode.Text & "'", txtMinorHead)
        txtMinorHead.SetFocus
    End Sub

'To display the Minor head detials from the detailed head tab.

    Private Sub txtMinorCodeByDetail_GotFocus()
           Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim RecMajorFromMinor As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long
           
           If gbSearchStr <> "" Then
                   Dim mStr As String
                   txtMinorCodeByDetail.Text = Trim(Token(gbSearchStr, " "))
                   txtMinorHeadByDetail.Text = Trim(gbSearchStr)
                   txtMinorCodeByDetail.Tag = gbSearchID
                   mSQL = "Select * From faMinorAccountHeads Where faMinorAccountHeads.intMinorAccountHeadID= " & gbSearchID
                   Set Rec = GetRecordSet(mSQL)
                   If Rec.RecordCount > 0 Then
                    mEditFlag = True
                    txtMinorCodeByDetail.Text = Rec!vchMinorAccountHeadCode
                    txtMinorHeadByDetail.Text = Rec!vchMinorAccountHead
                   End If
                    mSQL = "SELECT * "
                    mSQL = mSQL & " From faMinorAccountHeads  Left Join "
                    mSQL = mSQL & " faMajorAccountHeads ON faMinorAccountHeads.intMajorAccountHeadID = faMajorAccountHeads.intMajorAccountHeadID "
                    mSQL = mSQL & " WHERE faMinorAccountHeads.intMinorAccountHeadID = " & Val(txtMinorCodeByDetail.Tag)

                    Set Rec = GetRecordSet(mSQL)
                    If Rec.RecordCount > 0 Then
                        mEditFlag = True
                             txtMinorCodeByDetail.Text = Rec!vchMinorAccountHeadCode
                             txtMinorCodeByDetail.Tag = Rec!intMinorAccountHeadID
                             txtMinorHeadByDetail.Text = Rec!vchMinorAccountHead
                             txtMinorHeadByDetail.Tag = Rec!intMajorAccountHeadID
                             mTypeID = Rec!tinType
                             txtMajorCodeByDetail.Text = Rec!vchMajorAccountHeadCode
                             txtMajorHeadByDetail.Text = Rec!vchMajorAccountHead
                    End If
                   gbSearchStr = ""
                   gbSearchID = -1
                 
            ElseIf Val(txtMinorCodeByDetail.Tag) > -1 Then
                    mEditFlag = True
                    mSQL = "SELECT * "
                    mSQL = mSQL & " From faMinorAccountHeads LEFT JOIN "
                    mSQL = mSQL & " faMajorAccountHeads ON faMinorAccountHeads.intMajorAccountHeadID = faMajorAccountHeads.intMajorAccountHeadID "
                    mSQL = mSQL & " WHERE faMinorAccountHeads.intMinorAccountHeadID = " & Val(txtMinorCodeByDetail.Tag)
                  
                    Set Rec = GetRecordSet(mSQL)
                    If Rec.RecordCount > 0 Then
                        mEditFlag = True
                             txtMinorCodeByDetail.Text = Rec!vchMinorAccountHeadCode
                             txtMinorCodeByDetail.Tag = Rec!intMinorAccountHeadID
                             txtMinorHeadByDetail.Text = Rec!vchMinorAccountHead
                             txtMinorHeadByDetail.Tag = Rec!intMajorAccountHeadID
                             mTypeID = Rec!tinType
                             txtMajorCodeByDetail.Text = Rec!vchMajorAccountHeadCode
                             txtMajorHeadByDetail.Text = Rec!vchMajorAccountHead
                    End If
            Else
                    MsgBox "There is no minor Head to create the Detailed Head"
                    mEditFlag = False
                    txtMinorCodeByDetail.Tag = -1
                    txtMinorCodeByDetail.Text = ""
                    txtMinorHeadByDetail.Text = ""
                    mTypeID = 0
                    txtMajorCodeByDetail.Text = ""
                    txtMajorHeadByDetail.Text = ""
                    Exit Sub
            End If
                    Call FillVSGridDetailByMajorandMinor(mTypeID, Val(txtMajorCodeByDetail.Tag), Val(txtMinorCodeByDetail.Tag))
    End Sub

' Defines what should happen when the vsGrid in the detailed tab is clicked.

    Private Sub VSGridDetail_Click()
        If VSGridDetail.MouseRow > 0 Then
            Dim mTypeID As Long
            Dim objDB As New clsDB
            Dim RecMinor As New ADODB.Recordset
            Dim RecMajor As New ADODB.Recordset
            Dim recSchedule As New ADODB.Recordset
            Dim mCon As New ADODB.Connection
                mEditFlag = True
                txtDetailedCode.Text = mID(VSGridDetail.TextMatrix(VSGridDetail.Row, 1), 6, 4)
                txtTrimMinor.Text = mID(VSGridDetail.TextMatrix(VSGridDetail.Row, 1), 1, 5)
                txtDetailedHead.Text = VSGridDetail.TextMatrix(VSGridDetail.Row, 2)
                txtDetailedCode.Tag = VSGridDetail.TextMatrix(VSGridDetail.Row, 8)
                txtOpeningBalance.Text = VSGridDetail.TextMatrix(VSGridDetail.Row, 3)
                txtAlias.Text = VSGridDetail.TextMatrix(VSGridDetail.Row, 4)
                txtMinorCodeByDetail.Tag = VSGridDetail.TextMatrix(VSGridDetail.Row, 5)
                objDB.SetConnection mCon
                RecMinor.Open " select * from faMinorAccountHeads Where intMinorAccountHeadID=" & Val(txtMinorCodeByDetail.Tag), mCon
                If Not RecMinor.EOF Then
                    txtMinorCodeByDetail.Text = RecMinor!vchMinorAccountHeadCode
                    txtMinorHeadByDetail.Text = RecMinor!vchMinorAccountHead
                    txtDetailedCode_LostFocus
                End If
                
                txtMajorCodeByDetail.Tag = VSGridDetail.TextMatrix(VSGridDetail.Row, 6)
                RecMajor.Open "select * from faMajorAccountHeads Where intMajorAccountHeadID=" & Val(txtMajorCodeByDetail.Tag), mCon
                If Not RecMajor.EOF Then
                    txtMajorCodeByDetail.Text = RecMajor!vchMajorAccountHeadCode
                    txtMajorHeadByDetail.Text = RecMajor!vchMajorAccountHead
                End If
                txtSchedule.Tag = VSGridDetail.TextMatrix(VSGridDetail.Row, 7)
                recSchedule.Open "select * from faScheduleReports where intScheduleReportID=" & Val(txtSchedule.Tag), mCon
                If Not recSchedule.EOF Then
                    txtSchedule.Text = recSchedule!vchDescription
                End If
                mTypeID = VSGridDetail.TextMatrix(VSGridDetail.Row, 9)
                'Call txtDetailedCode_GotFocus
            End If
    End Sub

'Defines what should happen when the vsGrid in Major Head's tab is clicked.

    Private Sub VSGridMajor_Click()
            txtMajorCode.Text = mID(VSGridMajor.TextMatrix(VSGridMajor.Row, 1), 2, 8)
            txtMajorCode.Tag = VSGridMajor.TextMatrix(VSGridMajor.Row, 4)
            txtMajorHead.Text = VSGridMajor.TextMatrix(VSGridMajor.Row, 2)
            txtTrimFirstDigit.Text = mID(VSGridMajor.TextMatrix(VSGridMajor.Row, 1), 1, 1)
            txtMajorCode_LostFocus
    End Sub
  
  'Fills vsGrid in Major Head's tab .
  
    Private Sub FillVSGridMajor(mTypeID As Long)
            Dim objDB       As New clsDB
            Dim mCon        As New ADODB.Connection
            Dim Rec         As New ADODB.Recordset
            Dim mLoopCount  As Integer
            Dim mSQL        As String
            Dim mtinType    As Long
            
            
            VSGridMajor.Clear 1, 1   'Clear all rows excluding fixed rows
            VSGridMajor.Rows = 2     'Setting initial row size
            objDB.SetConnection mCon
            
                mSQL = "Select * From faMajorAccountHeads Where tinType =" & mTypeID
                Set Rec = GetRecordSet(mSQL)
                If Rec.RecordCount > 0 Then
                mLoopCount = 0
                While Not Rec.EOF
                     VSGridMajor.TextMatrix(mLoopCount + 1, 0) = mLoopCount + 1
                     VSGridMajor.TextMatrix(mLoopCount + 1, 1) = Rec.Fields(1).Value
                     VSGridMajor.TextMatrix(mLoopCount + 1, 2) = Rec.Fields(2).Value
                     VSGridMajor.TextMatrix(mLoopCount + 1, 3) = Rec.Fields(3).Value
                     VSGridMajor.TextMatrix(mLoopCount + 1, 4) = Rec.Fields(0).Value
                    
                     VSGridMajor.Rows = VSGridMajor.Rows + 1
                     Rec.MoveNext
                     mLoopCount = mLoopCount + 1
                Wend
                    VSGridMajor.Rows = VSGridMajor.Rows - 1
                End If
            
        End Sub
     
     ' this function fills the major head according to the type selected.
     
        Private Sub FormLoadMajor()
            Call FormInitialize
            If optIncome.Value Then
                Call FillVSGridMajor(1)
                txtTrimFirstDigit.Text = 1
            ElseIf optExpenditure.Value Then
                Call FillVSGridMajor(2)
                txtTrimFirstDigit.Text = 2
            ElseIf optLiability.Value Then
                Call FillVSGridMajor(3)
                txtTrimFirstDigit.Text = 3
            ElseIf optAsset.Value Then
                Call FillVSGridMajor(4)
                txtTrimFirstDigit.Text = 4
            End If
            
        End Sub
 
     'Fills vsGrid in Minor Head's tab .
 
    Private Sub FillVSGridMinor(mTypeID As Long)
            Dim objDB       As New clsDB
            Dim mCon        As New ADODB.Connection
            Dim Rec         As New ADODB.Recordset
            Dim mLoopCount  As Integer
            Dim mSQL        As String
            Dim mtinType    As Long
            
            
            VSGridMinor.Clear 1, 1   'Clear all rows excluding fixed rows
            VSGridMinor.Rows = 2     'Setting initial row size
            objDB.SetConnection mCon
            
                mSQL = "Select * From faMinorAccountHeads Where tinType =" & mTypeID
                Set Rec = GetRecordSet(mSQL)
                If Rec.RecordCount > 0 Then
                mLoopCount = 0
                While Not Rec.EOF
                     VSGridMinor.TextMatrix(mLoopCount + 1, 0) = mLoopCount + 1
                     VSGridMinor.TextMatrix(mLoopCount + 1, 1) = Rec.Fields(1).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 2) = Rec.Fields(2).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 3) = Rec.Fields(0).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 4) = Rec.Fields(3).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 5) = Rec!Fields(4).Value
                     VSGridMinor.Rows = VSGridMinor.Rows + 1
                     Rec.MoveNext
                     mLoopCount = mLoopCount + 1
                Wend
                    VSGridMinor.Rows = VSGridMinor.Rows - 1
                End If
        End Sub
  
  'Defines what should happen when the vsGrid in minor head's tab is clicked
  
        Private Sub VSGridMinor_Click()
            Dim mTypeID As Long
            Dim Rec As New ADODB.Recordset
            Dim mCon As New ADODB.Connection
            Dim objDB As New clsDB
            mEditFlag = True
            txtMinorCode.Text = mID(VSGridMinor.TextMatrix(VSGridMinor.Row, 1), 4, 6)
            txtMinorHead.Text = VSGridMinor.TextMatrix(VSGridMinor.Row, 2)
            txtMinorCode.Tag = VSGridMinor.TextMatrix(VSGridMinor.Row, 3)
            txtMajorCodeFromMinor.Tag = VSGridMinor.TextMatrix(VSGridMinor.Row, 4)
            objDB.SetConnection mCon
            Rec.Open " select * from faMajorAccountHeads where intMajorAccountHeadID=" & Val(txtMajorCodeFromMinor.Tag), mCon
                If Not Rec.EOF Then
                    txtMajorCodeFromMinor.Text = Rec!vchMajorAccountHeadCode
                    txtMajorHeadFromMinor.Text = Rec!vchMajorAccountHead
                    txtMinorCode_LostFocus
                End If
            mTypeID = Val(VSGridMinor.TextMatrix(VSGridMinor.Row, 5))
            txtTrimMajor.Text = mID(VSGridMinor.TextMatrix(VSGridMinor.Row, 1), 1, 3)
            
        End Sub
    
    'Displays the details of minor head when minor code text box is got focused.
    
        Private Sub txtMinorCode_GotFocus()
           Dim objAc As New clsAccounts
           Dim objDB As New clsDB
           Dim Rec As New ADODB.Recordset
           Dim RecMajorFromMinor As New ADODB.Recordset
           Dim mCon As New ADODB.Connection
           Dim mSQL As String
           Dim mTypeID As Long

           If gbSearchStr <> "" Then
                   Dim mStr As String
                   txtMajorCodeFromMinor.Text = Trim(Token(gbSearchStr, " "))
                   txtMajorHeadFromMinor.Text = Trim(gbSearchStr)
                   txtMajorCodeFromMinor.Tag = gbSearchID
                   mSQL = "Select * From faMajorAccountHeads Where faMajorAccountHeads.intMajorAccountHeadID= " & gbSearchID
                   Set Rec = GetRecordSet(mSQL)
                   If Rec.RecordCount > 0 Then
                    txtMajorCodeFromMinor.Text = Rec!vchMajorAccountHeadCode
                    txtMajorHeadFromMinor.Text = Rec!vchMajorAccount
                   End If
                   gbSearchStr = ""
                   gbSearchID = -1
                Else
                    Call dispMinorHead
                End If
                 txtMinorHead.SelStart = 0
                 txtMinorHead.SelLength = Len(txtMinorHead)
    End Sub

'Fills vsGrid in Minor Head's tab when a particular major head is selected.

    Private Sub FillVSGridMinorByMajor(mTypeID As Long, intMajorHeadID As Long)
            Dim objDB       As New clsDB
            Dim mCon        As New ADODB.Connection
            Dim Rec         As New ADODB.Recordset
            Dim mLoopCount  As Integer
            Dim mSQL        As String
            Dim mtinType    As Long
            
            
            VSGridMinor.Clear 1, 1   'Clear all rows excluding fixed rows
            VSGridMinor.Rows = 2     'Setting initial row size
            objDB.SetConnection mCon
            
                mSQL = "Select * From faMinorAccountHeads Where tinType =" & mTypeID & "and intMajorAccountHeadID =" & intMajorHeadID
                Set Rec = GetRecordSet(mSQL)
                If Rec.RecordCount > 0 Then
                mLoopCount = 0
                While Not Rec.EOF
                     VSGridMinor.TextMatrix(mLoopCount + 1, 0) = mLoopCount + 1
                     VSGridMinor.TextMatrix(mLoopCount + 1, 1) = Rec.Fields(1).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 2) = Rec.Fields(2).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 3) = Rec.Fields(0).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 4) = Rec.Fields(3).Value
                     VSGridMinor.TextMatrix(mLoopCount + 1, 5) = Rec.Fields(4).Value
                     VSGridMinor.Rows = VSGridMinor.Rows + 1
                     Rec.MoveNext
                     mLoopCount = mLoopCount + 1
                Wend
                    VSGridMinor.Rows = VSGridMinor.Rows - 1
                End If
            
        End Sub
    
    'Fills vsGrid in Detailed Head's tab when a particular major and minor heads are selected.
    
        Private Sub FillVSGridDetailByMajorandMinor(mTypeID As Long, mMajorHeadID As Long, mMinorHeadID As Long)
                Dim objDB       As New clsDB
                Dim mCon        As New ADODB.Connection
                Dim Rec         As New ADODB.Recordset
                Dim mLoopCount  As Integer
                Dim mSQL        As String
                Dim mtinType    As Long
        
        
                VSGridDetail.Clear 1, 1   'Clear all rows excluding fixed rows
                VSGridDetail.Rows = 2     'Setting initial row size
                objDB.SetConnection mCon
        
                    mSQL = "Select * From faAccountHeads Where tinType =" & mTypeID & "And intMajorAccountHeadID =" & mMajorHeadID & "And intMinorAccountHeadID= " & mMinorHeadID
                    Set Rec = GetRecordSet(mSQL)
                    If Rec.RecordCount > 0 Then
                    mLoopCount = 0
                    While Not Rec.EOF
                         VSGridDetail.TextMatrix(mLoopCount + 1, 0) = mLoopCount + 1
                         VSGridDetail.TextMatrix(mLoopCount + 1, 1) = Rec.Fields(1).Value
                         VSGridDetail.TextMatrix(mLoopCount + 1, 2) = Rec.Fields(2).Value
                         If IsNull(Rec.Fields(4).Value) Then
                         VSGridDetail.TextMatrix(mLoopCount + 1, 3) = ""
                         Else
                         VSGridDetail.TextMatrix(mLoopCount + 1, 3) = Rec.Fields(4).Value
                         End If
                         If IsNull(Rec.Fields(12)) Then
                         VSGridDetail.TextMatrix(mLoopCount + 1, 4) = ""
                         Else
                         VSGridDetail.TextMatrix(mLoopCount + 1, 4) = Rec.Fields(12).Value
                         End If
                         VSGridDetail.TextMatrix(mLoopCount + 1, 5) = Rec.Fields(5).Value
                         VSGridDetail.TextMatrix(mLoopCount + 1, 6) = Rec.Fields(6).Value
                         If IsNull(Rec.Fields(15).Value) Then
                         VSGridDetail.TextMatrix(mLoopCount + 1, 7) = ""
                         Else
                         VSGridDetail.TextMatrix(mLoopCount + 1, 7) = Rec.Fields(15).Value
                         End If
                         VSGridDetail.TextMatrix(mLoopCount + 1, 8) = Rec.Fields(0).Value
                         VSGridDetail.TextMatrix(mLoopCount + 1, 9) = Rec.Fields(8).Value
                         
                         VSGridDetail.Rows = VSGridDetail.Rows + 1
        
                         Rec.MoveNext
                         
                         mLoopCount = mLoopCount + 1
                    Wend
                        VSGridDetail.Rows = VSGridDetail.Rows - 1
                    End If
                
        End Sub
    
    'This Function does updation in account heads table as well as transaction child table. in "currentbalance" and "openingbalance" fields respectively when a transaction is posted against a particular head.
    
    Private Sub UpdateOPeningBalnce(ByVal mAccountHeadID As Integer)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        Dim mSQL As String
        Dim mVTransactions As Variant
        Dim mCurrentBalance As Double
        Dim mLoop As Long
        Dim mQuery As String
        Dim fltAmount As Double
        If objDB.SetConnection(mCon) Then
            Rec.Open "Select intTransactionID,intSerialNo,fltAmount,tinDebitOrCreditFlag FROM FATRANSACTIONCHILD Where intAccountHeadID= " & mAccountHeadID, mCon
            If Not Rec.EOF Then
                 mVTransactions = Rec.GetRows
            End If
            
            If IsArray(mVTransactions) Then
                If Rec.State = 1 Then
                    Rec.Close
                End If
                Rec.Open "Select fltCurrentBalance FROM faAccountHEads Where intAccountHeadID=" & mAccountHeadID, mCon
                If Not Rec.EOF Then
                    mCurrentBalance = IIf(IsNull(Rec!fltCurrentBalance), 0#, Rec!fltCurrentBalance)
                Else
                    mCurrentBalance = 0
                End If
                pbAccHead.Max = UBound(mVTransactions, 2) + 1
                pbAccHead.Value = 1
                pbAccHead.Visible = True
                For mLoop = 0 To UBound(mVTransactions, 2)
                    
                    mQuery = " Update faTransactionChild set fltOpeningBalance =" & mCurrentBalance & " where intTransactionID=" & mVTransactions(0, mLoop) & " and intSerialNo=" & mVTransactions(1, mLoop)
                    mCon.Execute mQuery
                    If mVTransactions(3, mLoop) = 1 Then
                        fltAmount = mVTransactions(3, mLoop)
                    Else
                        fltAmount = mVTransactions(3, mLoop) * (-1)
                    End If
                    mQuery = "Update faAccountHeads set fltCurrentBalance= " & mCurrentBalance + fltAmount & " Where intAccountHeadID= " & mAccountHeadID
                    mCon.Execute mQuery
                    If pbAccHead.Value < pbAccHead.Max Then
                        pbAccHead.Value = pbAccHead.Value + 1
                    End If
                    lblProgress.Caption = CStr(CInt((mLoop + 1) / pbAccHead.Max * 100)) & "% Complete"
                Next mLoop
            End If
        End If
           pbAccHead.Visible = False
           lblProgress.Caption = ""
    End Sub
    
    'This function displays the details of major head by MajorAccountheadCode
    
    Private Sub dispMajorHead()
            Dim mMajorCode As String
            Dim objMajorHead As New clsMajorAccountHead
            Dim mTypeID As Long
            
        mMajorCode = Trim(txtTrimFirstDigit.Text) & Trim(txtMajorCode.Text)
        If mMajorCode <> "" Then
            objMajorHead.SetMajorAccountHead (mMajorCode)
            If mEditFlag And objMajorHead.MajorAccountHeadID > -1 Then
                If Val(txtMajorCode.Tag) = objMajorHead.MajorAccountHeadID Then
                    Exit Sub
                End If
            ElseIf mEditFlag Then
                Exit Sub
            End If
            If objMajorHead.MajorAccountHeadID > 0 Then
                mEditFlag = True
                txtMajorCode.Tag = objMajorHead.MajorAccountHeadID
                txtTrimFirstDigit.Text = mID(objMajorHead.MajorAccountHeadCode, 1, 1)
                txtMajorCode.Text = mID(objMajorHead.MajorAccountHeadCode, 2, 8)
                txtMajorHead.Text = objMajorHead.MajorAccountHead
                mTypeID = objMajorHead.TypeID
            Else
                mEditFlag = False
                txtMajorCode.Tag = -1
                txtMajorHead.Text = ""
                mTypeID = -1
            End If
        End If
    End Sub
    
    'This function displays the details of minor head by MinorAccountheadCode

    Private Sub dispMinorHead()
        Dim mMinorCode As String
        Dim objMinorHead As New clsMinorAccountHead
       
        mMinorCode = Trim(txtTrimMajor.Text) & Trim(txtMinorCode.Text)
        If mMinorCode <> "" Then
            objMinorHead.SetMinorAccountHead (mMinorCode)
            If mEditFlag And objMinorHead.MinorAccountHeadID > -1 Then
                If Val(txtMinorCode.Tag) = objMinorHead.MinorAccountHeadID Then
                    Exit Sub
                End If
            ElseIf mEditFlag Then
                Exit Sub
            End If
            If objMinorHead.MinorAccountHeadID > 0 Then
                mEditFlag = True
                txtMajorCodeFromMinor.Text = objMinorHead.MajorAccountHeadCode
                txtMajorHeadFromMinor.Text = objMinorHead.MajorAccountHead
                txtMajorCodeFromMinor.Tag = objMinorHead.MajorAccountHeadID
                txtMinorCode.Text = mID(objMinorHead.MinorAccountHeadCode, 4, 6)
                txtMinorHead.Text = objMinorHead.MinorAccountHead
                txtMinorCode.Tag = objMinorHead.MinorAccountHeadID
                txtTrimMajor.Text = mID(objMinorHead.MinorAccountHeadCode, 1, 3)
            Else
                mEditFlag = False
                txtMinorHead.Text = ""
                txtMinorCode.Tag = -1
            End If
        End If
    End Sub
        
   'This function displays the details of detailed head by detailedAccountheadCode

    Private Sub dispDetailedHead()
        Dim mDetailedCode As String
        Dim mDetailedID As Long
        Dim objDetailedHead As New clsAccounts
        Dim mTypeID As Long
        mDetailedCode = Trim(txtTrimMinor.Text) & Trim(txtDetailedCode.Text)
        If mDetailedCode <> "" Then
                objDetailedHead.SetAccountCode (mDetailedCode)
                If mEditFlag And objDetailedHead.AccountHeadID > -1 Then
                    If Val(txtDetailedCode.Tag) = objDetailedHead.AccountHeadID Then
                        Exit Sub
                    End If
                ElseIf mEditFlag Then
                    Exit Sub
                End If
                If objDetailedHead.AccountHeadID > 0 Then
                    mEditFlag = True
                    txtDetailedCode.Tag = objDetailedHead.AccountHeadID
                    txtDetailedCode.Text = mID(objDetailedHead.AccountCode, 6, 4)
                    txtDetailedHead.Text = objDetailedHead.AccountHead
                    txtTrimMinor.Text = mID(objDetailedHead.AccountCode, 1, 5)
                    txtMinorCodeByDetail.Tag = objDetailedHead.MinorAccountHeadID
                    txtMinorCodeByDetail.Text = objDetailedHead.MinorAccountHeadCode
                    txtMinorHeadByDetail.Text = objDetailedHead.MinorAccountHead
                    txtMajorCodeByDetail.Tag = objDetailedHead.MajorAccountHeadID
                    txtMajorCodeByDetail.Text = objDetailedHead.MajorAccountHeadCode
                    txtMajorHeadByDetail.Text = objDetailedHead.MajorAccountHead
                    txtOpeningBalance.Text = objDetailedHead.OpeningBalance
                    txtAlias.Text = objDetailedHead.Alias
                    mTypeID = objDetailedHead.mType
                    If objDetailedHead.DebitOrCredit = 1 Then
                        optDebit = True
                    Else
                        optCredit = True
                    End If
                Else
                    mEditFlag = False
                    txtDetailedCode.Tag = -1
                    txtDetailedHead.Text = ""
                    txtOpeningBalance.Text = ""
                    txtAlias.Text = ""
                    mTypeID = -1
                    optDebit = False
                    optCredit = False
                End If
        End If
    End Sub

Private Sub SearchDetailedAccountHead()
        Dim mSQL As String
        Dim mTypeID As Variant
        If optIncome.Value Then mTypeID = 1
        If optExpenditure.Value Then mTypeID = 2
        If optLiability.Value Then mTypeID = 3
        If optAsset.Value Then mTypeID = 4
        If Not IsNumeric(mTypeID) Then mTypeID = Null
        
        Select Case mTypeID
            Case 1
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 1 "
            Case 2
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 2 "
            Case 3
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 3 "
            Case 4
                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where tinType= 4 "
         End Select
         
        frmSearchAccountHeads.SQLString = mSQL
        frmSearchAccountHeads.Show vbModal
        txtDetailedCode.SetFocus
End Sub
    

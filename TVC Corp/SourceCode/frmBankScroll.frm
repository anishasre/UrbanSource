VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBankScroll 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Scroll"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankScroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUnReconcile 
      BackColor       =   &H00DEEDDE&
      Caption         =   "Unreconcile"
      Height          =   330
      Left            =   9585
      TabIndex        =   56
      Top             =   90
      Width           =   1455
   End
   Begin VB.CheckBox chkEdit 
      BackColor       =   &H00DEEDDE&
      Caption         =   "Edit Saved Items"
      Height          =   330
      Left            =   11070
      TabIndex        =   55
      Top             =   90
      Width           =   1905
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H80000011&
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7830
      Width           =   915
   End
   Begin VB.TextBox txtRecNo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   7515
      TabIndex        =   42
      Top             =   7830
      Width           =   600
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000011&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7830
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000011&
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7830
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000011&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7830
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   -315
      Top             =   8100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabScroll 
      Height          =   5865
      Left            =   45
      TabIndex        =   17
      Top             =   1845
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   10345
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   16777215
      ForeColor       =   8388608
      TabCaption(0)   =   "Bank Scroll"
      TabPicture(0)   =   "frmBankScroll.frx":1CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vsGrid"
      Tab(0).Control(1)=   "lblBalance"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bank Scroll Export Wizard"
      TabPicture(1)   =   "frmBankScroll.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSteps"
      Tab(1).Control(1)=   "fraClearGrid"
      Tab(1).Control(2)=   "fraFinish"
      Tab(1).Control(3)=   "lblExcelFileName"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Bank Scroll Edit"
      TabPicture(2)   =   "frmBankScroll.frx":1D02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "vsGridEdit"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraSteps 
         Caption         =   "Step by Step"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   4965
         Left            =   -74730
         TabIndex        =   19
         Top             =   540
         Width           =   5100
         Begin MSComctlLib.ProgressBar pgrFillGrid 
            Height          =   150
            Left            =   90
            TabIndex        =   26
            Top             =   3375
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   265
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.CheckBox chkLoadExcel 
            Caption         =   " No"
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3915
            TabIndex        =   21
            Top             =   675
            Width           =   870
         End
         Begin VB.CheckBox chkClearGrid 
            Caption         =   " No"
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3915
            TabIndex        =   23
            Top             =   1440
            Width           =   870
         End
         Begin VB.CheckBox chkFillGrid 
            Caption         =   " No"
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3915
            TabIndex        =   25
            Top             =   2160
            Width           =   870
         End
         Begin VB.CheckBox chkFinish 
            Caption         =   " No"
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3825
            TabIndex        =   28
            Top             =   4410
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1. Loading the Excel File - "
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   270
            TabIndex        =   20
            Top             =   675
            Width           =   2310
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2. Do you want to clear the Grid - "
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   270
            TabIndex        =   22
            Top             =   1440
            Width           =   2880
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3. Filling the Grid"
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   270
            TabIndex        =   24
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4. Finish"
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   270
            TabIndex        =   27
            Top             =   4410
            Width           =   720
         End
      End
      Begin VB.Frame fraClearGrid 
         Caption         =   "Clear Grid"
         Height          =   1545
         Left            =   -68835
         TabIndex        =   30
         Top             =   1395
         Visible         =   0   'False
         Width           =   5820
         Begin VB.OptionButton optClearGrid 
            Caption         =   "Clear the Grid"
            Height          =   600
            Left            =   3195
            TabIndex        =   32
            Top             =   315
            Width           =   1635
         End
         Begin VB.OptionButton optDontClearGrid 
            Caption         =   "Don't Clear the Grid"
            Height          =   600
            Left            =   900
            TabIndex        =   31
            Top             =   315
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.CommandButton cmdClearNext 
            Caption         =   "Next >>"
            Height          =   420
            Left            =   3870
            TabIndex        =   34
            Top             =   945
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   420
            Left            =   2520
            TabIndex        =   33
            Top             =   945
            Width           =   1185
         End
      End
      Begin VB.Frame fraFinish 
         Caption         =   "Finish"
         Height          =   1185
         Left            =   -68835
         TabIndex        =   35
         Top             =   3060
         Width           =   5820
         Begin VB.CommandButton cmdFinish 
            Caption         =   "Finish"
            Height          =   420
            Left            =   2430
            TabIndex        =   36
            Top             =   405
            Width           =   1185
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   5415
         Left            =   -74955
         TabIndex        =   18
         Top             =   360
         Width           =   12435
         _cx             =   21934
         _cy             =   9551
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   10008483
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483624
         BackColorAlternate=   -2147483634
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmBankScroll.frx":1D1E
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   4
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
      Begin VSFlex8LCtl.VSFlexGrid vsGridEdit 
         Height          =   4920
         Left            =   45
         TabIndex        =   54
         Top             =   585
         Width           =   12795
         _cx             =   22569
         _cy             =   8678
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   14609886
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483624
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmBankScroll.frx":1E62
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
         TextStyleFixed  =   4
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
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -67260
         TabIndex        =   46
         Top             =   135
         Width           =   135
      End
      Begin VB.Label lblExcelFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- -"
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   -68790
         TabIndex        =   29
         Top             =   945
         Width           =   165
      End
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   13410
      Top             =   7965
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   45
      TabIndex        =   2
      Top             =   450
      Width           =   12975
      Begin VB.Frame fraYearMonth 
         BorderStyle     =   0  'None
         Caption         =   "Year Month"
         Height          =   645
         Left            =   5940
         TabIndex        =   47
         Top             =   135
         Width           =   1995
         Begin VB.TextBox txtYear 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   360
            Left            =   45
            TabIndex        =   51
            Top             =   225
            Width           =   570
         End
         Begin VB.CommandButton cmdYearUp 
            BackColor       =   &H8000000B&
            Caption         =   "Ù"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            Style           =   1  'Graphical
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   180
            Width           =   285
         End
         Begin VB.CommandButton cmdYearDown 
            BackColor       =   &H8000000B&
            Caption         =   "Ú"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   675
            Style           =   1  'Graphical
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   405
            Width           =   285
         End
         Begin VB.ComboBox cmbMonth 
            BackColor       =   &H80000018&
            Height          =   360
            ItemData        =   "frmBankScroll.frx":1F98
            Left            =   945
            List            =   "frmBankScroll.frx":1FC3
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   225
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            Height          =   240
            Left            =   180
            TabIndex        =   53
            Top             =   0
            Width           =   435
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            Height          =   240
            Left            =   990
            TabIndex        =   52
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "..."
         Height          =   330
         Left            =   5535
         TabIndex        =   13
         Top             =   630
         Width           =   330
      End
      Begin VB.TextBox txtBank 
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   4245
      End
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Left            =   9945
         TabIndex        =   9
         Top             =   405
         Width           =   1500
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Left            =   8100
         TabIndex        =   6
         Top             =   405
         Width           =   1500
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import From Excel"
         Height          =   465
         Left            =   10485
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   765
         Width           =   1770
      End
      Begin VB.CommandButton cmdClearGrid 
         Caption         =   "&Clear Grid"
         Height          =   420
         Left            =   3555
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   9585
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   405
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   16646145
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   11430
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   405
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   16646145
         CurrentDate     =   39612
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   9945
         TabIndex        =   8
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   8100
         TabIndex        =   5
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Excel File is Saankhya supplied Template"
         ForeColor       =   &H00004040&
         Height          =   240
         Left            =   6255
         TabIndex        =   15
         Top             =   900
         Width           =   4170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can add the scroll into Saankhya by typing also"
         ForeColor       =   &H8000000C&
         Height          =   240
         Left            =   1035
         TabIndex        =   14
         Top             =   990
         Width           =   4590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This will clear the contents in the Grid -"
         ForeColor       =   &H00004040&
         Height          =   240
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   3345
      End
   End
   Begin MSComctlLib.ProgressBar pgrSave 
      Height          =   150
      Left            =   45
      TabIndex        =   37
      Top             =   7515
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B A N K  S C R O L L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   4823
      TabIndex        =   1
      Top             =   90
      Width           =   2925
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record No"
      Height          =   240
      Left            =   6480
      TabIndex        =   41
      Top             =   7920
      Width           =   990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info Label - -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   1485
      TabIndex        =   44
      Top             =   7920
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackColor       =   &H0098B7A3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   12975
   End
   Begin VB.Label Label12 
      BackColor       =   &H006C8B77&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      TabIndex        =   38
      Top             =   7740
      Width           =   13020
   End
End
Attribute VB_Name = "frmBankScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private mScrollEdit As Boolean

    Public Property Let ScrolEditMode(mMode As Boolean)
        mScrollEdit = mMode
    End Property
    Private Sub chkEdit_Click()
        If chkEdit.value = vbChecked Then
            ScrolEditMode = True
        Else
            ScrolEditMode = False
        End If
        Call formInitialise
    End Sub

    Private Sub chkUnReconcile_Click()
        Me.Hide
        frmBankUnReconcile.Visible = True
        chkUnReconcile.value = vbUnchecked
    End Sub

    Private Sub cmbMonth_Click()
        If Trim(txtYear.Text) <> "" Then
            txtFromDate.Text = "01-" & cmbMonth.Text & "-" & txtYear.Text
            txtToDate.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(txtFromDate.Text))), "dd-MMM-yyyy")
            lblBalance.Caption = ""
            If val(txtBank.Tag) > 0 Then
                lblBalance.Caption = "Balance on " & txtFromDate.Text & "  : " & CStr(Format(GetScrollBalance, "#.00"))
            End If
        End If
    End Sub

    Private Sub cmdBank_Click()
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where faAccountHeads.intGroupID =2 And tinHiddenFlag = 0"
        frmSearchAccountHeads.Show vbModal
        txtBank.SetFocus
        If chkEdit.value = vbChecked Then
            vsGridEdit.Clear 1
            Call formInitialise
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Call formInitialise
    End Sub

    Private Sub cmdClearGrid_Click()
        Call ClearGrid
    End Sub

    Private Sub cmdClearNext_Click()
        '------------------------------------------------------------------------'
        '                          Clear the Grid                                '
        If optClearGrid.value = True Then
            Call ClearGrid
            chkClearGrid.value = 1
        End If
        '-----------------------------------------------------------------------'
        '                       Filling the Grid From Excel                     '
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim objAcc As New clsAccounts
        Dim mServerName As String
        Dim mCon As New ADODB.Connection
     
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mServerName = GetServerName(mCnn)
        
'        mSQL = "SELECT BankAccountHeadCode,BankEntryDate,Particulars,ChequeNo,DrAmount,CrAmount,Balance,colTemplate,B.intReconciliationID"
'        mSQL = mSQL + " FROM OPENROWSET('Microsoft.Jet.OLEDB.4.0'," & vbNewLine
'        mSQL = mSQL + " 'Excel 8.0;Database=" & lblExcelFileName.Caption & ";IMEX=1'," & vbNewLine
'        mSQL = mSQL + " 'SELECT * FROM [Scroll$]') S" & vbNewLine
'        mSQL = mSQL + " Left Join faBankReconciliationentries B On Cast(S.BankAccountHeadCode as varchar(10)) COLLATE Latin1_General_CI_AS Like B.vchBankAccountHeadCode COLLATE Latin1_General_CI_AS" & vbNewLine
'        mSQL = mSQL + " And S.BankEntryDate = B.dtBankEntryDate And S.Particulars COLLATE Latin1_General_CI_AS = B.vchParticulars COLLATE Latin1_General_CI_AS" & vbNewLine
'        mSQL = mSQL + " And S.DrAmount = B.fltDRAmount And S.CrAmount = B.fltCrAmount" & vbNewLine
'        mSQL = mSQL + " Where BankAccountHeadCode = '" & Left(txtBank.Text, 9) & "' And BankEntryDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "'"
        
'        mSQL = "Select BankAccountHeadCode,BankEntryDate,Particulars,ChequeNo,DrAmount,CrAmount,Balance,colTemplate From [Scroll$] Where BankEntryDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "'"
        '-----------------------Connction Changed-----------------------------'
'        mCon.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=;Initial Catalog=DB_Finance;Data Source=" & mServerName
        mSql = "Select BankAccountHeadCode,BankEntryDate,Particulars,ChequeNo,DrAmount,CrAmount,Balance,colTemplate From [Scroll$]"
        
        mCon.Open "Driver={Microsoft Excel Driver (*.xls)}; DBQ=" & lblExcelFileName.Caption & "; ReadOnly=False;"
        On Error GoTo Errs
        Rec.Open mSql, mCon, adOpenStatic, adLockReadOnly
'        Rec.Filter = " BankAccountHeadCode = '" & Left(txtBank.Text, 9) & "'"
        If vsGrid.Rows = 1 Then
            vsGrid.AddItem ""
        End If
        
        If Not (Rec.BOF And Rec.EOF) Then
            pgrFillGrid.Visible = True
            pgrFillGrid.Max = Rec.RecordCount
            pgrFillGrid.value = 0
            
            While Not Rec.EOF
                pgrFillGrid.value = pgrFillGrid.value + 1                                                     ''' Progress Bar Updation
                If Rec!BankAccountHeadCode = val(Left(txtBank.Text, 9)) Then                                       ''' Checking the Account Head Code is Correct
                    If CDate(Rec!BankEntryDate) >= CDate(txtFromDate.Text) And CDate(Rec!BankEntryDate) <= CDate(txtToDate.Text) Then ''' Date Checking
                        vsGrid.Cell(flexcpForeColor, vsGrid.Rows - 1, 0, , vsGrid.Cols - 1) = vbBlue      ''' Coloring Green From Excel
                        'vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(Rec!intReconciliationID), "", Rec!intReconciliationID)
        '                If vsGrid.TextMatrix(vsGrid.Rows - 1, 0) <> "" Then
        '                    vsGrid.Cell(flexcpForeColor, vsGrid.Rows - 1, 0, , vsGrid.Cols - 1) = &H40&               ''' Coloring For already Exists Red
        '                End If
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!BankAccountHeadCode), "", Rec!BankAccountHeadCode)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!BankEntryDate), "", Format(Rec!BankEntryDate, "dd-MMM-yyyy"))
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!Particulars), "", Rec!Particulars)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!ChequeNo), "", Rec!ChequeNo)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!DrAmount), 0, Rec!DrAmount)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!CrAmount), 0, Rec!CrAmount)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = IIf(IsNull(Rec!Balance), "", Rec!Balance)
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 8) = 1                                           '''From Excel Sheet Filling
                        
                        vsGrid.Rows = vsGrid.Rows + 1
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!BankAccountHeadCode), "", Rec!BankAccountHeadCode)   '' HeadCode for Next Row
                    End If
                End If
                Rec.MoveNext
            Wend
            chkFillGrid.value = 1
        Else
            MsgBox " Nothing to display Or Excel Sheet not in correct format", vbInformation
            Call formInitialise
        End If
        Rec.Close
        mCnn.Close
        mCon.Close
        '-----------------------------------------------------------------------'
        fraClearGrid.Visible = False
        fraFinish.Visible = True
        Exit Sub
Errs:
        MsgBox "Already Opened Or The Excel is Not in Correct Format" & vbNewLine & err.Description, vbInformation
        Call formInitialise
        mCnn.Close
        mCon.Close
    End Sub

    Private Sub cmdFinish_Click()
        fraFinish.Visible = False
        tabScroll.TabVisible(1) = False
        tabScroll.TabVisible(0) = True      '' Scroll Visible true
        tabScroll.Tab = 0
        cmdBank.Enabled = False
        fraYearMonth.Enabled = False
        cmdSave.Visible = True
    End Sub
    Private Sub FillGridForEdit()
        Dim objDb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim mCount    As Integer
        Dim mRecCnt    As Integer
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Set Dateformat DMY Select vchBankAccountHeadCode,dtBankEntrydate,vchParticulars,vchChequeNo,dtChequeDate,fltDrAmount,fltCrAmount,intReconciliationID,tnyReconciled"
        mSql = mSql + " From faBankReconciliationEntries Where isNull(tnyReconciled,0)=0 And dtBankEntryDate between '" & txtFromDate.Text & "' and '" & Format(txtToDate.Text, "dd-mmm-yyyy") & "' and  intBankAccountHeadID = '" & val(txtBank.Tag) & "'"
        'Rec.Open mSql, mCnn, adOpenKeyset, adLockOptimistic
        Rec.CursorLocation = adUseClient
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
        vsGridEdit.Rows = 1
        
        mRecCnt = Rec.RecordCount
        If Not (Rec.EOF Or Rec.BOF) Then
            vsGridEdit.Rows = Rec.RecordCount + 1
            vsGridEdit.Col = 1
            vsGridEdit.Row = 1
            vsGridEdit.ColSel = 8
            vsGridEdit.RowSel = vsGridEdit.Rows - 1
            mSql = Rec.GetString(, , vbTab, Chr(13))
            vsGridEdit.Clip = mSql
        End If
        Rec.Close
        For mCount = 1 To vsGridEdit.Rows - 1
            vsGridEdit.TextMatrix(mCount, 0) = mCount
        Next
    End Sub
    Private Sub cmdImport_Click()
        If mScrollEdit Then
            ''''-------------Fetch data For Edit------------------------------
            If val(txtBank.Tag) < 1 Then
                MsgBox "Please Select the Bank", vbInformation
                cmdBank.Enabled = True
                fraYearMonth.Enabled = True
                cmdBank.SetFocus
                Exit Sub
            End If
             If Trim(txtFromDate.Text) = "" Then
                MsgBox "The Date Field is Empty, Please Supply the From date", vbInformation
                txtFromDate.SetFocus
                Exit Sub
            End If
            If Trim(txtToDate.Text) = "" Then
                MsgBox "The Date Field is Empty, Please Supply the To date", vbInformation
                txtToDate.SetFocus
                Exit Sub
            End If
            Call FillGridForEdit
        Else
            If val(txtBank.Tag) < 1 Then
                MsgBox "Please Select the Bank", vbInformation
                cmdBank.Enabled = True
                fraYearMonth.Enabled = True
                cmdBank.SetFocus
                Exit Sub
            End If
            
            If Trim(txtFromDate.Text) = "" Then
                MsgBox "The Date Field is Empty, Please Supply the From date", vbInformation
                txtFromDate.SetFocus
                Exit Sub
            End If
            If Trim(txtToDate.Text) = "" Then
                MsgBox "The Date Field is Empty, Please Supply the To date", vbInformation
                txtToDate.SetFocus
                Exit Sub
            End If
            Call formInitialise
            
            cmdBank.Enabled = False
            fraYearMonth.Enabled = False
            
            tabScroll.TabVisible(0) = False     '' Scroll Visible False
            tabScroll.TabVisible(1) = True
            tabScroll.Tab = 1
            
            dlgOpenFile.ShowOpen                '' Open an Excel File
            If dlgOpenFile.FileName = "" Then
                MsgBox "You did not chose a file Name, Exiting the Wizard", vbInformation
                tabScroll.TabVisible(1) = False
                tabScroll.TabVisible(0) = True
                tabScroll.Tab = 0
                Exit Sub
            End If
            lblExcelFileName.Caption = dlgOpenFile.FileName
            chkLoadExcel.value = 1
            chkLoadExcel.Caption = " Yes"
            fraClearGrid.Visible = True
        End If
    End Sub

    Private Sub cmdNew_Click()
        chkEdit.value = vbUnchecked
        mScrollEdit = False
        txtBank.Text = ""
        txtBank.Tag = -1
        lblBalance.Caption = ""
        Call formInitialise
        Call ClearGrid
    End Sub

    Private Sub cmdRemove_Click()
        Dim mCnn                As New ADODB.Connection
        Dim objDb               As New clsDB
        Dim mFrom, mTo, mLoop   As Integer
        Dim mFromTo             As String
        Dim mSql                As String
        Dim mRecId              As Double
        Dim mCnt                As Integer
        Dim mOut                As Integer
        mFromTo = Trim(txtRecNo.Text)
        mFrom = val(Trim(Token(mFromTo, "-")))
        mTo = val(Trim(mFromTo))
        If mTo < mFrom Then
            mTo = mFrom
        End If
        If txtRecNo = "" Then
                MsgBox "Please Enter a Record No. To Remove data", vbInformation
                Exit Sub
            End If
        If chkEdit = vbChecked Then
            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
             ''---------------Remove Record From Databasse
            If vsGridEdit.TextMatrix(1, 1) = "" Then
                MsgBox "Record Does not Exists", vbInformation
                Exit Sub
            End If
            If txtRecNo = "" Then
                MsgBox "Please Enter a Record No. To Remove data", vbInformation
                Exit Sub
            End If
            
            If MsgBox("Are you sure ?" & vbNewLine & "You want to Remove the Record : " & txtRecNo.Text & vbNewLine & "Record Will Delete Permanently", vbInformation + vbYesNo) = vbYes Then
                For mCnt = mFrom To mTo
                    mRecId = vsGridEdit.TextMatrix(mCnt, 8)
                    mSql = "Delete From faBankReconciliationEntries Where intReconciliationID=" & mRecId
                    mCnn.Execute mSql, mOut
                Next
            End If
            Call FillGridForEdit
        Else
        
            ''---------------Remove Record From Grid Only
            mTo = IIf(mTo > vsGrid.Rows - 1, vsGrid.Rows - 1, mTo)
            If val(mFrom) > 0 And val(mFrom) < vsGrid.Rows Then
                If MsgBox("Are you sure ?" & vbNewLine & "You want to Remove the Record...", vbInformation + vbYesNo) = vbYes Then
                    For mLoop = mFrom To mTo
                        vsGrid.RemoveItem val(mFrom)            ''' Removing the Records From mFrom to mTo
                    Next mLoop
                    lblInfo.Caption = "Records From" & CStr(mFrom) & " - " & CStr(mTo) & " Removed"
                    txtRecNo.Text = ""
                End If
            Else
                MsgBox "Invalid Record Number(s), No Change"
            End If
            If vsGrid.Rows = 1 Then
                vsGrid.AddItem ""
                If val(txtBank.Tag) > 1 Then
                    vsGrid.Cell(flexcpText, 1, 1, vsGrid.Rows - 1, 1) = Left(Trim(txtBank.Text), 9)
                End If
            End If
        End If
    End Sub
    Private Sub EditReconcilation()
    Dim Rec         As New ADODB.Recordset
    Dim mCnn        As New ADODB.Connection
    Dim objDb       As New clsDB
    Dim objAcc      As New clsAccounts
    Dim aryIn       As Variant
    Dim mCrAmt      As Variant
    Dim mDrAmt      As Variant
    Dim mCnt        As Integer
    Dim mAccHead    As String

    objDb.SetConnection mCnn
    Dim arInn As Variant
    For mCnt = 1 To vsGridEdit.Rows - 1
        If vsGridEdit.TextMatrix(mCnt, 6) <> "" Then
            mDrAmt = vsGridEdit.TextMatrix(mCnt, 6)
        Else
            mDrAmt = 0
        End If
        If vsGridEdit.TextMatrix(mCnt, 7) <> "" Then
            mCrAmt = vsGridEdit.TextMatrix(mCnt, 7)
        Else
            mCrAmt = 0
        End If

        objAcc.SetAccountID (txtBank.Tag)
        mAccHead = objAcc.AccountCode
'        @intReconciliationID            bigint      ,
'        @intBankAccountHeadID       int     ,
'        @vchBankAccountHeadCode     varchar(100)    ,
'        @dtBankEntryDate            smalldatetime   ,
'        @vchParticulars             varchar(100)    ,
'        @vchChequeNo            varchar(15) ,
'        @dtChequeDate           smalldatetime   ,
'        @fltDrAmount                float       ,
'        @fltCrAmount                float       ,
'        @tnyOpening             tinyint,
'        @tnyType                tinyint = Null
       
        aryIn = Array(vsGridEdit.TextMatrix(mCnt, 8), _
                txtBank.Tag, _
                mAccHead, _
                vsGridEdit.TextMatrix(mCnt, 2), _
                vsGridEdit.TextMatrix(mCnt, 3), _
                Trim(vsGridEdit.TextMatrix(mCnt, 4)), _
                vsGridEdit.TextMatrix(mCnt, 5), _
                mDrAmt, _
                mCrAmt, Null)
        objDb.ExecuteSP "spSaveBankReconsilation", aryIn, , , mCnn

    Next
End Sub

    Private Sub cmdSave_Click()
    
''''Create Proc spSaveBankEntries                 -- Drop Proc spSaveBankEntries
''''(@vchBankAccountHeadCode varchar(10), @dtBankEntryDate smallDateTime,@vchParticulars varchar(300),
'''' @vchChequeNo varchar(20), @fltDrAmount numeric(18,2), @fltCrAmount numeric(18,2))
''''As
''''------------------------------------------------------------------------------------------------------------------------------------------
''''Declare @intReconciliationID numeric,@intBankAccountHeadID int
''''Select @intBankAccountHeadID = intAccountHeadID From faAccountHeads Where vchAccountHeadCode = @vchBankAccountHeadCode
''''Select @intReconciliationID = isNull(Max(intReconciliationID)+1,1) From faBankReconciliationEntries
''''------------------------------------------------------------------------------------------------------------------------------------------
''''Insert Into faBankReconciliationEntries
''''(intReconciliationID, intBankAccountHeadID, vchBankAccountHeadCode, dtBankEntryDate, vchParticulars, vchChequeNo, fltDrAmount, fltCrAmount, tnyOpening)
''''Values
''''(@intReconciliationID, @intBankAccountHeadID, @vchBankAccountHeadCode, @dtBankEntryDate, @vchParticulars, @vchChequeNo, @fltDrAmount, @fltCrAmount, 0)
        Dim mCnt    As Integer
        If chkEdit.value = vbChecked Then
        ''''--------------Update BankReconciliationDetails --------------------------
            If val(txtBank.Tag) < 1 Then
                MsgBox "Please Select A Bank ", vbInformation
                Exit Sub
            End If
            If vsGridEdit.Rows > 1 Then
                If vsGridEdit.TextMatrix(1, 1) = "" Then
                    MsgBox "Please Select An Item to Edit ", vbInformation
                    Exit Sub
                End If
            Else
                MsgBox "Please Select An Item to Edit ", vbInformation
                Exit Sub
            End If
            Call EditReconcilation
            MsgBox "SuccessFully Edited", vbInformation
            cmdSave.Caption = "Edit"
            cmdSave.Enabled = False
            cmdSave.Caption = "Edit"
        Else
            '--------Progress Bar Settings----------'
            pgrSave.Visible = True
            pgrSave.Max = vsGrid.Rows - 1
            pgrSave.Min = 0
            pgrSave.value = 0
            '---------------------------------------'
            Dim mCnn As New ADODB.Connection
            Dim mLoop As Long
            Dim objDb As New clsDB
            Dim mArray As Variant
            Dim mDate As String
            Dim mBankPassBookBalance As Double
            Dim mSql As String
            Dim Rec As New ADODB.Recordset
            Dim mMinVoucherDate As Variant
            
            mBankPassBookBalance = 0
            mDate = "31-Mar-1970"
            
            If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
                MsgBox "Connction Lost, Contact Administrator"
                Exit Sub
            End If
            
            On Error GoTo Errs
            mCnn.BeginTrans                     '   Transaction Begins
            Rec.Open "Select min(dtDate) MinDate From faVouchers", mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mMinVoucherDate = Rec!MinDate
            End If
            Rec.Close
            For mLoop = 1 To vsGrid.Rows - 1
                '------------------Validations-----------------'
                If Trim(vsGrid.TextMatrix(mLoop, 1)) <> "" And Trim(vsGrid.TextMatrix(mLoop, 2)) <> "" Then
                    If Len(Trim(vsGrid.TextMatrix(mLoop, 1))) <> 9 Then
                        MsgBox "The Headcode not correct format, Exiting", vbInformation
                        lblInfo.Caption = "Record " + CStr(mLoop) + "'s Headcode format Wrong"
                        GoTo Errs
                    End If
                    If IsDate(Trim(vsGrid.TextMatrix(mLoop, 2))) = False Then
                        MsgBox "The Date format is wrong", vbInformation
                        lblInfo.Caption = "Record " + CStr(mLoop) + "'s Date format Wrong"
                        GoTo Errs
                    End If
                    If val(Trim(vsGrid.Cell(flexcpText, mLoop, 5, mLoop, 6))) = 0 Then
                        MsgBox "Please Check the Amount fields", vbInformation
                        lblInfo.Caption = "Record " + CStr(mLoop) + "'s Amounts Wrong"
                        GoTo Errs
                    End If
                     'Select * From faBankReconciliationEntries Where vchBankAccountHeadCode = @vchBankAccountHeadCode And dtBankEntryDate = @dtBankEntryDate And isNull(vchChequeNo,'') = @vchChequeNo And isNull(fltDrAmount,0) = @fltDrAmount And isNull(fltCrAmount,0) = @fltCrAmount
    '                With vsGrid
    '                 mSQL = "Select * From faBankReconciliationEntries Where vchBankAccountHeadCode = '" & .TextMatrix(mLoop, 1) & "' And dtBankEntryDate = '" & Format(.TextMatrix(mLoop, 2), "dd-MMM-yyyy") & "' And isNull(vchChequeNo,'') = '" & val(.TextMatrix(mLoop, 4)) & "' And isNull(fltDrAmount,0) = " & val(.TextMatrix(mLoop, 5)) & " And isNull(fltCrAmount,0) = " & val(.TextMatrix(mLoop, 6))
    '                End With
    '                Rec.Open mSQL, mCnn
    '                If Not (Rec.EOF And Rec.BOF) Then
    '                    MsgBox "The Record already Exists  (" & mLoop & ")" & vbNewLine & " to Remove , Enter record Number in TextBox and Press Remove", vbInformation
    '                    txtRecNo.SetFocus
    '                    GoTo Errs
    '                End If
    '                Rec.Close
                    
                    '---------------------------------------------------------------'
                    '                              Saving                           '
                    With vsGrid
                    mDate = Format(vsGrid.TextMatrix(mLoop, 2), "dd-MMM-yyyy")
                    
                    mArray = Array(Trim(.TextMatrix(mLoop, 1)), _
                                    Format(.TextMatrix(mLoop, 2), "dd-MMM-yyyy"), _
                                    .TextMatrix(mLoop, 3), _
                                    .TextMatrix(mLoop, 4), _
                                    val(.TextMatrix(mLoop, 5)), _
                                    val(.TextMatrix(mLoop, 6)), _
                                    IIf(CDate(mDate) > CDate(mMinVoucherDate), 0, 1) _
                                    )
                   
                    objDb.ExecuteSP "spSaveBankEntries", mArray, , , mCnn
                    End With
                    
                    '----------------------------------------------------------------'
                    '                       Balance Calculating                      '
                    If CDate(mDate) > CDate(mMinVoucherDate) Then
                        mSql = "Select isNull(Sum(fltCrAmount),0)-isNull(Sum( fltDrAmount),0) fltAmount From faBankReconciliationEntries "
                        mSql = mSql + " Where tnyOpening = 0 AND intBankAccountHeadID = " & val(txtBank.Tag) & " And dtBankEntryDate <= '" & mDate & "'"
                    Else
                        mSql = "Select isNull(Sum(fltCrAmount),0)-isNull(Sum( fltDrAmount),0) fltAmount From faBankReconciliationEntries "
                        mSql = mSql + " Where tnyOpening = 1 AND intBankAccountHeadID = " & val(txtBank.Tag) & " And dtBankEntryDate <= '" & mDate & "'"
                    End If
                    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                    If Not (Rec.BOF And Rec.EOF) Then
                        mBankPassBookBalance = Rec!fltAmount
                    End If
                    Rec.Close
                    If CDate(mDate) > CDate(mMinVoucherDate) Then
                        mSql = "Select fltOpening * ((tinDebitOrCreditFlag*2)-1 )fltOpening From faBanks Where intAccountHeadID = " & val(txtBank.Tag)
                        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
                        If Not (Rec.BOF And Rec.EOF) Then
                            mBankPassBookBalance = Format(mBankPassBookBalance + Format(Rec!fltOpening, "0.00"), "#.00")
                        End If
                        Rec.Close
                    End If
                    If val(vsGrid.TextMatrix(mLoop, 7)) <> val(mBankPassBookBalance) Then
                        MsgBox "The Balance not Correct.", vbInformation
                        lblInfo.Caption = "Record " + CStr(mLoop) + "'s Balance is wrong"
                        GoTo Errs
                    End If
                    '------------------------------------------------------------------------'
                    '                       Progress Bar Incrementing                        '
                    pgrSave.value = pgrSave.value + 1
                    '------------------------------------------------------------------------'
                End If
            Next mLoop
    
    ''''''''''''''''                If vsGrid.Rows - 1 > mLoop + 1 Then
    ''''''''''''''''                    If mDate <> Format(vsGrid.TextMatrix(mLoop + 1, 2), "dd-MMM-yyyy") Then
    ''''''''''''''''                        If vsGrid.TextMatrix(mLoop, 7) <> mBankPassBookBalance Then
    ''''''''''''''''                            MsgBox "The Balance not Correct.", vbInformation
    ''''''''''''''''                            lblInfo.Caption = "Record " + CStr(mLoop) + "'s Balance is wrong"
    ''''''''''''''''                            GoTo Errs
    ''''''''''''''''                        End If
    ''''''''''''''''                        mBankPassBookBalance = 0
    ''''''''''''''''                    End If
    ''''''''''''''''                Else
    ''''''''''''''''                    If vsGrid.TextMatrix(mLoop, 7) <> mBankPassBookBalance Then
    ''''''''''''''''                        MsgBox "The Balance not Correct..", vbInformation
    ''''''''''''''''                        lblInfo.Caption = "Record " + CStr(mLoop) + "'s Balance is wrong"
    ''''''''''''''''                        GoTo Errs
    ''''''''''''''''                    End If
    ''''''''''''''''                End If
    
            mCnn.CommitTrans                    '   Transaction Commiting
            cmdSave.Visible = False
            lblInfo.Caption = "DONE SUCCESS FULLY"
            mCnn.Close
            Exit Sub
Errs:
            mCnn.RollbackTrans                  '   Transaction RollBacks
            mCnn.Close
            'MsgBox " The Procedure Cancelled", vbInformation
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdYearDown_Click()
        txtYear.Text = val(txtYear.Text) - 1
    End Sub

    Private Sub cmdYearUp_Click()
        txtYear.Text = val(txtYear.Text) + 1
    End Sub

    Private Sub dtpFrom_CloseUp()
        txtFromDate.Text = Format(dtpFrom.value, "dd-MMM-yyyy")
    End Sub

    Private Sub dtpTo_CloseUp()
        txtToDate.Text = Format(dtpTo.value, "dd-MMM-yyyy")
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        winXPC.InitIDESubClassing
        chkEdit.value = vbUnchecked
        mScrollEdit = False
        Call FillMonth
        txtYear.Text = gbFinancialYearID
        Call formInitialise
    End Sub

    Private Sub txtBank_GotFocus()
        If gbSearchStr <> "" Then
            txtBank.Tag = gbSearchID
            txtBank.Text = gbSearchStr
            vsGrid.Cell(flexcpText, 1, 1, vsGrid.Rows - 1, 1) = Trim(Token(gbSearchStr, " "))
            gbSearchStr = ""
            gbSearchID = -1
            lblBalance.Caption = ""
            If val(txtBank.Tag) > 0 Then
                lblBalance.Caption = "Balance on " & txtFromDate.Text & "  : " & CStr(Format(GetScrollBalance, "#.00"))
            End If
        End If
    End Sub
    
    Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, "/-")
    End Sub

    Private Sub txtFromDate_LostFocus()
        txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
    End Sub
    
    Private Sub txtRecNo_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, "-")
    End Sub
    
    Private Sub txtRecNo_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim mFrom, mTo, mLoop As Integer
        Dim mFromTo As String
        mFromTo = Trim(txtRecNo.Text)
        mFrom = val(Trim(Token(mFromTo, "-")))
        mTo = val(Trim(mFromTo))
        If mTo < mFrom Then
            mTo = mFrom
        End If
        mTo = IIf(mTo > vsGrid.Rows - 1, vsGrid.Rows - 1, mTo)
        If mFrom > 0 And mFrom < vsGrid.Rows Then
            vsGrid.Row = mFrom
            vsGrid.Select mFrom, 1, mTo, vsGrid.Cols - 1
            vsGrid.TopRow = val(txtRecNo.Text)                  '''     Moving to the Current Row  (.topRow is to move)
        End If
    End Sub

    Private Sub txtToDate_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, "/-")
    End Sub

    Private Sub txtToDate_LostFocus()
        txtToDate.Text = CheckDateInMMM(txtToDate.Text)
    End Sub
    
    Private Sub formInitialise()
        If mScrollEdit Then
            tabScroll.TabsPerRow = 1
            tabScroll.TabVisible(2) = True
            tabScroll.TabVisible(1) = False
            tabScroll.TabVisible(0) = False
            Label4.Visible = False
            dtpFrom.Enabled = True
            txtFromDate.Enabled = True
            txtToDate.Enabled = True
            dtpTo.Enabled = True
            dtpFrom.Enabled = True
            cmdYearUp.Enabled = False
            cmdYearDown.Enabled = False
            cmbMonth.Enabled = False
            cmdSave.Visible = True
            cmdImport.Caption = "Search"
            cmdSave.Caption = "Edit"

        Else
            tabScroll.TabsPerRow = 1
            tabScroll.TabVisible(1) = False
            tabScroll.TabVisible(0) = True
            tabScroll.Tab = 0
            tabScroll.TabVisible(2) = False
    
            cmdImport.Caption = "Import From Excel"
            cmdSave.Caption = "Save"
            dtpFrom.Enabled = False
            txtFromDate.Enabled = False
            txtToDate.Enabled = False
            dtpTo.Enabled = False
            dtpFrom.Enabled = False
            cmdYearUp.Enabled = True
            cmdYearDown.Enabled = True
            cmbMonth.Enabled = True
            cmdSave.Caption = "&Save"
            
            pgrSave.Visible = False
            pgrFillGrid.value = 0
            pgrFillGrid.value = 0
            
            dlgOpenFile.DefaultExt = "xls"
            dlgOpenFile.FileName = ""
            dlgOpenFile.Filter = "Microsoft Excel Workbooks(*.xls)|*.xls"
            
            optDontClearGrid.value = True
            
            chkLoadExcel.value = 0
            chkLoadExcel.Caption = " No"
            
            chkClearGrid.value = 0
            chkClearGrid.Caption = " No"
            
            chkFillGrid.value = 0
            chkFillGrid.Caption = " No"
            
            chkFinish.value = 0
            chkFinish.Caption = " No"
            
            lblExcelFileName.Caption = ""
            lblInfo.Caption = ""
            
            cmdBank.Enabled = True
            fraYearMonth.Enabled = True
            txtRecNo.Text = ""
            
            fraClearGrid.Visible = False
            fraFinish.Visible = False
            cmdSave.Visible = False
        End If
    End Sub

    Private Sub ClearGrid()
        '------------------------------------------------------------------------'
        '                          Clear the Grid                                '
        If MsgBox("This will clear all the rows of your grid" & vbNewLine & "Are you sure", vbInformation + vbYesNo) = vbYes Then
            vsGrid.Rows = 1
            vsGrid.Rows = 2
            vsGridEdit.Rows = 1
        End If
    End Sub

    Private Function GetServerName(mCnn As ADODB.Connection) As String
        Dim mServerName     As String
        Dim mStart          As Integer
        Dim mLength         As Integer
        Dim mEnd            As Integer
        
        mStart = InStr(1, mCnn.ConnectionString, "WSID=")
        mStart = mStart + 5
        mEnd = InStr(mStart, mCnn.ConnectionString, ";")
        mLength = mEnd - mStart
        mServerName = mID(mCnn.ConnectionString, mStart, mLength)
        GetServerName = mServerName
    End Function

    Private Sub txtYear_Change()
        If Trim(txtYear.Text) <> "" Then
            txtFromDate.Text = "01-" & cmbMonth.Text & "-" & txtYear.Text
            txtToDate.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(txtFromDate.Text))), "dd-MMM-yyyy")
            lblBalance.Caption = ""
            If val(txtBank.Tag) > 0 Then
                lblBalance.Caption = "Balance on " & txtFromDate.Text & "  : " & CStr(Format(GetScrollBalance, "#.00"))
            End If
        End If
    End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 2 Then
            If Trim(vsGrid.TextMatrix(Row, 2)) = "" Then
                vsGrid.TextMatrix(Row, 2) = ""
            Else
                vsGrid.TextMatrix(Row, 2) = CheckDateInMMM(Trim(vsGrid.TextMatrix(Row, 2)))
                If CDate(vsGrid.TextMatrix(Row, 2)) >= CDate(txtFromDate.Text) And CDate(vsGrid.TextMatrix(Row, 2)) <= CDate(txtToDate.Text) Then
                    vsGrid.Cell(flexcpForeColor, Row, Col) = vbBlue
                Else
                    MsgBox "Date Mismatch", vbInformation
                    vsGrid.Cell(flexcpForeColor, Row, Col) = vbRed
                End If
            End If
        End If
        If Col = 6 Then
            If val(vsGrid.TextMatrix(Row, 5)) <> 0 Then       ''      Debit Amount Validating
                vsGrid.TextMatrix(Row, 6) = 0
            End If
        End If
        If Col = 5 Then
            If val(vsGrid.TextMatrix(Row, 6)) <> 0 Then         ''      Debit Amount Validating
                vsGrid.TextMatrix(Row, 5) = 0
            End If
        End If
        If Trim(vsGrid.TextMatrix(Row, 5)) = "" Then        ''      Debit Amount Validating
            vsGrid.TextMatrix(Row, 5) = 0
        End If
        If Trim(vsGrid.TextMatrix(Row, 6)) = "" Then        ''      Credit Amount Validating
            vsGrid.TextMatrix(Row, 6) = 0
        End If
        If Trim(vsGrid.TextMatrix(Row, 7)) = "" Then        ''      Balance Amount Validating
            vsGrid.TextMatrix(Row, 7) = 0
        End If
    End Sub

    Private Sub vsGrid_Click()
        If Trim(vsGrid.TextMatrix(vsGrid.Row, 8)) <> "" Then
            vsGrid.Editable = flexEDNone
        ElseIf vsGrid.Col <> 1 Then
            vsGrid.Editable = flexEDKbdMouse
        Else
            vsGrid.Editable = flexEDNone
        End If
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If val(txtBank.Tag) < 1 Then
            KeyAscii = 0
            Exit Sub
        End If
        If Row > 0 Then
            If ValidateCells Then                                           ' Cell Validations
                If Row = vsGrid.Rows - 1 And vsGrid.Col = 7 Then     ' Finding the Last Row and Column to increment Rows
                    If KeyAscii = 13 Then
                        vsGrid.Rows = vsGrid.Rows + 1
                        vsGrid.Row = vsGrid.Rows - 1
                        vsGrid.Col = 2
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = vsGrid.TextMatrix(vsGrid.Rows - 2, 1)
                        cmdSave.Visible = True
                    End If
                End If
            Else
                lblInfo.Caption = "Complete the current row"
            End If
        End If
        '---------------------------------------------------------------'
        '                   column type validation                      '
        If Col = 1 Then
            Call KeyPressNumber(KeyAscii, "")
        ElseIf Col = 2 Then
            Call KeyPressNumber(KeyAscii, "/-")
        ElseIf Col = 4 Then
            Call KeyPressNumber(KeyAscii, "")
        ElseIf Col = 5 Then
            Call KeyPressNumber(KeyAscii, ".")
        ElseIf Col = 6 Then
            Call KeyPressNumber(KeyAscii, ".")
        ElseIf Col = 7 Then
            Call KeyPressNumber(KeyAscii, ".")
        End If
        '---------------------------------------------------------------'
    End Sub
'
'    Private Sub vsGrid_LeaveCell()
'        If vsGrid.Row > 0 Then
'            If ValidateCells Then                                           ' Cell Validations
'                If vsGrid.Row = vsGrid.Rows - 1 And vsGrid.Col = 7 Then     ' Finding the Last Row and Column to increment Rows
'                    vsGrid.Rows = vsGrid.Rows + 1
'                    vsGrid.Col = 2
'                End If
'            Else
'                lblInfo.Caption = "Complete the current row"
'            End If
'        End If
'    End Sub

    Function ValidateCells() As Boolean
        Dim mBoolValidateRow As Boolean
        Dim mLoop As Integer
        mBoolValidateRow = True
        For mLoop = 1 To vsGrid.Cols - 2        ' Avoiding the Hidden Columns
            If (Trim(vsGrid.TextMatrix(vsGrid.Row, mLoop)) = "") Then
                If Not (mLoop = 3 Or mLoop = 4) Then            '' Validation not required for Particulars and ChequeNo
                    mBoolValidateRow = False
                End If
            End If
        Next mLoop
        ValidateCells = mBoolValidateRow
    End Function
    Function ValidateEditGrid() As Integer
        
    
    End Function
    Private Function GetScrollBalance() As Double
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mBankPassBookBalance As Double
        Dim objDb As New clsDB
        mBankPassBookBalance = 0
        If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connction Lost, Contact Administrator"
            Exit Function
        End If
        
        mSql = "Select  Sum(ISNULL(fltCrAmount,0))- Sum(ISNULL(fltDrAmount,0)) fltAmount From faBankReconciliationEntries "
        mSql = mSql + " Where IsNull(tnyOpening,0) = 0 AND intBankAccountHeadID = " & val(txtBank.Tag) & " And dtBankEntryDate < '" & Trim(txtFromDate.Text) & "'"
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If Not (Rec.BOF And Rec.EOF) Then
            mBankPassBookBalance = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
        End If
        Rec.Close
        mSql = "Select fltOpening * ((tinDebitOrCreditFlag*2)-1 )fltOpening From faBanks Where intAccountHeadID = " & val(txtBank.Tag)
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If Not (Rec.BOF And Rec.EOF) Then
            mBankPassBookBalance = mBankPassBookBalance + Format(Rec!fltOpening, "0.00")
        End If
        Rec.Close
        mCnn.Close
        GetScrollBalance = mBankPassBookBalance
    End Function
    
    Private Sub FillMonth()
        cmbMonth.Clear
        
        cmbMonth.AddItem ("Apr")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem ("May")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem ("Jun")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem ("Jul")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem ("Aug")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem ("Sep")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem ("Oct")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem ("Nov")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem ("Dec")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
        cmbMonth.AddItem ("Jan")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem ("Feb")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem ("Mar")
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
        
        cmbMonth.ListIndex = 0
    End Sub
'
'    Function KeyPressNumber(mAscii As Integer, mExtraKeys As String)
'        '-------------------------------------------------------------------------------'
'        '      This Function is used to Give Extra Charectors to the TextEditor         '
'        '-------------------------------------------------------------------------------'
'        If Not (mAscii >= 48 And mAscii <= 57 Or mAscii = 8 Or InStr(0, mExtraKeys, CStr(mAscii)) > 1) Then
'            mAscii = 0
'        End If
'        KeyPressNumber = mAscii
'    End Function

    Private Sub vsGridEdit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If Col = 2 Then
            If Trim(vsGridEdit.TextMatrix(Row, 2)) = "" Then
                vsGridEdit.TextMatrix(Row, 2) = ""
            Else
                vsGridEdit.TextMatrix(Row, 2) = CheckDateInMMM(Trim(vsGridEdit.TextMatrix(Row, 2)))
                If CDate(vsGridEdit.TextMatrix(Row, 2)) >= CDate(txtFromDate.Text) And CDate(vsGridEdit.TextMatrix(Row, 2)) <= CDate(txtToDate.Text) Then
                Else
                    MsgBox "Date Mismatch", vbInformation
                End If
            End If
        ElseIf Col = 6 Then
            If val(vsGridEdit.TextMatrix(Row, 7)) <> 0 Then
               vsGridEdit.TextMatrix(Row, 6) = 0
            End If
        ElseIf Col = 7 Then
            If val(vsGridEdit.TextMatrix(Row, 6)) <> 0 Then
               vsGridEdit.TextMatrix(Row, 7) = 0
            End If
        End If
        If Trim(vsGridEdit.TextMatrix(Row, 6)) = "" Then        ''      Debit Amount Validating
            vsGridEdit.TextMatrix(Row, 6) = 0
        End If
        If Trim(vsGridEdit.TextMatrix(Row, 7)) = "" Then        ''      Credit Amount Validating
            vsGrid.TextMatrix(Row, 7) = 0
        End If
    End Sub
    Private Sub vsGridEdit_CellChanged(ByVal Row As Long, ByVal Col As Long)
        If Col = 0 Then Exit Sub
        If Col = 2 Then
            vsGridEdit.Editable = flexEDNone
            If vsGridEdit.TextMatrix(Row, 2) <> "" Then
                vsGridEdit.TextMatrix(Row, 2) = CheckDateInMMM(vsGridEdit.TextMatrix(Row, 2))
            End If
        End If
        If Col = 5 Then
        vsGridEdit.Editable = flexEDKbdMouse
        If vsGridEdit.TextMatrix(Row, 5) <> "" Then
            vsGridEdit.TextMatrix(Row, 5) = CheckDateInMMM(vsGridEdit.TextMatrix(Row, 5))
        End If
        End If
    End Sub

    Private Sub vsGridEdit_Click()
        With vsGridEdit
            If .Col = 0 Or .Col = 1 Then
                vsGridEdit.Editable = flexEDNone
            Else
                vsGridEdit.Editable = flexEDKbdMouse
            End If
        End With
    End Sub
    
  
    Private Sub vsGridEdit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Col = 6 Or Col = 7 Then
            If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
                KeyAscii = 0
            End If
        End If
    End Sub

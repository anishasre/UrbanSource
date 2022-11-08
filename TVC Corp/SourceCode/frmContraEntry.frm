VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmContraEntry 
   BackColor       =   &H00F9FFF9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contra Entry"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11760
      TabIndex        =   38
      Top             =   6120
      Width           =   11820
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00D6E0E0&
         Caption         =   "&TRANSFER CREDIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdReject 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "Reject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "Cance&L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00D6E0E0&
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
         Height          =   390
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00D6E0E0&
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
         Height          =   390
         Left            =   2805
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   11760
      TabIndex        =   37
      Top             =   0
      Width           =   11820
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   9480
      Top             =   6030
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E1EBEB&
      Height          =   825
      Left            =   0
      TabIndex        =   32
      Top             =   765
      Width           =   11820
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1755
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   5
         Top             =   315
         Width           =   1755
      End
      Begin VB.CommandButton cmdSearchVoucherNo 
         BackColor       =   &H00D6E0E0&
         Height          =   315
         Left            =   11340
         Picture         =   "frmContraEntry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtVoucherNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9555
         TabIndex        =   7
         Top             =   315
         Width           =   1755
      End
      Begin VB.ComboBox cmbTransactionType 
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
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6150
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblVoucherNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contra Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   0
         Top             =   345
         Width           =   1110
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EDF7F7&
      Height          =   315
      Left            =   0
      TabIndex        =   36
      Top             =   1500
      Width           =   11820
   End
   Begin VB.Frame fraBank 
      BackColor       =   &H00E1EBEB&
      Height          =   1710
      Left            =   0
      TabIndex        =   34
      Top             =   1725
      Width           =   11820
      Begin VB.ComboBox cmbCategory 
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
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1200
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.ComboBox cmbSource 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.CommandButton cmdAccoundHeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "..."
         Height          =   315
         Left            =   8385
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   315
      End
      Begin VB.TextBox txtIssuedDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9915
         MaxLength       =   50
         TabIndex        =   21
         Top             =   720
         Width           =   1440
      End
      Begin VB.TextBox txtInstDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9915
         MaxLength       =   50
         TabIndex        =   14
         Top             =   375
         Width           =   1440
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   500
         TabIndex        =   18
         Top             =   720
         Width           =   5475
      End
      Begin VB.TextBox txtAccountCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbInstruments 
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
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   345
         Width           =   4695
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6645
         MaxLength       =   50
         TabIndex        =   12
         Top             =   360
         Width           =   2070
      End
      Begin MSComCtl2.DTPicker dtpInstDate 
         Height          =   315
         Left            =   11370
         TabIndex        =   15
         Top             =   375
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Baltic"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   60751875
         CurrentDate     =   39291
      End
      Begin MSComCtl2.DTPicker dtpIssueDate 
         Height          =   315
         Left            =   11370
         TabIndex        =   35
         Top             =   1695
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Baltic"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   60751875
         CurrentDate     =   39291
      End
      Begin MSComCtl2.DTPicker dtpIssuedDate 
         Height          =   315
         Left            =   11370
         TabIndex        =   22
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Baltic"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   60751875
         CurrentDate     =   39291
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5280
         TabIndex        =   44
         Top             =   1200
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cr. A/c Head"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instruments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8880
         TabIndex        =   20
         Top             =   735
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inst. No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5985
         TabIndex        =   11
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inst. Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9090
         TabIndex        =   13
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDF7F7&
      Height          =   195
      Left            =   0
      TabIndex        =   39
      Top             =   3345
      Width           =   11820
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F5FFFF&
      Height          =   2715
      Left            =   0
      TabIndex        =   31
      Top             =   3390
      Width           =   11820
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   930
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1995
         Width           =   7470
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   1560
         Left            =   15
         TabIndex        =   23
         Top             =   390
         Width           =   11790
         _cx             =   20796
         _cy             =   2752
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         BackColorFixed  =   15595511
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
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
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmContraEntry.frx":00FA
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
         Editable        =   2
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
      Begin VB.TextBox txtDr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   9330
         TabIndex        =   27
         Top             =   2010
         Width           =   2130
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   24
         Top             =   2100
         Width           =   795
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00D6E0E0&
         Caption         =   "          Debit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   30
         TabIndex        =   33
         Top             =   150
         Width           =   11745
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8730
         TabIndex        =   26
         Top             =   2040
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmContraEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'*****************************************************************************************'
'* Application ID           :                                                            *'
'* Application Name         : Saankhya Double Entry                                      *'
'* Screen id                : Payments                                                   *'
'* Version No               : Ver 2.0.0                                                  *'
'* Form Designed By         : Achu                                                       *'
'* Created on               :                                                            *'
'* Coded By                 :                                                            *'
'* Coded on                 :                                                            *'
'* Reviewed By              :                                                            *'
'* Reviewed on              : 11-Sep-2007                                                *'
'* Purpose                  : Manual Tracking of Payments Voucher                        *'
'*                                                                                       *'
'*                                                                                       *'
'* Name of Database         : DB_Finance                                                 *'
'* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *'
'* Name of Table(s)         : faTransactions, faTransactionChild                         *'
'* Look up Table(s)         : faTransactionType, faTransactionChild, faAccountHeads      *'
'*                          : faBudgetCentre, faFunction, faFunctionaries, faFields      *'
'*                                                                                       *'
'* Stored Procedures        : spGetAccHead4Receipts, spSaveTrans,                        *'
'*                          : spSaveTransactionChild                                     *'
'*                          :                                                            *'
'*=======================================================================================*'

Option Explicit
    Private objCr               As New clsAccounts
    Private objBk               As New clsBank
    Private mdtLastRemittance   As Variant
    Private mCopiedAmount       As Variant
    Public mRemittanceModule      As Integer ' Set as moduleID From  frmListOfDailyCollection
    Dim mPreviousYearMode       As Variant
    Dim mPreviousYearRequestID  As Variant
    
    
    Public Sub DisplayReceiptDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        Dim mStatus         As Variant
        
        Call FormInitialize
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select tnyStatus From faVouchers"
        mSql = mSql + " Where Cast(intVoucherNo as varchar(20)) = '" & mVoucherNo & "'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), Null, Rec!tnyStatus)
        End If
        Rec.Close
        If mStatus = 0 Or IsNull(mStatus) Then
            mSql = "Select * From faVouchers"
            mSql = mSql + " Inner Join faTransactions On faTransactions.intVoucherId = faVouchers.intVoucherId"
            mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
            mSql = mSql + " Left Join faFunctions On fatransactions.intFunctionId = faFunctions.intFunctionId"
            mSql = mSql + " Left Join faFunctionaries On faTransactions.intFunctionaryId = faFunctionaries.intFunctionaryId"
            mSql = mSql + " Left Join faFunds On faFunds.intFundId = faTransactions.intFundId"
    '        mSQL = mSQL + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
            mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID"
            'mSQL = mSQL + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
            mSql = mSql + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
            mSql = mSql + " Left Join faAccountHeads On faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID"
            mSql = mSql + " Left Join faBanks On faVouchers.intKeyID1 = faBanks.intAccountHeadID"
            mSql = mSql + " Where Cast(faVouchers.intVoucherNo as varchar(20)) = '" & mVoucherNo & "'"
    '        mSQL = mSQL + " And faVouchers.tnyCancelFlag <> 1"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!dtDate <> gbTransactionDate Then ' AIBY : BLOCKED along with Date Change 09-Oct-2014
                    cmdSave.Enabled = False
                End If
            
            
            
'                cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
 '               cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                txtVoucherNo.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intVocherID
             
                txtReference.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                txtReference.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'                txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
'                txtFund.Tag = IIf(IsNull(Rec.Fields(34)), "", Rec.Fields(34)) 'intFundID
'                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
'                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
'                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
'                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                
                cmbInstruments.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                cmbInstruments.ItemData(cmbInstruments.ListIndex) = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                txtAccountCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
                
                If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
'                    txtAccountNo.Text = IIf(IsNull(Rec!vchAccountNumber), "", Rec!vchAccountNumber)
                    txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtIssuedDate.Text = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
                    txtInstDate.Text = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
'                    txtNameOfBank.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
'                    txtBranch.Text = IIf(IsNull(Rec!vchBranch), "", Rec!vchBranch)
                End If
                
                txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
    
                mSqlAccHeads = "Select * From faTransactionChild"
    '            mSqlAccHeads = mSqlAccHeads + " Inner Join faTransactionChild On faVoucherChild.intAccountHeadID = faTransactionChild.intAccountHeadID"
                mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & txtReference.Tag
                mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
                RecAccHeads.Open mSqlAccHeads, mCnn
                mRowCount = 1
                While Not Rec.EOF
                    While Not RecAccHeads.EOF
                        vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                        vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                        vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
    '                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
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
        Else
            MsgBox "Can't edit this entry", vbCritical
            Exit Sub
        End If
    End Sub
    
    Private Sub Calculate()
            Dim mLoopCount As Long
            Dim mCr As Currency
            For mLoopCount = 1 To vsGrid.Rows
                If val(vsGrid.TextMatrix(mLoopCount, 4)) > 0 Then
                    vsGrid.TextMatrix(mLoopCount, 4) = Format(val(vsGrid.TextMatrix(mLoopCount, 4)), "0.00")
                    mCr = mCr + val(vsGrid.TextMatrix(mLoopCount, 4))
                Else
                    Exit For
                End If
            Next mLoopCount
            txtDr.Text = Format(mCr, "0.00")
    End Sub

    Private Sub ShowSearchAccountHead()
            Dim mSql As String
            If cmbInstruments.ListIndex > 0 Then
                If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeTransferCredit Then
                    If gbLBPanchayat = 1 Then
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where vchAccountHeadcode in('" & gbAcHeadCodeTreasuryAccount6 & "','" & gbAcHeadCodeTreasuryAccount4 & "','" & gbAcHeadCodeTreasuryAccount5 & "','" & gbAcHeadCodeTreasuryAccount7 & "')"    ','" & gbAcHeadCodeTreasuryAccount3 & "','" & gbAcHeadCodeTreasuryAccount1 & "',
                    Else
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where vchAccountHeadcode in('" & gbAcHeadCodeTreasuryAccount6 & "','" & gbAcHeadCodeTreasuryAccount4 & "','" & gbAcHeadCodeTreasuryAccount5 & "','" & gbAcHeadCodeTreasuryAccount7 & "')"    ','" & gbAcHeadCodeTreasuryAccount3 & "','" & gbAcHeadCodeTreasuryAccount1 & "',
                    End If
                Else
                    Select Case cmbInstruments.ItemData(cmbInstruments.ListIndex)
                        Case 1 '[Cash]
                         mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.tinHiddenFlag = 0 AND  faAccountHeads.intGroupID = " & faCash
                        Case Else
                         mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID =" & faBank
                    End Select
                End If
                frmSearchAccountHeads.VoucherMode = 300
                frmSearchAccountHeads.SQLString = mSql
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.Show vbModal
                txtAccountCode.SetFocus
            Else
                MsgBox "Please select an Instrument", vbInformation
            End If
    End Sub

    Private Sub DisplayBankInfo(intAcID As Long)
    '    Dim objBank As clsBank
    '    objBank.SetBankInfoByAccID intAcID
    '    If objBank.BankID > 0 Then
    '        txtNameOfBank.Text = objBank.BankName
    '        txtBranch.Text = objBank.Branch
    '        txtAccountCode.Text = objBank.AccountNumber
    '    Else
    '        txtNameOfBank.Text = ""
    '        txtBranch.Text = ""
    '        txtAccountCode.Text = ""
    '    End If
    End Sub

    Private Sub FillGridCombo()
            Dim objdb As New clsDB
            Dim RecAccHead As New ADODB.Recordset
            Dim mItem As String
            
            RecAccHead.CursorLocation = adUseClient
            Set RecAccHead = GetRecordSet("spGetAccHead4Payments", adOpenStatic, adLockReadOnly)
            While Not RecAccHead.EOF
                mItem = mItem + "|" + RecAccHead!vchAccountHead
                RecAccHead.MoveNext
            Wend
            RecAccHead.Close
            'vsGrid.ColComboList(2) = mItem
    End Sub

    Private Sub FormInitialize()
        vsGrid.Rows = 1
        vsGrid.Rows = 50
        
        txtVoucherNo.Tag = ""
        cmbTransactionType.ListIndex = -1
        txtReference.Text = ""
        txtReference.Tag = ""
        
        'txtFunctionary.Text = "Accounts Department"
        'txtFunctionary.Tag = 4
        'txtFunction.Text = "Accounts"
        'txtFunction.Tag = 6
        
        
        'txtFund.Text = "General Fund"
        'txtFund.Tag = gbFundID
        
        txtAccountCode.Text = ""
        txtAccountCode.Tag = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        
        'txtNameOfBank.Text = ""
        'txtBranch.Text = ""
        'txtAccountNo.Text = ""
        txtRef.Text = ""
        
        cmbInstruments.ListIndex = -1
        dtpIssueDate.Value = Date
        dtpInstDate.Value = Date
        txtNarration.Text = ""
        txtDr.Text = ""
        
        txtDate.Text = gbTransactionDate
        txtDate.Locked = True
        
        mRemittanceModule = 0
        mdtLastRemittance = ""
        mCopiedAmount = ""
        
        vsGrid.Editable = flexEDKbdMouse
    End Sub



    Private Sub cmbInstruments_Click()
        txtAccountCode.Tag = -1
        txtAccountCode.Text = ""
        txtAccountHead.Text = ""
        'txtNameOfBank.Text = ""
        'txtBranch.Text = ""
        'txtAccountNo.Text = ""
        txtRef.Text = ""
        If cmbInstruments.ListIndex > 0 Then
            If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
                txtRef.Enabled = True
                txtInstDate.Enabled = True
                dtpInstDate.Enabled = True
                txtIssuedDate.Enabled = True
                dtpIssuedDate.Enabled = True
                'fraBank.Enabled = True
            Else
                txtRef.Enabled = False
                txtInstDate.Enabled = False
                dtpInstDate.Enabled = False
                txtIssuedDate.Enabled = False
                dtpIssuedDate.Enabled = False
                'fraBank.Enabled = False
            End If
        End If
    End Sub

    Private Sub cmbInstruments_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then PressTabKey
    End Sub

    Private Sub cmbInstruments_LostFocus()
        Dim objAcc As New clsAccounts
        objAcc.SetAccountCode (gbAcHeadCodeCash)
        If cmbInstruments.ListIndex > -1 Then    'Added on 04/10/2011
            If cmbInstruments.ItemData(cmbInstruments.ListIndex) = gbInstrumentCash Then
                txtAccountCode.Text = objAcc.AccountCode
                txtAccountHead.Text = objAcc.AccountHead
                txtAccountHead.Tag = objAcc.AccountHeadID
            Else
            objAcc.SetAccountID (gbDefaultBankID)
                txtAccountCode.Text = objAcc.AccountCode
                txtAccountHead.Text = objAcc.AccountHead
                txtAccountHead.Tag = objAcc.AccountHeadID
            End If
        End If
    End Sub

Private Sub cmbSource_Click()
    If cmbSource.ListIndex > 0 Then
        If cmbSource.ItemData(cmbSource.ListIndex) = 29 Then
            cmbCategory.ListIndex = 2
            cmbCategory.Enabled = False
            cmbCategory.Text = "SCP"
        ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 30 Then
            cmbCategory.ListIndex = 3
            cmbCategory.Enabled = False
            cmbCategory.Text = "TSP"
        ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 3 Then ' B-Fund
            cmbCategory.ListIndex = 0
            cmbCategory.Enabled = False
        ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 10 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 11 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 12 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 13 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 14 Then

            cmbCategory.Enabled = True
            cmbCategory.Text = "GENERAL"
      ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 2 Then ' Centrally Sponsored Scheme Fund
            cmbCategory.ListIndex = 1
            cmbCategory.Enabled = False
            cmbCategory.Text = "GENERAL"
        Else
            cmbCategory.ListIndex = 1
            cmbCategory.Enabled = False
            cmbCategory.Text = "GENERAL"
        End If
    End If
End Sub

    Private Sub cmbTransactionType_Validate(Cancel As Boolean)
        Dim mSql As String
        On Error Resume Next
        lblSource.Visible = False
        lblCategory.Visible = False
        cmbSource.Visible = False
        cmbCategory.Visible = False
        
        If cmbTransactionType.ListIndex > 0 Then
            If mPreviousYearMode <> 1 Then
            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactiontypeDailyCollection Then
                If cmdSave.Tag = 0 Then
                    frmListOfDailyCollection.Show vbModal
                End If
            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeTransferCredit Then
'''                lblSource.Visible = True
'''                lblCategory.Visible = True
'''                cmbSource.Visible = True
'''                cmbCategory.Visible = True

                If gbLBPanchayat = 1 Then
                    If CheckRequsitions(1490) = 1 Then
                        mSql = mSql + "Requisitions are pending to make payments" + vbCrLf
                        mSql = mSql + " Either make Payments for the requisition Or Cancel the Requisitions" + vbCrLf
                        MsgBox mSql, vbInformation
                        cmdSave.Enabled = False
                        cmdAccoundHeads.Enabled = False
                        Exit Sub
                    Else
                        cmdSave.Enabled = True
                        cmdAccoundHeads.Enabled = True
                    End If
                Else
                    If CheckRequsitions(1535) = 1 Then
                        mSql = mSql + "Requisitions are pending to make payments" + vbCrLf
                        mSql = mSql + " Either make Payments for the requisition Or Cancel the Requisitions" + vbCrLf
                        MsgBox mSql, vbInformation
                        cmdSave.Enabled = False
                        cmdAccoundHeads.Enabled = False
                        Exit Sub
                    Else
                        cmdSave.Enabled = True
                        cmdAccoundHeads.Enabled = True
                    End If
                End If
                cmdReport.Visible = True
                txtAccountHead.Text = ""
                txtAccountHead.Tag = -1
                txtAccountCode.Text = ""
                txtAccountCode.Tag = -1
                vsGrid.Clear 1, 0
            Else
                cmdReport.Visible = False
            End If
            End If
        End If
    End Sub
    Private Function CheckRequsitions(mAccountHeadID As Integer) As Integer
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mArrIn As Variant
        Dim objdb As New clsDB
        Dim mPendingPayments As Integer
        Dim mString As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If gbLBPanchayat = 1 Then
            If mAccountHeadID = 1418 Then
                mString = " AND intSourceID IN (4,31,32,33,34,35)"
            ElseIf mAccountHeadID = 1490 Then 'DF(General)
                mString = "  AND intSourceID IN (1,2,10,11,12,13,14,27,28,21) AND intFundCategoryID=1"
            ElseIf mAccountHeadID = 1494 Then 'SCP
                mString = "  AND intSourceID IN (29,2,10,11,12,13,14) AND intFundCategoryID=2"
            ElseIf mAccountHeadID = 1495 Then 'TSP
                mString = "  AND intSourceID IN (30,2,10,11,12,13,14) AND intFundCategoryID=3"
            ElseIf mAccountHeadID = 1491 Then 'Maintainance
                mString = "  AND intSourceID IN (16,17)"
            ElseIf mAccountHeadID = 1492 Then 'CFC-Award Grant
                mString = "  AND intSourceID IN (25)"
            ElseIf mAccountHeadID = 1493 Then 'KLGSDP Grant
                mString = "  AND intSourceID IN (26)"
            Else
                Exit Function
            End If
        Else
             If mAccountHeadID = 1512 Then
                mString = "  AND intSourceID IN (4,31,32,33,34,35)"
            ElseIf mAccountHeadID = 1535 Then  'DF(General)
                mString = "  AND intSourceID IN (1,2,10,11,12,13,14,27,28,21) AND intFundCategoryID=1"
            ElseIf mAccountHeadID = 1816 Then  'SCP
                mString = "  AND intSourceID IN (29,2,10,11,12,13,14) AND intFundCategoryID=2"
            ElseIf mAccountHeadID = 1817 Then  'TSP
                mString = "  AND intSourceID IN (30,2,10,11,12,13,14) AND intFundCategoryID=3"
            ElseIf mAccountHeadID = 1539 Then  'Maintainance
                mString = "  AND intSourceID IN (16,17)"
            ElseIf mAccountHeadID = 1755 Then  'CFC-Award Grant
                mString = " AND intSourceID IN (25)"
            ElseIf mAccountHeadID = 1756 Then  'KLGSDP Grant
                mString = "  AND intSourceID IN (26)"
            Else
                Exit Function
            End If
        End If
        
        
        mSql = " SELECT intID,intPayOrderID FROM faAllotments"
        mSql = mSql + " LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID=faAllotments.intID"
        mSql = mSql + " Where IsNull(faPayOrder.intAllotmentID, 0) = 0"
        mSql = mSql + " AND ISNULL(faAllotments.tnyStatus,0)<>2 AND faAllotments.intFinancialYearID=2015"
        mSql = mSql + " And isnull(tnyTypeID,0) not in (1,2,3)"   ' Added by Anisha On 25 Nov 2015
        mSql = mSql + " " & mString & " "
        
        mSql = mSql + " Union All"
        
        mSql = mSql + " SELECT intID,intPayOrderID FROM faPayOrder"
        mSql = mSql + " INNER JOIN faAllotments ON faPayOrder.intAllotmentID=faAllotments.intID"
        mSql = mSql + " Where IsNull(faPayOrder.intVoucherID, 0) = 0 And IsNull(faAllotments.tnyStatus, 0) <> 2"
        mSql = mSql + " AND faPayOrder.intFinancialYearID=2015"
        mSql = mSql + " " & mString & " "
        
      
''''        mSql = " SELECT * FROM faAllotments "
''''        mSql = mSql + " LEFT JOIN faPayOrder ON faPayOrder.intAllotmentID=faAllotments.intID"
''''        mSql = mSql + " WHERE ISNULL(faPayOrder.intAllotmentID,0)=0  AND ISNULL(faAllotments.tnyStatus,0)<>2"
''''        mSql = mSql + " AND faAllotments.intFinancialYearID=" & gbFinancialYearID
        
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            CheckRequsitions = 1
        Else
             CheckRequsitions = 0
        End If
        
        Rec.Close
''
''        If mPendingPayments = 1 Then
''            mSql = ""
''            mSql = mSql + "Requisition are pending to make payments" + vbCrLf
''            mSql = mSql + " Either make Payments for the requisition Or Cancel the Requisitions" + vbCrLf
''
''            MsgBox mSql, vbInformation
''            cmdSave.Enabled = False
''
''        Else
''            cmdSave.Enabled = True
''        End If
        
    End Function
    Private Function CheckTransferCredit(mAccountHeadID As Integer) As Integer
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mArrIn As Variant
        Dim objdb As New clsDB
        Dim mSourceOfFundID As Integer
        Dim objAcc As New clsAccounts

        
        If gbLBPanchayat = 1 Then
            If mAccountHeadID = 1494 Then 'SCP
                    mSourceOfFundID = 29
            ElseIf mAccountHeadID = 1495 Then 'TSP
                    mSourceOfFundID = 30
            ElseIf mAccountHeadID = 1492 Then 'CFC-Award Grant
                    mSourceOfFundID = 25
            ElseIf mAccountHeadID = 1493 Then 'KLGSDP Grant
                    mSourceOfFundID = 26
            End If
        Else
            If mAccountHeadID = 1816 Then 'SCP
                    mSourceOfFundID = 29
            ElseIf mAccountHeadID = 1817 Then 'TSP
                    mSourceOfFundID = 30
            ElseIf mAccountHeadID = 1755 Then 'CFC-Award Grant
                    mSourceOfFundID = 25
            ElseIf mAccountHeadID = 1756 Then 'KLGSDP Grant
                    mSourceOfFundID = 26
            End If
        
        End If
        
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "SELECT ISNULL(intVoucherID,0) intVoucherID,abs(fltAmount) fltAmount,intAllotmentID FROM faAllotmentLetters WHERE tnyGroupID=30 AND intSourceOfFundID= " & mSourceOfFundID & ""
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            CheckTransferCredit = 0
             While Not Rec.EOF
                If Rec!intVoucherID = 0 Then
                    MsgBox "Previous Transfer Credit Process Is Not Completed. Contra Entry Saving is Pending", vbInformation
                    If gbLBPanchayat = 1 Then
                        vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount2
                        Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount2)
                    Else
                        vsGrid.TextMatrix(1, 1) = gbAcHeadCodeTreasuryAccount2
                        Call objAcc.SetAccountCode(gbAcHeadCodeTreasuryAccount6)
                    End If
                    vsGrid.TextMatrix(1, 2) = objAcc.AccountHead
                    vsGrid.TextMatrix(1, 4) = val(Rec!fltAmount)
                    cmbSource.Tag = Rec!intAllotmentID
                    CheckTransferCredit = 1
                    vsGrid.Editable = flexEDNone
                End If
              Rec.MoveNext
              Wend
              
              
        Else
              CheckTransferCredit = 0
        End If
        Rec.Close
        
        mSql = "SELECT  intAllotmentID FROM faAllotmentLetters WHERE tnyGroupID=30 AND intSourceOfFundID= 1 AND ISNULL(intVoucherID,0)=0 "
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
            cmbCategory.Tag = Rec!intAllotmentID
        End If
        Rec.Close
'''        mSql = "SELECT intkeyID1,* FROM faVouchers "
'''        mSql = mSql + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
'''        If gbLBPanchayat = 1 Then
'''            mSql = mSql + " WHERE intTransactionTypeID=4010 AND intKeyID1=" & mAccountHeadID
'''        Else
'''            mSql = mSql + " WHERE intTransactionTypeID=4006 AND intKeyID1=" & mAccountHeadID
'''        End If
'''
'''        Rec.Open mSql, mCnn
'''        If Not (Rec.EOF And Rec.BOF) Then
'''            If Rec!intKeyID1 = mAccountHeadID Then
'''                CheckTransferCredit = 1
'''            Else
'''                CheckTransferCredit = 0
'''            End If
'''        End If

        
        
        mCnn.Close
     End Function
   
    Private Sub cmdAccoundHeads_Click()
        If cmbInstruments.ListIndex > 0 Then
            Call txtAccountCode_KeyDown(vbKeyF4, 0)
        Else
            MsgBox "Please select Instrument", vbInformation
            cmbInstruments.SetFocus
        End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdNew_Click()
        Call FormInitialize
        cmdSave.Enabled = True
        txtVoucherNo.Text = ""
        cmbInstruments.Locked = False
        vsGrid.Editable = flexEDKbdMouse
    End Sub

    Private Sub cmdReport_Click()
        frmViewAllotmentLetter.Mode = 10
        'frmViewAllotmentLetter.ArrayIn = Array(CStr(gbFinancialYearID))
        Unload Me
        frmViewAllotmentLetter.Show vbModal
    End Sub

'''    Private Sub cmdReject_Click()   'ADDED BY MINU FOR REJECTION ON 18/01/2011
'''        frmReject.Mode = 7
'''        'frmReject.RequestTypeID =
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''        cmdSave.Enabled = False
'''    End Sub

    Private Sub cmdSave_Click()
        Dim objAcc As New clsAccounts
        Dim objdb As New clsDB
        Dim objInstrument As New clsInstruments
        Dim mCnt    As Integer
        Dim mTotalDrs As Double
        
        Dim mYearID As Integer
        Dim mDate As Date
        
        
        '----------------------------------------------------'
        ' Validations
        '----------------------------------------------------'
        ' Transaction Type
        If cmbTransactionType.ListIndex < 1 Then
            MsgBox "Please Select the Contra Type"
            cmbTransactionType.SetFocus
            Exit Sub
        End If
        
'''        If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeTransferCredit Then
'''            If cmbSource.ListIndex < 1 Then
'''                MsgBox "Please Select the Source Of Fund"
'''                cmbSource.SetFocus
'''                Exit Sub
'''            End If
'''            If cmbCategory.ListIndex < 1 Then
'''                MsgBox "Please Select the Category"
'''                cmbCategory.SetFocus
'''                Exit Sub
'''            End If
'''        End If
        
        
        ' Debit Account Head
        objAcc.SetAccountCode (Trim(txtAccountCode.Text))
        If objAcc.AccountHeadID < 0 Then
            MsgBox "Select a Cash or Bank Account Head!", vbInformation
            txtAccountCode.SetFocus
            Exit Sub
        End If
        For mCnt = 1 To vsGrid.Rows - 1
            If txtAccountCode.Text = vsGrid.TextMatrix(mCnt, 1) Then
                MsgBox "Not a Valid Transaction,Please Check Account Heads", vbInformation
                Exit Sub
            End If
            If Trim(vsGrid.TextMatrix(mCnt, 1)) <> "" Then
                mTotalDrs = mTotalDrs + val(vsGrid.TextMatrix(mCnt, 4))
            End If
        Next
        
        '-------------------------'
        ' Debit and Credit Amount '
        '-------------------------'
        Call Calculate
        If val(txtDr.Text) <= 0 Then
            MsgBox "Check the Amount!!", vbInformation
            vsGrid.SetFocus
            Exit Sub
        End If
        If val(txtDr.Text) <> mTotalDrs Then
            MsgBox "Check the Amount ", vbInformation
            Exit Sub
        End If
        '-------------------------'
        ' Bank details required   '
        '-------------------------'
        objInstrument.SetInstrumentType (cmbInstruments.ItemData(cmbInstruments.ListIndex))
        If Len(txtRef.Text) < 0 Then
            MsgBox "Enter the Cheque or DD No.!", vbInformation
            txtRef.SetFocus
            Exit Sub
        End If
       
        '----------------------------------------------------'
        '  UPDATING DATABASE                                 '
        '----------------------------------------------------'
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        Dim Rec                 As New ADODB.Recordset
        Dim mCnn                As ADODB.Connection
        Dim mintTransactionID   As Long
        Dim mIntTransactionTypeID As Variant
        Dim mLoopCrl            As Long
        Dim mintByLedgerID      As Long
        
        Dim mintFundID          As Variant
        Dim mintFunctionID      As Variant
        Dim mintFunctionaryID   As Variant
        Dim mintFieldID         As Variant
        Dim mintBudgetCentreID  As Variant
        Dim mintProcessID       As Long
        
        Dim mintVoucherID       As Double
        Dim mInstrumentTypeID   As Variant
        Dim mInstrumentNo       As Variant
        Dim mBank          As Variant
        Dim mBranch             As Variant
        Dim mSql                As String
        Dim mInstrumentDate     As Variant
        'Dim mDate               As String
        Dim mExtModuleId        As Integer
        
        Dim mStr                As String
        
        
        mintFundID = Null
        mintFunctionID = Null
        mintFunctionaryID = Null
        mintFieldID = Null
        mintBudgetCentreID = Null
        
        mintProcessID = 0

        mintFundID = gbFundID
        mintFunctionaryID = gbFunctionaryAccountsDepartmentID
        mintFunctionID = gbFunctionAccountsID
        mintFieldID = Null
        mintBudgetCentreID = Null

        '-------------------------------------------------------'
        ' faVoucher
        '-------------------------------------------------------'
                If cmbInstruments.ListIndex > -1 Then
                    mInstrumentTypeID = cmbInstruments.ItemData(cmbInstruments.ListIndex)
                Else
                    mInstrumentTypeID = Null
                End If
                
                If cmbTransactionType.ListIndex > 0 Then
                    mIntTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                Else
                    MsgBox "Please Select Contra Type", vbApplicationModal
                    Exit Sub
                End If
                        
                If mInstrumentTypeID = 5 Then
                    If txtRef.Text = "" Then
                        MsgBox "Enter the Cheque No", vbCritical
                        txtRef.SetFocus
                        Exit Sub
                    End If
                    If txtAccountHead.Tag <> "" Then
                        objBk.SetBankInfoByAccID objCr.AccountHeadID
                        If objBk.BankAccountHeadID > -1 Then
                            mBank = objBk.BankName
                            mBranch = objBk.Branch
                        Else
                            mBank = ""
                            mBranch = ""
                        End If
                    End If
                    
                End If
                
                If mInstrumentTypeID = 1 Then
                    mInstrumentDate = Null
                    mInstrumentNo = Null
                Else
                    mInstrumentDate = txtInstDate.Text
                    mInstrumentNo = val(Trim(txtRef.Text))
                End If
                    
                '--------Date Change for Edit--------------
                If gDateValidation(CDate(txtDate.Text)) = False Then
                    If mPreviousYearMode = 0 Then
                    MsgBox "Please Enter Valid Date", vbApplicationModal
                    Exit Sub
                    End If
                End If
                    
                If txtDate.Text <> "" Then
                    mDate = Format(txtDate.Text, "dd/MMM/yy")
                Else
                    mDate = gbTransactionDate
                End If
                '----------------------------------------------
                    
                If vsGrid.TextMatrix(vsGrid.Row, 4) <> "" Then
                    If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then
                        MsgBox "Please select the Account Head", vbCritical
                        Exit Sub
                    End If
                End If
            
                '----------------------------------------------------'
                '  SET YEAR and DATE BY CHECKING PREVIOUS YEAR MODE  '
                '----------------------------------------------------'
                If mPreviousYearMode = 1 Then
                    mYearID = gbFinancialYearID - 1
                    If IsDate(txtDate.Text) Then
                        mDate = txtDate.Text
                    Else
                        MsgBox "Transaction Date is not specified!", vbInformation
                        Exit Sub
                    End If
                    cmdNew.Enabled = False
                Else
                    If IsDate(txtDate.Text) Then
                        mDate = txtDate.Text
                    Else
                        mDate = gbTransactionDate
                    End If
                    mYearID = gbFinancialYearID
                End If
                    
                

                If CDate(mDate) <= GetLastReconDate(txtAccountHead.Tag) Then
                    mStr = ""
                    mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                    mStr = mStr + " No new Transaction is allowed to Enter during the period."
                     MsgBox mStr, vbInformation
                     txtAccountHead.Text = ""
                     txtAccountHead.Tag = -1
                     Exit Sub
                End If

                
                '-------------------------------------------------------'
                ' Connection And Transaction Begins                     '
                '-------------------------------------------------------'
                objdb.SetConnection mCnn
                mCnn.BeginTrans
                On Error GoTo ErrRollBack:
                
                '-------------------------------------------------------------------'
                '         According to User wise & Transaction Type wise            '
                '                           Sinoj                                   '
                '-------------------------------------------------------------------'
                If gbLBType = 3 Or gbLBType = 4 Then    'For MUNICIPALITY AND COPORATION
                    If gbSeatGroupID = gbSeatGroupChiefCashier Or gbSeatGroupID = gbSeatGroupAccountsClerk Then '''' Checking User Types
                        If mIntTransactionTypeID = gbTransactionTypeContraContingentPension Or _
                            mIntTransactionTypeID = gbTransactionTypeContraRegularPension Or _
                            (mIntTransactionTypeID = gbTransactiontypeDailyCollection And IsNumeric(mCopiedAmount) And val(mCopiedAmount) <> val(txtDr.Text)) Then
                            '-----------------------------------------------------------'
                            '                       Demand Saving                       '
                            ' Generating Demand No in the Form of Contra Voucher Number with the Prefix of #
                            '-----------------------------------------------------------'
                            Dim mDemandNo As String
                            Dim mDemandID As Variant
                            mDemandNo = Trim(txtVoucherNo.Text)
                            If mDemandNo = "" Then
                                mSql = "Declare @vchrNo varchar(20)" & vbNewLine
                                mSql = mSql + "Exec spGetVoucherNo Null,30," & gbFinancialYearID & ",@vchrNo Out" & vbNewLine
                                mSql = mSql + "Select @vchrNo intVoucherNo"
                                Rec.Open mSql, mCnn
                                If Not (Rec.BOF And Rec.EOF) Then
                                    mDemandNo = Rec!intVoucherNo
                                End If
                                Rec.Close
                                Rec.Open "Select isNull(Count(*)+1,1) TotalCount From faIDemandTBL Where tnyExtModuleID = 25"
                                mDemandNo = CStr(CDbl(mDemandNo) + Rec!TotalCount)
                                mDemandNo = "#" + mDemandNo
                                Rec.Close
                            End If
                            
                            mDemandID = IIf(val(txtVoucherNo.Tag) < 1, Null, CStr(txtVoucherNo.Tag))
                            Dim mUDemand As uDemand
                            ' Saving to Demand TBL
                            With mUDemand
                                .intLBID = gbLocalBodyID
                                .tnyExtAppID = AppID.Saankhya
                                If mIntTransactionTypeID = gbTransactiontypeDailyCollection Then
                                    .tnyExtModuleID = mRemittanceModule
                                Else
                                    .tnyExtModuleID = 25
                                End If
                                .tnyDemandType = 30
                                .intTransactionTypeID = mIntTransactionTypeID
                                .intYearID = mYearID
                                .tnyPeriodID = Null
                                .dtDemandDate = IIf(IsDate(txtDate.Text), txtDate.Text, gbTransactionDate)
                                .numSubLedgerID = Null
                                .intKeyID = val(txtAccountHead.Tag)
                                .intKeyID2 = Null
                                .vchRemarks = Trim(txtNarration.Text)
                                .tnyStatus = 0
                                .intVoucherID = Null
                                .dtVoucherDate = Null
                                .tnyArrearFlag = Null
                                .dtExpiryDate = gbTransactionDate
                                .numDemandID = IIf(val(txtVoucherNo.Tag) < 1, Null, val(txtVoucherNo.Tag))
                                .intFinancialYearID = mYearID
                                .numSeatID = gbSeatID
                                .intSectionID = gbSectionID
                                .numUserID = gbUserID
                                .numCounterID = gbCounterID
                                .vchAdminNote = Null
                                .vchDemandNo = mDemandNo
                                .numZoneID = Null
                                .intWardNo = Null
                                .intDoorNo = Null
                                .vchDoorNo2 = Null
                                .numForwardedSeatID = Null
                                .dtDueDate = gbTransactionDate
                                .intInstrumentTypeID = mInstrumentTypeID
                                .vchInstrumentNo = mInstrumentNo
                                .dtInstrumentDate = mInstrumentDate
                                .vchDrawnFrom = Null
                                .vchDrawnPlace = Null
                                .tnyAccrualType = Null
                                .numLocationID = gbLocationID
                                .intFunctionID = Null
                                .intFunctionaryID = Null
                                .intSourceFundID = Null
    
                                arrInput = Array(.intLBID, _
                                                .tnyExtAppID, _
                                                .tnyExtModuleID, _
                                                .tnyDemandType, _
                                                .intTransactionTypeID, _
                                                .intYearID, _
                                                .tnyPeriodID, _
                                                .dtDemandDate, _
                                                .numSubLedgerID, _
                                                .intKeyID, _
                                                .intKeyID2, _
                                                .vchRemarks, _
                                                .tnyStatus, _
                                                .intVoucherID, _
                                                .dtVoucherDate, _
                                                .tnyArrearFlag, _
                                                .dtExpiryDate, _
                                                mDemandID, _
                                                .intFinancialYearID, _
                                                .numSeatID, _
                                                .intSectionID, _
                                                .numUserID, _
                                                .numCounterID, _
                                                .vchAdminNote, _
                                                .vchDemandNo, .numZoneID, .intWardNo, .intDoorNo, .vchDoorNo2, .numForwardedSeatID, .dtDueDate, .intInstrumentTypeID, .vchInstrumentNo, .dtInstrumentDate, .vchDrawnFrom, .vchDrawnPlace, .tnyAccrualType, .numLocationID, .intFunctionID, .intFunctionaryID, .intSourceFundID, .dtDemandDate, 1)
                            End With
                            objdb.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn
                            
                            mDemandID = arrOutPut(0, 0)
                            mDemandNo = arrOutPut(1, 0)
                            txtVoucherNo.Text = arrOutPut(1, 0)
                            mCnn.Execute "Delete From faIDemandChild Where numDemandID=" & mDemandID
                            ' Demand Saving To Child
                            Dim mUDemandChild As uDemandChild
                            For mLoopCrl = 1 To vsGrid.Rows - 1
                                With mUDemandChild
                                    .numDemandID = mDemandID
                                    .intLBID = gbLocalBodyID
                                    .tnySlNo = mLoopCrl
                                    If vsGrid.Cell(flexcpText, mLoopCrl, 1) <> "" Then
                                        objAcc.SetAccountCode (vsGrid.Cell(flexcpText, mLoopCrl, 1))
                                        .intAccountHeadID = objAcc.AccountHeadID
                                        .vchAccountHeadCode = objAcc.AccountCode
                                    Else
                                        Exit For
                                    End If
                                    .fltAmount = val(vsGrid.Cell(flexcpText, mLoopCrl, 4))
                                    .vchRemarks = ""
                                    .tnyStatus = 0
                                    .dtOnDate = gbTransactionDate
                                    .intYearID = mYearID
                                    .tnyPeriodID = 3
                                    .tnyArrearFlag = 0
                                    
                                    arrInput = Array(.numDemandID, _
                                                    .intLBID, _
                                                    .tnySlNo, _
                                                    .intAccountHeadID, _
                                                    .vchAccountHeadCode, _
                                                    .fltAmount, _
                                                    .vchRemarks, _
                                                    .tnyStatus, _
                                                    .dtOnDate, _
                                                    .intYearID, _
                                                    .tnyPeriodID, _
                                                    .tnyArrearFlag)
                                End With
                                objdb.ExecuteSP "spSaveIDemandChild", arrInput, , , mCnn
                            Next mLoopCrl
                            ' Demand Save To Demand Address
                            
                            arrInput = Array(mDemandID, _
                                            gbLocalBodyID, _
                                            Null, Null, Null, _
                                            Null, Null, Null, _
                                            Null, Null, Null, _
                                            Null, Null, Null, _
                                            Null, Null, Null, Null, IIf(mDemandID < 1, 0, 1))
                                
                            objdb.ExecuteSP "spSaveIDemandAddress", arrInput, , , mCnn
                            ''---------------------------------------------         Saving to faConfig
                            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactiontypeDailyCollection Then
                                If IsDate(mdtLastRemittance) Then
                                    mSql = " Update faConfig Set dtLastRemittance = '" & Format(mdtLastRemittance, "dd-MMM-yyyy") & "'"
                                    mCnn.Execute mSql
                                End If
                            End If
                            cmdSave.Enabled = False
                            mCnn.CommitTrans
                            Exit Sub
                        End If
                    End If
                End If
                '-------------------------------------------------------------------'       Demand Save Completed For Contra
                
                mExtModuleId = mRemittanceModule
                
                arrInput = Array( _
                    IIf(txtVoucherNo.Tag = "", -1, txtVoucherNo.Tag), _
                    gbLocalBodyID, _
                    Null, _
                    mIntTransactionTypeID, _
                    30, _
                    Null, _
                    Null, _
                    mDate, _
                    val(txtDr), _
                    cmbInstruments.ItemData(cmbInstruments.ListIndex), _
                    Trim(txtRef.Text), _
                    Format(mInstrumentDate, "DD/mmm/yy"), _
                    Trim(txtNarration), _
                    gbLocationID, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    gbUserID, _
                    gbCounterID, _
                    Null, _
                    val(txtAccountHead.Tag), Null, 115, _
                    mExtModuleId, _
                    mYearID, Null, Null, Null, mBank, mBranch, mintFundID, gbSeatID, gbSessionID, txtReference.Text, Null, Null, Null, Null, gbLocationID, Null, Null)
            
                objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID = arrOutPut(0, 0)
                    If arrOutPut(0, 0) <> "" Then
                        mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        End If
                        Rec.Close
                    End If
                Else
                    GoTo ErrRollBack:
                End If

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
                Dim mnumDemandID_9          As Variant
                Dim mfltAmount_10           As Double
                
                mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & mintVoucherID
                
                For mLoopCrl = 1 To vsGrid.Rows - 1
                    If vsGrid.Cell(flexcpText, mLoopCrl, 1) <> "" Then
                        
                        objAcc.SetAccountCode (vsGrid.Cell(flexcpText, mLoopCrl, 1))
                        
                        mintLocalBodyID_2 = gbLocalBodyID
                        mintSlNo_3 = mLoopCrl
                        
                        mintAccountHeadID_4 = objAcc.AccountHeadID
                        mtnyDebitOrCredit_5 = 1
                        mintYearID_6 = mYearID
                        mtnyPeriodID_7 = 3
                        mtnyArrearFlag_8 = 0
                        mnumDemandID_9 = Null
                        mfltAmount_10 = val(vsGrid.Cell(flexcpText, mLoopCrl, 4))
                        
                        arrInput = Array( _
                        mintVoucherID, _
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
                    Else
                        Exit For
                    End If
                Next mLoopCrl
                '-------------------------------------------------------'
                ' faVoucher Address
                '-------------------------------------------------------'
        
        
                '-------------------------------------------------------'
                ' faTransactions
                '-------------------------------------------------------'
                arrInput = Array(IIf(txtReference.Tag = "", -1, txtReference.Tag), _
                           gbLocalBodyID, _
                           mYearID, _
                           Format(mDate, "DD/MmM/YYYY"), _
                           0, _
                           mExtModuleId, _
                           mintFunctionID, _
                           mintFunctionaryID, _
                           mintFieldID, _
                           mintFundID, _
                           mintBudgetCentreID, _
                           txtNarration.Text, _
                           Null, _
                           0, _
                           "C", _
                           30, _
                           Null, _
                           Null, _
                           gbUserID, _
                           mintVoucherID _
                           )
                
                Rec.CursorLocation = adUseClient
                Call objdb.ExecuteSP("spSaveTransactions", arrInput, arrOutPut, , mCnn)
                
                '-------------------------------------------------------'
                ' faTransactionChild
                '-------------------------------------------------------'
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintTransactionID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack:
                End If
                objAcc.SetAccountCode (Trim(txtAccountCode.Text))
                If objAcc.AccountHeadID < 1 Then
                    GoTo ErrRollBack
                End If
                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mintTransactionID
                mintByLedgerID = objAcc.AccountHeadID
                arrInput = Array(mintTransactionID, _
                            1, _
                            objAcc.AccountHeadID, _
                            Format(val(txtDr.Text), "0.00"), _
                            0, _
                            Null, _
                            Trim(txtNarration.Text), _
                            mintFundID _
                            )
                objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                For mLoopCrl = 1 To vsGrid.Rows - 1
                     If Trim(vsGrid.TextMatrix(mLoopCrl, 1)) = "" Then
                         Exit For
                     End If
                     objAcc.SetAccountCode (Trim(vsGrid.TextMatrix(mLoopCrl, 1)))
                     If objAcc.AccountHeadID < 1 Then
                         GoTo ErrRollBack
                     End If
                     arrInput = Array(mintTransactionID, _
                             mLoopCrl + 1, _
                             objAcc.AccountHeadID, _
                             Format(val(vsGrid.TextMatrix(mLoopCrl, 4)), "0.00"), _
                             1, _
                             mintByLedgerID, _
                             Trim(vsGrid.TextMatrix(mLoopCrl, 3)), _
                             mintFundID _
                             )
                     objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                 Next mLoopCrl
                 If mIntTransactionTypeID = gbTransactiontypeDailyCollection Then
                    If val(cmdSave.Tag) = 1 Then     ' Accounts Officer
                        '---------------------------------------------------------------'
                        '                   Demand Table Updation                       '
                        mSql = "Update faIDemandTBL Set tnyStatus = 1, intVoucherID = " & mintVoucherID & " Where numDemandID = " & cmdSearchVoucherNo.Tag
                        mCnn.Execute mSql
                        '   Updating the Actual Contra
                        mCnn.Execute "Update faVouchers Set tnysync=Null,tnyVoucherGroupID = 1 Where intVoucherID = " & mintVoucherID
                        '---------------------------------'
                        '---------------------------------------------------------------'
                    End If
                 End If
        '--------------------------------------------------------------------------'    Save To Voucher Completes(Aiby)
        
        '-------------------------------------------------------------------' Sinoj
        '           Journal Creation For Treasury Type Transactions         '
        '-------------------------------------------------------------------'
        Dim mJvVoucherID As Long
        Dim mJvVoucherNo As String
        Dim mJVHead1Dr As Integer
        Dim mJVHead1Cr As Integer
        
        If mIntTransactionTypeID = gbTransactionTypeContraContingentPension Or _
            mIntTransactionTypeID = gbTransactionTypeContraRegularPension Then
            If mIntTransactionTypeID = gbTransactionTypeContraRegularPension Then
                objAcc.SetAccountCode (gbAcHeadCodeOtherReceivablesCur)
                mJVHead1Dr = objAcc.AccountHeadID           '431409901
                objAcc.SetAccountCode (gbAcHeadCodePensionAndGratuityPayable)
                mJVHead1Cr = objAcc.AccountHeadID           '350110500
            Else
                objAcc.SetAccountCode (gbAcHeadCodeContributionToPensionFundForContingentStaff)
                mJVHead1Dr = objAcc.AccountHeadID           '210300202
                objAcc.SetAccountCode (gbAcHeadCodePensionFundForContingentStaff)
                mJVHead1Cr = objAcc.AccountHeadID           '311700100
            End If
            '   Saving to Vouchers (JV)     '  [favoucher.intKeyID2 = mintVoucherID]
            mJvVoucherID = -1
            mSql = "Select intVoucherID From faVouchers Where intKeyID2 = '" & Trim(txtVoucherNo.Text) & "'"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mJvVoucherID = Rec!intVoucherID
            End If
            Rec.Close
            
            arrInput = Array( _
                        mJvVoucherID, _
                        gbLocalBodyID, _
                        Null, _
                        3001, _
                        40, _
                        Null, _
                        Null, _
                        mDate, _
                        val(txtDr), _
                        Null, _
                        Null, _
                        Null, _
                        Trim(txtNarration), _
                        gbLocationID, _
                        Null, _
                        Null, _
                        Null, _
                        Null, _
                        gbUserID, _
                        gbCounterID, _
                        Null, _
                        mJVHead1Dr, val(txtVoucherNo.Text), 115, _
                        1, _
                        mYearID, Null, Null, Null, mBank, mBranch, mintFundID, gbSeatID, gbSessionID, txtReference.Text)
        
            objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
            If IsNumeric(arrOutPut(0, 0)) Then
                mJvVoucherID = arrOutPut(0, 0)
                mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mJvVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mJvVoucherNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                End If
                Rec.Close
            End If
            '   Updating the Actual Contra
            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyVoucherGroupID = 1, intKeyID2  = " & mJvVoucherNo & " Where intVoucherID = " & mintVoucherID
            '---------------------------------'
            '   Saving to Voucher Child (JV)
            mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & mJvVoucherID
            arrInput = Array(mJvVoucherID, _
                            gbLocalBodyID, _
                            1, _
                            mJVHead1Cr, _
                            0, _
                            mYearID, _
                            3, _
                            0, _
                            Null, _
                            val(txtDr))
            objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
            '   Saving to Voucher Address (JV)
            
            '   Saving to Transactions (JV)
            Dim mJVTransactionID As Long
            mJVTransactionID = -1
            Rec.Open "Select intTransactionID From faTransactions Where intVoucherID = " & mJvVoucherID, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mJVTransactionID = Rec!intTransactionID
            End If
            Rec.Close
            arrInput = Array(mJVTransactionID, _
                           gbLocalBodyID, _
                           mYearID, _
                           Format(mDate, "DD/MMM/YYYY"), _
                           0, _
                           0, _
                           mintFunctionID, _
                           mintFunctionaryID, _
                           mintFieldID, _
                           mintFundID, _
                           mintBudgetCentreID, _
                           txtNarration.Text, _
                           Null, _
                           0, _
                           "JV", _
                           40, _
                           Null, _
                           Null, _
                           gbUserID, _
                           mJvVoucherID)
                
                Rec.CursorLocation = adUseClient
                Call objdb.ExecuteSP("spSaveTransactions", arrInput, arrOutPut, , mCnn)
                If IsNumeric(arrOutPut(0, 0)) Then
                    mJVTransactionID = arrOutPut(0, 0)
                Else
                    GoTo ErrRollBack:
                End If
            '   Saving to Transaction Child (JV)
            mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mJVTransactionID
            ' First Row
            arrInput = Array(mJVTransactionID, _
                             1, _
                             mJVHead1Dr, _
                             val(txtDr), _
                             1, _
                             Null, _
                             Null, _
                             mintFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            ' Second Row
            arrInput = Array(mJVTransactionID, _
                             2, _
                             mJVHead1Cr, _
                             val(txtDr), _
                             0, _
                             mJVHead1Dr, _
                             Null, _
                             mintFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
            '---------------------------------------------------------------'
            '                   Demand Table Updation                       '
            mSql = "Update faIDemandTBL Set tnyStatus = 1, intVoucherID = " & mintVoucherID & " Where numDemandID = " & cmdSearchVoucherNo.Tag
            mCnn.Execute mSql
            '---------------------------------------------------------------'
        End If
        '   Contra Entry List Updating
        With frmListOfContraEntries
            If .mGridRow > 0 Then
                .vsGrid.Cell(flexcpForeColor, .mGridRow, 7) = vbBlue
                .vsGrid.Cell(flexcpText, .mGridRow, 7) = mJvVoucherNo
                .vsGrid.Cell(flexcpText, .mGridRow, 0) = txtVoucherNo.Text
                .vsGrid.Cell(flexcpText, .mGridRow, 6) = txtDr.Text
                .vsGrid.Cell(flexcpText, .mGridRow, 9) = 1
            End If
        End With
        '-------------------------------------------------------------------'           Journal Completes Sinoj
        
        '-------------------------------------------------------'
        ' To Update faConfig(dtLastRemittance) in the case of Daily Collection from JSK     Poornima On 15 / 10 / 2010
        If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactiontypeDailyCollection Then
            If IsDate(mdtLastRemittance) Then
                mSql = " Update faConfig Set dtLastRemittance = '" & Format(mdtLastRemittance, "dd-MMM-yyyy") & "'"
                mCnn.Execute mSql
            End If
        End If
        '-------------------------------------------------------'
        '-------------------------------------------------------'
        ' Connection And Transaction Ends                       '
        '-------------------------------------------------------'
        mCnn.CommitTrans
        
        If mPreviousYearMode = 1 Then
            mSql = "Update faPendingTaskRequest SET tnyStatus = 8 WHERE intRequestID = " & mPreviousYearRequestID
            mCnn.Execute mSql
        End If
         
        '-----------------------------------------------------------
        ' To Update faBankSource tnyStatus=9 If Bank Balance is Zero
        '''Modified On 14 Dec 2015 By Anisha
        Dim mCn         As ADODB.Connection
        objdb.SetConnection mCn
        objdb.ExecuteSP "spUpdateBankSourceStatus", , , , mCn
        ''''Call SaveAllotmentLetters(mintVoucherID, mCnn)
        
        '-------------------------------------------------------'
        ' Save AllotmentLetter For Transfer Credit
        '-------------------------------------------------------'
        
        mSql = ""
        mSql = " UPDATE faAllotmentLetters SET intVoucherID=" & mintVoucherID & " WHERE intAllotmentID =" & val(cmbSource.Tag)
        mCnn.Execute mSql
        mSql = ""
        mSql = " UPDATE faAllotmentLetters SET intVoucherID=" & mintVoucherID & " WHERE intAllotmentID =" & val(cmbCategory.Tag)
        mCnn.Execute mSql
        
        '-----------------------------------------------------------
        
        '-------------------------------------------------------'
        ' To Print CotraVoucher                                 '
        'Call PrintVoucher(txtVoucherNo.Text)
        
        '-------------------------------------------------------'
        
        'Call FormInitialize
        
        
        cmdSave.Enabled = False
        mPreviousYearMode = 0
        mPreviousYearRequestID = Null
            
        frmViewVoucher.FormName = "PaymentVoucher"
        frmViewVoucher.ArrayIn = Array(CStr(mintVoucherID))
        frmViewVoucher.Show vbModal
        Exit Sub
        
ErrRollBack:
        txtVoucherNo.Text = ""
''        cmdSave.Enabled = False
        Debug.Print Error$
        mCnn.RollbackTrans
    End Sub
    Private Sub PrintVoucher(intVoucherID As Double)
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        'objDB.ExecuteSP spGetVoucherDetails
        
        
        'MsgBox (intVoucherID)
    End Sub
'Private Sub cmdSubLedger_Click()
'    Dim mSql As String
'    mSql = "Select vchSubAccountHead, intSubAccountID From faSubsidiaryAccounts Order By vchSubAccountHead"
'    Call PopulateList(lstMasters, mSql, , True, , True)
'    lstMasters.Tag = "6"
'    lstMasters.Visible = True
'    lstMasters.SetFocus
'End Sub

    Private Sub cmdSearchVoucherNo_Click()
        frmSearchPaymentVoucher.TransactionGroupId = 30
        frmSearchPaymentVoucher.Show vbModal
        txtVoucherNo.Text = gbSearchStr
        txtVoucherNo.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        txtVoucherNo.SetFocus
        If txtVoucherNo.Text <> "" Then
            txtVoucherNo_LostFocus
        End If
    End Sub

    Private Sub dtpInstDate_CloseUp()
        txtInstDate.Text = CheckDateInMMM(dtpInstDate.Value)
    End Sub

    Private Sub dtpInstDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    
    Private Sub dtpIssueDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    
    Private Sub dtpIssuedDate_CloseUp()
        txtIssuedDate.Text = CheckDateInMMM(dtpIssuedDate.Value)
    End Sub
    
    Private Sub Form_Activate()
        Me.Top = 0
        Me.Left = 0
    End Sub
    Private Sub Form_Load()
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        
        
        XPC.InitIDESubClassing
        mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType WHERE intGroupID = 30  AND (ISNULL(tnyHidden, 0) <> 1) Order By intTransactionTypeID "
        PopulateList cmbTransactionType, mSql, , True, True, True
        
        mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Order By vchInstrumentType "
        PopulateList cmbInstruments, mSql, "Cheque", True, True, True
        
        
         
        
        If gbLBPanchayat = 1 Then
            mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,25,26,27,28,10,11,12,13,14,19,21,29,30,41,42,43,44,45)"
        Else
            mSql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,19,21,25,26,27,28,29,30,41,42,43,44,45)"
        End If
        PopulateList cmbSource, mSql, , True, True, True, enuSourceString.Saankhya
        
        mSql = "SELECT vchTransactionCategory,intCategoryID FROM faTransactionCategory"
        PopulateList cmbCategory, mSql, True, True, True, True
        
        
        Call FillGridCombo
        vsGrid.ColComboList(1) = "|..."
        FormInitialize
        
        If mPreviousYearMode = 1 Then
            
            If objdb.SetConnection(mCnn) Then
                mSql = "SELECT * FROM faPendingTaskRequest WHERE intRequestID= " & mPreviousYearRequestID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    cmdNew.Enabled = False
                    txtDate.Text = DdMmmYy(Rec!dtTransactionDate)
                    txtDate.Enabled = False
                    txtDr.Tag = Rec!fltAmount
                    
                    
                    
                    Dim objAcc As New clsAccounts
                    'Dim i As Integer
                    'Dim mLastRow As Integer
                    'mLastRow = 0
                    If val(txtDr.Tag) > 0 Then
                        With frmContraEntry
                            
                            .copiedAmount = val(txtDr.Tag)
                            .cmdSearchVoucherNo.Tag = -1
                            .txtVoucherNo.Tag = -1
                            .txtVoucherNo.Text = ""
                            .txtReference.Text = ""
                            .txtReference.Tag = "" '
                            
                            If .cmbInstruments.ListCount > 0 Then
                                .cmbInstruments.Text = "Cash"
                            End If
                            
                            ' Credit Account Filling
                            objAcc.SetAccountCode (gbAcHeadCodeCash)
                            .txtAccountHead.Tag = objAcc.AccountHeadID
                            .txtAccountCode.Text = objAcc.AccountCode
                            .txtAccountHead.Text = objAcc.AccountHead
                            .txtRef.Text = ""
                            .txtIssuedDate.Text = ""
                            .txtInstDate.Text = ""
                            
         
                            'Grid Filling
                            .vsGrid.Rows = 1
                            .vsGrid.Rows = 10
                            Call objAcc.SetAccountID(gbDefaultBankID)
                            .vsGrid.TextMatrix(1, 1) = objAcc.AccountCode
                            .vsGrid.TextMatrix(1, 2) = objAcc.AccountHead
                            .vsGrid.TextMatrix(1, 4) = Format(txtDr.Tag, "0.00")
                            .txtDr.Text = Format(txtDr.Tag, "0.00")
                            .mRemittanceModule = 60 'Previous Year
                            
                        End With
                    Else
                        frmContraEntry.mRemittanceModule = 0
                        MsgBox "Requested transaction amount is not available..!", vbInformation
                    End If
                    
                    
                    
                Else
                    MsgBox "Didn't able to find Previous Year Transaction Request!", vbInformation
                    FormInitialize
                    Unload Me
                End If
                Rec.Close
            End If ' Connection
        End If
        
    End Sub


'Private Sub lstMasters_DblClick()
'    If lstMasters.ListIndex > -1 Then
'    gbSearchStr = lstMasters.Text
'    gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
'    Select Case val(lstMasters.Tag)
'        Case 1: txtFunction.SetFocus
'        Case 2: txtFunctionary.SetFocus
'        'Case 3: txtField.SetFocus
'        Case 4: txtFund.SetFocus
'        'Case 5: txtBudgetCentre.SetFocus
'        'Case 6: txtSubsidiaryLedger.SetFocus
'    End Select
'    End If
'End Sub

'Private Sub lstMasters_GotFocus()
'    Dim mWidth As Long
'    Dim mLeft As Long
'    Dim mTop As Long
'    Select Case val(lstMasters.Tag)
'        Case 1, 2, 4: mTop = 915: mWidth = 4000: mLeft = 2500
'        Case 6: mTop = 915: mWidth = 4000: mLeft = 4500
'    End Select
'    lstMasters.Top = mTop
'    lstMasters.Width = mWidth
'    lstMasters.Left = mLeft
'End Sub

'Private Sub lstMasters_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call PressTabKey
'        Call lstMasters_DblClick
'    End If
'End Sub

'Private Sub lstMasters_LostFocus()
'    If lstMasters.ListIndex > -1 Then
'        gbSearchStr = lstMasters.Text
'        gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
'    End If
'    lstMasters.Visible = False
'    Select Case val(lstMasters.Tag)
'        Case 1: txtFunction.SetFocus
'        Case 2: txtFunctionary.SetFocus
'        'Case 3: txtField.SetFocus
'        Case 4: txtFund.SetFocus
'        'Case 5: txtBudgetCentre.SetFocus
'        'Case 6: txtSubsidiaryLedger.SetFocus
'    End Select
'End Sub


    Private Sub txtAccountCode_GotFocus()
        Dim mDate       As Date
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtAccountCode.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            
           If CDate(txtDate.Text) <= GetLastReconDate(txtAccountHead.Tag) Then
                mStr = ""
                mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                mStr = mStr + " No new Transaction is allowed to Enter during the period."
                MsgBox mStr, vbInformation
                txtAccountHead.Text = ""
                txtAccountHead.Tag = -1
                txtAccountCode.Text = ""
                Exit Sub
            End If
        End If
        txtAccountCode.SelStart = 0
        txtAccountCode.SelLength = Len(txtAccountCode)
    End Sub

    Private Sub txtAccountCode_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            Call ShowSearchAccountHead
        End If
    End Sub

    Private Sub txtAccountCode_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub

    Private Sub txtAccountCode_LostFocus()
        Dim mChequeNo As Variant
        Dim mSql As String
        
        objCr.SetAccountCode Trim(txtAccountCode.Text)
        If objCr.AccountHeadID > 0 Then
            txtAccountHead.Text = objCr.AccountHead
            txtAccountCode.Text = objCr.AccountCode
            objBk.SetBankInfoByAccID objCr.AccountHeadID
            
            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeTransferCredit Then
                'If CheckTransferCredit(objCr.AccountHeadID) = 0 Then
                    If CheckRequsitions(objCr.AccountHeadID) = 1 Then
                        mSql = mSql + "Requisitions are pending to make payments" + vbCrLf
                        mSql = mSql + " Either make Payments for the requisition Or Cancel the Requisitions" + vbCrLf
                        MsgBox mSql, vbInformation
                        cmdSave.Enabled = False
                        Exit Sub
                    Else
                        If CheckTransferCredit(objCr.AccountHeadID) = 0 Then
                            cmdSave.Enabled = True
                            frmSourceFundSplitUp.TreasuryID = objCr.AccountHeadID
                            frmSourceFundSplitUp.Show vbModal
                        End If
                    End If
'''                Else
'''                    MsgBox "Already TRANSFER CREDITED", vbInformation
'''                    Exit Sub
                'End If
            End If
            
            If objBk.BankAccountHeadID > -1 Then
                'txtNameOfBank.Text = objBk.BankName
                'txtBranch.Text = objBk.Branch
                'txtAccountNo.Text = objBk.AccountNumber
                'mChequeNo = objBk.GetNeWChequeNumber
                'txtRef.Text = IIf(IsNull(mChequeNo), "", mChequeNo)
            Else
                'txtNameOfBank.Text = ""
                'txtBranch.Text = ""
                'txtAccountNo.Text = ""
                'txtRef.Text = ""
            End If
        Else
            txtAccountHead.Text = ""
            txtAccountCode.Text = ""
            txtAccountHead.Tag = ""
        End If
    End Sub

    Private Sub txtAccountHead_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then
            Call txtAccountCode_KeyDown(vbKeyF4, 0)
        End If
    End Sub

    Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub


'Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PressTabKey
'End Sub

'Private Sub txtBranch_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PressTabKey
'End Sub

'Private Sub txtFieldCode_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PressTabKey
'End Sub
'
'Private Sub txtBudgetCentre_GotFocus()
'    If gbSearchStr <> "" Then
'        Dim objBc As New clsBudgetCentre
'        objBc.SetBudgetCentreByID gbSearchID
'        If objBc.BudgetCentreID > -1 Then
'            txtBudgetCentre.Text = objBc.BudgetCentreCode
'            txtBudgetCentre.Tag = gbSearchID
'        End If
'        gbSearchStr = ""
'        gbSearchID = -1
'    End If
'End Sub

'Private Sub txtBudgetCentre_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'        Call cmdBudgetCentres_Click
'    End If
'End Sub

'Private Sub txtBudgetCentre_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PressTabKey
'End Sub
'
'Private Sub txtBudgetCentre_LostFocus()
'        '------------------------------------------------------'
'        ' Searches and Finding Function, Functionary and Field '
'        '------------------------------------------------------'
'        txtBudgetCentre = Trim(txtBudgetCentre)
'        If Len(Trim(txtBudgetCentre)) Then
'            Dim objBudCen As New clsBudgetCentre
'            objBudCen.SetBudgetCentre (txtBudgetCentre.Text)
'            If objBudCen.BudgetCentreID < 1 Then
'                txtBudgetCentre.Text = ""
'                txtBudgetCentre.Tag = ""
'            Else
'                txtBudgetCentre.Text = objBudCen.BudgetCentreCode
'                txtBudgetCentre.Tag = objBudCen.BudgetCentreID
'                txtFunction.Text = objBudCen.FunctionName
'                txtFunction.Tag = objBudCen.FunctionID
'                txtFunctionary.Text = objBudCen.FunctionaryName
'                txtFunctionary.Tag = objBudCen.FuntionaryID
'                txtField.Text = objBudCen.FieldName
'                txtField.Tag = objBudCen.FieldID
'            End If
'        End If
'End Sub
'
'Private Sub txtField_GotFocus()
'    If gbSearchStr <> "" Then
'        txtField.Text = gbSearchStr
'        txtField.Tag = gbSearchID
'        gbSearchStr = ""
'        gbSearchID = -1
'    End If
'End Sub

'Private Sub txtField_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'        Call cmdFields_Click
'    End If
'End Sub

    Private Sub txtField_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    
    Private Sub txtDate_DblClick()
''''        If txtVoucherNo.Tag <> "" Then
''''            If MsgBox("Do You want to Change the date", vbYesNo) = vbYes Then
''''                 txtDate.Locked = False
''''                 txtDate.SetFocus
''''            Else
''''                 txtDate.Locked = True
''''            End If
''''        End If
    End Sub

    Private Sub txtDate_LostFocus()
        txtDate.Text = CheckDateInMMM(txtDate.Text)
        If gDateValidation(CDate(txtDate.Text)) = False Then
            MsgBox "Please Enter Valid Date", vbApplicationModal
            Exit Sub
        End If
    End Sub

'    Private Sub txtFunction_GotFocus()
'        If gbSearchStr <> "" Then
'            txtFunction.Text = gbSearchStr
'            txtFunction.Tag = gbSearchID
'            gbSearchStr = ""
'            gbSearchID = -1
'        End If
'    End Sub
'
'    Private Sub txtFunction_KeyDown(KeyCode As Integer, Shift As Integer)
'        If KeyCode = vbKeyF4 Then
'            Call cmdFunctions_Click
'        End If
'    End Sub
'    Private Sub txtFunction_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then PressTabKey
'    End Sub
'    Private Sub txtFunctionary_GotFocus()
'        If gbSearchStr <> "" Then
'            txtFunctionary.Text = gbSearchStr
'            txtFunctionary.Tag = gbSearchID
'            gbSearchStr = ""
'            gbSearchID = -1
'        End If
'    End Sub
'    Private Sub txtFunctionary_KeyDown(KeyCode As Integer, Shift As Integer)
'        If KeyCode = vbKeyF4 Then
'            Call cmdFunctionaries_Click
'        End If
'    End Sub
'    Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then PressTabKey
'    End Sub
'    Private Sub txtFund_GotFocus()
'        If gbSearchStr <> "" Then
'            txtFund.Text = gbSearchStr
'            txtFund.Tag = gbSearchID
'            gbSearchStr = ""
'            gbSearchID = -1
'        End If
'    End Sub
'    Private Sub txtFund_KeyDown(KeyCode As Integer, Shift As Integer)
'        If KeyCode = vbKeyF4 Then
'            Call cmdFunds_Click
'        End If
'    End Sub
'    Private Sub txtFund_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then PressTabKey
'    End Sub
'Private Sub txtNameOfBank_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then PressTabKey
'End Sub

    Private Sub txtInstDate_LostFocus()
        If txtInstDate.Text <> "" Then
            txtInstDate.Text = CheckDateInMMM(txtInstDate.Text)
        End If
    End Sub
    
    Private Sub txtIssuedDate_LostFocus()
        If txtIssuedDate.Text <> "" Then
            txtIssuedDate.Text = CheckDateInMMM(txtIssuedDate.Text)
        End If
    End Sub
    
    Private Sub txtNarration_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    Private Sub txtRef_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
    End Sub
    
    Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
        Call KeyPressNumber(KeyAscii, "#-")
    End Sub

    Private Sub txtVoucherNo_LostFocus()
        If Trim(txtVoucherNo.Text) <> "" Then
            If mID(Trim(txtVoucherNo.Text), 1, 1) = "#" Then
                If mID(Trim(txtVoucherNo.Text), 2, 1) <> "3" Then
                    MsgBox "Invalid Contra Voucher Number", vbInformation
                    Exit Sub
                End If
            ElseIf mID(Trim(txtVoucherNo.Text), 1, 1) <> "3" Then
                MsgBox "Invalid Contra Voucher Number", vbInformation
                Exit Sub
            End If
            Call ListContraDemandOrVoucher(Trim(txtVoucherNo.Text))
        End If
    End Sub
    
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        Dim objBank     As New clsBank
            If vsGrid.Row > 1 Then
                If vsGrid.TextMatrix(vsGrid.Row - 1, 1) = "" Or _
                   vsGrid.TextMatrix(vsGrid.Row - 1, 2) = "" Or _
                   val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 Then
                   Cancel = True
                   Exit Sub
                End If
            End If
            If Len(gbSearchStr) Then
                vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
                vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
                vsGrid.Col = vsGrid.Col + 2
                vsGrid.Redraw = flexRDDirect
                gbSearchStr = ""
                gbSearchID = -1
            End If

    End Sub
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        If cmbInstruments.ListIndex > 0 And txtAccountCode.Text <> "" Then
            frmSearchAccountHeads.VoucherMode = 301     '       Contra Debit Mode
            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeTransferCredit Then
'''                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.vchAccountHeadCode in ('" & gbAcHeadCodeTreasuryAccount2 & "')"
'''                frmSearchAccountHeads.chkListAll.Enabled = False
'''                frmSearchAccountHeads.cmdsearch.Enabled = False
'''                frmSearchAccountHeads.Show vbModal
                
            Else
                If cmbInstruments.ItemData(cmbInstruments.ListIndex) = 1 Then
                    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where intGroupID  IN (2) And tinHiddenFlag <> 1 And intAccountHeadID <> " & val(txtAccountHead.Tag) & " Order by vchAccountHeadCode"
                    frmSearchAccountHeads.chkListAll.Enabled = False
                    frmSearchAccountHeads.Show vbModal
                Else
                    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where intGroupID  IN (1,2) And tinHiddenFlag <> 1 And intAccountHeadID <> " & val(txtAccountHead.Tag) & " Order by vchAccountHeadCode"
                    frmSearchAccountHeads.chkListAll.Enabled = False
                    frmSearchAccountHeads.Show vbModal
                End If
            End If
        Else
            MsgBox "Please select the Instrument Or Account Head", vbInformation
            cmbInstruments.SetFocus
        End If
        
    End Sub
    
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        Dim mStr As String
        Dim mAmt As Double
        If vsGrid.Col = 2 And Trim(vsGrid.Text) <> "" Then
            Set objAccHead = New clsAccounts
            If objAccHead.FindAccountByHead(Trim(vsGrid.Text)) Then
                vsGrid.TextMatrix(vsGrid.Row, 1) = objAccHead.AccountCode
            End If
        ElseIf vsGrid.Col = 1 And Trim(vsGrid.Text) <> "" Then
            Set objAccHead = New clsAccounts
            Call objAccHead.SetAccountCode(Trim(vsGrid.Text))

            If CDate(txtDate.Text) <= GetLastReconDate(objAccHead.AccountHeadID) Then
                mStr = ""
                mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                mStr = mStr + " No new Transaction is allowed to Enter during the period."
                MsgBox mStr, vbInformation
                vsGrid.TextMatrix(vsGrid.Row, 1) = ""
                vsGrid.TextMatrix(vsGrid.Row, 2) = ""
                Exit Sub
            End If

        ElseIf vsGrid.Col = 4 Then
            mAmt = Format(val(vsGrid.TextMatrix(vsGrid.Row, 4)), "#0")   ''Modified On 30-Aug-2014
            vsGrid.TextMatrix(vsGrid.Row, 4) = Format(mAmt, "0.00")
            Call Calculate
        End If
    End Sub
    
    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If vsGrid.Col = 4 Then
            vsGrid.EditMaxLength = 15
        End If
    End Sub

    Private Sub vsGrid_Validate(Cancel As Boolean)
        If vsGrid.Col = 3 Then
            If Len(vsGrid.TextMatrix(vsGrid.Row, 3)) > 100 Then
                vsGrid.TextMatrix(vsGrid.Row, 3) = Left(vsGrid.TextMatrix(vsGrid.Row, 3), 100)
            End If
        End If
    End Sub

    Public Sub ListContraDemandOrVoucher(mVoucherNo As String)
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        Dim mSql As String
        Dim mCount As Integer
        Dim mTotal As Double
        
        mTotal = 0
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya            '''' Connection to Saankhya ''''
        If InStr(1, mVoucherNo, "-") > 0 Or InStr(1, mVoucherNo, "#") > 0 Then
            mSql = "Select" & vbNewLine
            mSql = mSql + "    faIDemandTBL.numDemandID intVoucherID, vchDemandNo intVoucherNo,'' vchRefNo,Null intTransactionID,dtDemandDate dtDate,faInstrumentTypes.vchInstrumentType,faIDemandChild.intAccountHeadID,faIDemandChild.fltAmount," & vbNewLine
            mSql = mSql + "    faIDemandTBL.intInstrumentTypeID,faIDemandTBL.intKeyID intKeyID1,faIDemandTBL.vchInstrumentNo,faIDemandTBL.dtInstrumentDate,faIDemandTBL.vchRemarks vchDescription,vchTransactionType,faIDemandTBL.tnyStatus" & vbNewLine
            mSql = mSql + "From faIDemandTBL" & vbNewLine
            mSql = mSql + "Left Join faTransactionType On faIDemandTBL.intTransactionTypeID = faTransactionType.intTransactionTypeID" & vbNewLine
            mSql = mSql + "Inner Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID" & vbNewLine
            mSql = mSql + "Left Join faInstrumentTypes On faIDemandTBL.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID" & vbNewLine
            mSql = mSql + "Where faIDemandTBL.vchDemandNo = '" & mVoucherNo & "'"
        Else
            mSql = "Select" & vbNewLine
            mSql = mSql + "    faVouchers.intVoucherID,faVouchers.intVoucherNo,faVouchers.vchRefNo,faTransactions.intTransactionID,faVouchers.dtDate,faInstrumentTypes.vchInstrumentType,faVoucherChild.intAccountHeadID,faVoucherChild.fltAmount," & vbNewLine
            mSql = mSql + "    faVouchers.intInstrumentTypeID , faVouchers.intKeyID1, faVouchers.vchInstrumentNo, faVouchers.dtInstrumentDate, faVouchers.vchDescription,vchTransactionType,isNull(faVouchers.tnyVoucherGroupID,0) tnyStatus, isNull(faVouchers.intExternalModuleID,0) ModuleID" & vbNewLine
            mSql = mSql + "From faVouchers" & vbNewLine
            mSql = mSql + "Left Join faTransactions On faVouchers.intVoucherID = faTransactions.intVoucherID" & vbNewLine
            mSql = mSql + "Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID" & vbNewLine
            mSql = mSql + "Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID" & vbNewLine
            mSql = mSql + "Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID" & vbNewLine
            mSql = mSql + "Where faVouchers.tnyVoucherTypeID = 30 And Cast(faVouchers.intVoucherNo as varchar(20)) = '" & mVoucherNo & "'"
        End If
        mCount = 1
        Rec.Open mSql, mCnn
        If Not (Rec.BOF And Rec.EOF) Then
                
            '   Save Button Enable or Disable
            If Rec!tnyStatus < 1 Then            ''' Actually I used tnyVoucherGroupID as tnyStatus because tnyVoucherGroupID = 1, intKeyID2 = JV VouchrNo For This JV Required Contra
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
            If IsNull(Rec!vchTransactionType) Then
                cmbTransactionType.ListIndex = -1
            Else
                cmbTransactionType.Text = Rec!vchTransactionType
            End If
            
            If Rec!dtDate <> gbTransactionDate Then ' AIBY : BLOCKED along with Date Change 09-Oct-2014
                cmdSave.Enabled = False
            End If
            '----------------------------
            '****Reversed Contra*********
            If InStr(1, mVoucherNo, "#") <> 1 Then
                If Rec!ModuleID = 55 Then
                    MsgBox "You Are Not Allowed to Edit Reversed Contra Voucher", vbInformation
                    cmdSave.Enabled = False
                    'Exit Sub
                End If
            End If
            '----------------------------
            
            
            '''     Approving Officer can only Approve; No Editing; Editing can be Done By Cash clerck,Chief Cashiers
            cmdSearchVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID) 'intVocherID
            If cmdSave.Tag = 0 Then
                txtVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID) 'intVocherID
            Else
                txtVoucherNo.Tag = -1
            End If
            txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo) 'intVocherNo
            txtReference.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            txtReference.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
            txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            cmbInstruments.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
            txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
            Call objAcc.SetAccountID(txtAccountHead.Tag)
            txtAccountCode.Text = objAcc.AccountCode
            txtAccountHead.Text = objAcc.AccountHead
            If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
                txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                txtIssuedDate.Text = IIf(IsNull(Rec!dtInstrumentDate), gbTransactionDate, Rec!dtInstrumentDate)
                txtInstDate.Text = IIf(IsNull(Rec!dtInstrumentDate), gbTransactionDate, Rec!dtInstrumentDate)
            End If
            txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)

            While Not Rec.EOF
                vsGrid.Rows = vsGrid.Rows + 1
                Call objAcc.SetAccountID(val(Rec!intAccountHeadID))
                vsGrid.TextMatrix(mCount, 1) = objAcc.AccountCode
                vsGrid.TextMatrix(mCount, 2) = objAcc.AccountHead
                vsGrid.TextMatrix(mCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                mTotal = mTotal + IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                Rec.MoveNext
                mCount = mCount + 1
            Wend
            txtDr.Text = mTotal
        End If
        Rec.Close
    End Sub
    Private Function GetLastReconDate(intBankID As Integer) As Variant
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
   
    Public Property Let LastRemittanceDate(mData As Date)
        mdtLastRemittance = mData
    End Property

    Public Property Get LastRemittanceDate() As Date
        LastRemittanceDate = mdtLastRemittance
    End Property
    Public Property Let copiedAmount(mData As Variant)
        mCopiedAmount = mData
    End Property

    Public Property Get copiedAmount() As Variant
        copiedAmount = mCopiedAmount
    End Property

    Public Property Let PreviousYearMode(mData As Variant)
        mPreviousYearMode = mData
    End Property
    
    Public Property Let PreviousYearRequestID(mData As Variant)
        mPreviousYearRequestID = mData
    End Property


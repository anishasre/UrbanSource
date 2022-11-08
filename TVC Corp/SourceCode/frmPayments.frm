VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPayments 
   BackColor       =   &H00B0E1E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C8E7E7&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6210
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C8E7E7&
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
      Height          =   375
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6210
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8E7E7&
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
      Height          =   375
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6210
      Width           =   1215
   End
   Begin VB.ListBox lstMasters 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   11790
      TabIndex        =   56
      Top             =   2835
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DAF2F2&
      Height          =   1710
      Left            =   15
      TabIndex        =   53
      Top             =   -60
      Width           =   11850
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   3765
         TabIndex        =   61
         Top             =   600
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   39911
      End
      Begin VB.CommandButton cmdSearchPaymentOrder 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4935
         Picture         =   "frmPayments.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   930
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchVoucherNo 
         Height          =   315
         Left            =   3750
         Picture         =   "frmPayments.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtDate 
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
         Left            =   1680
         TabIndex        =   23
         Top             =   585
         Width           =   2040
      End
      Begin VB.TextBox txtFunctionary 
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
         Left            =   7665
         TabIndex        =   6
         Top             =   585
         Width           =   3735
      End
      Begin VB.TextBox txtFund 
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
         Left            =   7665
         TabIndex        =   4
         Top             =   255
         Width           =   3735
      End
      Begin VB.CommandButton cmdFunctionaries 
         BackColor       =   &H00C8E7E7&
         Caption         =   "..."
         Height          =   300
         Left            =   11430
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   585
         Width           =   285
      End
      Begin VB.CommandButton cmdFunctions 
         BackColor       =   &H00C8E7E7&
         Caption         =   "..."
         Height          =   300
         Left            =   11430
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   900
         Width           =   285
      End
      Begin VB.CommandButton cmdFunds 
         BackColor       =   &H00DAF2F2&
         Caption         =   "..."
         Height          =   300
         Left            =   11430
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   285
      End
      Begin VB.TextBox txtFunction 
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
         Left            =   7665
         TabIndex        =   8
         Top             =   915
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SEARCH"
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
         Left            =   3600
         TabIndex        =   54
         Top             =   12120
         Width           =   1155
      End
      Begin VB.TextBox txtVoucherNo 
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
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   2040
      End
      Begin VB.ComboBox cmbTransactionType 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1275
         Width           =   10050
      End
      Begin VB.TextBox txtPaymentOrderNo 
         BackColor       =   &H80000009&
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
         Left            =   1680
         TabIndex        =   2
         Top             =   930
         Width           =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7185
         TabIndex        =   26
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6855
         TabIndex        =   28
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6555
         TabIndex        =   27
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0E1E1&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Date"
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
         Left            =   465
         TabIndex        =   22
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1305
         Width           =   1275
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Order"
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
         Left            =   300
         TabIndex        =   25
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DAF2F2&
         Caption         =   "Voucher No"
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
         Left            =   630
         TabIndex        =   0
         Top             =   285
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAF2F2&
      Height          =   5100
      Left            =   -15
      TabIndex        =   52
      Top             =   1575
      Width           =   11850
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   -3570
         Top             =   5010
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   1560
         Left            =   60
         TabIndex        =   17
         Top             =   3015
         Width           =   11700
         _cx             =   20637
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
         ForeColor       =   1786190
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
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPayments.frx":01F4
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
      Begin VB.ComboBox cmbInstruments 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   2040
      End
      Begin VB.TextBox txtAccountHead 
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
         Left            =   3720
         MaxLength       =   500
         TabIndex        =   32
         Top             =   540
         Width           =   7680
      End
      Begin VB.TextBox txtDr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   9405
         TabIndex        =   50
         Top             =   4590
         Width           =   2100
      End
      Begin VB.Frame fraBank 
         BackColor       =   &H00DAF2F2&
         Height          =   1245
         Left            =   300
         TabIndex        =   51
         Top             =   855
         Width           =   6060
         Begin VB.TextBox txtRef 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4305
            MaxLength       =   50
            TabIndex        =   15
            Top             =   510
            Width           =   1665
         End
         Begin VB.TextBox txtBranch 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4305
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   13
            Top             =   180
            Width           =   1665
         End
         Begin VB.TextBox txtNameOfBank 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            Top             =   225
            Width           =   1935
         End
         Begin VB.TextBox txtAccountNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            Top             =   540
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtpDueDate 
            Height          =   300
            Left            =   4305
            TabIndex        =   40
            Top             =   810
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   61997059
            CurrentDate     =   39291
         End
         Begin MSComCtl2.DTPicker dtpIssueDate 
            Height          =   300
            Left            =   1365
            TabIndex        =   38
            Top             =   855
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   61997059
            CurrentDate     =   39291
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issued Date"
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
            Left            =   345
            TabIndex        =   37
            Top             =   870
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque No"
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
            Left            =   3330
            TabIndex        =   36
            Top             =   540
            Width           =   930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Left            =   3645
            TabIndex        =   34
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Left            =   855
            TabIndex        =   33
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblAccountNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
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
            Left            =   315
            TabIndex        =   35
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date"
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
            Left            =   3465
            TabIndex        =   39
            Top             =   840
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdProject 
         BackColor       =   &H00C8E7E7&
         Caption         =   "..."
         Height          =   300
         Left            =   11430
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1185
         Width           =   285
      End
      Begin VB.TextBox txtProject 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ML-TTRevathi"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7665
         TabIndex        =   43
         Top             =   1185
         Width           =   3750
      End
      Begin VB.TextBox txtSubsidiaryLedger 
         Enabled         =   0   'False
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
         Left            =   7665
         TabIndex        =   46
         Top             =   1620
         Width           =   3750
      End
      Begin VB.CommandButton cmdSubLedger 
         BackColor       =   &H00C8E7E7&
         Caption         =   "..."
         Height          =   300
         Left            =   11445
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1620
         Width           =   285
      End
      Begin VB.CommandButton cmdAccoundHeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00C8E7E7&
         Caption         =   "..."
         Height          =   300
         Left            =   11430
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   540
         Width           =   285
      End
      Begin VB.TextBox txtAccountCode 
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
         Left            =   1680
         TabIndex        =   31
         Top             =   540
         Width           =   2025
      End
      Begin VB.TextBox txtNarration 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1650
         MaxLength       =   500
         TabIndex        =   16
         Top             =   2115
         Width           =   4710
      End
      Begin VB.TextBox txtClaiment 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7665
         MultiLine       =   -1  'True
         TabIndex        =   48
         Text            =   "frmPayments.frx":02B8
         Top             =   1950
         Width           =   3750
      End
      Begin VB.Label lbBudgetAllocated 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00B0E1E1&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   3345
         TabIndex        =   58
         Top             =   4680
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00B0E1E1&
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Allotted :"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   120
         TabIndex        =   57
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instruments"
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
         Left            =   525
         TabIndex        =   29
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
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
         Left            =   8820
         TabIndex        =   49
         Top             =   4620
         Width           =   555
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "      Debit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   270
         Left            =   60
         TabIndex        =   55
         Top             =   2760
         Width           =   11700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub.Ledger"
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
         Left            =   6630
         TabIndex        =   45
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7005
         TabIndex        =   42
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/c Head"
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
         Left            =   795
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
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
         Left            =   690
         TabIndex        =   41
         Top             =   2145
         Width           =   885
      End
   End
   Begin VB.ListBox lstPaymentOrder 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   2070
      TabIndex        =   60
      Top             =   1620
      Visible         =   0   'False
      Width           =   3465
   End
End
Attribute VB_Name = "frmPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************************************************************************************
'* Application ID           :                                                            *
'* Application Name         : Saankhya Double Entry                                      *
'* Screen id                : Payments                                                   *
'* Version No               : Ver 2.0.0                                                  *
'* Form Designed By         : Achu                                                       *
'* Created on               :                                                            *
'* Coded By                 :                                                            *
'* Coded on                 :                                                            *
'* Reviewed By              :                                                            *
'* Reviewed on              : 11-Sep-2007                                                *
'* Purpose                  : Manual Tracking of Payments Voucher                        *
'*                                                                                       *
'*                                                                                       *
'* Name of Database         : DB_Finance                                                 *
'* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *
'* Name of Table(s)         : faTransactions, faTransactionChild                         *
'* Look up Table(s)         : faTransactionType, faTransactionChild, faAccountHeads      *
'*                          : faBudgetCentre, faFunction, faFunctionaries, faFields      *
'*                                                                                       *
'* Stored Procedures        : spGetAccHead4Receipts, spSaveTrans,                        *
'*                          : spSaveTransactionChild                                     *
'*                                                                                       *
'*=======================================================================================*
Option Explicit
    Private objCr As New clsAccounts
    Private objBk As New clsBank
    Public mSelect As Boolean
    
    Public Sub DisplayVoucherDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDb
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        Dim mStatus         As Variant
        
        Call FormInitialize
        
        If Not IsNumeric(mVoucherNo) Then
            MsgBox "Invalid Voucher Number!", vbInformation
            txtVoucherNo.SetFocus
            Exit Sub
        End If
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQL = " Select faVouchers.*, faFunctionaries.vchFunctionary, faFunctions.vchFunction, faFunds.vchFund, faInstrumentTypes.vchInstrumentType,"
        mSQL = mSQL + " faVoucherChild.*, faVoucherAddress.*, faTransactionType.vchTransactionType, "
        mSQL = mSQL + " faTransactions.intTransactionID, faTransactions.intFunctionaryID, faTransactions.intFunctionID, "
        mSQL = mSQL + " faAccountHeads.intAccountHeadID, faAccountHeads.vchAccountHeadCode, faAccountHeads.vchAccountHead "
        mSQL = mSQL + " From faVouchers Inner Join"
        mSQL = mSQL + " faTransactions On faTransactions.intVoucherID = faVouchers.intVoucherID Left Join"
        mSQL = mSQL + " faTransactionType On faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID Left Join"
        mSQL = mSQL + " faFunctionaries On faFunctionaries.intFunctionaryID = faTransactions.intFunctionaryID Left Join"
        mSQL = mSQL + " faFunctions On faFunctions.intFunctionID = faTransactions.intFunctionID Left Join"
        mSQL = mSQL + " faFunds On faFunds.intFundID = faVouchers.intFundID Left Join"
        mSQL = mSQL + " faInstrumentTypes On faInstrumentTypes.intInstrumentTypeID = faVouchers.intInstrumentTypeID Left Join"
        mSQL = mSQL + " faVoucherChild On faVoucherChild.intVoucherID = faVouchers.intVoucherID Left Join"
        mSQL = mSQL + " faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID Left Join "
        mSQL = mSQL + " faAccountHeads On faAccountHeads.intAccountHeadID = faVouchers.intKeyID1 "
        mSQL = mSQL + " Where faVouchers.intVoucherNo = " & mVoucherNo
        
        Rec.Open mSQL, mCnn
        If (Rec.EOF And Rec.BOF) Then
            Exit Sub
        End If
        
        If mStatus = 1 Then
            MsgBox "Can not Edit this Voucher!", vbInformation
            Exit Sub
        End If
        
        If Not IsNull(Rec!intTransactionTypeID) Then
            If Not IsNull(Rec!vchTransactionType) Then
                cmbTransactionType.Text = Rec!vchTransactionType
            Else
                cmbTransactionType.ListIndex = 0
            End If
        End If
        
        txtVoucherNo.Tag = Rec.Fields(0) 'intVoucherID
        txtDate.Text = DdMmmYy(Rec!dtDate)
        txtPaymentOrderNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
        txtDate.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
        
        txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
        txtFund.Tag = IIf(IsNull(Rec!intFundID), "", Rec!intFundID)
        txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
        txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
        txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
        txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
        
        cmbInstruments.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
        cmbInstruments.ItemData(cmbInstruments.ListIndex) = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
        txtAccountCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
        Call txtAccountCode_LostFocus
        txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
        dtpIssueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
        dtpDueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
        txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)

        mSqlAccHeads = "Select * From faTransactionChild"
        mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
        mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & txtDate.Tag
        mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
        RecAccHeads.Open mSqlAccHeads, mCnn
        
        mRowCount = 1
        While Not Rec.EOF
            While Not RecAccHeads.EOF
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
                vsGrid.Rows = vsGrid.Rows + 1
                mRowCount = mRowCount + 1
                RecAccHeads.MoveNext
            Wend
            Rec.MoveNext
        Wend
        RecAccHeads.Close
        Call Calculate
   
    Rec.Close
        
End Sub
    
    
    Public Sub DisplayReceiptDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objDB           As New clsDb
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        Dim mStatus         As Variant
        
        Call FormInitialize
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSQL = "Select tnyStatus From faVouchers"
        mSQL = mSQL + " Where intVoucherNo = " & mVoucherNo
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), Null, Rec!tnyStatus)
        End If
        Rec.Close
        If mStatus = 0 Or IsNull(mStatus) Then
            mSQL = "Select * From faVouchers"
            mSQL = mSQL + " Inner Join faTransactions On faTransactions.intVoucherId = faVouchers.intVoucherId"
            mSQL = mSQL + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
            mSQL = mSQL + " Left Join faFunctions On fatransactions.intFunctionId = faFunctions.intFunctionId"
            mSQL = mSQL + " Left Join faFunctionaries On faTransactions.intFunctionaryId = faFunctionaries.intFunctionaryId"
            mSQL = mSQL + " Left Join faFunds On faFunds.intFundId = faTransactions.intFundId"
    '        mSQL = mSQL + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
            mSQL = mSQL + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID"
            'mSQL = mSQL + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
            mSQL = mSQL + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
            mSQL = mSQL + " Left Join faAccountHeads On faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID"
            mSQL = mSQL + " Left Join faBanks On faVouchers.intKeyID1 = faBanks.intAccountHeadID"
            mSQL = mSQL + " Where faVouchers.intVoucherNo = " & mVoucherNo
    '        mSQL = mSQL + " And faVouchers.tnyCancelFlag <> 1"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                
                cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), " ", Rec!vchTransactionType)
                'cmbTransactionType.itemData(cmbTransactionType.ListIndex) = IIf(IsNull(Rec!intTransactionTypeID), " ", Rec!intTransactionTypeID)
                txtVoucherNo.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intVoucherID
                txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                txtPaymentOrderNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                'txtPaymentOrderNo.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                txtDate.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                
                txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
                txtFund.Tag = IIf(IsNull(Rec.Fields(34)), "", Rec.Fields(34)) 'intFundID
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                
                cmbInstruments.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                cmbInstruments.ItemData(cmbInstruments.ListIndex) = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                txtAccountCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
                
                If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
                    txtAccountNo.Text = IIf(IsNull(Rec!vchAccountNumber), "", Rec!vchAccountNumber)
                    txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    dtpIssueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
                    dtpDueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
                    txtNameOfBank.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
                    txtBranch.Text = IIf(IsNull(Rec!vchBranch), "", Rec!vchBranch)
                End If
                
                txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
    
                mSqlAccHeads = "Select * From faTransactionChild"
    '            mSqlAccHeads = mSqlAccHeads + " Inner Join faTransactionChild On faVoucherChild.intAccountHeadID = faTransactionChild.intAccountHeadID"
                mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & txtDate.Tag
                mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
                RecAccHeads.Open mSqlAccHeads, mCnn
                mRowCount = 1
                While Not Rec.EOF
                    While Not RecAccHeads.EOF
                        vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                        vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                        vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
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
        If val(txtDr.Text) > val(txtDr.Tag) Then
            'cmdSave.Enabled = False
        Else
            'cmdSave.Enabled = True
        End If
End Sub

Private Sub ShowSearchAccountHead()
        Dim mSQL As String
        If cmbInstruments.ListIndex > 0 Then
            Select Case cmbInstruments.ItemData(cmbInstruments.ListIndex)
                Case 1 '[Cash]
                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & faCash
'                Case 7 '[Treasury Bills]
'                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faBank
                Case Else
                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & faBank
            End Select
            frmSearchAccountHeads.SQLString = mSQL
            frmSearchAccountHeads.Show vbModal
            txtAccountCode.SetFocus
        End If
End Sub

Private Sub DisplayBankInfo(intAcID As Long)
    Dim objBank As clsBank
    objBank.SetBankInfoByAccID intAcID
    If objBank.BankID > 0 Then
        txtNameOfBank.Text = objBank.BankName
        txtBranch.Text = objBank.Branch
        txtAccountCode.Text = objBank.AccountNumber
    Else
        txtNameOfBank.Text = ""
        txtBranch.Text = ""
        txtAccountCode.Text = ""
    End If
End Sub

Private Sub FillGridCombo()
        Dim objDB As New clsDb
        Dim RecAccHead As New ADODB.Recordset
        Dim mItem As String

        RecAccHead.CursorLocation = adUseClient
        Set RecAccHead = GetRecordSet("spGetAccHead4Payments", adOpenStatic, adLockReadOnly)
        While Not RecAccHead.EOF
            mItem = mItem + "|" + RecAccHead!vchAccountHead
            RecAccHead.MoveNext
        Wend
        RecAccHead.Close
        vsGrid.ColComboList(2) = mItem
End Sub

Private Sub FormInitialize()
        vsGrid.Rows = 1
        vsGrid.Rows = 50

        cmbTransactionType.ListIndex = -1
        
        txtVoucherNo.Tag = ""
        txtDate.Text = ""
        txtDate.Tag = ""
        txtPaymentOrderNo.Text = ""
        txtPaymentOrderNo.Tag = ""
        
        txtFunctionary.Text = ""
        txtFunctionary.Tag = ""
        txtFunction.Text = ""
        txtFunction.Tag = ""


        txtFund.Text = ""
        txtFund.Tag = ""

        txtAccountCode.Text = ""
        txtAccountCode.Tag = ""
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""

        txtNameOfBank.Text = ""
        txtBranch.Text = ""
        txtAccountNo.Text = ""
        txtRef.Text = ""
        txtClaiment.Text = ""

        cmbInstruments.ListIndex = -1
        dtpIssueDate.Value = Date
        dtpDueDate.Value = Date
        txtNarration.Text = ""
        txtDr.Text = ""
        txtDr.Tag = ""
        
        txtProject.Text = ""
        txtProject.Tag = ""
        txtSubsidiaryLedger.Text = ""
        txtSubsidiaryLedger.Tag = ""
        txtClaiment.Text = ""

        dtpIssueDate.Value = gbTransactionDate
        dtpDueDate.Value = gbTransactionDate
        
        'On Error Resume Next
        cmbInstruments.Text = "Cheque"
        txtAccountCode.Text = "450210100"
        Call txtAccountCode_LostFocus
        txtFund.Tag = FindMasterID("faFunds", "intFundID", "vchFundCode", "0100")
        txtFund.Text = FindMaster("faFunds", "vchFund", "intFundID", val(txtFund.Tag))
        On Error GoTo 0
        
        mSelect = False
End Sub

Private Sub cmbField_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        End If
End Sub

Private Sub cmbFunction_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        End If
End Sub

Private Sub cmbFunctionary_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        End If
End Sub

Private Sub cmbFund_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            PressTabKey
        End If
End Sub

Private Sub cmbInstruments_Click()
    txtAccountCode.Text = ""
    txtAccountHead.Text = ""
    txtNameOfBank.Text = ""
    txtBranch.Text = ""
    txtAccountNo.Text = ""
    txtRef.Text = ""
    If cmbInstruments.ListIndex <> -1 Then
        If cmbInstruments.ItemData(cmbInstruments.ListIndex) <> 1 Then
            fraBank.Enabled = True
        Else
            fraBank.Enabled = False
        End If
    End If
End Sub

Private Sub cmbInstruments_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub cmdAccoundHeads_Click()
        Call txtAccountCode_KeyDown(vbKeyF4, 0)
End Sub

Private Sub cmdBudgetCentres_Click()
        Dim mSQL As String
        mSQL = "Select vchBudgetCentre, intBudgetCentreID From faBudgetCentres  Order By vchBudgetCentre"
        Call PopulateList(lstMasters, mSQL, , True, , True)
        lstMasters.Width = 495
        lstMasters.Left = 9330
        lstMasters.Tag = "5"
        lstMasters.Visible = True
        lstMasters.SetFocus
End Sub

Private Sub cmdFields_Click()
        Dim mSQL As String
        mSQL = "Select vchField, intFieldID From faFields Order By vchField"
        Call PopulateList(lstMasters, mSQL, , True, , True)
        lstMasters.Tag = "3"
        lstMasters.Visible = True
        lstMasters.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFunctionaries_Click()
        Dim mSQL As String
        mSQL = "Select vchFunctionary, intFunctionaryID From faFunctionaries Order By vchFunctionary"
        Call PopulateList(lstMasters, mSQL, , True, , True)
        lstMasters.Tag = "2"
        lstMasters.Visible = True
        lstMasters.SetFocus
End Sub

Private Sub cmdFunctions_Click()
        Dim mSQL As String
        mSQL = "Select vchFunction, intFunctionID From faFunctions Order By vchFunction"
        Call PopulateList(lstMasters, mSQL, , True, , True)
        lstMasters.Tag = "1"
        lstMasters.Visible = True
        lstMasters.SetFocus
End Sub

Private Sub cmdFunds_Click()
        Dim mSQL As String
        mSQL = "Select vchFund, intFundID From faFunds Where tnyActiveFlag = 1 Order By vchFund"
        Call PopulateList(lstMasters, mSQL, , True, , True)
        lstMasters.Tag = "4"
        lstMasters.Visible = True
        lstMasters.SetFocus

End Sub

Private Sub cmdNew_Click()
    Call FormInitialize
    cmdSave.Enabled = True
    txtVoucherNo.Text = ""
End Sub

Private Sub cmdProject_Click()
'   Modified by Cijith  '

''    Dim mSQL As String
''    lstMasters.Font = "ML-TTRevathi"
''    lstMasters.Left = 7695
''    lstMasters.Width = 3720
''    ''117100' Left Part
''    mSQL = "Select vchProjectName, Right(numProjectNo,6) From faSulekhaProjects Order By vchProjectName"
''    Call PopulateList(lstMasters, mSQL, , , , True)
''    lstMasters.Tag = 7
''    lstMasters.Visible = True
''    lstMasters.SetFocus

'       Aiby Sir Please Include the Project Search Window       '

End Sub

Private Sub cmdSave_Click()
        Dim objAcc As New clsAccounts
        Dim objDB As New clsDb
        Dim objInstrument As New clsInstruments
        '----------------------------------------------------'
        ' Validations
        '----------------------------------------------------'
        ' Debit Account Head
        objAcc.SetAccountCode (Trim(txtAccountCode.Text))
        If objAcc.AccountHeadID < 0 Then
            MsgBox "Select a Cash or Bank Account Head!", vbInformation
            txtAccountCode.SetFocus
            Exit Sub
        End If
        '-------------------------'
        ' Debit and Credit Amount '
        '-------------------------'
        Call Calculate
        If val(txtDr.Text) <= 0 Then
            MsgBox "Check the Amount!!", vbInformation
            vsGrid.SetFocus
            Exit Sub
        End If
        '-------------------------'
        ' Bank details required   '
        '-------------------------'
        objInstrument.SetInstrumentType (cmbInstruments.ItemData(cmbInstruments.ListIndex))
        If objInstrument.InstrumentTypeID = 5 Then
            If Len(txtNameOfBank.Text) < 0 Then
                MsgBox "Enter the name of Bank!", vbInformation
                txtNameOfBank.SetFocus
                Exit Sub
            End If
        ElseIf objInstrument.InstrumentTypeID < 1 Then
            cmbInstruments.ListIndex = 0
            MsgBox "Please choose the Instrument Type", vbInformation
            cmbInstruments.SetFocus
            Exit Sub
        End If

        If Len(txtRef.Text) < 0 Then
            MsgBox "Enter the Cheque or DD No.!", vbInformation
            txtNameOfBank.SetFocus
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
        Dim mLoopCrl            As Long
        Dim mintByLedgerID      As Long

        Dim mintFundID          As Variant
        Dim mintFunctionID      As Variant
        Dim mintFunctionaryID   As Variant
        Dim mintFieldID         As Variant
        Dim mintBudgetCentreID  As Variant
        Dim mintProcessID       As Long
        Dim mIntTransactionTypeID As Long
        

        Dim mintVoucherID       As Double
        Dim mInstrumentTypeID   As Variant
        Dim mSQL                As String

        mintFundID = Null
        mintFunctionID = Null
        mintFunctionaryID = Null
        mintFieldID = Null
        mintBudgetCentreID = Null

        mintProcessID = 0
        If val(txtFund.Tag) > 0 Then
            mintFundID = val(txtFund.Tag)
        Else
            MsgBox "Select Fund", vbInformation
            txtFund.SetFocus
            Exit Sub
        End If
        If val(txtFunctionary.Tag) > 0 Then
            mintFunctionaryID = val(txtFunctionary.Tag)
        Else
            MsgBox "Select Functionary", vbInformation
            txtFunctionary.SetFocus
            Exit Sub
        End If
        If val(txtFunction.Tag) > 0 Then
            mintFunctionID = val(txtFunction.Tag)
        Else
            MsgBox "Select Function", vbInformation
            txtFunction.SetFocus
            Exit Sub
        End If


        mintFieldID = Null
        mintBudgetCentreID = Null

        '-------------------------------------------------------'
        ' faVoucher
        '-------------------------------------------------------'
                If cmbInstruments.ListIndex > -1 Then
                    mInstrumentTypeID = cmbInstruments.ItemData(cmbInstruments.ListIndex)
                    cmbInstruments.Tag = cmbInstruments.ItemData(cmbInstruments.ListIndex)
                Else
                    mInstrumentTypeID = Null
                End If
                If mInstrumentTypeID = 5 Then
                    If txtRef.Text = "" Then
                        MsgBox "Please enter the Cheque No", vbCritical
                        txtRef.SetFocus
                        Exit Sub
                    End If
                End If
                If cmbTransactionType.ListIndex > 0 Then
                    mIntTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                Else
                    mIntTransactionTypeID = 0
                End If
                

'@intVoucherID_1    [bigint],
'@intLocalBodyID_2  [int],
'@intTransactionID_3    [bigint],
'@intTransactionTypeID_4    [int],
'@tnyVoucherTypeID_5    [tinyint],
'@intVoucherNo_6    [int],
'@intBookNo_7       [int],
'@dtDate_8      [smalldatetime],
'@fltAmount_9       [float],
'@intInstrumentTypeID_10 [int],
'@vchInstrumentNo_11    [varchar](50),
'@dtInstrumentDate_12   [smalldatetime],
'@vchDescription_13     [varchar](500),

'@numZoneID_14      [numeric],
'@numWardID_15      [numeric],
'@intDoorNoP1_16    [int],
'@vchDoorNoP2_17    [varchar](10),
'@vchDoorNoP3_18    [varchar](10),
'@intUserID_19      [int],
'@intCounterID_20   [int],

'@numSubLedgerID_21     [numeric],
'@intKeyID1_22      [int],
'@intKeyID2_23      [int],
'@intExternalApplicationID_24   [int],
'@intExternalModuleID_25    [int],
'@intFinancialYearID_26     [int],
'@tnyShiftID_27     [tinyint] = Null,
'@tnyPrintFlag_28   [tinyint] = Null,
'@tnyCancelFlag_29  [tinyint] = Null,
'
'@vchBank_33    [varchar](50)= Null,
'@vchBankPlace_34   [varchar](50)= Null,
'@intFundID_35  [int] = Null


                arrInput = Array( _
                IIf(txtVoucherNo.Tag = "", -1, txtVoucherNo.Tag), _
                gbLocalBodyID, _
                Null, _
                mIntTransactionTypeID, _
                20, _
                Null, _
                Null, _
                gbTransactionDate, _
                val(txtDr), _
                val(cmbInstruments.Tag), _
                Trim(txtRef.Text), _
                dtpDueDate.Value, _
                Trim(txtNarration), _
                Null, _
                Null, _
                Null, _
                Null, _
                Null, _
                gbUserID, _
                gbCounterID, _
                val(txtProject.Tag), _
                val(txtAccountHead.Tag), Null, 115, _
                1, _
                gbFinancialYearID, Null, Null, Null, txtNameOfBank.Text, txtBranch.Text, mintFundID, Null, Null, txtPaymentOrderNo.Text)

        '-------------------------------------------------------'
        ' Connection And Transaction Begins                     '
        '-------------------------------------------------------'
        objDB.SetConnection mCnn
        'mCnn.BeginTrans
        'On Error GoTo ErrRollBack:

                objDB.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
                If IsNumeric(arrOutPut(0, 0)) Then
                    mintVoucherID = arrOutPut(0, 0)
                    If arrOutPut(0, 0) <> "" Then
                        mSQL = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
                        Rec.Open mSQL, mCnn
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
                        mtnyDebitOrCredit_5 = 0
                        mintYearID_6 = gbFinancialYearID
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
                        objDB.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
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
                arrInput = Array(IIf(txtDate.Tag = "", -1, txtDate.Tag), _
                           gbLocalBodyID, _
                           gbFinancialYearID, _
                           Format(gbTransactionDate, "DD/MmM/YYYY"), _
                           0, _
                           0, _
                           mintFunctionID, _
                           mintFunctionaryID, _
                           mintFieldID, _
                           mintFundID, _
                           mintBudgetCentreID, _
                           txtNarration.Text, _
                           200, _
                           0, _
                           "P", _
                           20, _
                           Null, _
                           IIf(val(txtSubsidiaryLedger.Tag) > 0, val(txtSubsidiaryLedger.Tag), Null), _
                           gbUserID, _
                           mintVoucherID _
                           )

                Rec.CursorLocation = adUseClient
                Call objDB.ExecuteSP("spSaveTransactions", arrInput, arrOutPut, , mCnn)

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
                mintByLedgerID = objAcc.AccountHeadID
                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mintTransactionID
                arrInput = Array(mintTransactionID, _
                            1, _
                            objAcc.AccountHeadID, _
                            Format(val(txtDr.Text), "0.00"), _
                            0, _
                            Null, _
                            Trim(txtNarration.Text), _
                            mintFundID _
                            )
                objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
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
                     objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                 Next mLoopCrl

        '-------------------------------------------------------'
        ' Connection And Transaction Begins                     '
        '-------------------------------------------------------'
        'mCnn.CommitTrans
'        Call FormInitialize
        cmdSave.Enabled = False
        Exit Sub
ErrRollBack:
        Debug.Print Error$
        'mCnn.RollbackTrans
        
End Sub

Private Sub cmdSearchPaymentOrder_Click()
    Dim objDB As New clsDb
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    
    objDB.SetConnection mCnn
    mSQL = "Select (numPayOrderNo) , numPayOrderID from faPayOrder "
    mSQL = mSQL + " Where faPayOrder.tnyStatus=1 "
    'mSQL = mSQL + " Order By vchSubAccountHead"
    Call PopulateList(lstPaymentOrder, mSQL, , , , True)
        
        lstPaymentOrder.Visible = True
        lstPaymentOrder.ZOrder 0
        lstPaymentOrder.SetFocus
End Sub

Private Sub cmdSearchVoucherNo_Click()
    frmSearchPaymentVoucher.TransactionGroupId = 20
    
    frmSearchPaymentVoucher.Show vbModal
    txtVoucherNo.Text = gbSearchStr
    txtVoucherNo.Tag = gbSearchID
    gbSearchStr = ""
    gbSearchID = -1
    txtVoucherNo.SetFocus
End Sub

Private Sub cmdSubLedger_Click()
    On Error GoTo Err:
        Dim objSubLedger As New clsSubLedger
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchID = -1 Then Exit Sub
        txtSubsidiaryLedger.Text = gbSearchStr
        txtSubsidiaryLedger.Tag = gbSearchID
        objSubLedger.SetSubLedgerDetails (gbSearchID)
        txtClaiment.Visible = True
        txtClaiment.Text = IIf(IsNull(objSubLedger.HouseOrOffice), "", objSubLedger.HouseOrOffice)
        txtClaiment.Text = txtClaiment.Text + vbNewLine + IIf(IsNull(objSubLedger.LocalPlace), "", objSubLedger.LocalPlace)
        txtClaiment.Text = txtClaiment.Text + vbNewLine + IIf(IsNull(objSubLedger.MainPlace), "", objSubLedger.MainPlace)
        txtClaiment.Text = txtClaiment.Text + vbNewLine + IIf(IsNull(objSubLedger.Street), "", objSubLedger.Street)
        gbSearchID = -1
        gbSearchStr = ""
    Exit Sub
Err:
    MsgBox (Error$)
End Sub

Private Sub cmdSubLedger_LostFocus()
'    Dim mLength As Integer
'    Dim mSubLedgerCode As Double
'    Dim mSQL As String
'    Dim objDB As New clsDB
'    Dim Rec As New ADODB.Recordset
'    Dim mCnn As New ADODB.Connection
'
'    If txtSubsidiaryLedger.Text <> "" Then
'        mLength = InStr(CStr(txtSubsidiaryLedger.Text), " ")
'        mSubLedgerCode = mID(CStr(txtSubsidiaryLedger.Text), 1, mLength)
'            objDB.SetConnection mCnn
'        mSQL = "Select vchAddress1,vchAddress2,vchAddress3 from faSubsidiaryAccounts where vchSubAccountCode = '" & mSubLedgerCode & "'"
'        Rec.Open mSQL, mCnn
'        txtClaiment.Text = IIf(IsNull(Rec!vchAddress1), "", Rec!vchAddress1) + "      " + IIf(IsNull(Rec!vchAddress2), "", Rec!vchAddress2) + "      " + IIf(IsNull(Rec!vchAddress2), "", Rec!vchAddress2)
'        Rec.Close
'    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
    frmSearchPaymentVoucher.Show (1)
End Sub

Private Sub dtpDate_DropDown()
    txtDate.Text = dtpDate.Value
End Sub

Private Sub dtpDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub dtpIssueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Load()
    Dim mSQL As String
    mSQL = "Select vchTransactionType, intTransactionTypeID From faTransactionType WHERE intGroupID =20 Order By vchTransactionType "
    PopulateList cmbTransactionType, mSQL, , True, True, True

    mSQL = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Order By vchInstrumentType "
    PopulateList cmbInstruments, mSQL, "Cheque", True, True, True

    Call FillGridCombo
    vsGrid.ColComboList(1) = "|..."
    FormInitialize
    WindowsXPC1.InitIDESubClassing
End Sub


Private Sub lstMasters_DblClick()
    If lstMasters.ListIndex > -1 Then
    gbSearchStr = lstMasters.Text
    gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
    Select Case val(lstMasters.Tag)
        Case 1: txtFunction.SetFocus
        Case 2: txtFunctionary.SetFocus
        'Case 3: txtField.SetFocus
        Case 4: txtFund.SetFocus
        'Case 5: txtBudgetCentre.SetFocus
        Case 6: txtSubsidiaryLedger.SetFocus
        Case 7: txtProject.SetFocus
    End Select
    End If
End Sub

Private Sub lstMasters_GotFocus()
    Dim mWidth As Long
    Dim mLeft As Long
    Dim mTop As Long
    Select Case val(lstMasters.Tag)
        Case 1, 2, 4: mTop = 915: mWidth = 4000: mLeft = 2500
        Case 6: mTop = 915: mWidth = 4000: mLeft = 4500
        Case 7: mTop = 915: mWidth = 3720: mLeft = 7695
    End Select
    lstMasters.Top = mTop
    lstMasters.Width = mWidth
    lstMasters.Left = mLeft
End Sub

Private Sub lstMasters_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
        Call lstMasters_DblClick
    End If
End Sub

Private Sub lstMasters_LostFocus()
    If lstMasters.ListIndex > -1 Then
        gbSearchStr = lstMasters.Text
        gbSearchID = lstMasters.ItemData(lstMasters.ListIndex)
    End If
    lstMasters.Visible = False
    Select Case val(lstMasters.Tag)
        Case 1: txtFunction.SetFocus
        Case 2: txtFunctionary.SetFocus
        'Case 3: txtField.SetFocus
        Case 4: txtFund.SetFocus
        'Case 5: txtBudgetCentre.SetFocus
        Case 6: txtSubsidiaryLedger.SetFocus
    End Select
End Sub

Private Sub lstPaymentOrder_Click()

    Dim mSearchStr      As String
    Dim mSearchID       As Variant
    Dim mCharCnt        As Integer
    Dim mStrCnt         As Integer
    Dim mSQL            As String
    Dim mCnn            As New ADODB.Connection
    Dim rs              As New ADODB.Recordset
    Dim objDB           As New clsDb
    
        If lstPaymentOrder.ListIndex > -1 Then
            mSearchStr = lstPaymentOrder.Text
            mSearchID = lstPaymentOrder.ItemData(lstPaymentOrder.ListIndex)
            
            txtPaymentOrderNo.Text = mSearchStr
            txtPaymentOrderNo.Tag = mSearchID
            mSearchID = -1
            mSearchStr = ""
        End If
        lstPaymentOrder.Visible = False
        txtPaymentOrderNo.SetFocus
End Sub

Private Sub lstPaymentOrder_LostFocus()
    lstMasters.Font = "Areal"
    lstPaymentOrder.Visible = False
End Sub

Private Sub txtAccountCode_GotFocus()
    If gbSearchStr <> "" Then
        Dim mStr As String
        txtAccountCode.Text = Token(gbSearchStr, " ")
        txtAccountHead.Text = Trim(gbSearchStr)
        txtAccountHead.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
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
    If KeyAscii = 13 Then
        PressTabKey
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtAccountCode_LostFocus()
    Dim mChequeNo As Variant
    Dim mBalanceAmt As Variant
    objCr.SetAccountCode Trim(txtAccountCode.Text)
    If objCr.AccountHeadID > 0 Then
        
        txtAccountHead.Text = objCr.AccountHead
        txtAccountHead.Tag = objCr.AccountHeadID
        txtAccountCode.Text = objCr.AccountCode
        objBk.SetBankInfoByAccID objCr.AccountHeadID
        If objBk.BankAccountHeadID > -1 Then
            txtNameOfBank.Text = objBk.BankName
            txtBranch.Text = objBk.Branch
            txtAccountNo.Text = objBk.AccountNumber
            'mChequeNo = objBk.GetNeWChequeNumber
            'txtRef.Text = IIf(IsNull(mChequeNo), "", mChequeNo)
        Else
            txtNameOfBank.Text = ""
            txtBranch.Text = ""
            txtAccountNo.Text = ""
            txtRef.Text = ""
        End If
        
        mBalanceAmt = objCr.GetLedgerBalance(objCr.AccountHeadID)
        If Not IsNull(mBalanceAmt) Then
            txtDr.Tag = mBalanceAmt
        End If
    Else
        txtAccountHead.Text = ""
        txtAccountCode.Text = ""
    End If
End Sub

Private Sub txtAccountHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call txtAccountCode_KeyDown(vbKeyF4, 0)
    End If
End Sub

Private Sub txtAccountHead_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PressTabKey
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub txtFieldCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub
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

Private Sub txtBudgetCentre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdBudgetCentres_Click
    End If
End Sub

Private Sub txtBudgetCentre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub
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

Private Sub txtField_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdFields_Click
    End If
End Sub

Private Sub txtField_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub

Private Sub txtDate_LostFocus()
    If txtDate.Text <> "" Then
        txtDate.Text = CheckDateInMMM(txtDate.Text)
    End If
End Sub

Private Sub txtPaymentOrderNo_GotFocus()
'
'    Dim mPayOrderID As Long
'    Dim mPayOrderNo As String
'
'    Dim mCnn As New ADODB.Connection
'    Dim objDb As New clsDB
'    Dim Rec As New ADODB.Recordset
'    Dim mSql As String
'    Dim objAc As New clsAccounts
'
'    objDb.SetConnection mCnn
'    mPayOrderID = Val(txtPaymentOrderNo.Tag)
'
'    mSql = "SELECT fapayorder.numpayorderid,fapayorder.numpayorderno,faFunds.vchFund,faFunds.intFundID,fapayorder.vchbillno,fapayorder.dtbilldate,fapayorder.fltbillamount,fapayorder.vchpayto,fapayorder.dtDuedate, "
'    mSql = mSql + "fafunctions.vchfunction,fafunctions.intfunctionid,fafunctionaries.intfunctionaryid,fafunctionaries.vchfunctionary,fatransactiontype.vchtransactiontype,fapayorder.intInstrumentTypeID,fainstrumenttypes.vchinstrumenttype,faaccountheads.vchaccountheadcode, faaccountheads.vchaccounthead, fapayorder.vchparticulars,fapayorder.intcashorbankheadid,faPayOrder.intTransactionTypeID "
'    mSql = mSql + "From fapayOrderChild "
'    mSql = mSql + "Inner join faPayorder on fapayorderchild.numpayorderid=fapayorder.numpayorderid "
'    mSql = mSql + "Inner join fatransactions on fapayorderchild.numpayorderid=fatransactions.intVoucherID "
'    mSql = mSql + "Inner join fafunds on fatransactions.intFundID=FaFunds.intFundID "
'    mSql = mSql + "inner join fafunctions on fapayorder.intfunctionid=fafunctions.intfunctionid "
'    mSql = mSql + "inner join fafunctionaries on fapayorder.intfunctionaryid=fafunctionaries.intfunctionaryid "
'    mSql = mSql + "inner join faaccountheads on fapayorder.intcashorbankheadid=faaccountheads.intaccountheadid "
'    mSql = mSql + "inner join fainstrumenttypes on fapayorder.intinstrumenttypeid=fainstrumenttypes.intinstrumenttypeid "
'    mSql = mSql + "inner join fatransactiontype on fatransactiontype.inttransactiontypeid=fapayorder.inttransactiontypeid "
'    mSql = mSql + "where fapayorder.numpayorderid=" & Val(mPayOrderID)
'    Rec.Open mSql, mCnn
'
'    If Not (Rec.BOF And Rec.EOF) Then
'        txtFund.Text = Rec!vchFund
'        txtFund.Tag = Rec!intFundID
'        txtFunction = Rec!vchFunction
'        txtFunction.Tag = Rec!intFunctionID
'        txtFunctionary = Rec!vchFunctionary
'        txtFunctionary.Tag = Rec!intFunctionaryID
'        cmbTransactionType.Text = Rec!vchTransactionType
'        cmbTransactionType.Tag = Rec!intTransactionTypeID
'        txtAccountCode.Tag = Rec!intcashorbankheadid
'         objAc.SetAccountID (Val(txtAccountCode.Tag))
'            If (objAc.AccountHeadID) <> -1 Then
'                txtAccountCode.Text = objAc.AccountCode
'                txtAccountHead.Text = objAc.AccountHead
'            End If
'         txtAccountCode.SetFocus
'        Call txtAccountCode_GotFocus
'
'        dtpIssueDate.Value = Rec!dtbilldate
'        txtSubsidiaryLedger.Text = IIf(IsNull(Rec!vchPayTo), "", Rec!vchPayTo)
'        cmbInstruments.Text = Rec!vchInstrumentType
'        dtpDueDate.Value = Rec!dtDueDate
'
'   End If
'     mCnn.Close
'
'    Dim mLoop As Integer
'    Dim mSerialNo As Integer
'    If Val(cmbTransactionType.Tag) = 1001 Then
'        mSql = "SELECT "
'        mSql = mSql + "faPayOrderChild.numPayOrderID,faPayOrder.vchParticulars,faPayOrderchild.intAccountHeadID,faAccountHeads.vchAccountHead,faPayOrderchild.vchAccountHeadCode,fltAmount from faPayOrderChild  "
'        mSql = mSql + "inner join faAccountHeads On faPayorderChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadCode "
'        mSql = mSql + "inner join faPayOrder On faPayOrderChild.numPayOrderID=faPayOrder.numPayOrderID "
'        mSql = mSql + "where faPayOrderchild.vchAccountHeadCode like '350110200' and faPayOrderChild.tindebitorcreditflag=0 and fapayorderchild.numpayorderid=" & Val(mPayOrderID)
'        objDb.SetConnection mCnn
'        Rec.Open mSql, mCnn
'        'mLoop = 0
'        mSerialNo = 1
'        vsGrid.Clear 1, 1
'        If Not (Rec.EOF And Rec.BOF) Then
'            'mLoop = mLoop + 1
'            vsGrid.Row = 1
'            vsGrid.Rows = vsGrid.Rows + 1
'            vsGrid.TextMatrix(1, 0) = mSerialNo
'            vsGrid.TextMatrix(1, 1) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'            vsGrid.TextMatrix(1, 2) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
'            vsGrid.TextMatrix(1, 3) = IIf(IsNull(Rec!vchParticulars), "", Rec!vchParticulars)
'            vsGrid.TextMatrix(1, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'            vsGrid.TextMatrix(1, 5) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'            txtDr = Rec!fltAmount
'            'Rec.MoveNext
'            mSerialNo = mSerialNo + 1
'        End If
'
'        Rec.Close
'        mCnn.Close
'    End If
End Sub

Private Sub txtFunction_GotFocus()
    If gbSearchStr <> "" Then
        txtFunction.Text = gbSearchStr
        txtFunction.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub

Private Sub txtFunction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdFunctions_Click
    End If
End Sub
Private Sub txtFunction_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PressTabKey
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtFunctionary_GotFocus()
    If gbSearchStr <> "" Then
        txtFunctionary.Text = gbSearchStr
        txtFunctionary.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub
Private Sub txtFunctionary_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdFunctionaries_Click
    End If
End Sub
Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PressTabKey
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtFund_GotFocus()
    If gbSearchStr <> "" Then
        txtFund.Text = gbSearchStr
        txtFund.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub
Private Sub txtFund_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call cmdFunds_Click
    End If
End Sub
Private Sub txtFund_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PressTabKey
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtNameOfBank_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub
Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub
Private Sub txtProject_GotFocus()
    txtProject.Text = gbSearchStr
    txtProject.Tag = gbSearchID
    
    gbSearchStr = ""
    gbSearchID = -1
End Sub
Private Sub txtRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then PressTabKey
End Sub
Private Sub txtSubsidiaryLedger_GotFocus()
    If gbSearchStr <> "" Then
        txtSubsidiaryLedger.Text = gbSearchStr
        txtSubsidiaryLedger.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        Call txtSubsidiaryLedger_LostFocus
    End If
End Sub
Private Sub txtSubsidiaryLedger_LostFocus()
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDb
    Dim mSQL As String
    If val(txtSubsidiaryLedger.Tag) > 0 Then
        mSQL = "Select * From faSubsidiaryAccounts Where intSubAccountID = " & val(txtSubsidiaryLedger.Tag)
        Set Rec = GetRecordSet(mSQL)
        If Not (Rec.EOF And Rec.BOF) Then
            mSQL = Rec!vchSubAccountHead & "  [ " & Rec!vchSubAccountCode & " ]" & vbCrLf
            mSQL = mSQL + Rec!vchAddress1 & vbCrLf
            If Len(Rec!vchAddress2) Then
                mSQL = mSQL + Rec!vchAddress2 + vbCrLf
            End If
            If Len(Rec!vchAddress3) Then
                mSQL = mSQL + Rec!vchAddress3
            End If
            txtClaiment.Text = mSQL
        End If
    Else
        txtSubsidiaryLedger.Text = ""
        txtClaiment.Text = ""
    End If
End Sub
Private Sub txtVoucherNo_LostFocus()
    If txtVoucherNo.Text <> "" Then
       Call DisplayVoucherDetails(txtVoucherNo.Text)
    End If
'    Dim mSQL As String
'    Dim mCount As Integer
'    Dim mIndex As Integer
'    Dim mCnn As New ADODB.Connection
'    Dim Rec As New ADODB.Recordset
'    Dim objDb As New clsDB
'    If Trim(txtVoucherNo.Text) <> "" And IsNumeric(txtVoucherNo.Text) Then
'        mSQL = "Select "
'        mSQL = mSQL + " faVouchers.dtDate, "
'        mSQL = mSQL + "faFunds.intFundId,"
'        mSQL = mSQL + "faFunds.vchFund,"
'        mSQL = mSQL + "faFunctions.intFunctionId,"
'        mSQL = mSQL + "faFunctions.vchFunction,"
'        mSQL = mSQL + "faFunctionaries.intFunctionaryId,"
'        mSQL = mSQL + "faFunctionaries.vchFunctionary,"
'        mSQL = mSQL + "faVouchers.intInstrumentTypeId,"
'        mSQL = mSQL + "isnull(faTransactionChild.intByAccountHeadId,-1)as intByAccountHeadId,"
'        mSQL = mSQL + "faAccountHeads.intAccountHeadId,"
'        mSQL = mSQL + "faAccountHeads.vchAccountHeadCode,"
'        mSQL = mSQL + "faAccountHeads.vchAccountHead,"
'        mSQL = mSQL + "faVouchers.vchInstrumentNo,"
'        mSQL = mSQL + "faBanks.intBankId,"
'        mSQL = mSQL + "faBanks.vchBankName,"
'        mSQL = mSQL + "faBanks.vchAccountNumber,"
'        mSQL = mSQL + "faBanks.vchBranch,"
'        mSQL = mSQL + "faSubsidiaryAccounts.intSubAccountId,"
'        mSQL = mSQL + "faSubsidiaryAccounts.vchSubAccountCode+vchSubAccountHead as SubsidiaryAccount,"
'        mSQL = mSQL + "faSubsidiaryAccounts.vchAddress1+'|'+faSubsidiaryAccounts.vchAddress2+'|'+faSubsidiaryAccounts.vchAddress3 as vchAddress,"
'        mSQL = mSQL + "faTransactionChild.vchNarration,"
'        mSQL = mSQL + "faTransactionChild.fltAmount, faVouchers.dtInstrumentDate, "
'        mSQL = mSQL + "faInstrumentTypes.vchInstrumentType "
'    mSQL = mSQL + "From "
'        mSQL = mSQL + "faVouchers "
'    mSQL = mSQL + "Inner Join faTransactions "
'        mSQL = mSQL + "On faTransactions.intVoucherId=faVouchers.intVoucherId "
'    mSQL = mSQL + "Inner Join faTransactionChild "
'        mSQL = mSQL + "On faTransactions.intTransactionId=faTransactionChild.intTransactionId "
'    mSQL = mSQL + "Inner Join faFunctions "
'        mSQL = mSQL + "On fatransactions.intFunctionId=faFunctions.intFunctionId "
'    mSQL = mSQL + "Inner Join faFunctionaries "
'        mSQL = mSQL + "On faTransactions.intFunctionaryId=faFunctionaries.intFunctionaryId "
'    mSQL = mSQL + "Inner Join faFunds "
'        mSQL = mSQL + "On faFunds.intFundId=faTransactions.intFundId "
'    mSQL = mSQL + "Inner Join faAccountHeads "
'        mSQL = mSQL + "On faAccountHeads.intAccountHeadId=faTransactionChild.intAccountHeadId "
'    mSQL = mSQL + "Left Outer Join faInstrumentTypes "
'    mSQL = mSQL + "On faInstrumentTypes.intInstrumentTypeID=faVouchers.intInstrumentTypeID "
'    mSQL = mSQL + "Left Outer Join faSubsidiaryAccounts "
'        mSQL = mSQL + "On faSubsidiaryAccounts.intSubAccountId=fatransactions.NumSubLedgerId "
'    mSQL = mSQL + "Left Outer Join faBanks "
'        mSQL = mSQL + "On faTransactionChild.intAccountHeadId=faBanks.intAccountHeadId "
'    mSQL = mSQL + "Where "
'        mSQL = mSQL + " faVouchers.intVoucherNo= " & Trim(txtVoucherNo.Text)
'        mSQL = mSQL + " And faTransactions.intGroupId=20"
'            objDb.SetConnection mCnn
'            Rec.Open mSQL, mCnn
'            mCount = 1
'            If Rec.EOF And Rec.BOF Then
'                FormInitialize
'                cmdSave.Caption = "&Save"
'            Else
'                While Not Rec.EOF
'                    If Rec!intByAccountHeadID = -1 Then
'                        On Error Resume Next
'                        txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
'                        txtFund.Tag = IIf(IsNull(Rec!intFundID), "", Rec!intFundID)
'                        txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
'                        txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
'                        txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
'                        txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
'                        txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
'                        txtAccountCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'                        txtAccountCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'                        txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
'                        txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
'                        txtNameOfBank.Tag = IIf(IsNull(Rec!intBankID), "", Rec!intBankID)
'                        txtNameOfBank.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
'                        txtAccountNo.Text = IIf(IsNull(Rec!vchAccountNumber), "", Rec!vchAccountNumber)
'                        txtBranch.Text = IIf(IsNull(Rec!vchBranch), "", Rec!vchBranch)
'                        txtSubsidiaryLedger.Tag = IIf(IsNull(Rec!intSubAccountID), "", Rec!intSubAccountID)
'                        txtSubsidiaryLedger.Text = IIf(IsNull(Rec!SubsidiaryAccount), "", Rec!SubsidiaryAccount)
'                        txtClaiment.Text = IIf(IsNull(Rec!vchAddress), "", Rec!vchAddress)
'                        txtNarration.Text = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
'                        dtpDueDate.Value = Rec!dtInstrumentDate
'                        cmbInstruments.Tag = Rec!intInstrumentTypeID
'                        'Error
'                        'mIndex = SendMyMessage(cmbInstruments.hwnd, CB_FINDSTRING, -1, ByVal Rec!vchInstrumentType)
'                        'cmbInstruments.ListIndex = mIndex
'                        cmbInstruments.Text = Rec!vchInstrumentType
'                        'MsgBox cmbInstruments.ListIndex
'                    Else
'                        vsGrid.Rows = vsGrid.Rows + 1
'                        vsGrid.TextMatrix(mCount, 0) = mCount
'                        vsGrid.TextMatrix(mCount, 1) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'                        vsGrid.TextMatrix(mCount, 2) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
'                        vsGrid.TextMatrix(mCount, 3) = IIf(IsNull(Rec!vchNarration), "", Rec!vchNarration)
'                        vsGrid.TextMatrix(mCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
'                        mCount = mCount + 1
'                    End If
'                    Rec.MoveNext
'                Wend
'                cmbTransactionType.ListIndex = 1
'                cmdSave.Caption = "&Edit"
'                Call Calculate
'            End If
'    End If
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsGrid.Col = 1 Then
        If val(vsGrid.TextMatrix(vsGrid.Row, 5)) = 0 Then
            vsGrid.TextMatrix(vsGrid.Row, 1) = ""
        End If
    End If
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim objBjd As New clsBudgetCentre
    If vsGrid.Row > 1 Then
        If vsGrid.TextMatrix(vsGrid.Row - 1, 1) = "" Or _
           vsGrid.TextMatrix(vsGrid.Row - 1, 2) = "" Or _
           val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 Then

           Cancel = True
           Exit Sub
        End If
    End If
    
    If Col = 2 Then
        Cancel = True
    End If
    
    If Col = 4 Then
        If val(vsGrid.TextMatrix(vsGrid.Row, 5)) = 0 Then
            Cancel = True
        End If
    End If
    
    'MsgBox gbSearchStr & vsGrid.Row & "  " & vsGrid.Col
    If Len(gbSearchStr) Then
        vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        vsGrid.TextMatrix(vsGrid.Row, 5) = gbSearchID
        vsGrid.Col = vsGrid.Col + 2
        lbBudgetAllocated.Caption = Format(objBjd.GetBudgetAmount(IIf(txtFunction.Tag = "", 0, txtFunction.Tag), IIf(txtFunctionary.Tag = "", 0, txtFunctionary.Tag), vsGrid.TextMatrix(vsGrid.Row, 5)), "0.00")
        vsGrid.Redraw = flexRDDirect
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub
Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If cmbInstruments.ListIndex > -1 Then
        If cmbInstruments.ItemData(cmbInstruments.ListIndex) = 1 Then
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where intGroupID  IN (2) And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
        Else
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (2,3,4) AND tinHiddenFlag <> 1 Order by vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
        End If
    Else
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where  tinType IN (2,3,4) AND tinHiddenFlag <> 1 Order by vchAccountHeadCode"
        frmSearchAccountHeads.Show vbModal
    End If
End Sub
Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        If vsGrid.Col = 2 And Trim(vsGrid.Text) <> "" Then
            Set objAccHead = New clsAccounts
            If objAccHead.FindAccountByHead(Trim(vsGrid.Text)) Then
                vsGrid.TextMatrix(vsGrid.Row, 1) = objAccHead.AccountCode
            End If
        ElseIf vsGrid.Col = 4 Then
            vsGrid.TextMatrix(vsGrid.Row, 4) = Format(val(vsGrid.TextMatrix(vsGrid.Row, 4)), "0.00")
            Call Calculate
        End If
End Sub
Private Sub vsGrid_Validate(Cancel As Boolean)
        If vsGrid.Col = 3 Then
            If Len(vsGrid.TextMatrix(vsGrid.Row, 3)) > 100 Then
                vsGrid.TextMatrix(vsGrid.Row, 3) = Left(vsGrid.TextMatrix(vsGrid.Row, 3), 100)
            End If
        End If
End Sub


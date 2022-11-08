VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPendingTasks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PendingTasks"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15420
   Icon            =   "frmPendingTasks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicCaption 
      BackColor       =   &H80000009&
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   15360
      TabIndex        =   53
      Top             =   0
      Width           =   15420
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previoud Year Pending Tasks Request"
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
         Left            =   180
         TabIndex        =   54
         Top             =   45
         Width           =   15000
      End
   End
   Begin VB.Frame fraSave 
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      TabIndex        =   46
      Top             =   6210
      Width           =   6045
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   405
         Left            =   2910
         TabIndex        =   59
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdApprove 
         Caption         =   "Approve"
         Height          =   405
         Left            =   2910
         TabIndex        =   58
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   4020
         TabIndex        =   41
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   405
         Left            =   1815
         TabIndex        =   40
         Top             =   150
         Width           =   915
      End
   End
   Begin VB.CheckBox chkNonPlan 
      Caption         =   "Non-Plan"
      Height          =   285
      Left            =   1125
      TabIndex        =   45
      Top             =   7155
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraMain 
      Height          =   5955
      Left            =   0
      TabIndex        =   44
      Top             =   315
      Width           =   6045
      Begin VB.CommandButton cmdInstNo 
         Caption         =   "..."
         Height          =   330
         Left            =   3720
         TabIndex        =   14
         Top             =   1905
         Width           =   255
      End
      Begin VB.CommandButton cmdKeyID 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   20
         Top             =   2655
         Width           =   285
      End
      Begin VB.TextBox txtKeyID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   19
         Top             =   2655
         Width           =   3750
      End
      Begin VB.CommandButton cmdSourceFund 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   26
         Top             =   3405
         Width           =   285
      End
      Begin VB.TextBox txtSourceFund 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   25
         Top             =   3405
         Width           =   3750
      End
      Begin VB.CommandButton cmdProject 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   32
         Top             =   4155
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtProject 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   31
         Top             =   4155
         Visible         =   0   'False
         Width           =   3750
      End
      Begin VB.CommandButton cmdCategory 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   29
         Top             =   3780
         Width           =   285
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   28
         Top             =   3780
         Width           =   3750
      End
      Begin VB.CommandButton cmdImpo 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   35
         Top             =   4530
         Width           =   285
      End
      Begin VB.TextBox txtImpo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   34
         Top             =   4530
         Width           =   3750
      End
      Begin VB.TextBox txtInstrumentDate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   16
         Top             =   2280
         Width           =   1875
      End
      Begin VB.TextBox txtInstrumentNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   13
         Top             =   1905
         Width           =   1875
      End
      Begin VB.CommandButton cmdInstrumentType 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   11
         Top             =   1500
         Width           =   285
      End
      Begin VB.TextBox txtInstrumentType 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   10
         Top             =   1500
         Width           =   3750
      End
      Begin VB.CommandButton cmdExpdHead 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   23
         Top             =   3030
         Width           =   285
      End
      Begin VB.TextBox txtExpdHead 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   22
         Top             =   3030
         Width           =   3750
      End
      Begin VB.CommandButton cmdTransactionType 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         TabIndex        =   8
         Top             =   1125
         Width           =   285
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   7
         Top             =   1125
         Width           =   3750
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1815
         TabIndex        =   37
         Top             =   5010
         Width           =   1770
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1815
         TabIndex        =   39
         Top             =   5370
         Width           =   3750
      End
      Begin VB.CommandButton cmdTasks 
         Caption         =   "..."
         Height          =   315
         Left            =   5580
         TabIndex        =   2
         Top             =   375
         Width           =   285
      End
      Begin VB.TextBox txtPendingTask 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   375
         Width           =   3750
      End
      Begin VB.TextBox txtTrnDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1815
         TabIndex        =   4
         Top             =   735
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker dtTrnDate 
         Height          =   330
         Left            =   3705
         TabIndex        =   5
         Top             =   735
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   43555
      End
      Begin MSComCtl2.DTPicker dtInstrumentDate 
         Height          =   345
         Left            =   3645
         TabIndex        =   17
         Top             =   2280
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   609
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   41357
      End
      Begin VB.Label lblKeyName 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   18
         Top             =   2715
         Width           =   1485
      End
      Begin VB.Label lblFund 
         Alignment       =   1  'Right Justify
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   225
         TabIndex        =   24
         Top             =   3480
         Width           =   1560
      End
      Begin VB.Label lblProject 
         Alignment       =   1  'Right Justify
         Caption         =   "Project"
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
         Left            =   255
         TabIndex        =   30
         Top             =   4215
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblCategory 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   27
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label lblIMPO 
         Alignment       =   1  'Right Justify
         Caption         =   "IMPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   255
         TabIndex        =   33
         Top             =   4590
         Width           =   1530
      End
      Begin VB.Label lblInstDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Inst Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   2325
         Width           =   1485
      End
      Begin VB.Label lblInstNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Inst No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   12
         Top             =   1965
         Width           =   1230
      End
      Begin VB.Label lblInstType 
         Alignment       =   1  'Right Justify
         Caption         =   "Instrument Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   285
         TabIndex        =   9
         Top             =   1545
         Width           =   1500
      End
      Begin VB.Label lblExpHead 
         Alignment       =   1  'Right Justify
         Caption         =   "Expenditure Head"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   3090
         Width           =   1575
      End
      Begin VB.Label lblTranType 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   6
         Top             =   1185
         Width           =   1515
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   570
         TabIndex        =   36
         Top             =   5070
         Width           =   1215
      End
      Begin VB.Label lblRemarks 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   555
         TabIndex        =   38
         Top             =   5370
         Width           =   1230
      End
      Begin VB.Label lblTransactionDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   3
         Top             =   765
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pending Task"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   285
         TabIndex        =   0
         Top             =   405
         Width           =   1500
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   510
      Left            =   6090
      TabIndex        =   43
      Top             =   7035
      Width           =   9240
      Begin VB.ComboBox cmbSearchTask 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   135
         Width           =   2730
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   615
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   150
         Width           =   1830
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         Height          =   270
         Left            =   90
         TabIndex        =   56
         Top             =   195
         Width           =   630
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   6675
      Left            =   6090
      TabIndex        =   42
      Top             =   360
      Width           =   9255
      _cx             =   16325
      _cy             =   11774
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      BackColorBkg    =   16777215
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPendingTasks.frx":1CCA
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
      Begin VSFlex8LCtl.VSFlexGrid vsTasks 
         Height          =   5580
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         _cx             =   6588
         _cy             =   9842
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
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
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
         Rows            =   18
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPendingTasks.frx":1E19
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   10
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
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   14985
      Top             =   7875
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label8 
      Caption         =   "Finished"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4995
      TabIndex        =   52
      Top             =   7260
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Cancelled"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4995
      TabIndex        =   51
      Top             =   7455
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Approved"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4995
      TabIndex        =   50
      Top             =   7050
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FAEFEE&
      Height          =   195
      Left            =   4770
      TabIndex        =   49
      Top             =   7245
      Width           =   195
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E7F9E7&
      Height          =   195
      Left            =   4770
      TabIndex        =   48
      Top             =   7035
      Width           =   195
   End
   Begin VB.Label Label2 
      BackColor       =   &H0090AAFF&
      Height          =   180
      Left            =   4770
      TabIndex        =   47
      Top             =   7470
      Width           =   195
   End
End
Attribute VB_Name = "frmPendingTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mCurFinancialYear   As Integer
    Dim mPreFinancialYear   As Integer
    Dim mPreStartDate       As Date
    Dim mPreEndDate         As Date
    Dim mID                 As Integer
    Dim mValidateTrnDate    As Date
    Dim mCancelFlag     As Boolean

    Public Sub SetInputField(TaskID As Integer)
    
        Dim mExtractedStatus As Integer
        Dim mMsg As String
        
        mMsg = ""
        mMsg = mMsg + "Previous year's Source wise transactions are all closed by Secretary" & vbCrLf
        mMsg = mMsg + "by brought down Source wise balances to new financial year by declaring the Source wise balances are correct." & vbCrLf
        mMsg = mMsg + "" & vbCrLf
        mMsg = mMsg + "Further changes in previous year's source wise transaction will" & vbCrLf
        mMsg = mMsg + "make difference in Current year's Source wise allocations, thus this functionality is no more permitted in Previous year transactions." & vbCrLf

    
        cmdProject.Enabled = True
        txtAmount.Enabled = True
        txtTrnDate.Enabled = True
        dtTrnDate.Enabled = True
    
        fraMain.Enabled = True
        lblTransactionDate.Caption = "Transaction Date"
        lblTranType.Visible = False
        
        txtTransactionType.Visible = False
        cmdTransactionType.Visible = False
        
        lblInstType.Visible = False
        txtInstrumentType.Visible = False
        cmdInstrumentType.Visible = False
        
        lblInstNo.Visible = False
        txtInstrumentNo.Visible = False
    
        lblInstDate.Visible = False
        txtInstrumentDate.Visible = False
        dtInstrumentDate.Visible = False
        
        lblKeyName.Visible = False
        txtKeyID.Visible = False
        cmdKeyID.Visible = False
        
        lblExpHead.Visible = False
        txtExpdHead.Visible = False
        cmdExpdHead.Visible = False
        
        lblFund.Visible = False
        txtSourceFund.Visible = False
        cmdSourceFund.Visible = False
        
        lblCategory.Visible = False
        lblCategory.Caption = "Category"
        txtCategory.Visible = False
        cmdCategory.Visible = False
        
        lblProject.Visible = False
        txtProject.Visible = False
        cmdProject.Visible = False
        
        lblIMPO.Visible = False
        txtImpo.Visible = False
        cmdImpo.Visible = False
        cmdInstNo.Visible = False
        
        Select Case TaskID
            Case 1 ' Letter of Authority
                
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
                lblTranType.Visible = True
                txtTransactionType.Visible = True
                cmdTransactionType.Visible = True
                
                lblInstNo.Caption = "Allotment NO:"
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                
                'lblInstDate.Visible = True
                'txtInstrumentDate.Visible = True
                'dtInstrumentDate.Visible = True
            
            Case 2 'Cancel Letter of Authority
            
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
                lblInstNo.Caption = "Allotment NO:"
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                cmdInstNo.Visible = True
                
            Case 3 ' Requistion
                
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
                
                lblTransactionDate.Caption = "Requisition Date"
                lblFund.Visible = True
                
                txtInstrumentNo.Visible = True
                cmdInstNo.Visible = True
                
                txtSourceFund.Visible = True
                cmdSourceFund.Visible = True
                cmdSourceFund.Enabled = False
                
                lblCategory.Visible = True
                txtCategory.Visible = True
                cmdCategory.Visible = True
                cmdCategory.Enabled = False
                    
                lblProject.Visible = True
                txtProject.Visible = True
                cmdProject.Visible = True
                
            Case 4 ' Allotment B Fund
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
            Case 5 ' Demand/Receipts
                lblInstType.Visible = True
                txtInstrumentType.Visible = True
                
                lblTranType.Visible = True
                txtTransactionType.Visible = True
                cmdTransactionType.Visible = True
                cmdInstrumentType.Visible = True
                
'                lblInstNo.Visible = True
'                txtInstrumentNo.Visible = True
'                'cmdInstNo.Visible = True
'
'                lblInstDate.Visible = True
'                txtInstrumentDate.Visible = True
                
                'lblKeyName.Visible = True
                'txtKeyID.Visible = True
                'cmdKeyID.Visible = True
                
            Case 6 ' Reverse Entry
                lblInstType.Visible = True
                txtInstrumentType.Visible = True
                
                lblTranType.Visible = True
                txtTransactionType.Visible = True
                
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                'cmdInstNo.Visible = True
                
                lblInstDate.Visible = True
                txtInstrumentDate.Visible = True
                
                lblKeyName.Visible = True
                txtKeyID.Visible = True
                
                cmdKeyID.Visible = True
                
            Case 7 ' PayOrder
                lblTranType.Visible = True
                txtTransactionType.Visible = True
                cmdTransactionType.Visible = True
                lblExpHead.Visible = True
                txtExpdHead.Visible = True
                cmdExpdHead.Visible = True
                
            Case 8 ' PayOrder Cancellation
                lblInstNo.Caption = "PayOrder NO:"
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                cmdInstNo.Visible = True
            Case 9 ' Interrupted Receipt
            
            Case 10 'Cancel Requisition
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
                lblInstNo.Caption = "RequisitionNo:"
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                cmdInstNo.Visible = True
                
             Case 11 'Pay Order Approval
                lblInstNo.Caption = "PayOrder NO:"
                lblInstNo.Visible = True
                txtInstrumentNo.Visible = True
                cmdInstNo.Visible = True
                lblTranType.Visible = True
                txtTransactionType.Visible = True
                txtTransactionType.Locked = True
            Case 12 'OBRP
            Case 13 'Requisition BFund
                
                mExtractedStatus = GetStatusFlag
                If mExtractedStatus = 2 Then
                    MsgBox mMsg, vbInformation
                    txtPendingTask.Tag = ""
                    txtPendingTask.Text = ""
                    Exit Sub
                End If
                
                
                lblTransactionDate.Caption = "Requisition Date"
                lblFund.Visible = True
                txtSourceFund.Visible = True
                cmdSourceFund.Visible = True
                cmdSourceFund.Enabled = True
                
                lblCategory.Visible = True
                lblCategory.Caption = "Scheme"
                txtCategory.Visible = True
                cmdCategory.Visible = True
                cmdCategory.Enabled = True
'                lblProject.Visible = True
'                txtProject.Visible = True
'                cmdProject.Visible = True
            Case 14 ' CONTRA ENTRY
            
            Case 15 ' JOURNAL ENTRY
            
            Case 16 ' UNAUTHORIZED DRAWAL
            
                lblTransactionDate.Caption = "Requisition Date"
                lblFund.Visible = True
                txtSourceFund.Visible = True
                cmdSourceFund.Visible = True
                cmdSourceFund.Enabled = True
                
                lblCategory.Visible = True
                txtCategory.Visible = True
                cmdCategory.Visible = True
                cmdCategory.Enabled = True
            Case 16 ' E bill voucher
            
        End Select
        fraMain.Refresh
    End Sub
        
    Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        Dim mTrAccHeadId As Integer
        
        If objdb.SetConnection(mCnn) Then
            mSql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                GetStatusFlag = -1
            End If
            Rec.Close
        End If
    End Function

Private Sub cmbMonth_Click()
    Call FillGrid
End Sub

Private Sub cmbSearchTask_Click()
    Call FillGrid
End Sub

    Private Sub cmdApprove_Click()
        Dim mSql As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        If val(cmdTasks.Tag) = 0 Then
            cmdApprove.Enabled = False
            Exit Sub
        End If
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Update faPendingTaskRequest set tnyStatus=2,numApprovedUser=" & gbUserID & ",dtApprovedDate='" & DdMmmYy(gbTransactionDate) & "' Where intRequestId=" & cmdTasks.Tag
            mCnn.Execute mSql
            mCnn.Close
        End If
        cmdApprove.Enabled = False
        Call FillGrid
    End Sub

    Private Sub cmdCancel_Click()
        Dim mSql As String
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            If val(txtPendingTask.Tag) > 0 Then
                If mCancelFlag = False Then
                    If MsgBox("Do you Want to Cancel this Request ..", vbYesNo, "Saankhya") = vbYes Then
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            mSql = "Update faPendingTaskRequest set tnyStatus=4,numApprovedUser=" & gbUserID & ",dtApprovedDate='" & gbTransactionDate & "' Where intRequestId=" & cmdTasks.Tag
                        Else
                            mSql = "Update faPendingTaskRequest set tnyStatus=4 Where intRequestId=" & cmdTasks.Tag
                        End If
                        mCnn.Execute mSql
                    End If
                Else
                    If MsgBox("Do you Want to Undo this Request ..", vbYesNo, "Saankhya") = vbYes Then
                        mSql = "Update faPendingTaskRequest set tnyStatus=2 Where intRequestId=" & cmdTasks.Tag
                        mCnn.Execute mSql
                    End If
                End If
            End If
            mCnn.Close
        End If
        Call FillGrid
    End Sub

    Private Sub cmdCategory_Click()
        gbSearchID = -1
        gbSearchStr = ""
        Select Case val(txtPendingTask.Tag)
        Case Is = 13
            frmSearchMasters.QrySP = StoredProcedure
            frmSearchMasters.SQLQry = "spSelectDepSchemePro"
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            If gbSearchStr <> "" Then
                txtCategory.Text = gbSearchStr
                txtCategory.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
            txtCategory.SetFocus
        Case Is = 16
            If val(txtSourceFund.Tag) = 1 Then
                frmSearchMasters.QrySP = Qyery
                frmSearchMasters.SQLQry = "SELECT intCategoryID,vchTransactionCategory FROM faTransactionCategory"
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.Show vbModal
                If gbSearchStr <> "" Then
                    txtCategory.Text = gbSearchStr
                    txtCategory.Tag = gbSearchID
                End If
                gbSearchStr = ""
                gbSearchID = -1
                txtCategory.SetFocus
            Else
                txtCategory.Text = "General"
                txtCategory.Tag = 1
                cmdCategory.Enabled = False
            End If
        End Select
        
    End Sub
    
    Private Sub cmdExpdHead_Click()
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
            If val(txtTransactionType.Tag) < 1 Then
                MsgBox "Please Select TransactionType", vbApplicationModal
                Exit Sub
            Else
                If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = True Then
                
                    mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
                    mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                    mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag)
                    mSql = mSql + " And faTransactionTypeChild.tinDebitOrCredit = 1 And faAccountHeads.tinHiddenFlag = 0"
                    mSql = mSql + " And faTransactionTypeChild.tnyListID = 1 And faAccountHeads.intGroupID is Null"
                    mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode"
                    'mSQL = mSQL + " Order By faTransactionTypeChild.intOrder"
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Rec.BOF Or Rec.EOF Then
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where intGroupID is Null And tinHiddenFlag = 0" ' Where tinType IN (3)"
                    End If
                    frmSearchAccountHeads.SQLString = mSql
                    frmSearchAccountHeads.VoucherMode = 400
                    frmSearchAccountHeads.Show vbModal
                    If gbSearchStr <> "" Then
                        txtExpdHead.Text = gbSearchStr
                        txtExpdHead.Tag = gbSearchID
                        gbSearchID = -1
                        gbSearchStr = ""
                    End If
                End If
            End If
    End Sub

    Private Sub cmdIMPO_Click()
        gbSearchID = -1
        gbSearchStr = ""
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intFunctionaryID,vchFunctionary From faFunctionaries Where vchFunctionaryCode >= 310000 Order by vchFunctionary"
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtImpo.Text = gbSearchStr
            txtImpo.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdInstNo_Click()
        Dim objAllotment As New clsAllotmentLetter
        Dim objAcc      As New clsAccounts
        
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mRow As Integer
        Dim mSql As String
        
        Select Case val(txtPendingTask.Tag)
            Case 2 ' Cancel Letter of Authority
                gbSearchStr = ""
                gbSearchID = -1
                frmSearchLetterOfAuthority.Show vbModal
                
                mSql = "Select * from faAllotmentLetters WHERE intAllotmentID = " & gbSearchID
                objdb.SetConnection mCnn
                Rec.Open mSql, mCnn
                If Not (Rec.BOF And Rec.EOF) Then
                    txtAmount.Text = Rec!fltAmount
                    txtTrnDate.Text = Rec!dtAllotmentDate
                    txtInstrumentNo.Text = gbSearchStr
                    txtInstrumentNo.Tag = gbSearchID
                Else
                    txtAmount.Text = ""
                    txtTrnDate.Text = ""
                    txtInstrumentNo.Text = ""
                    txtInstrumentNo.Tag = ""
                End If
                Rec.Close
                gbSearchStr = ""
                gbSearchID = -1
                
            Case 3 ' Requistion
                gbSearchStr = ""
                gbSearchID = -1
                frmSearchRequesition.PreviousYearTaskID = 3
                frmSearchRequesition.PreviousYearMode = 1
                frmSearchRequesition.Show vbModal
                
                objdb.SetConnection mCnn
                
                mSql = " SELECT intID,vchRequisitionNo,dtRequisitionDate,faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID,"
                mSql = mSql + " faSubSidiaryAccountHeads.vchName,suSourceOfFund.intSourceFundID,suSourceOfFund.vchSourceFundName,"
                mSql = mSql + " faTransactionCategory.intCategoryID,faTransactionCategory.vchTransactionCategory, "
                mSql = mSql + " vchProjectNO, numProjectID ,fltRequestedAmt"
                mSql = mSql + " FROM faAllotments "
                mSql = mSql + " INNER JOIN faSubSidiaryAccountHeads on faSubSidiaryAccountHeads.intSubsidiaryAccountHeadID=faAllotments.intImplementingOfficersID"
                mSql = mSql + " INNER JOIN suSourceOfFund on suSourceOfFund.intSourceFundID=faAllotments.intSourceID"
                mSql = mSql + " LEFT JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faAllotments.intFundCategoryID"
                mSql = mSql + " WHERE vchRequisitionNo = '" & gbSearchStr & "'"
                'mSql = mSql + " And tnyStatus = 1 "
                Rec.Open mSql, mCnn
                If Not (Rec.BOF And Rec.EOF) Then
                    txtProject.Text = Rec!vchProjectNo
                    txtProject.Tag = Rec!numProjectID
                    
                    txtSourceFund.Text = Rec!vchSourceFundName
                    txtSourceFund.Tag = Rec!intSourceFundID
                    
                    txtCategory.Text = Rec!vchTransactionCategory
                    txtCategory.Tag = Rec!intCategoryID
                    
                    txtAmount.Text = Rec!fltRequestedAmt
                    txtTrnDate.Text = DdMmmYy(Rec!dtRequisitionDate)
                    cmdProject.Enabled = False
                    txtAmount.Enabled = False
                    txtTrnDate.Enabled = False
                    dtTrnDate.Enabled = False
                End If
                Rec.Close
                
                txtInstrumentNo.Text = gbSearchStr
                txtInstrumentNo.Tag = gbSearchID
                
            Case Is = 11 ' Pay Order Approval [Previous Year]
                '                gbSearchCode = ""
                '                gbSearchID = -1
                '                frmSearchPaymentOrder.PendingTask = 1
                '                frmSearchPaymentOrder.chkListToApprove.value = 0
                '                frmSearchPaymentOrder.txtDateFrom = "1-Mar-" & gbFinancialYearID
                '                frmSearchPaymentOrder.txtDateTo = "31-Mar-" & gbFinancialYearID
                '                frmSearchPaymentOrder.Staus = 0
                '                frmSearchPaymentOrder.Show vbModal
                '                frmSearchPaymentOrder.ZOrder (0)
                '                txtInstrumentNo.Text = gbSearchCode
                '                txtInstrumentNo.Tag = gbSearchID

                gbSearchStr = ""
                gbSearchID = -1
                frmSearchPaymentOrder.PendingTask = 1
                frmSearchPaymentOrder.Show vbModal
                txtInstrumentNo.Text = gbSearchStr
                txtInstrumentNo.Tag = gbSearchID
                Call GetPODetails(gbSearchID)
            Case Is = 7 'Pay Order
                frmListOfAllotmentLetters.PreviousYearMode = 1
                If val(txtTransactionType.Tag) = 1201 Or val(txtTransactionType.Tag) = 1391 Then
                    frmListOfAllotmentLetters.UnAuthorizedDrawal = 1
                End If
                frmListOfAllotmentLetters.Show vbModal
                If gbSearchID <> -1 Then
                    txtInstrumentNo.Text = gbSearchCode
                    txtInstrumentNo.Tag = gbSearchID
                    
                    objAllotment.SetAllotment (txtInstrumentNo.Tag)
                    
                    txtSourceFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
                    txtSourceFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), "", objAllotment.SourceOfFundID)
                    txtImpo.Text = IIf(IsNull(objAllotment.ImplementingOfficer), "", objAllotment.ImplementingOfficer)
                    txtImpo.Tag = IIf(IsNull(objAllotment.ImplementingOfficersID), "", objAllotment.ImplementingOfficersID)
                    txtAmount.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)
                    txtProject.Tag = IIf(IsNull(objAllotment.ProjectID), "", objAllotment.ProjectID)
                    lblIMPO.Visible = True
                    txtImpo.Visible = True
                    lblFund.Visible = True
                    txtExpdHead.Tag = objAllotment.GrossAccountHeadID
                    objAcc.SetAccountID (val(txtExpdHead.Tag))
                    txtExpdHead.Text = objAcc.AccountCode + objAcc.AccountHead
                    txtSourceFund.Visible = True
                    Dim objProject As New clsProject
                    
                    objProject.SetProject val(txtProject.Tag), gbFinancialYearID - 1
                    If Not IsNull(objProject.ProjectID) Then
                        txtProject.Text = IIf(IsNull(objProject.ProjectSerialNo), "", objProject.ProjectSerialNo)
                        'txtProject.Tag = IIf(IsNull(objProject.ProjectID), "", objProject.ProjectID)
                        txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
                        txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
                        txtProject.Visible = True
                        lblProject.Visible = True
                        txtCategory.Visible = True
                        lblCategory.Visible = True
                    End If
                    gbSearchID = -1
                    gbSearchStr = ""
                    gbSearchCode = ""
                    
                End If
            Case Is = 8 'Pay Order Cancel
                frmSearchPaymentOrder.Staus = 1
                frmSearchPaymentOrder.PendingTask = 8
                frmSearchPaymentOrder.chkListToApprove.Visible = True
                frmSearchPaymentOrder.Show vbModal
                If gbSearchID > 0 Then
                    txtInstrumentNo.Tag = gbSearchID
                    txtInstrumentNo.Text = gbSearchStr
                    Call GetPODetails(gbSearchID)
                    gbSearchID = -1
                    gbSearchStr = ""
                    If IsDate(mValidateTrnDate) Then
                        If IsDate(txtTrnDate.Text) Then
                            If CDate(txtTrnDate) < mValidateTrnDate Then
                                MsgBox "Cancellation Date Can not be less then the Pay Order Date", vbInformation
                                txtTrnDate.Text = ""
                                txtTrnDate.SetFocus
                            End If
                        Else
                            txtTrnDate.Text = ""
                        End If
                    Else
                        txtInstrumentNo.Text = ""
                        txtInstrumentNo.Tag = ""
                    End If
                    If Trim(txtInstrumentNo) <> "" Then
                    objdb.SetConnection mCnn
                    mSql = "SELECT * FROM faPendingTaskRequest WHERE intTaskID = 8 And isNull(tnystatus,0)<>4 AND vchInstrumentNo = " & Trim(txtInstrumentNo)
                    Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
                    If Not (Rec.BOF And Rec.EOF) Then
                        MsgBox "This Pay Order is already selected in Pay Order Cancel Request in Pending Task", vbInformation
                        txtInstrumentNo.Text = ""
                    End If
                    Rec.Close
                    End If
                    
                End If
                
            Case Is = 10 'Cancel Requisitions
                frmSearchRequesition.PreviousYearTaskID = 10
                frmSearchRequesition.PreviousYearMode = 1
                frmSearchRequesition.Show vbModal
                txtInstrumentNo.Text = gbSearchStr
                txtInstrumentNo.Tag = gbSearchID
                objAllotment.SetAllotment (txtInstrumentNo.Tag)
                'txtAmount.Text = IIf(IsNull(objAllotment.amount), "", objAllotment.amount)
                txtAmount.Text = IIf(IsNull(objAllotment.RequisitionAmount), "", objAllotment.RequisitionAmount)
                txtInstrumentNo.SetFocus
        End Select
    End Sub
    
    Private Sub GetPODetails(mPoID)
        Dim mSql As String
        Dim Rec    As New ADODB.Recordset
        Dim mRec    As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSeatID As Double
        Dim mSeatName As String
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            
            mSql = "Select * From faPendingTaskRequest Where intTaskID= " & val(txtPendingTask.Tag) & "  And tnyStatus<>4 And numDemandID=" & mPoID
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If (Rec.EOF Or Rec.BOF) Then
                mSql = "Select *,isNull(faPayOrder.vchDescription,' ') Descrip,dbo.fnGetSeat(numSeatID) seatName From faPayOrder "
                mSql = mSql + " Inner Join faPayOrderChild on faPayOrderChild.intpayOrderID=faPayOrder.intpayOrderID"
                mSql = mSql + " Inner Join faTransactionType On faPayOrder.intTransactionTypeID=faTransactionType.intTransactionTypeID"
                mSql = mSql + " Where faPayOrderChild.intSlNo=1 And faPayOrder.intPayOrderID=" & mPoID
                Set mRec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (mRec.EOF Or mRec.BOF) Then
                    mValidateTrnDate = mRec!dtPayOrderDate
                    txtTransactionType.Text = mRec!vchTransactionType
                    txtTransactionType.Tag = mRec!intTransactionTypeID
                    txtAmount.Text = mRec!numAmount
                    txtAmount.Enabled = False
                    txtRemarks.Text = mRec!Descrip
                    mSeatID = mRec!numSeatID
                    mSeatName = mRec!SeatName
                End If
                If mSeatID <> gbSeatID Then
                    MsgBox "The Pay Order is Generated at Seat " + CStr(mSeatName) + ".. Please Request through Generated seat", vbInformation
                    txtInstrumentNo.Text = ""
                    txtInstrumentNo.Tag = -1
                End If
            Else
                MsgBox "This Pay Order Already Selected", vbInformation
                txtInstrumentNo.Text = ""
                txtInstrumentNo.Tag = -1
            End If
        End If
    End Sub
    Private Sub cmdInstrumentType_Click()
        Dim mSql    As String
        gbSearchID = -1
        gbSearchStr = ""
        If val(txtPendingTask.Tag) = 5 Or val(txtPendingTask.Tag) = 6 Or val(txtPendingTask.Tag) = 9 Then
            If val(txtPendingTask.Tag) = 5 Then
                mSql = "Select intInstrumentTypeID,vchInstrumentType  From faInstrumentTypes Where intInstrumentTypeID not in(10,1)"
            ElseIf val(txtPendingTask.Tag) = 9 Or val(txtPendingTask.Tag) = 6 Then
              mSql = "Select intInstrumentTypeID,vchInstrumentType From faInstrumentTypes Where intInstrumentTypeID<>10"
            End If
            frmSearchMasters.SQLQry = mSql
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            txtInstrumentType.Text = gbSearchStr
            txtInstrumentType.Tag = gbSearchID
        End If
    End Sub
    
    Private Sub cmdKeyID_Click()
        If val(txtPendingTask.Tag) = 5 Then
            
        ElseIf val(txtPendingTask.Tag) = 6 Then ' REVERSE ENTRY
            Dim mSql    As String
            Dim Rec     As New ADODB.Recordset
            Dim mCnn    As New ADODB.Connection
            Dim objdb   As New clsDB
            Dim mStatus     As Integer
            Dim mRevSatatus As Integer
            Dim mAdjJv  As Boolean
            
            frmSearchVouchers.PreviousYearMode = 1
            frmSearchVouchers.CheckMode = 10
            
            frmSearchVouchers.chkInterrupted.Visible = False
            frmSearchVouchers.chkContra.Enabled = False
            frmSearchVouchers.chkJournal.Enabled = False
            frmSearchVouchers.chkPayment.Enabled = False
            
            frmSearchVouchers.Show vbModal
            If gbSearchID <> -1 Then
                mStatus = CheckReverseRequestExist(gbSearchID)
                If mStatus = 0 Or mStatus = 1 Or mStatus = 2 Then
                    MsgBox "Request Already Exists", vbInformation
                    txtKeyID.Text = ""
                    txtKeyID.Tag = ""
                    Exit Sub
                End If
                mRevSatatus = ReverseStatus(gbSearchID)
                If mRevSatatus <> 0 Then
                    txtKeyID.Text = ""
                    txtKeyID.Tag = ""
                    Exit Sub
                End If
                mAdjJv = AdvJournal(gbSearchCode)
                If mAdjJv = True Then
                    MsgBox "This Voucher has done Adjustment Entry.. U can't Reverse.."
                    txtKeyID.Text = ""
                    txtKeyID.Tag = ""
                    Exit Sub
                End If
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
                mSql = "Select tnyVoucherTypeID,isNull(intInstrumentTypeID,100) intInstrumentTypeID,intFinancialYearID,dtDate,numLinkKeyID From faVouchers Where intVoucherID=" & gbSearchID
                Rec.Open mSql, mCnn
                If Not Rec.EOF Then
                    If Not (IsNull(Rec!numLinkKeyID)) Then
                        MsgBox "This is Reversed Voucher", vbInformation
                        gbSearchCode = ""
                        gbSearchStr = ""
                        gbSearchID = -1
                        txtKeyID.Text = ""
                        txtKeyID.Tag = ""
                        Exit Sub
                    End If
                    If CDate(Rec!dtDate) > CDate(gbTransactionDate) Then
                        MsgBox "Request Date must be Greater/Equal to Voucher Date", vbInformation
                        gbSearchCode = ""
                        gbSearchStr = ""
                        gbSearchID = -1
                        txtKeyID.Text = ""
                        txtKeyID.Tag = ""
                        Exit Sub
                    End If
                    
                    If Rec!intFinancialYearID <> gbFinancialYearID - 1 Then
                        MsgBox "Sorry!.. This Voucher is not in the Current Financial Year", vbInformation
                        gbSearchCode = ""
                        gbSearchStr = ""
                        gbSearchID = -1
                        txtKeyID.Text = ""
                        txtKeyID.Tag = ""
                        Exit Sub
                    End If
                    If Rec!tnyVoucherTypeID = 20 Then
                        MsgBox "Payment Voucher is Not Allowed To Reverse", vbInformation
                        gbSearchCode = ""
                        gbSearchStr = ""
                        gbSearchID = -1
                        txtKeyID.Text = ""
                        txtKeyID.Tag = ""
                        Exit Sub
                    ElseIf Rec!tnyVoucherTypeID = 30 Then
                            If CheckReverseRequestExist(gbSearchID) = 1 Then
                                MsgBox "Already sent Request for this Voucher", vbInformation
                            ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                                MsgBox "This Voucher Already Reversed", vbInformation
                            Else
                                Call GetVoucherDetails(gbSearchID)
                            End If
                    ElseIf Rec!tnyVoucherTypeID = 40 Then
                            If AutoJournalCheck Then
                                MsgBox "This Journal is AutoGenerated Not allowed to Reverse"
                                Exit Sub
                            End If
                            If CheckReverseRequestExist(gbSearchID) = 1 Then
                                MsgBox "Already sent Request for this Voucher", vbInformation
                            ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                                Call GetVoucherDetails(gbSearchID)
                                MsgBox "Voucher Already Reversed", vbInformation
                            Else
                                Call GetVoucherDetails(gbSearchID)
                            End If
                    ElseIf Rec!tnyVoucherTypeID = 10 Then
                        Call GetVoucherDetails(gbSearchID)
                            
                            'NOTE:: BLOCKED BY AIBY ON 8th May,2013
                            'If Rec!intInstrumentTypeID = 5 Then
                            '    If CheckReverseRequestExist(gbSearchID) = 1 Then
                            '    ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            '        Call GetVoucherDetails(gbSearchID)
                            '    Else
                            '        Call GetVoucherDetails(gbSearchID)
                            '    End If
                            'Else
                            '    If CheckReverseRequestExist(gbSearchID) = 1 Then
                            '    ElseIf CheckReverseRequestExist(gbSearchID) = 2 Then
                            '        Call GetVoucherDetails(gbSearchID)
                            '    Else
                            '        Call GetVoucherDetails(gbSearchID)
                            '    End If
                            'End If
                            
                    End If
                End If
            Else

            End If
            gbSearchCode = ""
            gbSearchStr = ""
            gbSearchID = -1
            End If
    End Sub
    Private Function AutoJournalCheck() As Boolean
        Dim objdb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim mKeyID2 As Variant
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select intKeyID2 from faVouchers Where tnyVoucherTypeID=40 And intVoucherID=" & val(txtKeyID.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mKeyID2 = IIf(IsNull(Rec!intKeyID2), 0, Rec!intKeyID2)
            If mKeyID2 <> 0 Then
                AutoJournalCheck = True
            End If
        End If
    End Function
    Public Function GetVoucherDetails(ByVal intVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mSql        As String
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim objdb       As New clsDB
            Dim objRev      As New clsReverseProcess
            

            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where intVoucherID = " & intVoucherID
                mSql = "Select * From faVouchers "
                mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID "
                mSql = mSql + " LEFT JOIN faInstrumentTypes ON faInstrumentTypes.intInstrumentTypeID = faVouchers.intInstrumentTypeID "
                mSql = mSql + " Where intVoucherID = " & intVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                
                    'Select Case Rec!tnyVoucherTypeID
                    '    Case 10:
                    '        lblVoucherType.Caption = "Receipt Voucher"
                    '    Case 20:
                    '        lblVoucherType.Caption = "Payment Voucher"
                    '    Case 30:
                    '        lblVoucherType.Caption = "Contra Voucher"
                    '    Case 40:
                    '        lblVoucherType.Caption = "Journal Voucher"
                    'End Select
                    'lblVoucherType.Tag = Rec!tnyVoucherTypeID
                    
                    txtKeyID.Tag = Rec!intVoucherID
                    txtTransactionType.Tag = Rec!intTransactionTypeID
                    txtTransactionType.Text = Rec!vchTransactionType
                    txtKeyID.Text = Rec!intVoucherNo
                    txtTrnDate.Tag = Rec!dtDate
                    txtInstrumentType.Tag = Rec!intInstrumentTypeID
                    txtInstrumentType.Text = Rec!vchInstrumentType
                    txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtInstrumentDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    If IsDate(txtInstrumentDate) Then txtInstrumentDate.Text = DdMmmYy(txtInstrumentDate)
                    txtAmount.Text = Rec!fltAmount
                    txtAmount.Enabled = False
                    If IsDate(txtTrnDate) And IsDate(txtTrnDate.Tag) Then
                        If CDate(txtTrnDate.Text) < CDate(txtTrnDate.Tag) Then
                            MsgBox "Reverse Entry Date could not be Older than the Actual Transaction Date", vbInformation
                            txtTrnDate.Text = ""
                            Exit Function
                        End If
                    End If
                    
                    'intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                    
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function AdvJournal(ByVal VrNo As Double) As Boolean
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        AdvJournal = False
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faVouchers Where tnyVoucherGroupID=2 And numLinkKeyID = " & VrNo
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                AdvJournal = True
            End If
        End If
    End Function
    Private Function CheckReverseRequestExist(ByVal VchID As Double) As Integer
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            If objdb.SetConnection(mCnn) Then
                mSql = " Select tnyStatus from faReverseEntry "
                mSql = mSql + " Inner Join faReverseEntryChild On faReverseEntry.intRequestID = faReverseEntryChild.intRequestID "
                mSql = mSql + " Where intVoucherID =  " & VchID
                mSql = mSql + " And tnyStatus<>4"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyStatus = 0 Then      'Requested
                        CheckReverseRequestExist = 0
                    ElseIf Rec!tnyStatus = 1 Then  ' Approved
                        CheckReverseRequestExist = 1
                    ElseIf Rec!tnyStatus = 2 Then   'Reversed
                        CheckReverseRequestExist = 2
                    Else 'Cancelled Status=4
                        CheckReverseRequestExist = 4
                    End If
                    Exit Function
                Else
                    CheckReverseRequestExist = 5  'NOT EXISTS IN THE TABLE
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact Your System Administrator"
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    Private Function ReverseStatus(ByVal VchID As Double) As Integer
         On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            ReverseStatus = 0
            If objdb.SetConnection(mCnn) Then
                mSql = "Select isNull(tnyReversed,0) tnyReversed,isNull(numLinkKeyID,0) numLinkKeyID ,* From faVouchers "
                mSql = mSql + " Where intVoucherID =  " & VchID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!tnyReversed = 1 And Rec!numLinkKeyID = 0 Then

                        ReverseStatus = 1
                    ElseIf Rec!tnyReversed = 1 And Rec!numLinkKeyID <> 0 Then

                        ReverseStatus = 2
                    End If
                End If
            End If
               Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Sub cmdNew_Click()
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is CommandButton Then
                ctrl.Tag = ""
            End If
        Next
        
        cmdTasks.Tag = -1
        cmdSave.Caption = "Save"
'        Call FormInitialize
        cmdCancel.Enabled = True
        cmdCancel.Caption = "Cancel"
        txtAmount.Enabled = True
        If GetLFAStatus = True Then
         
            MsgBox "AFS is Submitted to LFA, Further Modification is not Possible", vbInformation
        Else
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                'cmdApprove.Enabled = True
            Else
                cmdSave.Enabled = True
                fraMain.Enabled = True
            End If
        End If
    End Sub


    Private Sub cmdProject_Click()
        frmSearchProjects.PreviousYearMode = 1
        frmSearchProjects.Show vbModal
        txtProject.SetFocus
    End Sub

    Private Sub cmdSave_Click()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim arrIn   As Variant
        Dim mAllotmentNo As String
        Dim intRequestID_1     As Variant
        Dim dtRequestdate_2     As Variant
        Dim intTaskID_3     As Variant
        Dim dtTransactionDate_4     As Variant
        Dim intSourceofFundID_5     As Variant
        Dim intCategoryID_6     As Variant
        Dim numIMPO_7       As Variant
        Dim intTransactionTypeId_8      As Variant
        Dim intInstrumentTypeId_9       As Variant
        Dim vchInstrumentNo_10      As Variant
        Dim dtInstrumentDate_11     As Variant
        Dim fltAmount_12    As Variant
        Dim vchRemarks_13       As Variant
        Dim numDemandID_14      As Variant
        Dim intVoucherID_15    As Variant
        Dim tnystatus_16   As Variant
        Dim intKeyId_17     As Variant
        Dim numUserID_18    As Variant
        Dim numSeatID_19    As Variant
        Dim numApprovedUser_20      As Variant
        Dim dtApprovedDate_21       As Variant
        Dim intYearID_22    As Variant
        Dim numProjectId_23    As Variant
        Dim intExpenditureHead_24    As Variant
        
        cmdSave.Enabled = True
        If val(txtPendingTask.Tag) > 0 Then
            If SaveValidate Then
                intRequestID_1 = -1
                dtRequestdate_2 = gbTransactionDate
                intTaskID_3 = val(txtPendingTask.Tag)
                dtTransactionDate_4 = txtTrnDate.Text
                
                fltAmount_12 = val(txtAmount.Text)
                vchRemarks_13 = Trim(txtRemarks.Text)
              
                numUserID_18 = gbUserID
                numSeatID_19 = gbSeatID
                intYearID_22 = gbFinancialYearID - 1
                
                
                If gbLBID = 167 Then    ''''TO SKIP THE APPROVER VALIDATION FOR TRIVANDRUM CORPORATION ONLY
                      tnystatus_16 = 2
                Else
                      tnystatus_16 = 0
                End If
                        
                dtTransactionDate_4 = Format(txtTrnDate.Text, "dd/mmm/yyyy")
                    Select Case (val(txtPendingTask.Tag))
                    Case 1
                        mAllotmentNo = txtKeyID.Text
                        intSourceofFundID_5 = val(txtSourceFund.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                    Case 2 ' Letter of Authority Cancallation
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        intKeyId_17 = val(txtInstrumentNo.Tag)
                        dtInstrumentDate_11 = txtTrnDate.Text
                        mAllotmentNo = txtKeyID.Text
                        intSourceofFundID_5 = val(txtSourceFund.Tag)
                    Case 3 ' Requisition
                        'numDemandID_14 = val(txtProject.Tag)
                        numProjectId_23 = val(txtProject.Tag)
                        intSourceofFundID_5 = val(txtSourceFund.Tag)
                        intCategoryID_6 = val(txtCategory.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        intKeyId_17 = val(txtInstrumentNo.Tag)
                    Case 4
                        
                    Case 5
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                        intInstrumentTypeId_9 = val(txtInstrumentType.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        dtInstrumentDate_11 = Format(txtInstrumentDate.Text, "dd/mmm/yyyy")
                    Case 6 'R E V E R S E   E N T R Y
                        intInstrumentTypeId_9 = val(txtInstrumentType.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        If IsDate(txtInstrumentDate) Then
                            dtInstrumentDate_11 = Format(txtInstrumentDate.Text, "dd/mmm/yyyy")
                        End If
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                        intKeyId_17 = val(txtKeyID.Tag)
                    Case 7 'Payment Order
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                        intExpenditureHead_24 = val(txtExpdHead.Tag)
                        If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Or val(txtTransactionType.Tag) = 1201 Or val(txtTransactionType.Tag) = 1391 Then
                            vchInstrumentNo_10 = txtInstrumentNo.Text
                            intSourceofFundID_5 = val(txtSourceFund.Tag)
                            intCategoryID_6 = val(txtCategory.Tag)
                            'numProjectId_23 = val(txtProject.Tag)
                        End If
                        intKeyId_17 = val(txtInstrumentNo.Tag) ' RequisitionID
                    Case 8 'Payment Order Cancel
                            
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                        intInstrumentTypeId_9 = val(txtInstrumentType.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                        dtInstrumentDate_11 = Format(txtInstrumentDate.Text, "dd/mmm/yyyy")
                        
                    Case 9
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                    Case 10  'Cancel Requisitions
                        intKeyId_17 = val(txtInstrumentNo.Tag)
                        vchInstrumentNo_10 = Trim(txtInstrumentNo.Text)
                    Case 11 ' Payment Order Approval
                        vchInstrumentNo_10 = txtInstrumentNo.Text
                        numDemandID_14 = val(txtInstrumentNo.Tag)
                        intTransactionTypeId_8 = val(txtTransactionType.Tag)
                    Case 13  'Reqisition B Fund
                        intSourceofFundID_5 = val(txtSourceFund.Tag)
                        intCategoryID_6 = val(txtCategory.Tag)
                    Case 14 ' Contra Entry
                    
                    Case 15 ' Journal Entry
                    
                    Case 16 ' UnAuthorized Drawal
                        intSourceofFundID_5 = val(txtSourceFund.Tag)
                        intCategoryID_6 = val(txtCategory.Tag)
                    Case 17 ' E bill Voucher
                        intCategoryID_6 = val(txtCategory.Tag)
                    End Select
                    
                    If val(cmdTasks.Tag) > 0 Then
                        intRequestID_1 = val(cmdTasks.Tag)
                    Else
                        intRequestID_1 = -1
                    End If
       
    
                    arrIn = Array(intRequestID_1, _
                                dtRequestdate_2, _
                                intTaskID_3, _
                                dtTransactionDate_4, _
                                intSourceofFundID_5, _
                                intCategoryID_6, _
                                numIMPO_7, _
                                intTransactionTypeId_8, _
                                intInstrumentTypeId_9, _
                                vchInstrumentNo_10, _
                                dtInstrumentDate_11, _
                                fltAmount_12, _
                                vchRemarks_13, _
                                numDemandID_14, _
                                intVoucherID_15, _
                                tnystatus_16, _
                                intKeyId_17, _
                                numUserID_18, _
                                numSeatID_19, _
                                numApprovedUser_20, _
                                dtApprovedDate_21, _
                                intYearID_22, _
                                numProjectId_23, _
                                intExpenditureHead_24)
                    On Error GoTo ErrorHandler
                    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                        objdb.ExecuteSP "spSavePendingTaskRequest", arrIn, , , mCnn, adCmdStoredProc
                        MsgBox "Saved Successfully", vbApplicationModal
                        
                        FormInitialize
                        FillGrid
                    End If
            End If
        End If
        Exit Sub
ErrorHandler:
        MsgBox "Didn't able to save request", vbInformation
    End Sub
    Private Function SaveValidate() As Boolean
        
        
        SaveValidate = True
        If val(txtPendingTask.Tag) < 1 Then
           MsgBox "Please Select Task Name", vbApplicationModal
           SaveValidate = False
           Exit Function
        End If
        If Not IsDate(txtTrnDate.Text) Then
           MsgBox "Please Enter Transaction date", vbApplicationModal
           SaveValidate = False
           Exit Function
        End If
        If Trim(txtRemarks.Text) = "" Then
           MsgBox "Please Enter Remarks", vbApplicationModal
           SaveValidate = False
           Exit Function
        End If
        If val(txtPendingTask.Tag) <> 17 Then
            If val(txtAmount.Text) = 0 Then
               MsgBox "Please Enter Amount", vbApplicationModal
               SaveValidate = False
               Exit Function
            End If
        End If
        Select Case val(txtPendingTask.Tag)
            Case Is = 3
                 If txtProject.Text = "" Then
'                    MsgBox "Please Enter Project", vbApplicationModal
'                    SaveValidate = False
'                    Exit Function
                End If
            
            Case Is = 5
                If val(txtTransactionType.Tag) < 1 Then
                    MsgBox "Please Select Transaction Type", vbApplicationModal
                    SaveValidate = False
                    Exit Function
                End If
                
                If val(txtInstrumentType.Tag) < 0 Then
                    MsgBox "Please Specify the Instrument Type!", vbInformation
                    SaveValidate = False
                    Exit Function
                End If
            Case Is = 7 'Payment Order
                
                If Trim(txtInstrumentNo.Text) = "" Then
                    If val(txtTransactionType.Tag) <> 1391 Then ' DEV.Exp OTHER THAN CAPITAL OTHER THAN PROJECT
                        'MsgBox "Please Enter the Instrument No", vbApplicationModal
                        'SaveValidate = False
                        'Exit Function
                    End If
                End If
                If val(txtExpdHead.Tag) < 1 Then
                    MsgBox "Please Select Expenditure Head", vbApplicationModal
                    SaveValidate = False
                    Exit Function
                End If
                If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
'                    If txtInstrumentNo.Text = "" Then
'                        MsgBox "Please Select Allotment No", vbApplicationModal
'                        SaveValidate = False
'                        Exit Function
'                    End If
                End If
            Case Is = 8 'Payment Order Cancellation
                If txtInstrumentNo.Text = "" Then
                    MsgBox "Please Select PayOrder No", vbApplicationModal
                    SaveValidate = False
                    Exit Function
                End If
            Case Is = 11
                If txtInstrumentNo.Text = "" Then
                   MsgBox "Please Select PayOrder No", vbApplicationModal
                   SaveValidate = False
                   Exit Function
                End If
                If CDate(txtTrnDate.Text) < CDate(mValidateTrnDate) Then
                   MsgBox "Transaction date must be Greater than PayOrder Date ", vbApplicationModal
                   SaveValidate = False
                   Exit Function
                End If
        End Select
        
    End Function
    Private Sub cmdSourceFund_Click()
        gbSearchID = -1
        gbSearchStr = ""
        
      
        Select Case val(txtPendingTask.Tag)
        Case Is = 13
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund Where intSourceFundID=3"
            frmSearchMasters.Show vbModal
            If gbSearchStr <> "" Then
                txtSourceFund.Text = gbSearchStr
                txtSourceFund.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Case Is = 16
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund Where intSourceFundID In (1,4,16,17,25,26,27,28,10,11,12,13,14,29,30)"
            frmSearchMasters.Show vbModal
            If gbSearchStr <> "" Then
                txtSourceFund.Text = gbSearchStr
                txtSourceFund.Tag = gbSearchID
            End If
            
            If val(txtSourceFund.Tag) = 1 Then
                cmdCategory.Enabled = True
            ElseIf val(txtSourceFund.Tag) = 29 Then
                txtCategory.Text = "SCP"
                txtCategory.Tag = 2
                cmdCategory.Enabled = False
            ElseIf val(txtSourceFund.Tag) = 30 Then
                txtCategory.Text = "TSP"
                txtCategory.Tag = 3
                cmdCategory.Enabled = False
            Else
                txtCategory.Text = "General"
                txtCategory.Tag = 1
                cmdCategory.Enabled = False
            End If

            gbSearchStr = ""
            gbSearchID = -1
        End Select
    End Sub
    Private Sub cmdTasks_Click()
        Call cmdNew_Click
        vsTasks.Visible = True
    End Sub
    Private Sub cmdTransactionType_Click()
        Dim mSql As String
        
        gbSearchStr = ""
        gbSearchID = -1
        
        Select Case val(txtPendingTask.Tag)
            Case Is = 1
                'If AuthorityOrAllotment = "Authority" Or AuthorityOrAllotment = "OpeningAuthority" Then
                If gbLBPanchayat = 1 Then
                    mSql = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID In(108,109,110,125,126,155,166,168,169,170,171,119,120,121,122,123,174) Order By vchTransactionType"
                Else
                    mSql = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID In(108,109,110,125,126,155,166,168,169,170,171,174) Order By vchTransactionType"
                End If
                'frmSearchTransactionType.ModeOfTransaction = 2
                frmSearchTransactionType.StrQuery = mSql
                frmSearchTransactionType.Show vbModal
            Case Is = 5
               mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
               mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
               mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
               mSql = mSql + " Where (faTransactionType.intGroupID = 10) and faTransactionType.inttransactionTypeID NOT IN (9996,9997,9998)"
               mSql = mSql + " And isNull(tnyHidden,0)=0"
               mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
               frmSearchTransactionType.ModeOfTransaction = 1
               frmSearchTransactionType.StrQuery = mSql
               frmSearchTransactionType.Show vbModal
            Case 6
               frmSearchTransactionType.ModeOfTransaction = 1
               frmSearchTransactionType.Show vbModal
            Case 7, 8
               frmSearchTransactionType.ModeOfTransaction = 2
               frmSearchTransactionType.Show vbModal
        End Select
        
        If Not gbSearchStr = "" Then
            txtTransactionType.Text = gbSearchStr
            txtTransactionType.Tag = gbSearchID
            cmdTransactionType.SetFocus
        Else
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
            txtTransactionType.SetFocus
        End If
        
        If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Or val(txtTransactionType.Tag) = 1201 Then 'Or val(txtTransactionType.Tag) = 1391 Then :: BLOCKED ON 16-JAN-2014
            lblInstNo.Caption = "Allotment No"
            lblInstNo.Visible = True
            txtInstrumentNo.Visible = True
            cmdInstNo.Visible = True
        '        ElseIf val(txtTransactionType.Tag) = 1391 Then ':: ADDED ON 2-FEB-2014
        '            lblInstNo.Caption = "Instrument No"
        '            lblInstNo.Visible = True
        '            txtInstrumentNo.Visible = True
        '            cmdInstNo.Visible = True
        Else
            lblInstNo.Caption = ""
            lblInstNo.Visible = False
            txtInstrumentNo.Visible = False
            cmdInstNo.Visible = False
        End If
        gbSearchStr = ""
        gbSearchID = -1
        
    End Sub

    Private Sub Command1_Click()
        If mID = 8 Then
            mID = 0
        Else
            mID = mID + 1
        End If
        Call SetInputField(mID)
    End Sub

    Private Sub dtInstrumentDate_CloseUp()
        txtInstrumentDate.Text = dtInstrumentDate.Value
    End Sub
    Private Sub dtTrnDate_CloseUp()
        If dtTrnDate.Value >= mPreStartDate And dtTrnDate.Value <= mPreEndDate Then
           txtTrnDate.Text = Format(dtTrnDate.Value, "dd/mmm/yyyy")
        Else
            MsgBox "Invalid Date", vbApplicationModal
        End If
        txtTrnDate.SetFocus
        'Call CheckLastPostingDate
    End Sub
    Private Sub CheckLastPostingDate()   '-----------------LAST POSTING VALIDATION------------------
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim Rec As New Recordset
        Dim dtPostingDate As Date
        
        Call SetgbLastPostingDate
        
        objdb.SetConnection mCnn
        mSql = "SELECT MAX(dtPostingDate) dtPostingDate FROM faPostingIndex WHERE tnyStage=2"
        Set Rec = GetRecordSet(mSql)
        If Not (Rec.BOF And Rec.EOF) Then
            dtPostingDate = Format(Rec!dtPostingDate, "dd-mmm-yyyy")
            If CDate(mPreEndDate) <= CDate(dtPostingDate) Then
                MsgBox "Transactions Locked for the Month!!!No More Transactions Is Possible for Current Date And less", vbInformation
                txtTrnDate.Text = ""
                cmdSave.Enabled = False
                cmdNew.Enabled = False
                vsGrid.Enabled = False
                Exit Sub
            End If
        End If
    End Sub
    Private Sub Form_Activate()
        
        Call FillGrid
        
        'vsTasks.RowHidden(7) = True ' DEMAND/RECEIPTS
        'vsTasks.RowHidden(8) = True ' REVERSE ENTRY
        'vsTasks.RowHidden(1) = True ' LETTER OF AUTHORITY
        'vsTasks.RowHidden(2) = True ' CANCEL LETTER OF AUTHORIYT
        'vsTasks.RowHidden(6) = True ' LETTER OF ALLOTMENT - B FUND
        'vsTasks.RowHidden(11) = True ' PAY ORDER CANCELLATION
        
        vsTasks.RowHidden(12) = True ' INTERRUPTED RECEIPTS
        vsTasks.RowHidden(13) = True ' REQUISITION B-Fund
        
        'vsTasks.RowHidden(14) = True ' CONTRA ENTRY
        'vsTasks.RowHidden(15) = True ' JOURNAL ENTRY
        Call FormInitialize
    End Sub

    Private Sub Form_Load()
     Exit Sub
        Dim PreFinEndDate As Date
        'Me.WindowState = vbMaximized
        WindowsXPC1.InitIDESubClassing
        Call FillCombo
        Call FillComboTask
        Call GetLFAStatus
        Call SetInputField(0)
        Call FormInitialize
        Call FillGrid
        
        Call CheckLastPostingDate
        PreFinEndDate = DateAdd("d", -1, gbStartingDate)
        dtTrnDate.Value = PreFinEndDate
    End Sub
    Private Sub FillCombo()
        cmbMonth.AddItem ""
        cmbMonth.ItemData(0) = 0
        
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(1) = 3
        
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(2) = 2
        
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(3) = 3
        
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(4) = 12
        
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(5) = 11
        
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(6) = 10
        
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(7) = 9
        
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(8) = 8
        
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(9) = 7
        
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(10) = 6
        
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(11) = 5
        
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(12) = 4

    End Sub
    Private Sub FillComboTask()
        Dim mCnt As Integer
      
        cmbSearchTask.AddItem ""
        cmbSearchTask.ItemData(0) = 0
        For mCnt = 1 To vsTasks.Rows - 1
            cmbSearchTask.AddItem vsTasks.TextMatrix(mCnt, 1)
            cmbSearchTask.ItemData(mCnt) = val(vsTasks.TextMatrix(mCnt, 0))
        Next
    End Sub
    Private Sub FormInitialize()
        Dim objdb         As New clsDB
        Dim mSql          As String
        Dim mCnn          As New ADODB.Connection
        Dim Rec           As New ADODB.Recordset
        Dim ctrl          As Control
        
        mSql = " Select * From faFinancialYear Where tinCurrentFinancialYearFlag=1"
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                mCurFinancialYear = Rec!intFinancialYear
                mPreFinancialYear = mCurFinancialYear - 1
            Else
                MsgBox "Current Financial Year Not Set", vbApplicationModal
                Exit Sub
            End If
            Rec.Close
            
            mSql = " Select * From faFinancialYear Where intFinancialYear=" & mPreFinancialYear
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If (Rec.EOF And Rec.BOF) Then
               MsgBox "Previous Financial Year Not Set", vbApplicationModal
               Exit Sub
            Else
                mPreStartDate = Rec!dtStartingDate
                mPreEndDate = Rec!dtEndingDate
            End If
            lblcaption.Caption = "Pending Tasks Request For The Year " & mPreFinancialYear & "-" & mPreFinancialYear + 1
        End If
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdSave.Visible = True
            cmdApprove.Visible = False
            cmdNew.Enabled = True
        ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Visible = False
            cmdApprove.Visible = True
            cmdCancel.Caption = "Reject"
            cmdNew.Enabled = False
        End If
        
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is OptionButton Then
                ctrl.Value = False
            ElseIf TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                ctrl.Tag = ""
            End If
        Next
        Call SetInputField(0)
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            fraMain.Enabled = False
        Else
            fraMain.Enabled = True
        End If
    End Sub
    
    Private Sub Form_Paint()
        Call FillGrid
    End Sub

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtAmount_LostFocus()
        Dim mMsg As String
        Dim mAmt As Double
        If val(txtAmount.Text) > 0 Then
            mAmt = val(txtAmount.Text)
            txtAmount.Text = Format(txtAmount.Text, "0.00")
        End If
        
        Select Case val(txtPendingTask.Tag)
            Case Is = 3 'Requisition
                    If val(txtAmount.Text) > val(txtAmount.Tag) Then
                        mMsg = "Balance Amount avaliable for " & txtSourceFund.Text & vbCrLf
                        mMsg = mMsg + " for this project is  Only Rs. " & Format(val(txtAmount.Tag), "0.00")
                        MsgBox mMsg, vbInformation
                        txtAmount.Text = ""
                        txtAmount.SetFocus
                        Exit Sub
                    End If
                    
                    
                    Dim mSql        As String
                    Dim objdb       As New clsDB
                    Dim mCnn        As New ADODB.Connection
                    Dim Rec         As New ADODB.Recordset
                    Dim mTotAmount  As Variant
                    Dim mUtilizedAmt As Variant
            
                    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                        If val(txtSourceFund.Tag) = 3 Then
                            mSql = " Select sum(fltAmount) TotAmount from faAllotmentLetters"
                            mSql = mSql + "     Where intSourceOfFundID = " & val(txtSourceFund.Tag) & " And intSchemeID=" & txtCategory.Tag & " AND faAllotmentLetters.tnyStatus <> 8 And intFinancialYearID = " & gbFinancialYearID - 1
                            mSql = mSql + "     AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                        Else
                            mSql = " Select A.TotAmount TotAmount From"
                            mSql = mSql + " (   Select sum(fltAmount) TotAmount from faAllotmentLetters"
                            mSql = mSql + "     Where intSourceOfFundID = " & val(txtSourceFund.Tag) & " And faAllotmentLetters.tnyStatus <> 8 And intFinancialYearID = " & gbFinancialYearID - 1
                            mSql = mSql + "     AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                            mSql = mSql + "    Union All"
                            mSql = mSql + "     Select sum(fltAmount) TotAmount from faExtractAllotments"
                            mSql = mSql + "     Where intSourceOfFundID = " & val(txtSourceFund.Tag) & "  And intFinancialYearID = " & gbFinancialYearID - 1
                            mSql = mSql + " )A"
                        End If
                        Rec.Open mSql, mCnn
                        
                        If Not (Rec.EOF And Rec.BOF) Then
                            While Not Rec.EOF
                                mTotAmount = mTotAmount + IIf(IsNull(Rec!TotAmount), 0, Rec!TotAmount)
                            Rec.MoveNext
                            Wend
                        End If
                        Rec.Close
                        If val(txtSourceFund.Tag) = 3 Then
                            mSql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & val(txtSourceFund.Tag) & "  And  intSchemeID = " & txtCategory.Tag & " And tnyStatus  in (0,1)  And intFinancialYearID = " & gbFinancialYearID - 1
                            mSql = mSql + " AND ISNULL(tnyTypeID,0) NOT IN (1,2)"
                        Else
                            mSql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & val(txtSourceFund.Tag) & "  And tnyStatus  in (0,1)  And intFinancialYearID = " & gbFinancialYearID - 1
                            mSql = mSql + " AND ISNULL(tnyTypeID,0) NOT IN (1,2)"
                        End If
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                           mUtilizedAmt = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                        End If
                        Rec.Close
                        If val(mUtilizedAmt + val(txtAmount.Text)) > mTotAmount Then
                            MsgBox "Amount Exceed (The total amount alloted is only Rs." & Format(mTotAmount, "0.00") & ")", vbInformation
                            Exit Sub
                        End If
                    End If
                    
                    
        End Select
        
    End Sub

    Private Sub txtCategory_GotFocus()
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mRAmount    As Long
        
        If val(txtSourceFund.Tag) > 0 Then
            If val(txtCategory.Tag) > 0 Then
                If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                    'mSql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & txtSourceFund.Tag
                    'mSql = mSql + " And tnyStatus not in (8,0) And intSchemeID = " & txtCategory.Tag & " And intFinancialYearID = " & gbFinancialYearID - 1 & " "
                    
                    mSql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & txtSourceFund.Tag
                    mSql = mSql + " And tnyStatus not in (8,0) And intCategoryID = " & txtCategory.Tag & " And intFinancialYearID = " & gbFinancialYearID - 1 & " "
                    mSql = mSql + "     AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                       mRAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                    End If
                    
                    txtAmount.Tag = Abs(val(mRAmount))
                    If val(txtAmount.Text) >= 0 Then
                        If val(txtAmount.Text) > val(txtAmount.Tag) Then
                            mSql = " Balance Available for " & txtSourceFund.Text & "  & vbCrLf  'Amount allocated"
                            mSql = mSql + " is Rs. " & Format(val(txtAmount.Tag), "0.00")
                            MsgBox mSql
                            Exit Sub
                        End If
                    End If
                End If
                Rec.Close
            End If
       End If
    End Sub

    Private Sub txtInstrumentNo_LostFocus()
        If CheckCancelRequisitionStatus(val(txtInstrumentNo.Tag)) = 1 Then
            MsgBox "Request Already Given", vbInformation, "Saankhya"
            txtInstrumentNo.Text = ""
            txtInstrumentNo.Tag = ""
            txtInstrumentNo.SetFocus
            txtAmount.Text = ""
        End If
    End Sub





    Private Sub txtProject_GotFocus()
         Dim objProj As New clsProject
            Dim objProFund As New clsProjectFund
            Dim mProjectID As Variant
            Dim mSourceOfFundID As Variant
            Dim mSubsectorID As Integer
            Dim mintCategoryID As Integer
            Dim mCol As Collection
            Dim mRow As Integer
            
            
            mProjectID = gbSearchStr
            mSourceOfFundID = gbSearchID
            If val(gbSearchStr) > 0 Then
                objProj.SetProject mProjectID, gbFinancialYearID - 1
                If objProj.ProjectID > 0 Then
                    txtProject.Text = "[" & objProj.ProjectSerialNo & "]" & objProj.ProjectNameEnglish
                    'txtProjectNo.Text = objProj.ProjectSerialNo
                    txtProject.Tag = objProj.ProjectID
                    
                    txtCategory.Text = objProj.Category
                    txtCategory.Tag = objProj.ProjCatID
                    
                    txtSourceFund.Text = objProj.FindSourceOfFund(mSourceOfFundID)
                    txtSourceFund.Tag = objProj.SourceOfFundID
                    
                    Set mCol = objProj.GetFundDetails(CInt(gbFinancialYearID - 1), objProj.ProjectID)
                    For mRow = 1 To mCol.count
                        Set objProFund = mCol.Item(mRow)
                        If objProFund.SourceOfFundID = mSourceOfFundID Then
                            txtAmount.Tag = objProFund.SourceWiseAmount
                            Exit For
                        End If
                    Next mRow
                    
                End If
    
                
                    Dim mCnnAmt    As New ADODB.Connection
                    Dim RecAmt     As New ADODB.Recordset
                    Dim objAmt     As New clsDB
                    Dim mSQLAmt    As String
                    Dim mUtilizedAmt As Variant
                    Dim mSql As String
            
                    mUtilizedAmt = 0
                    objAmt.CreateNewConnection mCnnAmt, enuSourceString.Saankhya
                    mSQLAmt = "select fltRequestedAmt from faAllotments where tnyStatus<>2 And numProjectID= " & val(txtProject.Tag) & "  And intFinancialYearID=" & gbFinancialYearID - 1 & " And intSourceID=" & val(mSourceOfFundID) & " "
                    mSQLAmt = mSQLAmt + " AND ISNULL(tnyTypeID,0) NOT IN (1)"
                    RecAmt.Open mSQLAmt, mCnnAmt
                    If Not (RecAmt.EOF And RecAmt.BOF) Then
                        While Not (RecAmt.EOF)
                            mUtilizedAmt = mUtilizedAmt + IIf(IsNull(RecAmt!fltRequestedAmt), 0, RecAmt!fltRequestedAmt)
                            RecAmt.MoveNext
                        Wend
                    End If
                    RecAmt.Close
            
                    txtAmount.Tag = Abs(val(txtAmount.Tag) - mUtilizedAmt)
                    If val(txtAmount.Text) >= 0 Then
                        If val(txtAmount.Text) > val(txtAmount.Tag) Then
                            mSql = " Balance Available for " & txtSourceFund.Text & " in this Project" & vbCrLf  'Amount allocated
                            mSql = mSql + " is Rs. " & Format(val(txtAmount.Tag), "0.00")
                            MsgBox mSql
                            'txtAmountRequested.SetFocus
                            Exit Sub
                        End If
                    End If
                    Exit Sub
                
            End If
    End Sub



    Private Sub txtTrnDate_LostFocus()
        If Len(Trim(txtTrnDate)) Then
            txtTrnDate.Text = CheckDateInMMM(txtTrnDate.Text)
        End If
        
        If IsDate(txtTrnDate.Text) Then
            'If CDate(txtTrnDate.Text) >= mPreStartDate And CDate(txtTrnDate.Text) <= mPreEndDate Then
            If CDate(txtTrnDate.Text) < DateAdd("yyyy", -1, gbStartingDate) Or CDate(txtTrnDate.Text) > DateAdd("yyyy", -1, gbEndingDate) Then
                MsgBox "Please Enter a Date betwwen Previous financialYear", vbApplicationModal
                txtTrnDate.Text = ""
                Exit Sub
            End If
        Else
            txtTrnDate.Text = ""
        End If
        
        Select Case val(txtPendingTask.Tag)
            Case Is = 6
                If IsDate(txtTrnDate) And IsDate(txtTrnDate.Tag) Then
                    If CDate(txtTrnDate.Text) < CDate(txtTrnDate.Tag) Then
                        MsgBox "Reverse Entry Date could not be Older than the Actual Transaction Date", vbInformation
                        txtTrnDate.Text = ""
                        Exit Sub
                    End If
                End If
            Case Is = 8 ' PAY ORDER CANCELLATION
                If IsDate(mValidateTrnDate) Then
                        If IsDate(txtTrnDate.Text) Then
                            If CDate(txtTrnDate) < mValidateTrnDate Then
                                MsgBox "Cancellation Date Can not be less then the Pay Order Date", vbInformation
                                txtTrnDate.Text = ""
                                txtTrnDate.SetFocus
                            End If
                        Else
                            txtTrnDate.Text = ""
                        End If
                    Else
                        txtInstrumentNo.Text = ""
                        txtInstrumentNo.Tag = ""
                    End If
            Case Else
            
            
        End Select
        
        
    End Sub

    Private Sub vsGrid_Click()
        If vsGrid.Row > 0 Then
            Call FillDetails(val(vsGrid.TextMatrix(vsGrid.Row, 8)))
        End If
    End Sub
    Private Sub vsGrid_DblClick()
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        
        If vsGrid.Row > 0 Then
            If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 8 Then ' TASK PROCESSING COMPLETED
                MsgBox "Process Completed", vbApplicationModal
                Exit Sub
            End If
            
            ' NOTE: vsGrid.TextMatrix(vsGrid.Row, 10)) = 2 => APPROVED TASK LIST
            
            If val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 2 And gbSeatGroupID <> gbSeatGroupAccountsSuperintended Then
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    Select Case val(vsGrid.TextMatrix(vsGrid.Row, 9))
                        Case 1 ' Letter of Authority
                            frmAllotmentLetter.AuthorityOrAllotment = "Authority"
                            frmAllotmentLetter.PreviousYearMode = 1
                            frmAllotmentLetter.PreviousYearTaskID = 1
                            frmAllotmentLetter.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmAllotmentLetter.LoadMode = 10
                            frmAllotmentLetter.Show vbModal
                        Case 2 'LETTER OF AUTHORITY CANCELLATION
                            GoTo ApproverCheck:
                        Case 3 ' REQUISTION
                            
                            frmRequisition.PreviousYearMode = 1
                            frmRequisition.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmRequisition.Show vbModal
                            
                        Case 4 ' LETTTER OF ALLOTMENT [B-FUND]
                            frmAllotmentLetter.AuthorityOrAllotment = "Allotment"
                            frmAllotmentLetter.PreviousYearMode = 1
                            frmAllotmentLetter.PreviousYearTaskID = 4
                            frmAllotmentLetter.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            'frmAllotmentLetter.cmbTransactionTypes = "B Fund - State Sponsored Scheme Funds"
                            frmAllotmentLetter.LoadMode = 10
                            frmAllotmentLetter.Show vbModal
                            
                        Case 5 ' Demand/Receipt
                            frmDemandInterface.PreviousYearMode = 1
                            frmDemandInterface.PendingTaskReqID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmDemandInterface.Show vbModal
                        Case 6
                            frmReverseRequest.PreviousYearMode = 1
                            frmReverseRequest.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmReverseRequest.Show vbModal
                        Case 7 ' Payment Order
                            frmPaymentOrder.PendingTask = 2
                            frmPaymentOrder.PendingTaskReqID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmPaymentOrder.Show
                        Case 8 ' Payment Order Cancel
                            frmPayOrderCancellations.PreviousYearMode = 1
                            
                            frmPayOrderCancellations.PendingTaskReqID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmPayOrderCancellations.cmdPayorderSearch.Enabled = False
                            frmPayOrderCancellations.txtPayOrderNo.Text = txtInstrumentNo.Text
                            Call frmPayOrderCancellations.txtPayOrderNo_LostFocus
                            frmPayOrderCancellations.Show vbModal
                        Case 10 ' Cancel Requisition
                            frmCancelRequisitions.PreviousYearMode = 1
                            frmCancelRequisitions.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmCancelRequisitions.Show vbModal
                        Case 11 'Payment Order Approval
                            If gbLBPanchayat = 1 Then
                                frmPaymentOrder.PendingTask = 2
                                frmPaymentOrder.FillPayOrder (val(txtInstrumentNo.Tag))
                                frmPaymentOrder.PendingTaskReqID = vsGrid.TextMatrix(vsGrid.Row, 8)
                                frmPaymentOrder.ListLoaded = True  ' To inform this From ( frmViewPaymentOrder ) is loaded
                                frmPaymentOrder.Show
                            End If
                        Case 13 ' Requisition for BFund
                            frmRequisition.PreviousYearMode = 1
                            frmRequisition.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmRequisition.Show vbModal
                        Case 14 ' Contra Voucher
                            frmContraEntry.PreviousYearMode = 1
                            frmContraEntry.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmContraEntry.Show
                            frmContraEntry.ZOrder (0)
                         Case 15 ' Journal Voucher
                            frmJournalEntry.PreviousYearMode = 1
                            frmJournalEntry.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmJournalEntry.Show
                            frmJournalEntry.ZOrder (0)
                        Case 16  ' UnAuthorized Drawal
                            frmRequisition.PreviousYearMode = 1
                            frmRequisition.LoadMode = 10
                            frmRequisition.PreviousYearRequestID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                            frmRequisition.Show vbModal
                        Case 17  ' E Bill Voucher
                            frmWebExtracts.mPreviousYearMode = 1
                            frmWebExtracts.Show
                            frmWebExtracts.ZOrder (0)
                     End Select
                ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Or val(vsGrid.TextMatrix(vsGrid.Row, 10)) = 2 Then
                 '   If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 17 Then
                            frmWebExtracts.mPreviousYearMode = 1
                            frmWebExtracts.Show
                            frmWebExtracts.ZOrder (0)
                 '   End If
                ElseIf gbSeatGroupID = gbSeatGroupAssistantSecretary Or gbSeatGroupID = gbSeatGroupHeadClerk Then
                        If gbLBPanchayat = 1 Then
                            frmPaymentOrder.PendingTask = 1
                            frmPaymentOrder.FillPayOrder (val(txtInstrumentNo.Tag))
                            frmPaymentOrder.PendingTaskReqID = vsGrid.TextMatrix(vsGrid.Row, 8)
                            frmPaymentOrder.ListLoaded = True  ' To inform this From ( frmViewPaymentOrder ) is loaded
                            frmPaymentOrder.Show
                        End If
                Else
                    If val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 11 Or val(vsGrid.TextMatrix(vsGrid.Row, 9)) = 2 Then
                        GoTo ApproverCheck:
                    End If
                    MsgBox "Only Accountant can Do further process", vbApplicationModal
                End If
                    
            Else ' SECRETARY/ACCOUNTS OFFICER
ApproverCheck:
                Select Case val(vsGrid.TextMatrix(vsGrid.Row, 9))
                    Case Is = 2 'LETTER OF AUTHORITY CANCELLATION
                        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
                            Call CancelLetterOfAuthority
                            'If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            '    If MsgBox("Cancel this Letter of Authority?", vbYesNo + vbInformation) = vbYes Then
                            '        mSql = "Update faAllotmentLetters SET tnyStatus = 8 WHERE intAllotmentID = " & val(txtInstrumentNo.Tag)
                            '        objdb.SetConnection mCnn
                            '        mCnn.Execute mSql
                            '
                            '        mSql = "Update faPendingTaskRequest SET tnyStatus = 8 Where intRequestID=" & val(cmdTasks.Tag)
                            '        mCnn.Execute mSql
                            '    End If
                            'End If
                        ElseIf gbLBType = 3 Or gbLBType = 4 Then
                            Call CancelLetterOfAuthority
                            
                            'If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                            '    If MsgBox("Cancel this Letter of Authority?", vbYesNo + vbInformation) = vbYes Then
                            '        mSql = "Update faAllotmentLetters SET tnyStatus = 8 WHERE intAllotmentID = " & val(txtInstrumentNo.Tag)
                            '        objdb.SetConnection mCnn
                            '        mCnn.Execute mSql
                            '        mSql = "Update faPendingTaskRequest SET tnyStatus = 8 Where intRequestID=" & val(cmdTasks.Tag)
                            '        mCnn.Execute mSql
                            '    End If
                            'End If
                        End If
                    Case Is = 8
                        If gbLBType = 4 Then
                            If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                                frmPayOrderCancellations.PreviousYearMode = 1
                                frmPayOrderCancellations.PendingTaskReqID = val(vsGrid.TextMatrix(vsGrid.Row, 8))
                                frmPayOrderCancellations.cmdPayorderSearch.Enabled = False
                                frmPayOrderCancellations.txtPayOrderNo.Text = txtInstrumentNo.Text
                                Call frmPayOrderCancellations.txtPayOrderNo_LostFocus
                                frmPayOrderCancellations.Show vbModal
                            End If
                        End If
                        Exit Sub
                    
                    End Select
                

                If val(vsGrid.TextMatrix(vsGrid.Row, 9)) <> 11 Then   ' Previous Year's Payment Order Approval
                    If gbSeatGroupID <> gbSeatGroupAccountsOfficer Then
                        MsgBox "Request Pending For Approval", vbApplicationModal
                        Exit Sub
                    End If
                Else
STEP1:
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                        frmPaymentOrder.PendingTask = 1
                        frmPaymentOrder.FillPayOrder (val(txtInstrumentNo.Tag))
                        frmPaymentOrder.PendingTaskReqID = vsGrid.TextMatrix(vsGrid.Row, 8)
                        frmPaymentOrder.ListLoaded = True  ' To inform this From ( frmViewPaymentOrder ) is loaded
                        frmPaymentOrder.Show
                    Else
                        MsgBox "Only Approver can do further process", vbApplicationModal
                    End If
                 End If
            End If
             
        End If
    End Sub
    
    Private Sub CancelLetterOfAuthority()
        Dim mCnn    As New ADODB.Connection
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mStatus As Variant
    
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select tnyStatus From faAllotmentLetters Where intAllotmentID = " & val(txtInstrumentNo.Tag)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            End If
            Rec.Close
            
            If mStatus <> "" Then
                If (mStatus = 0) Then 'Letter of Authority/Allotment is not approved
                    mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & val(txtInstrumentNo.Tag)
                    MsgBox "Letter of Authority Cancelled", vbInformation
                    'cmdCancellAllotment.Enabled = False
                End If
                If (mStatus = 1) Then 'Letter of Authority/Allotment is  approved
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        mSql = "Select faVouchers.intVoucherNo,faIDemandTBL.numDemandID,faIDemandTBL.vchDemandNo,*"
                        mSql = mSql + " From faAllotmentLetters"
                        mSql = mSql + " Inner Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                        mSql = mSql + " Inner Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                        mSql = mSql + " Where intAllotmentID = " & val(txtInstrumentNo.Tag)
                        mSql = mSql + " And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                        Rec.Open mSql, mCnn 'Checking whether Receipt is taken against the Letter of Authority
                        If Not (Rec.EOF And Rec.BOF) Then
                            Rec.Close
                            mSql = "Select faVouchers.intVoucherNo,faIDemandTBL.numDemandID,faIDemandTBL.vchDemandNo,*"
                            mSql = mSql + " ,faVouchers.tnyStatus Status, faVouchers.tnyCancelFlag CancelFlag"
                            mSql = mSql + " From faAllotmentLetters"
                            mSql = mSql + " Inner Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                            mSql = mSql + " Inner Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                            mSql = mSql + " Left Join faReverseEntryChild On faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                            mSql = mSql + " Left Join faReverseEntry On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID And faReverseEntry.tnyStatus = 2 "
                            mSql = mSql + " Where intAllotmentID = " & val(txtInstrumentNo.Tag)
                            mSql = mSql + " And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                            Rec.Open mSql, mCnn 'Checking whether Receipt is Reversed
                            If Not (Rec.EOF And Rec.BOF) Then
                                If (IIf(IsNull(Rec!Status), 0, Rec!Status) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Or IIf(IsNull(Rec!tnyReversed), 0, Rec!tnyReversed) = 1 Then
                                    mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & val(txtInstrumentNo.Tag) 'Cancelling the Letter of Authority/Allotment
                                    mSql = "Update faPendingTaskRequest SET tnyStatus = 8 Where intRequestID=" & val(cmdTasks.Tag)
                                    mCnn.Execute mSql
                                    MsgBox "Letter of Authority Cancelled", vbInformation
                                    'cmdCancellAllotment.Enabled = False
                                Else
                                    MsgBox "Please Reverse/Cancel the Receipt Issued", vbInformation
                                    Exit Sub
                                End If
                            End If
                            Rec.Close
                        Else
                             'If Me.AuthorityOrAllotment = "Authority" Then
                                 If val(txtInstrumentNo.Tag) > 0 Then
                                     mCnn.Execute "Update faIDemandTBL Set tnyStatus = 9 Where numDemandID = " & val(txtInstrumentNo.Tag) 'Cancelling the Letter of Authority/Allotment
                                 End If
                             'End If
                             mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & val(txtInstrumentNo.Tag) 'Cancelling the Letter of Authority/Allotment
                             mSql = "Update faPendingTaskRequest SET tnyStatus = 8 Where intRequestID=" & val(cmdTasks.Tag)
                             mCnn.Execute mSql
                             MsgBox "Letter of Authority Cancelled", vbInformation
                             'cmdCancellAllotment.Enabled = False
                         End If
                     Else
                        MsgBox "You are not authorized to Cancel this Approved Letter of Authority", vbInformation
                     End If
                End If
            End If
        Else
            MsgBox "Connection to Finance does not exist, Please contact System Administrator", vbInformation
        End If
    
    End Sub
    
    Private Sub FillDetails(mReqID)
        Dim objdb           As New clsDB
        Dim mSql            As String
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim objTrType       As New clsTransactionType
        Dim objAcc          As New clsAccounts
        Dim mRequestID      As Integer
        Dim mRec            As New ADODB.Recordset
        Dim objAllotment    As New clsAllotmentLetter
        
        mRequestID = mReqID
        mCancelFlag = False
        Call cmdNew_Click
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faPendingTaskRequest "
            mSql = mSql + " Where intRequestID = " & mReqID
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                Call SetInputField(Rec!intTaskID)
                Select Case Rec!intTaskID
                        Case 1
                            
                            txtPendingTask.Text = "Letter of Authority"
                            txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                            objTrType.SetTransactionType (IIf(IsNull(Rec!intTransactionTypeID), 0, Rec!intTransactionTypeID))
                            txtTransactionType.Tag = objTrType.TransactionTypeID 'Rec!intTransactionTypeID
                            txtTransactionType.Text = objTrType.TransactionType
                        
                        Case 2
                            txtPendingTask.Text = "Cancel Letter of Authority"
                           
                            txtTrnDate.Text = DdMmmYy(Rec!dtTransactionDate)
                            txtInstrumentNo.Text = Rec!vchInstrumentNo
                            txtInstrumentNo.Tag = Rec!intKeyID
                            txtAmount.Text = Rec!fltAmount
                            txtRemarks.Text = Rec!vchRemarks
                        Case 3
                            txtPendingTask.Text = "Requisition"
                            Call GetRequisitionDetails(IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID), IIf(IsNull(Rec!intSourceOfFundID), 0, Rec!intSourceOfFundID))
                        Case 4
                            txtPendingTask.Text = "Letter of Allotment (B Fund)"
                        Case 5
                            If Not IsNull(Rec!intInstrumentTypeID) Then
                                txtPendingTask.Text = "Demand/Receipt"
                                txtInstrumentType.Text = FindMaster("faInstrumentTypes", "vchInstrumentType", "intInstrumentTypeID", Rec!intInstrumentTypeID)
                                txtInstrumentType.Tag = Rec!intInstrumentTypeID
                            End If
                            If Not IsNull(Rec!intTransactionTypeID) Then
                                objTrType.SetTransactionType (Rec!intTransactionTypeID)
                                txtTransactionType.Tag = Rec!intTransactionTypeID
                                txtTransactionType.Text = objTrType.TransactionType
                            End If
                        Case 6
                            txtPendingTask.Text = "Reverse Entry"
                            txtPendingTask.Tag = Rec!intTaskID
                            Call SetInputField(val(txtPendingTask.Tag))
                            Call GetVoucherDetails(Rec!intKeyID)
                        Case 7
                            txtPendingTask.Text = "Pay Order"
                            txtPendingTask.Tag = Rec!intTaskID
                            Call SetInputField(val(txtPendingTask.Tag))
                           
                            objTrType.SetTransactionType (Rec!intTransactionTypeID)
                            txtTransactionType.Tag = Rec!intTransactionTypeID
                            txtTransactionType.Text = objTrType.TransactionType
                            txtExpdHead.Tag = Rec!intExpenditureHead
                            objAcc.SetAccountID (val(txtExpdHead.Tag))
                            txtExpdHead.Text = objAcc.AccountCode + " " + objAcc.AccountHead
                            If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
                               lblInstNo.Caption = "Allotment No"
                               lblInstNo.Visible = True
                               txtInstrumentNo.Visible = True
                               cmdInstNo.Visible = True
                            End If
                            
'                            Dim objProject As New clsProject
'                            objProject.SetProject val(txtProject.Tag), gbFinancialYearID - 1
'                            If Not IsNull(objProject.ProjectID) Then
'                                txtProject.Text = IIf(IsNull(objProject.ProjectSerialNo), "", objProject.ProjectSerialNo)
'                                'txtProject.Tag = IIf(IsNull(objProject.ProjectID), "", objProject.ProjectID)
'                                txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
'                                txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
'                                txtProject.Visible = True
'                                lblProject.Visible = True
'                                txtCategory.Visible = True
'                                lblCategory.Visible = True
'                            End If
                            
                            If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
'                                txtInstrumentNo.Text = Rec!vchInstrumentNo
'                                mSQL = "Select intID From faAllotments Where vchAllotmentNo=" & txtInstrumentNo.Text
'                                Set mRec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
'                                If Not (mRec.EOF And mRec.BOF) Then
'                                    txtInstrumentNo.Tag = mRec!intID
'                                    objAllotment.SetAllotment (txtInstrumentNo.Tag)
'
'                                    txtSourceFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
'                                    txtSourceFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), "", objAllotment.SourceOfFundID)
'                                    txtImpo.Text = IIf(IsNull(objAllotment.ImplementingOfficer), "", objAllotment.ImplementingOfficer)
'                                    txtImpo.Tag = IIf(IsNull(objAllotment.ImplementingOfficersID), "", objAllotment.ImplementingOfficersID)
'                                    txtAmount.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)
'                                    txtProject.Tag = IIf(IsNull(objAllotment.ProjectID), "", objAllotment.ProjectID)
'                                    lblIMPO.Visible = True
'                                    txtImpo.Visible = True
'                                    lblFund.Visible = True
'                                    txtExpdHead.Tag = objAllotment.GrossAccountHeadID
'                                    objAcc.SetAccountID (val(txtExpdHead.Tag))
'                                    txtExpdHead.Text = objAcc.AccountCode + objAcc.AccountHead
'                                    txtSourceFund.Visible = True
'
'                                    objProject.SetProject (val(txtProject.Tag))
'                                    If Not IsNull(objProject.ProjectID) Then
'                                        txtProject.Text = IIf(IsNull(objProject.ProjectSerialNo), "", objProject.ProjectSerialNo)
'                                        'txtProject.Tag = IIf(IsNull(objProject.ProjectID), "", objProject.ProjectID)
'                                        txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
'                                        txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
'                                        txtProject.Visible = True
'                                        lblProject.Visible = True
'                                        txtCategory.Visible = True
'                                        lblCategory.Visible = True
'                                    End If
                                'End If
                            End If
                        Case 8 'Payment Order Cancel
                            txtPendingTask.Text = "Pay Order Cancel"
                            txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                            If Rec!tnyStatus = 8 Then
                                 If (POCancelValidation = True) Then mCancelFlag = False Else mCancelFlag = True
                            End If
                        Case 9 '
                            txtPendingTask.Text = "Interrupted Receipt"
                        Case 10
                            txtPendingTask.Text = "Cancel Requisition"
                            txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                            txtInstrumentNo.Tag = Rec!intKeyID
                        Case 11 'PayOrder Approval
                            txtPendingTask.Text = "Pay Order Approval"
                            txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                            txtInstrumentNo.Tag = Rec!numDemandID
                            objTrType.SetTransactionType (Rec!intTransactionTypeID)
                            txtTransactionType.Tag = Rec!intTransactionTypeID
                            txtTransactionType.Text = objTrType.TransactionType
                        Case 12
                            txtPendingTask.Text = "OBRP"
                        Case 13
                            txtPendingTask.Text = "Requisition BFund"
                            Call GetBFundRequisitionDetails(mRequestID)
                        Case 14
                            txtPendingTask.Text = "Contra Entry"
                        Case 15
                            txtPendingTask.Text = "Journal Entry"
                        Case 16
                            txtPendingTask.Text = "UnAuthorized Drawal"
                            Call GetUnAuthorizedDrawalRequisitionDetails(mRequestID)
                        Case 17
                            txtPendingTask.Text = "E-Bill Voucher"
                    End Select
                    
                    txtTrnDate.Text = Format(Rec!dtTransactionDate, "dd/mmm/yyyy")
                    txtPendingTask.Tag = Rec!intTaskID
                    cmdTasks.Tag = Rec!intRequestID
                    txtAmount.Text = Rec!fltAmount
                    txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    
                    If Rec!tnyStatus = 0 Then
                        cmdNew.Enabled = True
                        cmdSave.Caption = "Edit"
                        cmdSave.Enabled = True
                        cmdCancel.Enabled = True
                        cmdApprove.Enabled = True
                    ElseIf Rec!tnyStatus = 1 Then
                        cmdApprove.Enabled = True
                        cmdCancel.Enabled = True
                         cmdNew.Enabled = True
                    ElseIf Rec!tnyStatus = 2 Then
                        cmdNew.Enabled = True
                        cmdSave.Enabled = False
                        cmdApprove.Enabled = False
                        cmdCancel.Enabled = False
                        fraMain.Enabled = False
                    ElseIf Rec!tnyStatus = 4 Then
                        cmdSave.Enabled = False
                        cmdApprove.Enabled = False
                        cmdNew.Enabled = True
                        cmdCancel.Enabled = False
                    ElseIf Rec!tnyStatus = 8 Then
                        cmdSave.Enabled = False
                        cmdApprove.Enabled = False
                        cmdNew.Enabled = True
                        
                        ' Modified On 29 May 2015
                        If val(txtPendingTask.Tag) = 8 And mCancelFlag = True Then
                            cmdCancel.Enabled = True
                        Else
                            cmdCancel.Enabled = False
                        End If
                    End If
                    
            End If
            mCnn.Close
        End If
    End Sub
    
    Private Function POCancelValidation() As Boolean
        Dim objdb         As New clsDB
        Dim mSql          As String
        Dim mCnn          As New ADODB.Connection
        Dim Rec           As New ADODB.Recordset
        Dim mCnt          As Integer
        Dim mPONo           As Variant
        Dim mRev            As Integer
        mPONo = txtInstrumentNo.Text
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select * From faVouchers Where intKeyID2=" & mPONo
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
               Do While Not (Rec.EOF)
                    mRev = IIf(IsNull(Rec!tnyReversed), 0, Rec!tnyReversed)
                    If mRev = 1 Then
                        POCancelValidation = True
                        Exit Do
                    End If
                    POCancelValidation = False
                    Rec.MoveNext
                    
                Loop
            Else
                POCancelValidation = False
            End If
        End If
    End Function
    Private Sub FillGrid()
        Dim objdb         As New clsDB
        Dim mSql          As String
        Dim mCnn          As New ADODB.Connection
        Dim Rec           As New ADODB.Recordset
        Dim mCnt            As Integer
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mSql = "Select *,case When tnyStatus=0 then 'Pending For Approval'"
            mSql = mSql + " When tnyStatus=1 then 'Pending For Approval' "
            mSql = mSql + " When tnyStatus=2 then 'Approved' "
            mSql = mSql + " When tnyStatus=8 then 'Finished' "
            mSql = mSql + " When tnyStatus=4 then 'Cancel/Reject' end as pStatus From faPendingTaskRequest"
            mSql = mSql + " Where intYearID = " & gbFinancialYearID - 1
            If cmbMonth.ListIndex > 0 Then
                mSql = mSql + "  AND Month(dtTransactionDate) = " & cmbMonth.ItemData(cmbMonth.ListIndex) & " "
            End If
            If cmbSearchTask.ListIndex > 0 Then
                mSql = mSql + "  AND intTaskID = " & cmbSearchTask.ItemData(cmbSearchTask.ListIndex) & " "
            End If
            mSql = mSql + " Order By tnyStatus Asc, dtTransactionDate,intTaskID"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            mCnt = 0
            vsGrid.Rows = 1
            If Not (Rec.EOF And Rec.BOF) Then
                While Not (Rec.EOF)
                    
                    mCnt = mCnt + 1
                    If vsGrid.Rows = mCnt Then
                        vsGrid.Rows = vsGrid.Rows + 1
                    End If
                    
                    Select Case Rec!intTaskID
                        Case 1
                            vsGrid.TextMatrix(mCnt, 0) = "Letter of Authority"
                        Case 2
                            vsGrid.TextMatrix(mCnt, 0) = "Cancel Letter of Authority"
                        Case 3
                            vsGrid.TextMatrix(mCnt, 0) = "Requisition"
                        Case 4
                            vsGrid.TextMatrix(mCnt, 0) = "Letter of Allotment (B Fund)"
                        Case 5
                            vsGrid.TextMatrix(mCnt, 0) = "Demand/Receipt"
                        Case 6
                            vsGrid.TextMatrix(mCnt, 0) = "Reverse Entry"
                        Case 7
                            vsGrid.TextMatrix(mCnt, 0) = "Pay Order"
                        Case 8
                            vsGrid.TextMatrix(mCnt, 0) = "Pay Order Cancel"
                        Case 9
                            vsGrid.TextMatrix(mCnt, 0) = "Pay Order Cancellation"
                        Case 10
                            vsGrid.TextMatrix(mCnt, 0) = "Cancel Requisition"
                        Case 11
                            vsGrid.TextMatrix(mCnt, 0) = "Pay Order Approval"
                        Case 12
                            vsGrid.TextMatrix(mCnt, 0) = "OBRP"
                        Case 13
                            vsGrid.TextMatrix(mCnt, 0) = "Requisition B-Fund"
                        Case 14
                            vsGrid.TextMatrix(mCnt, 0) = "Contra Entry"
                        Case 15
                            vsGrid.TextMatrix(mCnt, 0) = "Journal Entry"
                        Case 16
                            vsGrid.TextMatrix(mCnt, 0) = "UnAuthorized Drawal"
                        Case 17
                            vsGrid.TextMatrix(mCnt, 0) = "E-Bill Voucher"
                    End Select
                    
                    vsGrid.TextMatrix(mCnt, 1) = Rec!dtTransactionDate
                    vsGrid.TextMatrix(mCnt, 2) = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    vsGrid.TextMatrix(mCnt, 3) = IIf(IsNull(Rec!vchInstrumentNo), 0, Rec!vchInstrumentNo)
                    vsGrid.TextMatrix(mCnt, 4) = IIf(IsNull(Rec!dtInstrumentDate), 0, Rec!dtInstrumentDate)
                    vsGrid.TextMatrix(mCnt, 5) = Rec!pStatus
                    vsGrid.TextMatrix(mCnt, 8) = Rec!intRequestID
                    vsGrid.TextMatrix(mCnt, 9) = Rec!intTaskID
                    vsGrid.TextMatrix(mCnt, 10) = Rec!tnyStatus
                    If val(vsGrid.TextMatrix(mCnt, 10)) = 2 Then
                        vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 10) = &HE7F9E7
                    ElseIf val(vsGrid.TextMatrix(mCnt, 10)) = 4 Then
                        vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 10) = &H90AAFF
                    ElseIf val(vsGrid.TextMatrix(mCnt, 10)) = 8 Then     '''COMPLETED
                        vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 10) = &HFAEFEE
                    Else
                        vsGrid.Cell(flexcpBackColor, mCnt, 0, mCnt, 10) = &H80000005
                    End If
                    Rec.MoveNext
                    
                Wend
                Rec.Close
            End If
            mCnn.Close
        End If
    End Sub

    Private Sub vsTasks_DblClick()
        If vsTasks.Row > 0 Then
            txtPendingTask.Text = vsTasks.TextMatrix(vsTasks.Row, 1)
            txtPendingTask.Tag = vsTasks.TextMatrix(vsTasks.Row, 0)
            If val(txtPendingTask.Tag) = 1 Or val(txtPendingTask.Tag) = 2 Or val(txtPendingTask.Tag) = 3 Or val(txtPendingTask.Tag) = 4 _
            Or val(txtPendingTask.Tag) = 10 Or val(txtPendingTask.Tag) = 13 Then
                MsgBox "ACR is Blocked because it done through Saankhya Web", vbApplicationModal
                cmdSave.Enabled = False
            ElseIf val(txtPendingTask.Tag) = 7 Or val(txtPendingTask.Tag) = 11 Then
                cmdSave.Enabled = True
            Else
            
'                MsgBox "All transactions in Pending Task is Blocked", vbApplicationModal
'                cmdSave.Enabled = False
            End If
            vsTasks.Visible = False
            cmdKeyID.Enabled = True
            Call SetInputField(val(txtPendingTask.Tag))
            cmdTasks.SetFocus
        End If
    End Sub
    Private Sub GetRequisitionDetails(numProjectID As Variant, intSourceOfFundID As Variant)
        Dim objProj As New clsProject
        Dim objProFund As New clsProjectFund
        Dim mProjectID As Variant
        Dim mSourceOfFundID As Variant
        
        mProjectID = numProjectID
        mSourceOfFundID = intSourceOfFundID
        If val(mProjectID) > 0 Then
            objProj.SetProject mProjectID, gbFinancialYearID - 1
            If objProj.ProjectID > 0 Then
                txtProject.Text = "[" & objProj.ProjectSerialNo & "]" & objProj.ProjectNameEnglish
                txtProject.Tag = objProj.ProjectID
                
                txtCategory.Text = objProj.Category
                txtCategory.Tag = objProj.ProjCatID
                
                txtSourceFund.Text = objProj.FindSourceOfFund(mSourceOfFundID)
                txtSourceFund.Tag = objProj.SourceOfFundID
            End If
        End If
    End Sub
    Private Sub GetBFundRequisitionDetails(mReqID As Integer)
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mReqID > 0 Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = " Select * from faPendingTaskRequest"
                mSql = mSql + " Inner Join suSourceOfFund On suSourceOfFund.intSourceFundID=faPendingTaskRequest.intSourceOfFundID"
                mSql = mSql + " Inner Join faDepSchPro On faDepSchPro.intID = faPendingTaskRequest.intCategoryID"
                mSql = mSql + " Where intRequestID = " & mReqID
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtSourceFund.Visible = True
                    lblCategory.Visible = True
                    lblCategory.Caption = "Scheme"
                    txtCategory.Visible = True
                    txtSourceFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    txtCategory.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Sub
    
    Private Sub GetUnAuthorizedDrawalRequisitionDetails(mReqID As Integer)   '*************UNAUTHORIZED DRAWAL*********************
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mReqID > 0 Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = " Select * from faPendingTaskRequest"
                mSql = mSql + " Inner Join suSourceOfFund On suSourceOfFund.intSourceFundID=faPendingTaskRequest.intSourceOfFundID"
                mSql = mSql + " Inner Join faTransactionCategory On faTransactionCategory.intCategoryID = faPendingTaskRequest.intCategoryID"
                mSql = mSql + " Where intRequestID = " & mReqID
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtSourceFund.Visible = True
                    lblCategory.Visible = True
                    txtCategory.Visible = True
                    txtSourceFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Sub
    
    Private Function CheckCancelRequisitionStatus(mID As Integer)
        Dim mSql        As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mID > 0 Then
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = "Select * From faPendingTaskRequest Where intTaskID = " & val(txtPendingTask.Tag) & "AND intKeyID=" & mID
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    CheckCancelRequisitionStatus = 1
                Else
                    CheckCancelRequisitionStatus = 0
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Function
    Public Function GetLFAStatus()
'        If frmViewBalanceSheet.GetStatus0fLFA = True Then
'            MsgBox "AFS is Submitted to LFA, Further Modification is not Possible", vbInformation
'            cmdTasks.Enabled = False
'            cmdSave.Enabled = False
'            cmdNew.Enabled = False
'        Else
'            cmdTasks.Enabled = True
'            cmdSave.Enabled = True
'        End If
    End Function

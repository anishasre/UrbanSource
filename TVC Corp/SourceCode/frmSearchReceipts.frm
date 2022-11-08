VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSearchReceipts 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Receipts"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   12390
   Icon            =   "frmSearchReceipts.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12390
   Begin VB.TextBox txtcheque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9765
      TabIndex        =   40
      Top             =   2160
      Width           =   1845
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   39
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtAccountHead 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1815
      Width           =   5565
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9765
      TabIndex        =   11
      Top             =   1815
      Width           =   1845
   End
   Begin VB.CommandButton cmdSearchHead 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8565
      TabIndex        =   3
      Top             =   1830
      Width           =   330
   End
   Begin VB.TextBox txtAccountHeadCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1815
      Width           =   1290
   End
   Begin VB.TextBox txtDoorNo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9765
      TabIndex        =   9
      Top             =   1485
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar pbSearch 
      Height          =   225
      Left            =   165
      TabIndex        =   32
      Top             =   2805
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txtBookNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6735
      TabIndex        =   31
      Top             =   1485
      Width           =   915
   End
   Begin VB.Frame fraApplication 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1538
      TabIndex        =   27
      Top             =   90
      Width           =   8745
      Begin VB.OptionButton optBackUp 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Saankhya DE Backup Server"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5715
         TabIndex        =   38
         Top             =   180
         Width           =   2730
      End
      Begin VB.OptionButton optSahatha 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Sahatha"
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
         Left            =   4425
         TabIndex        =   36
         Top             =   210
         Width           =   1005
      End
      Begin VB.OptionButton optSaankhya 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Saankhya"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2925
         TabIndex        =   29
         Top             =   180
         Width           =   1125
      End
      Begin VB.OptionButton optSaankhyaDoubleEntry 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Saankhya Double Entry"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   375
         TabIndex        =   28
         Top             =   180
         Width           =   2280
      End
   End
   Begin VB.TextBox txtDoorNo2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10680
      TabIndex        =   10
      Top             =   1485
      Width           =   930
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -870
      TabIndex        =   26
      Top             =   6570
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtWard 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9765
      TabIndex        =   8
      Top             =   1155
      Width           =   1845
   End
   Begin VB.ComboBox cmbSeat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1695
      TabIndex        =   2
      Top             =   1485
      Width           =   3315
   End
   Begin VB.TextBox txtVoucherNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6735
      TabIndex        =   6
      Top             =   1485
      Width           =   1830
   End
   Begin VB.ComboBox cmbTransactionType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1695
      TabIndex        =   1
      Top             =   1155
      Width           =   3315
   End
   Begin VB.ComboBox cmbDepartment 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1695
      TabIndex        =   0
      Top             =   825
      Width           =   3315
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9765
      TabIndex        =   7
      Top             =   825
      Width           =   1845
   End
   Begin VB.TextBox txtToDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6735
      TabIndex        =   5
      Top             =   1155
      Width           =   1830
   End
   Begin VB.TextBox txtFromDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6735
      TabIndex        =   4
      Top             =   825
      Width           =   1830
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   11685
      Top             =   6915
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   8580
      TabIndex        =   14
      Top             =   810
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      _Version        =   393216
      Format          =   66453505
      CurrentDate     =   39697
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   8580
      TabIndex        =   15
      Top             =   1155
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      _Version        =   393216
      Format          =   66453505
      CurrentDate     =   39698
   End
   Begin VSFlex8LCtl.VSFlexGrid vsDetails 
      Height          =   3510
      Left            =   165
      TabIndex        =   16
      Top             =   3705
      Width           =   11985
      _cx             =   21140
      _cy             =   6191
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
      BackColor       =   14349042
      ForeColor       =   -2147483640
      BackColorFixed  =   14349042
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
      BackColorAlternate=   14349042
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchReceipts.frx":1CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VSFlex8LCtl.VSFlexGrid vsDetailsOld 
      Height          =   3555
      Left            =   165
      TabIndex        =   30
      Top             =   3375
      Width           =   12045
      _cx             =   21246
      _cy             =   6271
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
      BackColor       =   14349042
      ForeColor       =   -2147483640
      BackColorFixed  =   14349042
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
      BackColorAlternate=   14349042
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchReceipts.frx":1E3B
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VSFlex8LCtl.VSFlexGrid vsDetailsSahatha 
      Height          =   3555
      Left            =   165
      TabIndex        =   37
      Top             =   3030
      Width           =   12045
      _cx             =   21246
      _cy             =   6271
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
      BackColor       =   14349042
      ForeColor       =   -2147483640
      BackColorFixed  =   14349042
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
      BackColorAlternate=   14349042
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSearchReceipts.frx":1F58
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Label lblcheque 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No"
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
      Left            =   9030
      TabIndex        =   41
      Top             =   2190
      Width           =   690
   End
   Begin VB.Label lblAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   9030
      TabIndex        =   34
      Top             =   1845
      Width           =   690
   End
   Begin VB.Label lblAccountHead 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head "
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
      Left            =   450
      TabIndex        =   33
      Top             =   1860
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   3870
      Left            =   75
      Top             =   2730
      Width           =   12255
   End
   Begin VB.Shape Shape1 
      Height          =   2640
      Left            =   60
      Top             =   60
      Width           =   12255
   End
   Begin VB.Label lblDoorNo 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Door No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9030
      TabIndex        =   25
      Top             =   1500
      Width           =   690
   End
   Begin VB.Label lblWard 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Ward "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9240
      TabIndex        =   24
      Top             =   1155
      Width           =   465
   End
   Begin VB.Label lblReceiptNo 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Receipt No"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5730
      TabIndex        =   23
      Top             =   1500
      Width           =   2040
   End
   Begin VB.Label lblSeatName 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Seat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1260
      TabIndex        =   22
      Top             =   1545
      Width           =   375
   End
   Begin VB.Label lblTransactionType 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Transaction Type"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   21
      Top             =   1170
      Width           =   1515
   End
   Begin VB.Label lblSection 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   990
      TabIndex        =   20
      Top             =   825
      Width           =   645
   End
   Begin VB.Label lblToDate 
      BackColor       =   &H00DAF2F2&
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5970
      TabIndex        =   19
      Top             =   1155
      Width           =   690
   End
   Begin VB.Label lblFromDate 
      BackColor       =   &H00DAF2F2&
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5745
      TabIndex        =   18
      Top             =   825
      Width           =   915
   End
   Begin VB.Label lblName 
      BackColor       =   &H00DAF2F2&
      Caption         =   "Name"
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
      Left            =   9225
      TabIndex        =   17
      Top             =   825
      Width           =   495
   End
End
Attribute VB_Name = "frmSearchReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim gSeatID     As Variant
    
    '*********************************************************************************************'
    '           Form to Search Receipts from DB_Finance, Accounts and Receipts Databases          '
    '*********************************************************************************************'
    Public Sub DisplayReceiptDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim mYearID         As Variant
        Dim mPeriodID       As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
               
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in  DB_Finance (Saankhya Double Entry)'''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error GoTo err
        
        If optSaankhyaDoubleEntry.Value = True Then
            frmReceipt.txtBookNo.Visible = False
            frmReceipt.lblReceiptNo.Caption = "Receipt No"
            frmReceipt.lblReceiptNo.Left = 3960
            frmReceipt.lblReceiptNo.Top = 825
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            
            mSql = "Select *,faVouchers.intVoucherID[VoucherID] From faVouchers"
            mSql = mSql + " Left Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
            mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
            mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
            mSql = mSql + " Left Join faSection On faTransactionType.intSectionID=faSection.intSectionID"
            mSql = mSql + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
            mSql = mSql + " Left Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
            'mSQL = mSQL + " Or  faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID "
            mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
            mSql = mSql + " Where faVouchers.intVoucherID=" & mVoucherNo
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Not IsNull(Rec!tnyCancelFlag) Then
                    If Rec!tnyCancelFlag = 1 Then
                        frmReceipt.lblMessage.Visible = True
                        frmReceipt.lblMessage.Caption = "This is a Cancelled Receipt"
                        frmReceipt.Timer1.Enabled = True
                    End If
                End If
                frmReceipt.txtReceiptNo.Text = vsDetails.TextMatrix(vsDetails.Row, 4)
                frmReceipt.txtReceiptNo.Tag = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                    
                frmReceipt.txtSection.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
                frmReceipt.txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                frmReceipt.txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                frmReceipt.txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                
                frmReceipt.txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                frmReceipt.txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                frmReceipt.txtInstNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                frmReceipt.txtDated.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                frmReceipt.txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                frmReceipt.txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
                
                If IsNull(Rec!chvZoneNameEnglish) = False Then
                    frmReceipt.txtZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
                End If
                frmReceipt.txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                frmReceipt.txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                frmReceipt.txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                frmReceipt.txtRefNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                
                frmReceipt.txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                frmReceipt.txtInit1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
                frmReceipt.txtInit2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
                frmReceipt.txtInit3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
                frmReceipt.txtInit4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
                frmReceipt.txtHouse.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                frmReceipt.txtStreet.Text = IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
                frmReceipt.txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                frmReceipt.txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                frmReceipt.txtPost.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
                frmReceipt.txtPin.Text = IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
                frmReceipt.txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                
                frmReceipt.txtAdvance.Text = IIf(IsNull(Rec!fltAdvAmtAdj), 0, Rec!fltAdvAmtAdj)
                frmReceipt.txtDescription.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                
                mSqlAccHeads = "Select * From faVoucherChild"
                mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                '-----------------------------------------
                'Added By Anisha On 30.09.10 to Diplay Period
                mSqlAccHeads = mSqlAccHeads + " left Join faPeriodicity On faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
                '-------------------------------------------
                mSqlAccHeads = mSqlAccHeads + " Where intVoucherID=" & frmReceipt.txtReceiptNo.Tag
                RecAccHeads.Open mSqlAccHeads, mCnn
                mRowCount = 1
                While Not Rec.EOF
                    While Not RecAccHeads.EOF
                        frmReceipt.vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                        frmReceipt.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                        
                        ''''''''''''''''''''''''To be Removed'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If frmReceipt.txtTransactionType.Tag = 12 And RecAccHeads!vchAccountHeadCode = 140130400 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 0) = "140130200"
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 1) = "Fees for Delayed Registration - Birth & DeathCertificate"
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
                        mPeriodID = IIf(IsNull(RecAccHeads!tnyPeriodID), "", RecAccHeads!tnyPeriodID)
                        mYearID = IIf(IsNull(RecAccHeads!intYearID), 0, RecAccHeads!intYearID)
                        If mYearID <> 0 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 2) = mYearID & "-" & mYearID + 1
                        End If
                        
                        '-----------------------------------------
                        'Added By Anisha On 30.09.10 to Diplay Period
'                        If mPeriodID = 1 Then
'                            frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = "1st Half"
'                        End If
'                        If mPeriodID = 2 Then
'                            frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = "2nd Half"
'                        End If
'                        If mPeriodID = 3 Then
'                            frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = "Full Year"
'                        End If
                        
                        frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchPeriodicity), "", RecAccHeads!vchPeriodicity)
                        '--------------------------------------------------------
                        mArrearFlag = IIf(IsNull(RecAccHeads!tnyArrearFlag), "", RecAccHeads!tnyArrearFlag)
                        If mArrearFlag = 0 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        End If
                        If mArrearFlag = 1 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        End If
                        frmReceipt.vsGrid.Rows = frmReceipt.vsGrid.Rows + 1
                        mRowCount = mRowCount + 1
                        RecAccHeads.MoveNext
                    Wend
                    Rec.MoveNext
                Wend
                RecAccHeads.Close
                Call Calculate
            End If
            mCnn.Close
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in DB_Accounts (Saankhya)''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        
        If optSaankhya.Value = True Then
            frmReceipt.txtBookNo.Visible = True
            frmReceipt.lblReceiptNo.Caption = "Book/ReceiptNo"
            frmReceipt.lblReceiptNo.Left = 3510
            frmReceipt.lblReceiptNo.Top = 825
            frmReceipt.txtDoorNo1.Width = 1785
            If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaOld)) Then
            
                mSql = "Select * From TblReceipt"
                mSql = mSql + " Left Join TB_DetailedHead_MST On TblReceipt.Tohead=TB_DetailedHead_MST.intDetailedHead_ID"
                mSql = mSql + " Where TblReceipt.Id=" & mVoucherNo
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    frmReceipt.txtAccountHead.Text = IIf(IsNull(Rec!chvHead), "", Rec!chvHead)
                End If
                Rec.Close
                
                mSql = "Select * From TblReceipt"
                mSql = mSql + " Left Join TB_Department_MST On TblReceipt.DepartmentId=TB_Department_MST.intDeptId"
                mSql = mSql + " Left Join TB_InstrumentType_MST On TblReceipt.InstrumentType=TB_InstrumentType_MST.intInstrumentTypeID"
                mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
                mSql = mSql + " Left Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
                mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                mSql = mSql + " Where TblReceipt.Id =" & mVoucherNo
                mSql = mSql + " Order By TB_Transaction_MST.PeriodYear"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    frmReceipt.txtBookNo.Text = IIf(IsNull(Rec!BookNo), "", Rec!BookNo)
                    frmReceipt.txtReceiptNo.Text = IIf(IsNull(Rec!ReceiptNO), "", Rec!ReceiptNO)
                    frmReceipt.txtReceiptNo.Tag = mVoucherNo
                        
                    frmReceipt.txtSection.Text = IIf(IsNull(Rec!chvDeptName), "", Rec!chvDeptName)
                    frmReceipt.txtTransactionType.Text = IIf(IsNull(Rec!chvHead), "", Rec!chvHead)
                    frmReceipt.txtDate.Text = IIf(IsNull(Rec!ReceiptDate), "", Rec!ReceiptDate)
                    
    '                frmReceipt.txtAccountHead.Text = IIf(IsNull(Rec!chvHead), "", Rec!chvHead)
                    frmReceipt.txtInstrument.Text = IIf(IsNull(Rec!chvInstrumentType), "", Rec!chvInstrumentType)
                    frmReceipt.txtInstNo.Text = IIf(IsNull(Rec!InstrumentNo), "", Rec!InstrumentNo)
                    frmReceipt.txtDated.Text = IIf(IsNull(Rec!InstrumentDate), "", Rec!InstrumentDate)
    '                frmReceipt.txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
    '                frmReceipt.txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
                    
    '                If IsNull(Rec!chvZoneNameEnglish) = False Then
    '                    frmReceipt.txtZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
    '                End If
                    frmReceipt.txtWardNo.Text = IIf(IsNull(Rec!wardNo), "", Rec!wardNo)
                    frmReceipt.txtDoorNo1.Text = IIf(IsNull(Rec!HouseNo), "", Rec!HouseNo)
    '                frmReceipt.txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
    '                frmReceipt.txtRefNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                    
                    frmReceipt.txtName.Text = IIf(IsNull(Rec!Payee), "", Rec!Payee)
                    frmReceipt.txtInit1.Text = IIf(IsNull(Rec!PayeeIni1), "", Rec!PayeeIni1)
                    frmReceipt.txtInit2.Text = IIf(IsNull(Rec!PayeeIni2), "", Rec!PayeeIni2)
                    frmReceipt.txtInit3.Text = IIf(IsNull(Rec!PayeeIni3), "", Rec!PayeeIni3)
                    frmReceipt.txtInit4.Text = IIf(IsNull(Rec!PayeeIni4), "", Rec!PayeeIni4)
                    frmReceipt.txtHouse.Text = IIf(IsNull(Rec!HouseName), "", Rec!HouseName)
                    frmReceipt.txtStreet.Text = IIf(IsNull(Rec!Address), "", Rec!Address)
                    frmReceipt.txtLocalPlace.Text = IIf(IsNull(Rec!LocalPlace), "", Rec!LocalPlace)
                    frmReceipt.txtMainPlace.Text = IIf(IsNull(Rec!MainPlace), "", Rec!MainPlace)
    '                frmReceipt.txtPost.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
    '                frmReceipt.txtPin.Text = IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
    '                frmReceipt.txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    
                    frmReceipt.txtDescription.Text = IIf(IsNull(Rec!Narration), "", Rec!Narration)
                    
                    mSqlAccHeads = "Select * From TblReceipt"
                    mSqlAccHeads = mSqlAccHeads + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
                    mSqlAccHeads = mSqlAccHeads + " Left Join TB_DetailedHead_MST On TB_Transaction_MST.HeadId=TB_DetailedHead_MST.intDetailedHead_ID"
                    mSqlAccHeads = mSqlAccHeads + " Left Join TblYear On TB_Transaction_MST.PeriodYear=TblYear.Id"
                    mSqlAccHeads = mSqlAccHeads + " Left Join TB_PeriodType_MST On TB_Transaction_MST.Period=TB_PeriodType_MST.intPeriodId"
                    mSqlAccHeads = mSqlAccHeads + " Where TblReceipt.Id=" & frmReceipt.txtReceiptNo.Tag
                    mSqlAccHeads = mSqlAccHeads + " Order By TB_Transaction_MST.PeriodYear"
                    RecAccHeads.Open mSqlAccHeads, mCnn
                    mRowCount = 1
                    While Not Rec.EOF
                        While Not RecAccHeads.EOF
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!intNumCode), "", RecAccHeads!intNumCode)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!chvHead), "", RecAccHeads!chvHead)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!Year), "", RecAccHeads!Year)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!chvPeriod), "", RecAccHeads!chvPeriod)
                            mArrearFlag = IIf(IsNull(RecAccHeads!intArrearFlag), "", RecAccHeads!intArrearFlag)
                            If mArrearFlag = 0 Then
                                frmReceipt.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!CrAmount), "", RecAccHeads!CrAmount)
                            End If
                            If mArrearFlag = 1 Then
                                frmReceipt.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!CrAmount), "", RecAccHeads!CrAmount)
                            End If
                            frmReceipt.vsGrid.Rows = frmReceipt.vsGrid.Rows + 1
                            mRowCount = mRowCount + 1
                            RecAccHeads.MoveNext
                        Wend
                        Rec.MoveNext
                    Wend
                    RecAccHeads.Close
                    Call Calculate
                End If
                Rec.Close
                mCnn.Close
            Else
                MsgBox "Connection To Accounts does not exit, Please contact your System Administrator", vbInformation
            End If
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in Receipts (Sahatha)'''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        
        If optSahatha.Value = True Then
            frmReceipt.txtBookNo.Visible = True
            frmReceipt.lblReceiptNo.Caption = "Book/ReceiptNo"
            frmReceipt.lblReceiptNo.Left = 3510
            frmReceipt.lblReceiptNo.Top = 825
            frmReceipt.txtDoorNo1.Width = 1785
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Sahatha)) Then
    
                mSql = "Select * From TblReceipt"
                mSql = mSql + " Left Join tblAccountHeads On TblReceipt.Tohead=tblAccountHeads.ID"
                mSql = mSql + " Where TblReceipt.Id=" & mVoucherNo
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    frmReceipt.txtAccountHead.Text = IIf(IsNull(Rec!Head), "", Rec!Head)
                End If
                Rec.Close
                
                If cmbTransactionType.ListIndex = -1 Then cmbTransactionType.ListIndex = 0
                If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 1 Then
                    mSql = "Select * From TblReceipt"
                    mSql = mSql + " Inner Join tblReceiptBuildings On tblReceipt.Id=tblReceiptBuildings.ReceiptID"
                    mSql = mSql + " Inner Join Department On TblReceipt.DepartmentID=Department.ID"
                    mSql = mSql + " Inner Join TblAccountHeads On TblReceipt.ToHead=TblAccountHeads.Id"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
    '                mSQL = mSQL + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
                    mSql = mSql + " Where TblReceipt.ID = " & mVoucherNo
                Else
                    mSql = "Select * From TblReceipt"
                    mSql = mSql + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ID"
                    mSql = mSql + " Inner Join Department On TblReceipt.DepartmentID=Department.ID"
                    mSql = mSql + " Inner Join TblAccountHeads On TblReceipt.ToHead=TblAccountHeads.Id"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                    mSql = mSql + " Where TblReceipt.ID = " & mVoucherNo
                End If
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    frmReceipt.txtBookNo.Text = IIf(IsNull(Rec!BookNo), "", Rec!BookNo)
                    frmReceipt.txtReceiptNo.Text = IIf(IsNull(Rec!ReceiptNO), "", Rec!ReceiptNO)
                    frmReceipt.txtReceiptNo.Tag = mVoucherNo
    
                    frmReceipt.txtSection.Text = IIf(IsNull(Rec!Name), "", Rec!Name)
    '                frmReceipt.txtTransactionType.Text = IIf(IsNull(Rec!chvHead), "", Rec!chvHead)
                    frmReceipt.txtDate.Text = IIf(IsNull(Rec!ReceiptDate), "", Rec!ReceiptDate)
    
    '                frmReceipt.txtAccountHead.Text = IIf(IsNull(Rec!chvHead), "", Rec!chvHead)
    '                frmReceipt.txtInstrument.Text = IIf(IsNull(Rec!chvInstrumentType), "", Rec!chvInstrumentType)
    '                frmReceipt.txtInstNo.Text = IIf(IsNull(Rec!InstrumentNo), "", Rec!InstrumentNo)
    '                frmReceipt.txtDated.Text = IIf(IsNull(Rec!InstrumentDate), "", Rec!InstrumentDate)
    '                frmReceipt.txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
    '                frmReceipt.txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
    
    '                If IsNull(Rec!chvZoneNameEnglish) = False Then
    '                    frmReceipt.txtZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
    '                End If
                    frmReceipt.txtWardNo.Text = IIf(IsNull(Rec!wardNo), "", Rec!wardNo)
                    frmReceipt.txtDoorNo1.Text = IIf(IsNull(Rec!HouseNo), "", Rec!HouseNo)
    '                frmReceipt.txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
    '                frmReceipt.txtRefNo.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
    
                    frmReceipt.txtName.Text = IIf(IsNull(Rec!Payee), "", Rec!Payee)
                    frmReceipt.txtInit1.Text = IIf(IsNull(Rec!PayeeIni1), "", Rec!PayeeIni1)
                    frmReceipt.txtInit2.Text = IIf(IsNull(Rec!PayeeIni2), "", Rec!PayeeIni2)
                    frmReceipt.txtInit3.Text = IIf(IsNull(Rec!PayeeIni3), "", Rec!PayeeIni3)
                    frmReceipt.txtInit4.Text = IIf(IsNull(Rec!PayeeIni4), "", Rec!PayeeIni4)
                    frmReceipt.txtHouse.Text = IIf(IsNull(Rec!HouseName), "", Rec!HouseName)
                    frmReceipt.txtStreet.Text = IIf(IsNull(Rec!Address), "", Rec!Address)
                    frmReceipt.txtLocalPlace.Text = IIf(IsNull(Rec!LocalPlace), "", Rec!LocalPlace)
                    frmReceipt.txtMainPlace.Text = IIf(IsNull(Rec!MainPlace), "", Rec!MainPlace)
    '                frmReceipt.txtPost.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
    '                frmReceipt.txtPin.Text = IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
    '                frmReceipt.txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 1 Then
                        frmReceipt.txtDescription.Text = IIf(IsNull(Rec!Description), "", Rec!Description)
                    Else
                        frmReceipt.txtDescription.Text = IIf(IsNull(Rec!Narration), "", Rec!Narration)
                    End If
                    
                    mSqlAccHeads = "Select * From TblReceipt"
                    mSqlAccHeads = mSqlAccHeads + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
                    mSqlAccHeads = mSqlAccHeads + " Left Join TblAccountHeads On TblReceiptChild.HeadId=TblAccountHeads.ID"
                    mSqlAccHeads = mSqlAccHeads + " Left Join TblYear On TblReceiptChild.PeriodYear=TblYear.Id"
                    mSqlAccHeads = mSqlAccHeads + " Where TblReceipt.Id=" & frmReceipt.txtReceiptNo.Tag
    '                mSqlAccHeads = mSqlAccHeads + " Order By TB_Transaction_MST.PeriodYear"
                    RecAccHeads.Open mSqlAccHeads, mCnn
                    mRowCount = 1
                    While Not Rec.EOF
                        While Not RecAccHeads.EOF
    '                        frmReceipt.vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!intNumCode), "", RecAccHeads!intNumCode)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!Head), "", RecAccHeads!Head)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!Year), "", RecAccHeads!Year)
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!Period), "", RecAccHeads!Period)
                            mArrearFlag = IIf(IsNull(RecAccHeads!ArrearFlg), "", RecAccHeads!ArrearFlg)
                            If mArrearFlag = 0 Then
                                frmReceipt.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!Amount), "", RecAccHeads!Amount)
                            End If
                            If mArrearFlag = 1 Then
                                frmReceipt.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!Amount), "", RecAccHeads!Amount)
                            End If
                            frmReceipt.vsGrid.Rows = frmReceipt.vsGrid.Rows + 1
                            mRowCount = mRowCount + 1
                            RecAccHeads.MoveNext
                        Wend
                        Rec.MoveNext
                    Wend
                    RecAccHeads.Close
                    Call Calculate
                End If
                Rec.Close
                mCnn.Close
            Else
                MsgBox "Connection To Receipts does not exit, Please contact your System Administrator", vbInformation
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Public Sub Calculate()
        Dim mAmtArrear As Double
        Dim mAmtCurrent As Double
        Dim mCount As Long
        For mCount = 1 To frmReceipt.vsGrid.Rows - 1
            If val(frmReceipt.vsGrid.TextMatrix(mCount, 4)) Then
                mAmtArrear = mAmtArrear + val(frmReceipt.vsGrid.Cell(flexcpText, mCount, 4))
            Else
                mAmtCurrent = mAmtCurrent + val(frmReceipt.vsGrid.Cell(flexcpText, mCount, 5))
            End If
        Next
        frmReceipt.txtTotalArrear.Text = Format(mAmtArrear, "0.00")
        frmReceipt.txtTotalCurrent.Text = Format(mAmtCurrent, "0.00")
        frmReceipt.txtTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
        frmReceipt.txtRoundOff.Text = Format(RoundOffAdjustment(val(frmReceipt.txtTotal)), "0.00")
        frmReceipt.txtTotal.Text = Format(val(frmReceipt.txtTotal) + val(frmReceipt.txtRoundOff) - val(frmReceipt.txtAdvance), "0.00")
        
''''        If optSaankhyaDoubleEntry.Value = True Then
''''            If Val(frmReceipt.txtTotal.Text) <> vsDetails.TextMatrix(vsDetails.Row, 7) Then
''''                frmReceipt.lblMessage.Visible = True
''''                frmReceipt.lblMessage.Caption = "The Transaction is not Correct"
''''                frmReceipt.Timer1.Enabled = True
''''            Else
''''                frmReceipt.lblMessage.Visible = False
''''            End If
''''        ElseIf optSaankhya.Value = True Then
''''            If Val(frmReceipt.txtTotal.Text) <> vsDetailsOld.TextMatrix(vsDetailsOld.Row, 7) Then
''''                frmReceipt.lblMessage.Visible = True
''''                frmReceipt.lblMessage.Caption = "The Transaction is not Correct"
''''                frmReceipt.Timer1.Enabled = True
''''            Else
''''                frmReceipt.lblMessage.Visible = False
''''            End If
''''        ElseIf optSahatha.Value = True Then
''''            If Val(frmReceipt.txtTotal.Text) <> vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 7) Then
''''                frmReceipt.lblMessage.Visible = True
''''                frmReceipt.lblMessage.Caption = "The Transaction is not Correct"
''''                frmReceipt.Timer1.Enabled = True
''''            Else
''''                frmReceipt.lblMessage.Visible = False
''''            End If
''''        End If
'        If Val(frmReceipt.txtAdvance.Text) Then
'            frmReceipt.txtAdvance.Visible = True
'            frmReceipt.lblAdvance.Visible = True
'        Else
'            frmReceipt.txtAdvance.Visible = False
'            frmReceipt.txtAdvance.Visible = False
'            frmReceipt.txtAdvance.Text = ""
'        End If
    End Sub

    Private Sub FormInitialize()
        Dim mCrl As Control
        
        vsDetails.Clear 1, 1
        vsDetailsOld.Clear 1, 1
        vsDetailsSahatha.Clear 1, 1
        vsDetails.Rows = 1
        vsDetailsOld.Rows = 1
        vsDetailsSahatha.Rows = 1
        For Each mCrl In Me
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
    End Sub
    
    Private Sub FillvsDetails(Rec As ADODB.Recordset, mCount As Double)
        Dim mRowCount       As Double
        Dim mSerialNo       As Double
        Dim mID             As Double
        Dim mAmount         As Variant
        Dim mNum            As Variant
        Dim mVoucherID      As Double
        Dim mLoop           As Integer
        
        On Error GoTo err
        mRowCount = 1
        mSerialNo = 1
        frmSearchReceipts.Enabled = False
        If optSaankhyaDoubleEntry.Value = True Then
            mAmount = 0
            pbSearch.Value = 0
            vsDetails.Clear 1, 1
            vsDetails.Rows = 1
            vsDetails.Rows = mCount + 2
            If mCount <> 0 Then
                pbSearch.Max = mCount
            End If
            While Not Rec.EOF
'                If (Rec!intVoucherID = mVoucherID) Then
'                    vsDetails.Rows = vsDetails.Rows - 1
'                    GoTo MOV
'                End If
                
                vsDetails.TextMatrix(mRowCount, 0) = mSerialNo
                vsDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                vsDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intWardNo) Or Rec!intWardNo = 0, "", Rec!intWardNo)
                vsDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!intDoorNo) Or Rec!intDoorNo = 0, "", Rec!intDoorNo) & "   " & IIf(IsNull(Rec!vchDoorNo2) Or Rec!vchDoorNo2 = 0, "", Rec!vchDoorNo2)
                vsDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                vsDetails.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!dtDate), "", CheckDateInMMM(Rec!dtDate))
                vsDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                vsDetails.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltAmount), "", Format(Rec!fltAmount, "0.00"))
                vsDetails.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                vsDetails.TextMatrix(mRowCount, 10) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                '*************MODIFIED BY SABEEN*************************
                vsDetails.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!tnyCancelFlag), "", Rec!tnyCancelFlag)
                If Rec!tnyCancelFlag = 1 Then
                       For mLoop = 0 To vsDetails.Cols - 1
                            vsDetails.Cell(flexcpBackColor, mRowCount, mLoop) = &HE0E0E0
                        Next mLoop
                End If
                '*********************************************************
                
                
                
                
                mVoucherID = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                mAmount = mAmount + CDbl(vsDetails.TextMatrix(mRowCount, 7))
                mRowCount = mRowCount + 1
                mSerialNo = mSerialNo + 1
                Rec.MoveNext
                If pbSearch.Value < pbSearch.Max + 1 Then
                    pbSearch.Value = pbSearch.Value + 1
                End If
            Wend
            vsDetails.MergeRow(mRowCount) = True
            vsDetails.Cell(flexcpAlignment, mRowCount, , , 7) = 7 'Align Right
            vsDetails.TextMatrix(mRowCount, 1) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 2) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 3) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 4) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 5) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 6) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetails.TextMatrix(mRowCount, 7) = "Total Collection -- " + Format(mAmount, "0.00")
            If IsNull(mAmount) Or mAmount = "" Then
                vsDetails.TextMatrix(mRowCount, 7) = Format(0, "0.00")
            End If
            vsDetails.Cell(flexcpBackColor, mRowCount, , , 7) = vbRed
            vsDetails.Cell(flexcpForeColor, mRowCount, , , 7) = vbYellow
            vsDetails.Cell(flexcpFontBold, mRowCount, , , 7) = True
        End If
        
        If optSaankhya.Value = True Then
            mAmount = 0
            pbSearch.Value = 0
            If mCount <> 0 Then
                pbSearch.Max = mCount
            End If
            vsDetailsOld.Clear 1, 1
            vsDetailsOld.Rows = 1
            vsDetailsOld.Rows = mCount + 2
            While Not Rec.EOF
                If (Rec!id) = mID Then
                    vsDetailsOld.Rows = vsDetailsOld.Rows - 1
                    GoTo lbl
                End If
                vsDetailsOld.TextMatrix(mRowCount, 0) = mSerialNo
                vsDetailsOld.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!Payee), "", Rec!Payee)
                vsDetailsOld.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!wardNo), "", Rec!wardNo)
                vsDetailsOld.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!HouseNo), "", Rec!HouseNo)
                vsDetailsOld.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!BookNo), "", Rec!BookNo)
                vsDetailsOld.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!ReceiptNO), "", Rec!ReceiptNO)
                vsDetailsOld.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!ReceiptDate), "", Rec!ReceiptDate)
                vsDetailsOld.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!TotalAmount), "", Format(Rec!TotalAmount, "0.00"))
                vsDetailsOld.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!BookId), "", Rec!BookId)
                'vsDetailsOld.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                mAmount = mAmount + CDbl(vsDetailsOld.TextMatrix(mRowCount, 7))
                mID = IIf(IsNull(Rec!id), "", Rec!id)
                mRowCount = mRowCount + 1
                mSerialNo = mSerialNo + 1
'                vsDetailsOld.Rows = vsDetailsOld.Rows + 1
lbl:            Rec.MoveNext
                If pbSearch.Value < pbSearch.Max + 1 Then
                    pbSearch.Value = pbSearch.Value + 1
                End If
            Wend
            vsDetailsOld.MergeRow(mRowCount) = True
            vsDetailsOld.Cell(flexcpAlignment, mRowCount, , , 7) = 7 'Align Right
            vsDetailsOld.TextMatrix(mRowCount, 1) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 2) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 3) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 4) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 5) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 6) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsOld.TextMatrix(mRowCount, 7) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            If IsNull(mAmount) Or mAmount = "" Then
                vsDetailsOld.TextMatrix(mRowCount, 7) = Format(0, "0.00")
            End If
            vsDetailsOld.Cell(flexcpBackColor, mRowCount, , , 7) = vbRed
            vsDetailsOld.Cell(flexcpForeColor, mRowCount, , , 7) = vbYellow
            vsDetailsOld.Cell(flexcpFontBold, mRowCount, , , 7) = True
        End If
        
        If optSahatha.Value = True Then
            mAmount = 0
            pbSearch.Value = 0
            If mCount <> 0 Then
                pbSearch.Max = mCount
            End If
            vsDetailsSahatha.Clear 1, 1
            vsDetailsSahatha.Rows = 1
            vsDetailsSahatha.Rows = mCount + 2
            While Not Rec.EOF
                vsDetailsSahatha.TextMatrix(mRowCount, 0) = mSerialNo
                vsDetailsSahatha.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!Payee), "", Rec!Payee)
                vsDetailsSahatha.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!wardNo), "", Rec!wardNo)
                vsDetailsSahatha.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!HouseNo), "", Rec!HouseNo)
                vsDetailsSahatha.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!BookNo), "", Rec!BookNo)
                vsDetailsSahatha.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!ReceiptNO), "", Rec!ReceiptNO)
                vsDetailsSahatha.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!ReceiptDate), "", Rec!ReceiptDate)
                vsDetailsSahatha.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!SumAmount), "", Format(Rec!SumAmount, "0.00"))
                vsDetailsSahatha.TextMatrix(mRowCount, 8) = IIf(IsNull(Rec!BookId), "", Rec!BookId)
                'vsDetailsSahatha.TextMatrix(mRowCount, 9) = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                mAmount = mAmount + CDbl(vsDetailsSahatha.TextMatrix(mRowCount, 7))
                mID = IIf(IsNull(Rec!id), "", Rec!id)
                mRowCount = mRowCount + 1
                mSerialNo = mSerialNo + 1
'                vsDetailsOld.Rows = vsDetailsOld.Rows + 1
                Rec.MoveNext
                If pbSearch.Value < pbSearch.Max + 1 Then
                    pbSearch.Value = pbSearch.Value + 1
                End If
            Wend
            vsDetailsSahatha.MergeRow(mRowCount) = True
            vsDetailsSahatha.Cell(flexcpAlignment, mRowCount, , , 7) = 7 'Align Right
            vsDetailsSahatha.TextMatrix(mRowCount, 1) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 2) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 3) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 4) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 5) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 6) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            vsDetailsSahatha.TextMatrix(mRowCount, 7) = "Total Collection -- " + CStr(Format(mAmount, "0.00"))
            
            If IsNull(mAmount) Or mAmount = "" Then
                vsDetailsSahatha.TextMatrix(mRowCount, 7) = Format(0, "0.00")
            End If
            vsDetailsSahatha.Cell(flexcpBackColor, mRowCount, , , 7) = vbRed
            vsDetailsSahatha.Cell(flexcpForeColor, mRowCount, , , 7) = vbYellow
            vsDetailsSahatha.Cell(flexcpFontBold, mRowCount, , , 7) = True
        End If
        frmSearchReceipts.Enabled = True
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmbDepartment_Click()
        Dim mCnn        As New ADODB.Connection
        Dim mSql        As String
        Dim objdb       As New clsDB
'        Dim Rec         As New ADODB.Recordset
'        Dim mcount      As Double
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        '''''''''''''''''''''''''Searching in  DB_Finance (Saankhya Double Entry)'''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
        If optSaankhyaDoubleEntry.Value = True Or optBackUp.Value = True Then
''            If (cmbDepartment.itemData(cmbDepartment.ListIndex) = 8) Then
'                txtWard.Enabled = True
'                txtDoorNo1.Enabled = True
'                txtDoorNo2.Enabled = True
''            Else
''                txtWard.Enabled = False
''                txtDoorNo1.Enabled = False
''                txtDoorNo2.Enabled = False
''            End If
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'            mSql = "Select Count(intVoucherID) as Count From faVouchers"
'            mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'            mSql = mSql + " Where faTransactionType.intSectionID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mcount = Rec!Count
'            End If
'            Rec.Close
'
'            mSql = "Select * From faVouchers "
'            mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
'    '        mSQL = mSQL + " Inner Join faIDemandTBL On faVouchers.intVoucherID=faIDemandTBL.intVoucherID"
'            mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'            mSql = mSql + " Where faTransactionType.intSectionID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
'            mSql = mSql + " Order By dtDate Desc"
'            Rec.Open mSql, mCnn
'            Call FillvsDetails(Rec, mcount)
'            Rec.Close
            If cmbDepartment.ItemData(cmbDepartment.ListIndex) = 99 Then
                mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 Order By vchTransactionType"
            Else
                mSql = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intSectionID='" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "'"
            End If
            PopulateList cmbTransactionType, mSql, , True, , True, enuSourceString.Saankhya
            mCnn.Close
        End If
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        '''''''''''''''''''''''''Searching in DB_Accounts (Saankhya)''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'        If optSaankhya.Value = True Then
'            objDb.CreateNewConnection mCnn, enuSourceString.SaankhyaOld
'            mSql = "Select Count(Id) as Count From TblReceipt"
'            mSql = mSql + " Where DepartmentID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mcount = Rec!Count
'            End If
'            Rec.Close
'
'            mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,TblReceiptBook.BookNo,ReceiptNo,ReceiptDate,TotalAmount,BookId From TblReceipt"
'            mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'            mSql = mSql + " Where DepartmentID='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            mSql = mSql + " Order By ReceiptDate Desc"
'            Rec.Open mSql, mCnn
'            Call FillvsDetails(Rec, mcount)
'            Rec.Close
'            mCnn.Close
'        End If
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        '''''''''''''''''''''''''Searching in Receipts (Sahatha)''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'        If optSahatha.Value = True Then
'            objDb.CreateNewConnection mCnn, enuSourceString.Sahatha
'            mSql = "Select Count(Id) As Count From TblReceipt"
'            mSql = mSql + " Where DepartmentId='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mcount = Rec!Count
'            End If
'            Rec.Close
'
'            mSql = "Select TblReceipt.Id,BookId,Payee,WardNo,HouseNo,TblReceiptBook.BookNo,ReceiptNo,ReceiptDate,Sum(Amount) As SumAmount From TblReceipt"
'            mSql = mSql + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
'            mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'            mSql = mSql + " Where DepartmentId='" & cmbDepartment.itemData(cmbDepartment.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            mSql = mSql + " Group By TblReceipt.Id,BookId,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TblReceiptBook.BookNo"
'            mSql = mSql + " Order By ReceiptDate Desc"
'            Rec.Open mSql, mCnn
'            Call FillvsDetails(Rec, mcount)
'            Rec.Close
'            mCnn.Close
'        End If
    End Sub
    
    Private Sub cmbSeat_KeyPress(KeyAscii As Integer)
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub cmbSeat_LostFocus()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        
        On Error GoTo err
        gSeatID = ""
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select numSeatID From DB_Masters..GL_Seats Where DB_Masters..GL_Seats.chvSeatTitle='" & cmbSeat.Text & "'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            gSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
        End If
        Rec.Close
        mCnn.Close
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub cmbTransactionType_Click()
'        Dim objDb       As New clsDB
'        Dim mCnn        As New ADODB.Connection
'        Dim Rec         As New ADODB.Recordset
'        Dim mSql        As String
''        Dim mRowCount   As Integer
'        Dim mcount      As Double
'        Dim mSerialNo   As Integer
'        Dim mStatus     As Variant
'
''        Call Forminitialize
''        txtWard.Enabled = False
''        txtDoorNo1.Enabled = False
''        txtDoorNo2.Enabled = False
'        mcount = 0
'
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        '''''''''''''''''''''''''Searching in  DB_Finance (Saankhya Double Entry)'''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        If optSaankhyaDoubleEntry.Value = True Then
'            If cmbDepartment.ListIndex = -1 Then
'                MsgBox "Please select the Department", vbInformation
'                cmbDepartment.SetFocus
'                Exit Sub
'            End If
'            objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
''            If (cmbDepartment.itemData(cmbDepartment.ListIndex) = 8) Then
'                txtWard.Enabled = True
'                txtDoorNo1.Enabled = True
'                txtDoorNo2.Enabled = True
''            End If
'
'            mSql = "Select Count(intVoucherID) as Count From faVouchers"
'            mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'            mSql = mSql + " Where faTransactionType.intTransactionTypeID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'            mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mcount = Rec!Count
'            End If
'            Rec.Close
'
'            mSql = "Select * From faVouchers "
'            mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
'    '            mSQL = mSQL + " Left Join faIDemandTBL On faVouchers.intVoucherID=faIDemandTBL.intVoucherID"
'            mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'            mSql = mSql + " Where faTransactionType.intTransactionTypeID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'            mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
'            mSql = mSql + " Order By dtDate Desc"
'            Rec.Open mSql, mCnn
'            Call FillvsDetails(Rec, mcount)
'            Rec.Close
'            mCnn.Close
'        End If
'
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        '''''''''''''''''''''''''Searching in DB_Accounts (Saankhya)''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        If optSaankhya.Value = True Then
'            objDb.CreateNewConnection mCnn, enuSourceString.SaankhyaOld
'            If cmbTransactionType.itemData(cmbTransactionType.ListIndex) = 1 Then
'                txtWard.Enabled = True
'                txtDoorNo1.Enabled = True
''                txtDoorNo2.Enabled = True
'            End If
'
'            If cmbTransactionType.itemData(cmbTransactionType.ListIndex) = 2 Or cmbTransactionType.itemData(cmbTransactionType.ListIndex) = 3 Then
'                txtWard.Enabled = True
'                txtDoorNo1.Enabled = True
''                txtDoorNo2.Enabled = True
'            End If
'
'            mSql = "Select Count(TblReceipt.Id) as Count From TblReceipt"
'            mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
'            mSql = mSql + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
'            mSql = mSql + " Where TB_DetailedHead_MST.intDetailedHead_ID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            Rec.Open mSql, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mcount = Rec!Count
'            End If
'            Rec.Close
'
'            mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TotalAmount,TblReceiptBook.BookNo From TblReceipt"
'            mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
'            mSql = mSql + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
'            mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'            mSql = mSql + " Where TB_DetailedHead_MST.intDetailedHead_ID='" & cmbTransactionType.itemData(cmbTransactionType.ListIndex) & "'"
'            mSql = mSql + " And TblReceipt.CancelFlag<>1"
'            mSql = mSql + " Order By TblReceipt.Id  Asc , ReceiptDate Desc"
'            Rec.Open mSql, mCnn
'            Call FillvsDetails(Rec, mcount)
'            Rec.Close
'            mCnn.Close
'        End If
'
'        If optSahatha.Value = True Then
'            If cmbTransactionType.itemData(cmbTransactionType.ListIndex) = 1 Then
'                txtWard.Enabled = True
'                txtDoorNo1.Enabled = True
'            End If
'        End If
    End Sub
    
    Private Sub cmdPrint_Click()
        
         '***************ADDED BY SABEEN***************************
        
        Dim aryIn As Variant
        Dim mDepartment         As String
        Dim mTransactionType    As String
        Dim mAccountHeadID      As String
        Dim mFromDate           As String
        Dim mToDate             As String
        Dim mName               As String
        Dim mSeatID             As String
        Dim mBookId             As String
        Dim mVoucherNo          As String
        Dim mAmount             As String
        Dim mWard               As String
        Dim mDoorNo1            As String
        Dim mDoorNo2            As String
        
            If cmbSeat.Text = "" Then
                mSeatID = "%"
            Else
                mSeatID = gSeatID
            End If
            
            If cmbDepartment.ListIndex < 1 Then
                mDepartment = "%"
            Else
                mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
            End If
            
            If cmbTransactionType.ListIndex < 1 Then
                mTransactionType = "%"
            Else
                mTransactionType = CStr(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
            End If
            
            If txtAccountHeadCode.Text = "" Then
                mAccountHeadID = "%"
            Else
                mAccountHeadID = txtAccountHeadCode.Tag
            End If
            
            If txtVoucherNo.Text = "" Then
                mVoucherNo = "%"
            Else
                mVoucherNo = CStr(txtVoucherNo.Text)
        
            End If
            
            If txtAmount.Text = "" Then
                mAmount = "%"
            Else
                mAmount = CStr(txtAmount.Text)
            End If
            
            If txtName.Text = "" Then
                mName = "%"
            Else
                mName = "%" + CStr(txtName.Text) + "%"
            End If
            
            If txtDoorNo1.Text <> "" Then
                 mDoorNo1 = CStr(txtDoorNo1.Text)
            Else
                mDoorNo1 = "%"
                
            End If
            
            If txtDoorNo2.Text <> "" Then
                mDoorNo2 = CStr(txtDoorNo2.Text)
            Else
                mDoorNo2 = "%"
            End If
            
            If txtWard.Text = "" Then
                mWard = "%"
            Else
                mWard = CStr(txtWard.Text)
            End If
            If txtFromDate.Text = "" Then
                MsgBox "Please Select From Date", vbApplicationModal
                Exit Sub
            End If
            If txtToDate.Text = "" Then
                MsgBox "Please Select To Date", vbApplicationModal
                Exit Sub
            End If
        If txtFromDate.Text <> "" And txtToDate.Text <> "" Then
           
            aryIn = Array(CDate(txtFromDate.Text), CDate(txtToDate.Text), mSeatID, mDepartment, mTransactionType, mVoucherNo, mAmount, mDoorNo1, mDoorNo2, mWard, mName)
            frmViewVoucher.ArrayIn = aryIn
            frmViewVoucher.FormName = "frmSearchReceipts"
            frmViewVoucher.Show vbModal
        End If
        '*********************************************************
    End Sub

    Private Sub cmdsearch_Click()
        Dim objdb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim mCnnBkUp            As New ADODB.Connection
        Dim mSql                As String
        Dim Rec                 As New ADODB.Recordset
        Dim mCount              As Double
        Dim mDepartment         As String
        Dim mTransactionType    As String
        Dim mAccountHeadID      As String
        Dim mFromDate           As String
        Dim mToDate             As String
        Dim mName               As String
        Dim mSeatID             As String
        Dim mBookId             As String
        Dim mVoucherNo          As String
        Dim mAmount             As String
        Dim mWard               As String
        Dim mDoorNo1            As String
        Dim mDoorNo2            As String
        Dim mChequeNo           As String
        On Error GoTo err
        If cmbDepartment.ListIndex = -1 Then cmbDepartment.ListIndex = 0
        If cmbTransactionType.ListIndex = -1 Then cmbTransactionType.ListIndex = 0

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in  DB_Finance (Saankhya Double Entry)'''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        vsDetails.Clear 1, 1
        vsDetailsOld.Clear 1, 1
        vsDetailsSahatha.Clear 1, 1
        vsDetails.Rows = 1
        vsDetailsOld.Rows = 1
        vsDetailsSahatha.Rows = 1
        
        If optSaankhyaDoubleEntry.Value = True Or optBackUp.Value = True Then
            
            If optSaankhyaDoubleEntry.Value = True Then
                objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            ElseIf optBackUp.Value = True Then
                objdb.CreateNewConnection mCnn, enuSourceString.SaankhyaBackUp
            End If
  
            vsDetails.Clear 1, 1
            vsDetails.Rows = 1
            
            If cmbSeat.Text = "" Then
                mSeatID = "%"
            Else
                mSeatID = gSeatID
            End If
            
            If cmbDepartment.ListIndex < 1 Then
                mDepartment = "%"
            Else
                mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
            End If
            
            If cmbTransactionType.ListIndex < 1 Then
                mTransactionType = "%"
            Else
                mTransactionType = CStr(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
            End If
            
            If txtAccountHeadCode.Text = "" Then
                mAccountHeadID = "%"
            Else
                mAccountHeadID = txtAccountHeadCode.Tag
            End If
            
            If txtFromDate.Text = "" Then
                If txtVoucherNo.Text = "" Then
                    mSql = "Select dtStartingDate From  faFinancialYear Where tinCurrentFinancialYearFlag=1"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mFromDate = IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
                    End If
                    Rec.Close
                Else
                
                End If
            Else
                mFromDate = txtFromDate.Text
            End If
            
            If txtToDate.Text = "" Then
                If txtVoucherNo.Text = "" Then
                    mSql = "Select dtEndingDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mToDate = IIf(IsNull(Rec!dtEndingDate), "", CheckDateInMMM(Rec!dtEndingDate))
                    End If
                    Rec.Close
                Else
                End If
            Else
                mToDate = txtToDate.Text
            End If
            
            If txtVoucherNo.Text = "" Then
                mVoucherNo = "%"
            Else
                mVoucherNo = CStr(txtVoucherNo.Text)
                If txtVoucherNo.Text = "" Then

                mVoucherNo = "%"

            Else
                mVoucherNo = CStr(txtVoucherNo.Text)
'''                mSql = "Select  top 1 dtStartingdate From faFinancialYear Order By intFinancialYearId Asc"
'''                Rec.Open mSql, mCnn
'''                If Not (Rec.EOF And Rec.BOF) Then
'''                    mFromDate = IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
'''                End If
'''                Rec.Close
            End If

            End If
            
            If txtAmount.Text = "" Then
                mAmount = "%"
            Else
                mAmount = CStr(txtAmount.Text)
'''                mSql = "Select  top 1 dtStartingdate From faFinancialYear Order By intFinancialYearId Asc"
'''                Rec.Open mSql, mCnn
'''                If Not (Rec.EOF And Rec.BOF) Then
'''                    mFromDate = IIf(IsNull(Rec!dtStartingDate), "", CheckDateInMMM(Rec!dtStartingDate))
'''                End If
'''                Rec.Close
            End If
            
            If txtName.Text = "" Then
                mName = ""
            Else
                mName = CStr(txtName.Text)
            End If
            If txtcheque.Text = "" Then
                mChequeNo = ""
            Else
                mChequeNo = CStr(txtcheque.Text)
            End If
        
'            If (cmbDepartment.itemData(cmbDepartment.ListIndex) = 8) Then
'                mSQL = "Select Count(faVouchers.intVoucherID) as Count From faVouchers  "
'                mSQL = mSQL + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID "
''                If txtAmount.Text = "" Then
''                    mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
''                End If
'                mSQL = mSQL + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'                If mAccountHeadID <> "%" Then
'                    mSQL = mSQL + " Inner Join faVoucherChild On faVouchers.intVoucherID = faVoucherChild.intVoucherID"
'                    mSQL = mSQL + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
'                End If
'                mSQL = mSQL + " Where faVouchers.numSeatID LIKE '" & mSeatID & "'"
'                mSQL = mSQL + " And faTransactionType.intSectionID LIKE '" & mDepartment & "'"
'                mSQL = mSQL + " And faVouchers.intTransactionTypeID LIKE '" & mTransactionType & "'"
''                If txtAmount.Text = "" Then
''                    mSql = mSql + " And faVoucherChild.intAccountHeadID LIKE '" & mAccountHeadID & "'"
''                End If
'                mSQL = mSQL + " And dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                mSQL = mSQL + " And intVoucherNo LIKE '" & mVoucherNo & "'"
'                mSQL = mSQL + " And faVouchers.fltAmount LIKE '" & mAmount & "'"
'                mSQL = mSQL + " And faVoucherAddress.vchName LIKE '" & "%" & mName & "%" & "'"
'                'mSql = mSql + " And faVoucherAddress.intWardNo LIKE '" & "%" & mWard & "%" & "'"
'                'mSQL = mSQL + " And faVoucherAddress.intDoorNo LIKE '" & "%" & mDoorNo1 & "%" & "'"
'                If txtWard.Text <> "" Then
'                    mSQL = mSQL + " And faVoucherAddress.intWardNo =" & txtWard.Text
'                End If
'                If txtDoorNo1.Text <> "" Then
'                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
'                        mSQL = mSQL + " And faVoucherAddress.intDoorNo LIKE '" & txtDoorNo1.Text & "%" & "'"
'                    Else
'                        mSQL = mSQL + " And faVoucherAddress.intDoorNo =" & txtDoorNo1.Text
'                    End If
'                End If
'                If txtDoorNo2.Text <> "" Then
'                    If mID(txtDoorNo2.Text, 1, 1) = "%" Then
'                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 LIKE '" & txtDoorNo2.Text & "%" & "'"
'                    Else
'                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 ='" & txtDoorNo2.Text & "'"
'                    End If
'                End If
'                If mAccountHeadID <> "%" Then
'                    mSQL = mSQL + " And faVoucherChild.intAccountHeadID =" & mAccountHeadID
'                End If
'                mSQL = mSQL + " And faVouchers.tnyCancelFlag<>1"
''            Else
''                mSql = "Select Count(faVouchers.intVoucherID) as Count From faVouchers  "
''                mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID "
''                mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'''                If txtAmount.Text = "" Then
'''                    mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
'''                End If
''                mSql = mSql + " Where faVouchers.numSeatID LIKE '" & mSeatID & "'"
''                mSql = mSql + " And faTransactionType.intSectionID LIKE '" & mDepartment & "'"
''                mSql = mSql + " And faVouchers.intTransactionTypeID LIKE '" & mTransactionType & "'"
'''                If txtAmount.Text = "" Then
'''                    mSql = mSql + " And faVoucherChild.intAccountHeadID LIKE '" & mAccountHeadID & "'"
'''                End If
''                mSql = mSql + " And dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
''                mSql = mSql + " And intVoucherNo LIKE '" & mVoucherNo & "'"
''                mSql = mSql + " And faVouchers.fltAmount LIKE '" & mAmount & "'"
''                mSql = mSql + " And faVoucherAddress.vchName LIKE '" & "%" & mName & "%" & "'"
''                mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
''            End If
'            Rec.Open mSQL, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mCount = Rec!count
'            End If
'            Rec.Close
            
'            If (cmbDepartment.itemData(cmbDepartment.ListIndex) = 8) Then
                mSql = "Select *,faVouchers.intVoucherID[VoucherID] From faVouchers"
                mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
'                If txtAmount.Text = "" Then
'                    mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
'                End If
                mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
                If mAccountHeadID <> "%" Then
                    mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID = faVoucherChild.intVoucherID"
                    mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
                End If
                mSql = mSql + " Where tnyVoucherTypeID = 10 AND faVouchers.numSeatID LIKE'" & mSeatID & "'"
                If mDepartment <> "%" Then
                    'mSQL = mSQL + " And faTransactionType.intSectionID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And faTransactionType.intSectionID =" & mDepartment
                End If
'                If txtAmount.Text = "" Then
'                    mSql = mSql + " And faVoucherChild.intAccountHeadID LIKE '" & mAccountHeadID & "'"
'                End If
                If mTransactionType <> "%" Then
                    'mSQL = mSQL + " And faVouchers.intTransactionTypeID LIKE '" & mTransactionType & "'"
                    mSql = mSql + " And faVouchers.intTransactionTypeID =" & mTransactionType
                End If
                If mFromDate <> "" Then
                    mSql = mSql + " And dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                End If
                mSql = mSql + " And intVoucherNo LIKE '" & mVoucherNo & "'"
                If mAmount = "%" Then
                    mSql = mSql + " And faVouchers.fltAmount LIKE '" & mAmount & "'"
                Else
                    mSql = mSql + " And faVouchers.fltAmount = " & mAmount
                End If
                If mName <> "" Then
                    mSql = mSql + " And faVoucherAddress.vchName LIKE '" & "%" & mName & "%" & "'"
                End If
    '            mSql = mSql + " And faVoucherAddress.intWardNo LIKE '" & "%" & mWard & "%" & "'"
    '            mSQL = mSQL + " And faVoucherAddress.intDoorNo LIKE '" & "%" & mDoorNo1 & "%" & "'"
                If txtWard.Text <> "" Then
                    mSql = mSql + " And faVoucherAddress.intWardNo =" & txtWard.Text
                End If
                If txtDoorNo1.Text <> "" Then
                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
                        mSql = mSql + " And faVoucherAddress.intDoorNo LIKE '" & txtDoorNo1.Text & "%" & "'"
                    Else
                        mSql = mSql + " And faVoucherAddress.intDoorNo =" & txtDoorNo1.Text
                    End If
                End If
                If txtDoorNo2.Text <> "" Then
                    If mID(txtDoorNo2.Text, 1, 1) = "%" Then
                        mSql = mSql + " And faVoucherAddress.vchDoorNo2 LIKE '" & txtDoorNo2.Text & "%" & "'"
                    Else
                        mSql = mSql + " And faVoucherAddress.vchDoorNo2 ='" & txtDoorNo2.Text & "'"
                    End If
                End If
                If mChequeNo <> "" Then
                    mSql = mSql + " And vchInstrumentNo like '%" & mChequeNo & "%'"
                End If
                If mAccountHeadID <> "%" Then
                    mSql = mSql + " And faVoucherChild.intAccountHeadID =" & mAccountHeadID
                End If
                'mSQL = mSQL + " And faVouchers.tnyCancelFlag<>1"
                mSql = mSql + " Order By faVouchers.dtDate Desc,faVouchers.intVoucherNo Desc"
'            Else
'                mSql = "Select * From faVouchers"
'                mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
''                If txtAmount.Text = "" Then
''                    mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
''                End If
'                mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
'                mSql = mSql + " Where faVouchers.numSeatID LIKE'" & mSeatID & "'"
'                mSql = mSql + " And faTransactionType.intSectionID LIKE '" & mDepartment & "'"
'                mSql = mSql + " And faVouchers.intTransactionTypeID LIKE '" & mTransactionType & "'"
''                If txtAmount.Text = "" Then
''                    mSql = mSql + " And faVoucherChild.intAccountHeadID LIKE '" & mAccountHeadID & "'"
''                End If
'                mSql = mSql + " And dtDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                mSql = mSql + " And intVoucherNo LIKE '" & mVoucherNo & "'"
'                mSql = mSql + " And faVouchers.fltAmount LIKE '" & mAmount & "'"
'                mSql = mSql + " And faVoucherAddress.vchName LIKE '" & "%" & mName & "%" & "'"
'                mSql = mSql + " And faVouchers.tnyCancelFlag<>1"
'                mSql = mSql + " Order By faVouchers.dtDate Desc"
'            End If
'            If CDate(mFromDate) < CDate("01-Apr-2010") Then
'                Rec.Open mSql, mCnnBkUp, adOpenStatic, adLockPessimistic
'            Else
                Rec.Open mSql, mCnn, adOpenStatic, adLockPessimistic
'            End If
            
            If Rec.EOF Or Rec.BOF Then
                MsgBox "No Records Exist!!", vbInformation
                Exit Sub
            End If
            Call FillvsDetails(Rec, Rec.RecordCount)
            Rec.Close
            mCnn.Close
        End If
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in DB_Accounts (Saankhya)''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        If optSaankhya.Value = True Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaOld)) Then
                vsDetailsOld.Clear 1, 1
                vsDetailsOld.Rows = 1
                
                If cmbDepartment.ListIndex < 1 Then
                    mDepartment = "%"
                Else
                    mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
                End If
                
                If cmbTransactionType.ListIndex < 1 Then
                    mTransactionType = "%"
                Else
                    mTransactionType = CStr(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
                End If
                
                If txtFromDate.Text = "" Then
                    mSql = "Select YearStartDate From  TblYear Where Id=1"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mFromDate = IIf(IsNull(Rec!YearStartDate), "", CheckDateInMMM(Rec!YearStartDate))
                    End If
                    Rec.Close
                Else
                    mFromDate = txtFromDate.Text
                End If
                
                If txtToDate.Text = "" Then
                    mSql = "Select YearEndDate From TblYear Where Id=26"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mToDate = IIf(IsNull(Rec!YearEndDate), "", CheckDateInMMM(Rec!YearEndDate))
                    End If
                    Rec.Close
                Else
                    
                    mToDate = txtToDate.Text
                End If
                
                If txtBookNo.Text = "" Then
                    mBookId = "%"
                Else
                    mBookId = CStr(txtBookNo.Text)
                End If
                
                If txtVoucherNo.Text = "" Then
                    mVoucherNo = "%"
                Else
                    mVoucherNo = CStr(txtVoucherNo.Text)
                End If
                
                If txtAmount.Text = "" Then
                    mAmount = "%"
                Else
                    mAmount = CStr(txtAmount.Text)
                End If
                
                If txtName.Text = "" Then
                    mName = ""
                Else
                    mName = CStr(txtName.Text)
                End If

'                If (cmbTransactionType.ListIndex = 1) Then
'                    mSQL = "Select Count(TblReceipt.Id) As Count From TblReceipt"
'                    mSQL = mSQL + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
'                    mSQL = mSQL + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
'                    mSQL = mSQL + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'                    mSQL = mSQL + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
'                    mSQL = mSQL + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
'                    mSQL = mSQL + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                    mSQL = mSQL + " And BookNo LIKE '" & mBookId & "'"
'                    mSQL = mSQL + " And ReceiptNo LIKE '" & mVoucherNo & "'"
'                    mSQL = mSQL + " And Payee LIKE '" & "%" & mName & "%" & "'"
'                    mSQL = mSQL + " And TotalAmount LIKE '" & mAmount & "'"
'                    If txtWard.Text <> "" Then
'                        mSQL = mSQL + " And TblReceipt.WardNo =" & txtWard.Text
'                    End If
'                    If txtDoorNo1.Text <> "" Then
'    '                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
'                            mSQL = mSQL + " And TblReceipt.HouseNo LIKE '" & txtDoorNo1.Text & "%" & "'"
'    '                    Else
'    '                        mSQL = mSQL + " And TblReceipt.HouseNo =" & txtDoorNo1.Text
'    '                    End If
'                    End If
'    '                If txtDoorNo2.Text <> "" Then
'    '                    If mID(txtDoorNo2.Text, 1, 1) = "%" Then
'    '                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 LIKE '" & txtDoorNo2.Text & "%" & "'"
'    '                    Else
'    '                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 ='" & txtDoorNo2.Text & "'"
'    '                    End If
'    '                End If
'                    mSQL = mSQL + " And TblReceipt.CancelFlag<>1"
'                ElseIf (cmbTransactionType.ListIndex = 2 Or cmbTransactionType.ListIndex = 3) Then
'                    mSQL = "Select Count(TblReceipt.Id) As Count From TblReceipt"
'                    mSQL = mSQL + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
'                    mSQL = mSQL + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
'                    mSQL = mSQL + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'                    mSQL = mSQL + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
'                    mSQL = mSQL + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
'                    mSQL = mSQL + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                    mSQL = mSQL + " And BookNo LIKE '" & mBookId & "'"
'                    mSQL = mSQL + " And ReceiptNo LIKE '" & mVoucherNo & "'"
'                    mSQL = mSQL + " And Payee LIKE '" & "%" & mName & "%" & "'"
'                    mSQL = mSQL + " And TotalAmount LIKE '" & mAmount & "'"
'                    If txtWard.Text <> "" Then
'                        mSQL = mSQL + " And TblReceipt.WardNo =" & txtWard.Text
'                    End If
'                    If txtDoorNo1.Text <> "" Then
'        '                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
'                            mSQL = mSQL + " And TblReceipt.HouseNo = '" & txtDoorNo1.Text & "'"
'        '                    Else
'        '                        mSQL = mSQL + " And TblReceipt.HouseNo =" & txtDoorNo1.Text
'        '                    End If
'                    End If
'                    mSQL = mSQL + " And TblReceipt.CancelFlag<>1"
'                Else
'                    mSQL = "Select Count(TblReceipt.Id) As Count From TblReceipt"
'                    mSQL = mSQL + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
'                    mSQL = mSQL + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
'                    mSQL = mSQL + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'                    mSQL = mSQL + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
'                    mSQL = mSQL + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
'                    mSQL = mSQL + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                    mSQL = mSQL + " And BookNo LIKE '" & mBookId & "'"
'                    mSQL = mSQL + " And ReceiptNo LIKE '" & mVoucherNo & "'"
'                    mSQL = mSQL + " And Payee LIKE '" & "%" & mName & "%" & "'"
'                    mSQL = mSQL + " And TotalAmount LIKE '" & mAmount & "'"
'                    mSQL = mSQL + " And TblReceipt.CancelFlag<>1"
'                End If
'                Rec.Open mSQL, mCnn
'                If Not (Rec.EOF And Rec.BOF) Then
'                    mCount = Rec!count
'                End If
'                Rec.Close
                
                If (cmbTransactionType.ListIndex = 1) Then
                    mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TotalAmount,TblReceiptBook.BookNo From TblReceipt"
                    mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
                    mSql = mSql + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                    mSql = mSql + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
                    mSql = mSql + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                    mSql = mSql + " And BookNo LIKE '" & mBookId & "'"
                    mSql = mSql + " And ReceiptNo LIKE '" & mVoucherNo & "'"
                    mSql = mSql + " And Payee LIKE '" & "%" & mName & "%" & "'"
                    mSql = mSql + " And TotalAmount LIKE '" & mAmount & "'"
                    If txtWard.Text <> "" Then
                        mSql = mSql + " And TblReceipt.WardNo =" & txtWard.Text
                    End If
                    If txtDoorNo1.Text <> "" Then
        '                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
                            mSql = mSql + " And TblReceipt.HouseNo LIKE '" & txtDoorNo1.Text & "%" & "'"
        '                    Else
        '                        mSQL = mSQL + " And TblReceipt.HouseNo =" & txtDoorNo1.Text
        '                    End If
                    End If
        '                If txtDoorNo2.Text <> "" Then
        '                    If mID(txtDoorNo2.Text, 1, 1) = "%" Then
        '                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 LIKE '" & txtDoorNo2.Text & "%" & "'"
        '                    Else
        '                        mSQL = mSQL + " And faVoucherAddress.vchDoorNo2 ='" & txtDoorNo2.Text & "'"
        '                    End If
        '                End If
                    mSql = mSql + " And TblReceipt.CancelFlag<>1"
                    mSql = mSql + " Order By TblReceipt.Id  Asc , ReceiptDate Desc"
                ElseIf (cmbTransactionType.ListIndex = 2 Or cmbTransactionType.ListIndex = 3) Then
                    mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TotalAmount,TblReceiptBook.BookNo From TblReceipt"
                    mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
                    mSql = mSql + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                    mSql = mSql + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
                    mSql = mSql + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                    mSql = mSql + " And BookNo LIKE '" & mBookId & "'"
                    mSql = mSql + " And ReceiptNo LIKE '" & mVoucherNo & "'"
                    mSql = mSql + " And Payee LIKE '" & "%" & mName & "%" & "'"
                    mSql = mSql + " And TotalAmount LIKE '" & mAmount & "'"
                    If txtWard.Text <> "" Then
                        mSql = mSql + " And TblReceipt.WardNo =" & txtWard.Text
                    End If
                    If txtDoorNo1.Text <> "" Then
        '                    If mID(txtDoorNo1.Text, 1, 1) = "%" Then
                            mSql = mSql + " And TblReceipt.HouseNo = '" & txtDoorNo1.Text & "'"
        '                    Else
        '                        mSQL = mSQL + " And TblReceipt.HouseNo =" & txtDoorNo1.Text
        '                    End If
                    End If
                    mSql = mSql + " And TblReceipt.CancelFlag<>1"
                    mSql = mSql + " Order By TblReceipt.Id  Asc , ReceiptDate Desc"
                Else
                    mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TotalAmount,TblReceiptBook.BookNo From TblReceipt"
                    mSql = mSql + " Inner Join TB_Transaction_MST On TblReceipt.Id=TB_Transaction_MST.ReceiptId"
                    mSql = mSql + " Inner Join TB_DetailedHead_MST On TB_Transaction_MST.HeadID=TB_DetailedHead_MST.intDetailedHead_ID"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                    mSql = mSql + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And TB_DetailedHead_MST.intDetailedHead_ID LIKE '" & mTransactionType & "'"
                    mSql = mSql + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                    mSql = mSql + " And BookNo LIKE '" & mBookId & "'"
                    mSql = mSql + " And ReceiptNo LIKE '" & mVoucherNo & "'"
                    mSql = mSql + " And Payee LIKE '" & "%" & mName & "%" & "'"
                    mSql = mSql + " And TotalAmount LIKE '" & mAmount & "'"
                    mSql = mSql + " And TblReceipt.CancelFlag<>1"
                    mSql = mSql + " Order By TblReceipt.Id  Asc , ReceiptDate Desc"
                End If
                Rec.Open mSql, mCnn, adOpenStatic, adLockPessimistic
                If Rec.EOF Or Rec.BOF Then
                    MsgBox "No Records Exist!!", vbInformation
                    Exit Sub
                End If
                Call FillvsDetails(Rec, Rec.RecordCount)
                Rec.Close
                mCnn.Close
            Else
                MsgBox "Connection To Accounts does not exit, Please contact your System Administrator", vbInformation
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''Searching in Receipts (Sahatha)''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If optSahatha.Value = True Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Sahatha)) Then
            
                If cmbDepartment.ListIndex < 1 Then
                    mDepartment = "%"
                Else
                    mDepartment = CStr(cmbDepartment.ItemData(cmbDepartment.ListIndex))
                End If
                
                If cmbTransactionType.ListIndex < 1 Then
                    mTransactionType = "%"
                Else
                    mTransactionType = CStr(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
                End If
                
                If txtFromDate.Text = "" Then
                    mSql = "Select YearStartDate From  TblYear Where Id=1"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mFromDate = IIf(IsNull(Rec!YearStartDate), "", CheckDateInMMM(Rec!YearStartDate))
                    End If
                    Rec.Close
                Else
                    mFromDate = txtFromDate.Text
                End If
                
                If txtToDate.Text = "" Then
                    mSql = "Select YearEndDate From TblYear Where Id=18"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mToDate = IIf(IsNull(Rec!YearEndDate), "", CheckDateInMMM(Rec!YearEndDate))
                    End If
                    Rec.Close
                Else
                    mToDate = txtToDate.Text
                End If
                
                If txtBookNo.Text = "" Then
                    mBookId = "%"
                Else
                    mBookId = CStr(txtBookNo.Text)
                End If
                
                If txtVoucherNo.Text = "" Then
                    mVoucherNo = "%"
                Else
                    mVoucherNo = CStr(txtVoucherNo.Text)
                End If
                
                If txtAmount.Text = "" Then
                    mAmount = "%"
                Else
                    mAmount = CStr(txtAmount.Text)
                End If
                
                If txtName.Text = "" Then
                    mName = ""
                Else
                    mName = CStr(txtName.Text)
                End If
                
'                If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 1 Then
'                    mSQL = "Select Count(TblReceipt.Id) As Count From TblReceipt"
'                    mSQL = mSQL + " Inner Join tblReceiptBuildings On tblReceipt.Id=tblReceiptBuildings.ReceiptID"
'                    mSQL = mSQL + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'    '                mSQL = mSQL + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
'                    mSQL = mSQL + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
'                    mSQL = mSQL + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                    mSQL = mSQL + " And BookNo LIKE '" & mBookId & "'"
'                    mSQL = mSQL + " And ReceiptNo LIKE '" & mVoucherNo & "'"
'                    mSQL = mSQL + " And Payee LIKE '" & "%" & mName & "%" & "'"
'                    mSQL = mSQL + " And Convert(VarChar,Convert(float,Amount)) LIKE '" & mAmount & "'"
'                    If txtWard.Text <> "" Then
'                        mSQL = mSQL + " And TblReceipt.WardNo =" & txtWard.Text
'                    End If
'                    If txtDoorNo1.Text <> "" Then
'                        mSQL = mSQL + " And TblReceipt.HouseNo LIKE '" & txtDoorNo1.Text & "%" & "'"
'                    End If
'                    mSQL = mSQL + " And TblReceipt.CancelFlag<>1"
'                Else
'                    mSQL = "Select Count(TblReceipt.Id) As Count From TblReceipt"
'    '                mSQL = mSQL + " Inner Join tblReceiptBuildings On tblReceipt.Id=tblReceiptBuildings.ReceiptID"
'                    mSQL = mSQL + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
'                    mSQL = mSQL + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
'                    mSQL = mSQL + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
'                    mSQL = mSQL + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
'                    mSQL = mSQL + " And BookNo LIKE '" & mBookId & "'"
'                    mSQL = mSQL + " And ReceiptNo LIKE '" & mVoucherNo & "'"
'                    mSQL = mSQL + " And Payee LIKE '" & "%" & mName & "%" & "'"
'                    mSQL = mSQL + " And Convert(VarChar,Convert(float,Amount)) LIKE '" & mAmount & "'"
'                    mSQL = mSQL + " And TblReceipt.CancelFlag<>1"
'    '                mSql = mSql + " Group By TblReceipt.ID"
'                End If
'                Rec.Open mSQL, mCnn
'                If Not (Rec.EOF And Rec.BOF) Then
'                    mCount = Rec!count
'                End If
'                Rec.Close
                
                If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = 1 Then
                    mSql = "Select TblReceipt.Id,Payee,tblReceiptBuildings.WardNo,tblReceiptBuildings.HouseNo,ReceiptNo,ReceiptDate,TblReceiptBook.BookNo,Sum(Amount) As SumAmount From TblReceipt"
                    mSql = mSql + " Inner Join tblReceiptBuildings On tblReceipt.Id=tblReceiptBuildings.ReceiptID"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
    '                mSQL = mSQL + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
                    mSql = mSql + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                    mSql = mSql + " And BookId LIKE '" & mBookId & "'"
                    mSql = mSql + " And ReceiptNo LIKE '" & mVoucherNo & "'"
                    mSql = mSql + " And Payee LIKE '" & "%" & mName & "%" & "'"
                    mSql = mSql + " And Convert(VarChar,Convert(float,Amount)) LIKE '" & mAmount & "'"
                    If txtWard.Text <> "" Then
                        mSql = mSql + " And TblReceipt.WardNo =" & txtWard.Text
                    End If
                    If txtDoorNo1.Text <> "" Then
                        mSql = mSql + " And TblReceipt.HouseNo LIKE '" & txtDoorNo1.Text & "%" & "'"
                    End If
                    mSql = mSql + " And TblReceipt.CancelFlag<>1"
                    mSql = mSql + " Group By TblReceipt.Id,Payee,tblReceiptBuildings.WardNo,tblReceiptBuildings.HouseNo,BookId,ReceiptNo,ReceiptDate,TblReceiptBook.BookNo"
                    mSql = mSql + " Order By  TblReceipt.Id Asc,ReceiptDate Desc"
                Else
                    mSql = "Select TblReceipt.Id,Payee,WardNo,HouseNo,ReceiptNo,ReceiptDate,TblReceiptBook.BookNo,Sum(Amount) As SumAmount From TblReceipt"
    '                mSQL = mSQL + " Inner Join tblReceiptBuildings On tblReceipt.Id=tblReceiptBuildings.ReceiptID"
                    mSql = mSql + " Inner Join TblReceiptChild On TblReceipt.Id=TblReceiptChild.ReceiptId"
                    mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId=TblReceiptBook.Id"
                    mSql = mSql + " Where TblReceipt.DepartmentID LIKE '" & mDepartment & "'"
                    mSql = mSql + " And ReceiptDate BETWEEN '" & mFromDate & "' AND '" & mToDate & "'"
                    mSql = mSql + " And BookNo LIKE '" & mBookId & "'"
                    mSql = mSql + " And ReceiptNo LIKE '" & mVoucherNo & "'"
                    mSql = mSql + " And Payee LIKE '" & "%" & mName & "%" & "'"
                    mSql = mSql + " And Convert(VarChar,Convert(float,Amount)) LIKE '" & mAmount & "'"
                    mSql = mSql + " And TblReceipt.CancelFlag<>1"
                    mSql = mSql + " Group By TblReceipt.Id,Payee,WardNo,HouseNo,BookId,ReceiptNo,ReceiptDate,TblReceiptBook.BookNo"
                    mSql = mSql + " Order By  TblReceipt.Id Asc,ReceiptDate Desc"
                End If
                Rec.Open mSql, mCnn, adOpenStatic, adLockPessimistic
                If Rec.EOF Or Rec.BOF Then
                    MsgBox "No Records Exist!!", vbInformation
                    Exit Sub
                End If
                Call FillvsDetails(Rec, Rec.RecordCount)
                Rec.Close
                mCnn.Close
            Else
                MsgBox "Connection To Receipts does not exit, Please contact your System Administrator", vbInformation
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub cmdSearchHead_Click()
        Dim mIndex  As Long
        Dim mSql    As String
        
        On Error GoTo err
        If optSaankhyaDoubleEntry.Value = True Then
            If cmbTransactionType.ListIndex > -1 Then
                mIndex = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
            End If
            If mIndex > 0 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                mSql = mSql + " Where intTransactionTypeID = " & mIndex & " Order By faTransactionTypeChild.intOrder"
                frmSearchAccountHeads.SQLString = mSql '"Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            Else
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            End If
            frmSearchAccountHeads.Show vbModal
            
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    txtAccountHeadCode.Text = objAccHead.AccountCode
                    txtAccountHead.Text = objAccHead.AccountHead
                    txtAccountHeadCode.Tag = objAccHead.AccountHeadID
    '                vsGrid.TextMatrix(Row, 1) = objAccHead.AccountHead
    '                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
                gbSearchStr = ""
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

'    Private Sub cmdUpdate_Click()
'        Dim mCnn                As New ADODB.Connection
'        Dim Rec                 As New ADODB.Recordset
'        Dim objDb               As New clsDB
'        Dim mSQL                As String
'        Dim mDVoucherID         As Variant
'        Dim mVVoucherID         As Variant
'        Dim mDTransactionTypeID As Variant
'        Dim mVTransactionTypeID As Variant
'
'        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mSQL = "Select faIDemandTBL.intTransactionTypeID,faVouchers.intTransactionTypeID,faIDemandTBL.intVoucherID,faVouchers.intVoucherID From faIDemandTBL"
'        mSQL = mSQL + " Inner Join faVouchers On faIDemandTBL.intVoucherID=faVouchers.intVoucherID"
'        Rec.Open mSQL, mCnn
'        While Not Rec.EOF
'            mDTransactionTypeID = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0))
'            mVTransactionTypeID = IIf(IsNull(Rec.Fields(1)), "", Rec.Fields(1))
'            mDVoucherID = IIf(IsNull(Rec.Fields(2)), "", Rec.Fields(2))
'            mVVoucherID = IIf(IsNull(Rec.Fields(3)), "", Rec.Fields(3))
'            mSQL = "Update faVouchers"
'            mSQL = mSQL + " Set intTransactionTypeID=" & mDTransactionTypeID
'            mSQL = mSQL + " Where intVoucherID=" & mDVoucherID
'
'            mCnn.Execute mSQL
'            Rec.MoveNext
'        Wend
'    End Sub

    Private Sub dtpFromDate_CloseUp()
        txtFromDate.Text = CheckDateInMMM(dtpFromDate.Value)
    End Sub

    Private Sub dtpToDate_CloseUp()
        txtToDate.Text = CheckDateInMMM(dtpToDate.Value)
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCount      As Double
'        Dim mRowCount   As Integer
        Dim mSerialNo   As Integer
        Dim mStatus     As Variant
    
        frmSearchReceipts.Width = 11940
        frmSearchReceipts.Height = 7050
'        XPC.InitIDESubClassing
'        txtWard.Enabled = False
'        txtDoorNo1.Enabled = False
'        txtDoorNo2.Enabled = False
        cmdUpdate.Visible = False
        cmdSearchHead.Enabled = False
        optSaankhyaDoubleEntry.Value = True
        PopulateList cmbSeat, "SELECT chvSeatTitle,numSeatID FROM GL_Seats ORDER BY chvSeatTitle", , , True, , enuSourceString.DBMaster
    End Sub

    Private Sub optBackUp_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        
        cmbDepartment.Clear
        cmbTransactionType.Clear
        cmbTransactionType.Enabled = True
        Call FormInitialize
        pbSearch.Value = 0
        vsDetails.Visible = True
        vsDetailsOld.Visible = False
        vsDetailsSahatha.Visible = False
        vsDetails.Left = 165
        vsDetails.Top = 3030
'        txtWard.Enabled = False
'        txtDoorNo1.Enabled = False
'        txtDoorNo2.Enabled = False
        cmbSeat.Enabled = True
        txtBookNo.Visible = False
        cmdSearchHead.Enabled = True
        lblReceiptNo.Left = 5730
        lblReceiptNo.Top = 1500
        lblReceiptNo.Caption = "Receipt No"
        txtDoorNo1.Width = 915
        cmdSearch.Enabled = True
        If optBackUp.Value = True Then
            '''''''''''''''''''''''''''''''''To Search in Finance Backup Database''''''''''''''''''''''''''''''''''''''''''''
            If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaBackUp)) Then
                mSql = "Select vchSectionName,intSectionID From faSection"
                PopulateList cmbDepartment, mSql, , True, , True, enuSourceString.SaankhyaBackUp
                mCnn.Close
            Else
                MsgBox "Connection To Finance BackUp does not exit, Please contact your System Administrator", vbInformation
                cmdSearch.Enabled = False
            End If
        End If
    End Sub

    Private Sub optSaankhya_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
                
        cmbDepartment.Clear
        cmbTransactionType.Clear
        cmbTransactionType.Enabled = True
        Call FormInitialize
        pbSearch.Value = 0
        vsDetailsOld.Visible = True
        vsDetails.Visible = False
        vsDetailsSahatha.Visible = False
        vsDetailsOld.Left = 165
        vsDetailsOld.Top = 3030
'        txtWard.Enabled = False
'        txtDoorNo1.Enabled = False
'        txtDoorNo2.Enabled = False
        cmbSeat.Enabled = False
        cmdSearchHead.Enabled = False
        txtBookNo.Visible = True
        txtBookNo.TabIndex = 5
        lblReceiptNo.Left = 5055
        lblReceiptNo.Top = 1500
        lblReceiptNo.Caption = "BookNo/ReceiptNo"
        txtDoorNo1.Width = 1845
        cmdSearch.Enabled = True
        If optSaankhya.Value = True Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaOld)) Then
                mSql = "Select chvDeptName,intDeptId From TB_Department_MST"
                PopulateList cmbDepartment, mSql, , True, , True, enuSourceString.SaankhyaOld
                
                mSql = "Select chvHead,intDetailedHead_ID From TB_DetailedHead_MST"
                PopulateList cmbTransactionType, mSql, , True, , True, enuSourceString.SaankhyaOld
                mCnn.Close
            Else
                MsgBox "Connection To Accounts does not exit, Please contact your System Administrator", vbInformation
                cmdSearch.Enabled = False
            End If
        End If
    End Sub

    Private Sub optSaankhyaDoubleEntry_Click()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        
        cmbDepartment.Clear
        cmbTransactionType.Clear
        cmbTransactionType.Enabled = True
        Call FormInitialize
        pbSearch.Value = 0
        vsDetails.Visible = True
        vsDetailsOld.Visible = False
        vsDetailsSahatha.Visible = False
        vsDetails.Left = 165
        vsDetails.Top = 3030
'        txtWard.Enabled = False
'        txtDoorNo1.Enabled = False
'        txtDoorNo2.Enabled = False
        cmbSeat.Enabled = True
        txtBookNo.Visible = False
        cmdSearchHead.Enabled = True
        lblReceiptNo.Left = 5730
        lblReceiptNo.Top = 1500
        lblReceiptNo.Caption = "Receipt No"
        txtDoorNo1.Width = 915
        cmdSearch.Enabled = True
        If optSaankhyaDoubleEntry.Value = True Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSql = "Select vchSectionName,intSectionID From faSection"
                PopulateList cmbDepartment, mSql, , True, , True, enuSourceString.Saankhya
                mCnn.Close
            Else
                MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
                cmdSearch.Enabled = False
            End If
        End If
    End Sub

    Private Sub optSahatha_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        
        cmbDepartment.Clear
        cmbTransactionType.Clear
'        cmbTransactionType.Enabled = False
        cmbSeat.Enabled = False
        cmdSearchHead.Enabled = False
        Call FormInitialize
        pbSearch.Value = 0
        vsDetails.Visible = False
        vsDetailsOld.Visible = False
        vsDetailsSahatha.Visible = True
        vsDetailsSahatha.Left = 165
        vsDetailsSahatha.Top = 3030
        'txtWard.Enabled = False
        'txtDoorNo1.Enabled = False
        txtDoorNo1.Width = 1845
        'txtDoorNo2.Enabled = False
        txtBookNo.Visible = True
        txtBookNo.TabIndex = 5
        lblReceiptNo.Left = 5055
        lblReceiptNo.Top = 1500
        lblReceiptNo.Caption = "BookNo/ReceiptNo"
        txtDoorNo1.Width = 1845
        cmdSearch.Enabled = True
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Sahatha)) Then
        
            mSql = "Select Name,Id From Department"
            PopulateList cmbDepartment, mSql, , True, , True, enuSourceString.Sahatha
            
            mSql = "Select DemandType,Id From DemandTypes Where Id=1"
            PopulateList cmbTransactionType, mSql, , True, , True, enuSourceString.Sahatha
        Else
            MsgBox "Connection To Receipts does not exit, Please contact your System Administrator", vbInformation
            cmdSearch.Enabled = False
        End If
    End Sub

    Private Sub txtAccountHead_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then
            txtAccountHead.Text = ""
        End If
    End Sub
    
    Private Sub txtAccountHeadCode_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 46 Then
            txtAccountHeadCode.Text = ""
        End If
    End Sub

    Private Sub txtAmount_LostFocus()
        txtAmount.Text = Format(txtAmount.Text, ".00")
    End Sub

    Private Sub txtBookNo_GotFocus()
        txtBookNo.SelStart = 0
        txtBookNo.SelLength = Len(txtBookNo.Text)
    End Sub

    Private Sub txtBookNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtDoorNo1_GotFocus()
        txtDoorNo1.SelStart = 0
        txtDoorNo1.SelLength = Len(txtDoorNo1.Text)
    End Sub
    
    Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
        If optSaankhyaDoubleEntry.Value = True Then
            If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8 Or KeyAscii = Asc("%")) Then
                KeyAscii = 0
            End If
        End If
    End Sub

    Private Sub txtDoorNo2_GotFocus()
        txtDoorNo2.SelStart = 0
        txtDoorNo2.SelLength = Len(txtDoorNo2.Text)
    End Sub

    Private Sub txtFromDate_GotFocus()
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate)
    End Sub

    Private Sub txtFromDate_LostFocus()
        If Trim(txtFromDate.Text) <> "" Then
            txtFromDate.Text = CheckDateInMMM(txtFromDate.Text)
        End If
    End Sub

    Private Sub txtName_GotFocus()
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName)
    End Sub
   
    Private Sub txtToDate_GotFocus()
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate)
    End Sub

    Private Sub txtToDate_LostFocus()
        If Trim(txtToDate.Text) <> "" Then
            txtToDate.Text = CheckDateInMMM(txtToDate.Text)
        End If
    End Sub
    
    Private Sub txtVoucherNo_GotFocus()
        txtVoucherNo.SelStart = 0
        txtVoucherNo.SelLength = Len(txtVoucherNo.Text)
    End Sub

    Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtWard_GotFocus()
        txtWard.SelStart = 0
        txtWard.SelLength = Len(txtWard.Text)
    End Sub

    Private Sub txtWard_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsDetails_Click()
        vsDetails.SelectionMode = flexSelectionByRow
    End Sub

    Private Sub vsDetails_DblClick()
        On Error GoTo err
        If vsDetails.TextMatrix(vsDetails.Row, 4) <> "" Then
            If IsNumeric(vsDetails.TextMatrix(vsDetails.Row, 4)) Then
                frmReceipt.Visible = True
                frmReceipt.vsGrid.Clear 1, 1
                Call DisplayReceiptDetails(vsDetails.TextMatrix(vsDetails.Row, 8))
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub vsDetailsOld_Click()
        vsDetailsOld.SelectionMode = flexSelectionByRow
    End Sub

    Private Sub vsDetailsOld_DblClick()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mID         As String

        On Error GoTo err
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SaankhyaOld)) Then
            If vsDetailsOld.TextMatrix(vsDetailsOld.Row, 4) <> "" And vsDetailsOld.TextMatrix(vsDetailsOld.Row, 5) <> "" Then
                If IsNumeric(vsDetailsOld.TextMatrix(vsDetailsOld.Row, 4)) And IsNumeric(vsDetailsOld.TextMatrix(vsDetailsOld.Row, 5)) Then
                    mSql = "Select TblReceipt.Id From TblReceipt"
                    'mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookID = TblReceiptBook.Id"
                    mSql = mSql + " Where BookId =" & vsDetailsOld.TextMatrix(vsDetailsOld.Row, 8)
                    mSql = mSql + " And ReceiptNo =" & vsDetailsOld.TextMatrix(vsDetailsOld.Row, 5)
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mID = IIf(IsNull(Rec!id), "", Rec!id)
                    End If
                    Rec.Close
                    frmReceipt.Visible = True
                    frmReceipt.vsGrid.Clear 1, 1
                    Call DisplayReceiptDetails(mID)
                End If
            End If
        Else
            MsgBox "Connection To Accounts does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub vsDetailsSahatha_Click()
        vsDetailsSahatha.SelectionMode = flexSelectionByRow
    End Sub

    Private Sub vsDetailsSahatha_DblClick()
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mSql        As String
        Dim mID         As String

        On Error GoTo err
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Sahatha)) Then
        
            If vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 4) <> "" And vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 5) <> "" Then
                If IsNumeric(vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 4)) And IsNumeric(vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 5)) Then
                    mSql = "Select TblReceipt.Id From TblReceipt"
                    'mSql = mSql + " Inner Join TblReceiptBook On TblReceipt.BookId = TblReceiptBook.Id"
                    mSql = mSql + " Where BookId =" & vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 8)
                    mSql = mSql + " And ReceiptNo =" & vsDetailsSahatha.TextMatrix(vsDetailsSahatha.Row, 5)
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mID = IIf(IsNull(Rec!id), "", Rec!id)
                    End If
                    Rec.Close
                    frmReceipt.Visible = True
                    frmReceipt.vsGrid.Clear 1, 1
                    Call DisplayReceiptDetails(mID)
                End If
            End If
        Else
            MsgBox "Connection To Reciepts does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub


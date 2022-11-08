VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmReceipt 
   BackColor       =   &H00DAF2F2&
   Caption         =   "Receipt"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   Icon            =   "frmReceipt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11820
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   2430
      Left            =   30
      TabIndex        =   29
      Top             =   1530
      Width           =   11745
      _cx             =   20717
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReceipt.frx":1CCA
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
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   11640
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   11640
         Top             =   2040
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   3675
         TabIndex        =   72
         Top             =   1995
         Visible         =   0   'False
         Width           =   4290
      End
   End
   Begin VB.Frame fraTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   6885
      TabIndex        =   54
      Top             =   3975
      Width           =   4920
      Begin VB.TextBox txtRoundOff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF2F2&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   570
         Width           =   690
      End
      Begin VB.TextBox txtAdvance 
         Alignment       =   1  'Right Justify
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   570
         Width           =   1725
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   900
         Width           =   1725
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   1215
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   1200
         Width           =   3450
      End
      Begin VB.TextBox txtTotalCurrent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   1725
      End
      Begin VB.TextBox txtTotalArrear 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   270
         Width           =   1710
      End
      Begin VB.TextBox txtAdminNote 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1215
         TabIndex        =   55
         Top             =   2235
         Width           =   3450
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
         Left            =   465
         TabIndex        =   66
         Top             =   585
         Width           =   705
      End
      Begin VB.Label lblAdvance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Amt."
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
         Left            =   1935
         TabIndex        =   65
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label10 
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
         Left            =   2040
         TabIndex        =   64
         Top             =   930
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Description:"
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
         Left            =   285
         TabIndex        =   63
         Top             =   1200
         Width           =   930
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
         Left            =   810
         TabIndex        =   62
         Top             =   270
         Width           =   375
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
         Left            =   240
         TabIndex        =   61
         Top             =   2250
         Width           =   960
      End
   End
   Begin VB.Frame fraTransactionType 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   30
      TabIndex        =   43
      Top             =   15
      Width           =   6780
      Begin VB.TextBox txtBookNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4935
         TabIndex        =   70
         Top             =   810
         Width           =   840
      End
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   53
         Top             =   210
         Width           =   4905
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   47
         Top             =   510
         Width           =   4905
      End
      Begin VB.TextBox txtOutDoorStaff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   46
         Top             =   1110
         Visible         =   0   'False
         Width           =   4905
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   810
         Width           =   1755
      End
      Begin VB.TextBox txtReceiptNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4935
         TabIndex        =   44
         Top             =   810
         Width           =   1680
      End
      Begin VB.Label lblReceiptNo 
         BackColor       =   &H00DAF2F2&
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
         Height          =   255
         Left            =   3960
         TabIndex        =   71
         Top             =   825
         Width           =   2070
      End
      Begin VB.Label lblSection 
         BackColor       =   &H00DAF2F2&
         Caption         =   "Section"
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
         Height          =   225
         Left            =   1005
         TabIndex        =   52
         Top             =   210
         Width           =   645
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
         Left            =   195
         TabIndex        =   50
         Top             =   510
         Width           =   1485
      End
      Begin VB.Label lblOutDoorStaff 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   8
         Left            =   495
         TabIndex        =   49
         Top             =   1170
         Width           =   1170
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
         Left            =   1275
         TabIndex        =   48
         Top             =   870
         Width           =   405
      End
   End
   Begin VB.Frame fraAccountHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   6840
      TabIndex        =   30
      Top             =   0
      Width           =   4950
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
         Left            =   1065
         TabIndex        =   36
         Top             =   510
         Width           =   3765
      End
      Begin VB.TextBox txtDated 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         TabIndex        =   35
         Top             =   810
         Width           =   1470
      End
      Begin VB.TextBox txtInstNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1065
         TabIndex        =   34
         Top             =   810
         Width           =   1740
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
         Left            =   1065
         TabIndex        =   33
         Top             =   210
         Width           =   3765
      End
      Begin VB.TextBox txtBank 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1065
         TabIndex        =   32
         Top             =   1110
         Width           =   1740
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         TabIndex        =   31
         Top             =   1110
         Width           =   1470
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
         Left            =   90
         TabIndex        =   42
         Top             =   540
         Width           =   915
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
         Left            =   2835
         TabIndex        =   41
         Top             =   840
         Width           =   525
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
         Left            =   330
         TabIndex        =   40
         Top             =   825
         Width           =   675
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
         Left            =   255
         TabIndex        =   39
         Top             =   240
         Width           =   750
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
         Left            =   540
         TabIndex        =   38
         Top             =   1110
         Width           =   450
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
         Left            =   2835
         TabIndex        =   37
         Top             =   1140
         Width           =   495
      End
   End
   Begin VB.Frame fraDemandDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   30
      TabIndex        =   0
      Top             =   3975
      Width           =   6840
      Begin VB.TextBox txtDoorNo1 
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
         Left            =   855
         MaxLength       =   50
         TabIndex        =   2
         Top             =   855
         Width           =   1095
      End
      Begin VB.TextBox txtRefNo 
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
         Left            =   870
         MaxLength       =   50
         TabIndex        =   69
         Top             =   1410
         Width           =   1770
      End
      Begin VB.TextBox txtZone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   855
         TabIndex        =   68
         Top             =   195
         Width           =   1785
      End
      Begin VB.TextBox txtForwardedSeat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         TabIndex        =   51
         Top             =   1740
         Width           =   1770
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
         Height          =   300
         Left            =   3750
         MaxLength       =   30
         TabIndex        =   15
         Top             =   2205
         Width           =   1635
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
         TabIndex        =   14
         Top             =   1875
         Width           =   975
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
         TabIndex        =   13
         Top             =   1875
         Width           =   1635
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
         Left            =   6315
         MaxLength       =   1
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
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
         Left            =   5985
         MaxLength       =   1
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   225
         Width           =   315
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
         Left            =   5655
         MaxLength       =   1
         TabIndex        =   10
         Top             =   225
         Width           =   315
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
         Left            =   5325
         MaxLength       =   1
         TabIndex        =   9
         Top             =   225
         Width           =   315
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
         TabIndex        =   8
         Top             =   1545
         Width           =   2880
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
         TabIndex        =   7
         Top             =   1215
         Width           =   2880
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
         TabIndex        =   6
         Top             =   885
         Width           =   2880
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
         TabIndex        =   5
         Top             =   555
         Width           =   2880
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
         Left            =   3750
         MaxLength       =   100
         TabIndex        =   4
         Top             =   225
         Width           =   1560
      End
      Begin VB.TextBox txtDoorNo2 
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
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   3
         Top             =   855
         Width           =   675
      End
      Begin VB.TextBox txtWardNo 
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
         Left            =   855
         MaxLength       =   3
         TabIndex        =   1
         Top             =   525
         Width           =   1785
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
         Left            =   255
         TabIndex        =   28
         Top             =   1785
         Width           =   555
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
         TabIndex        =   27
         Top             =   2250
         Width           =   690
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
         TabIndex        =   26
         Top             =   1920
         Width           =   420
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
         TabIndex        =   25
         Top             =   1905
         Width           =   315
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
         TabIndex        =   24
         Top             =   1575
         Width           =   765
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
         TabIndex        =   23
         Top             =   1230
         Width           =   825
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
         TabIndex        =   22
         Top             =   930
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
         TabIndex        =   21
         Top             =   600
         Width           =   960
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
         Left            =   3285
         TabIndex        =   20
         Top             =   285
         Width           =   405
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
         Left            =   465
         TabIndex        =   19
         Top             =   255
         Width           =   375
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
         Left            =   255
         TabIndex        =   18
         Top             =   870
         Width           =   585
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
         Left            =   210
         TabIndex        =   17
         Top             =   540
         Width           =   630
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
         Left            =   390
         TabIndex        =   16
         Top             =   1425
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '*********************************************************************************************'
    '                                       Form to view the Receipt                              '
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

        frmReceipt.vsGrid.Clear 1, 1
        txtBookNo.Visible = False
        lblReceiptNo.Caption = "Receipt No"
        lblReceiptNo.Left = 3960
        lblReceiptNo.Top = 825
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = "Select * From faVouchers"
        mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
        mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
        mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        ''mSql = mSql + " Inner Join faSection On faTransactionType.intSectionID=faSection.intSectionID"
        mSql = mSql + " Inner Join faCounters On faCounters.intCounterID=faVouchers.intCounterID"
        mSql = mSql + " Inner Join faSection On faCounters.intSectionID=faSection.intSectionID"
        mSql = mSql + " Inner Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
        mSql = mSql + " Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
        'mSQL = mSQL + " Or  faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID "
        mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
        mSql = mSql + " Where intVoucherNo=" & mVoucherNo
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtReceiptNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            txtReceiptNo.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                
            txtSection.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
            txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
            txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            
            txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
            txtInstNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            txtDated.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            txtBank.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
            txtPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
            
            If IsNull(Rec!chvZoneNameEnglish) = False Then
                txtZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
            End If
            txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
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
            
            txtAdvance.Text = IIf(IsNull(Rec!fltAdvAmtAdj), 0, Rec!fltAdvAmtAdj)
            txtDescription.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            mSqlAccHeads = "Select * From faVoucherChild"
            mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
            '-----------------------------------------
            'Added By Anisha On 30.09.10 to Diplay Period
            mSqlAccHeads = mSqlAccHeads + " left Join faPeriodicity On faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
            '-------------------------------------------
            mSqlAccHeads = mSqlAccHeads + " Where intVoucherID=" & txtReceiptNo.Tag
            RecAccHeads.Open mSqlAccHeads, mCnn
            mRowCount = 1
            While Not Rec.EOF
                While Not RecAccHeads.EOF
                    vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                    
                    ''''''''''''''''''''''''To be Removed'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If txtTransactionType.Tag = 12 And RecAccHeads!vchAccountHeadCode = 140130400 Then
                        vsGrid.TextMatrix(mRowCount, 0) = "140130200"
                        vsGrid.TextMatrix(mRowCount, 1) = "Fees for Delayed Registration - Birth & DeathCertificate"
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    mPeriodID = IIf(IsNull(RecAccHeads!tnyPeriodID), "", RecAccHeads!tnyPeriodID)
                    mYearID = IIf(IsNull(RecAccHeads!intYearID), 0, RecAccHeads!intYearID)
                    If mYearID <> 0 Then
                        vsGrid.TextMatrix(mRowCount, 2) = mYearID & "-" & mYearID + 1
                    End If
                    '---------------------------------------------------------
'                    If mPeriodID = 1 Then
'                        vsGrid.TextMatrix(mRowCount, 3) = "1st Half"
'                    End If
'                    If mPeriodID = 2 Then
'                        vsGrid.TextMatrix(mRowCount, 3) = "2nd Half"
'                    End If
'                    If mPeriodID = 3 Then
'                        vsGrid.TextMatrix(mRowCount, 3) = "Full Year"
'                    End If
                     frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchPeriodicity), "", RecAccHeads!vchPeriodicity)
                    '--------------------------------------------------------
                    mArrearFlag = IIf(IsNull(RecAccHeads!tnyArrearFlag), 0, RecAccHeads!tnyArrearFlag)
                    If mArrearFlag = 0 Then
                        vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                    End If
                    If mArrearFlag = 1 Then
                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
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
        mCnn.Close
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
        txtTotalArrear.Text = Format(mAmtArrear, "0.00")
        txtTotalCurrent.Text = Format(mAmtCurrent, "0.00")
        txtTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
        txtRoundOff.Text = Format(RoundOffAdjustment(val(txtTotal)), "0.00")
        txtTotal.Text = Format(val(txtTotal) + val(txtRoundOff) - val(txtAdvance), "0.00")
    End Sub

    
    Private Sub FormInitialize()
        Dim mCrl As Control
        
        vsGrid.Clear 1, 0
        For Each mCrl In Me
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        frmReceipt.Width = 11940
        frmReceipt.Height = 7155
        Call FormInitialize
    End Sub

    Private Sub Timer1_Timer()
        lblMessage.Visible = True
        Timer2.Enabled = True
    End Sub

    Private Sub Timer2_Timer()
        lblMessage.Visible = False
        Timer1.Enabled = True
    End Sub
    
    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        KeyAscii = 0
    End Sub

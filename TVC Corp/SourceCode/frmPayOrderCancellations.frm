VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPayOrderCancellations 
   BackColor       =   &H00EDF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " P a y m e n t   o r d e r   C a n c e l l a t i o n s"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayOrderCancellations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraProceedings 
      BackColor       =   &H00EDF7F7&
      Caption         =   "Proceedings"
      Height          =   2970
      Left            =   0
      TabIndex        =   80
      Top             =   5160
      Width           =   4935
      Begin VB.TextBox txtProceedingsRemarks 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1080
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   1680
         Width           =   2970
      End
      Begin VB.TextBox txtProceddingsNo 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   85
         Top             =   840
         Width           =   2580
      End
      Begin VB.CommandButton cmdProceedings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   840
         Width           =   270
      End
      Begin VB.TextBox txtProceedingsDate 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   82
         Top             =   1200
         Width           =   2580
      End
      Begin VB.TextBox txtCancelStatus 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   360
         Width           =   3165
      End
      Begin MSComCtl2.DTPicker dtpProcDate 
         Height          =   315
         Left            =   4275
         TabIndex        =   87
         Top             =   1320
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61276161
         CurrentDate     =   39910
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   855
         TabIndex        =   90
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedings Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedings No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   88
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1080
         TabIndex        =   83
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14115
      TabIndex        =   34
      Top             =   4770
      Width           =   14115
      Begin VB.TextBox txtDemandNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10710
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   45
         Width           =   1860
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demand No"
         Height          =   195
         Left            =   9630
         TabIndex        =   36
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payorder Cancellations"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   630
         TabIndex        =   35
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDF7F7&
      Caption         =   "Pay Order Details"
      Height          =   4020
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   14115
      Begin VB.TextBox txtReqSeat 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   3555
         Width           =   1635
      End
      Begin VB.TextBox txtReqUser 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   3555
         Width           =   2715
      End
      Begin VB.TextBox txtAppSeat 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   3555
         Width           =   1635
      End
      Begin VB.TextBox txtAppUser 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   10755
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   3555
         Width           =   2715
      End
      Begin VB.CommandButton cmdPayorderSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   240
         Width           =   270
      End
      Begin VB.CommandButton cmdForwardedSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1755
         Width           =   270
      End
      Begin VB.TextBox txtForwardedSeat 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1755
         Width           =   4200
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   630
         Left            =   9045
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   4200
      End
      Begin VB.ComboBox cmbCancelReasons 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1440
         Width           =   4470
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   3165
      End
      Begin VB.TextBox txtNetAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1125
         Width           =   1680
      End
      Begin VB.TextBox txtPayee 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   855
         Width           =   4470
      End
      Begin VB.TextBox txtSourceofFund 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1170
         Width           =   4200
      End
      Begin VB.CommandButton cmdSearchSourceofFund 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12195
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1170
         Width           =   270
      End
      Begin VB.TextBox txtFunction 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   4200
      End
      Begin VB.TextBox txtFunctionary 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   4200
      End
      Begin VB.CommandButton cmdSearchFunction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12195
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   900
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13275
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   630
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   495
         Width           =   270
      End
      Begin VB.TextBox txtPaymentType 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   4200
      End
      Begin VB.TextBox txtPayorderDate 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   6
         Top             =   315
         Width           =   1230
      End
      Begin VB.TextBox txtPayOrderNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   270
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpDueDate 
         Height          =   315
         Left            =   5220
         TabIndex        =   7
         Top             =   315
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61276161
         CurrentDate     =   39910
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGridRecoveries 
         Height          =   1275
         Left            =   2025
         TabIndex        =   31
         Top             =   2160
         Width           =   5190
         _cx             =   9155
         _cy             =   2249
         Appearance      =   1
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayOrderCancellations.frx":000C
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
      Begin VSFlex8LCtl.VSFlexGrid vsGriidVouchers 
         Height          =   1275
         Left            =   9045
         TabIndex        =   33
         Top             =   2160
         Width           =   4245
         _cx             =   7488
         _cy             =   2249
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483646
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayOrderCancellations.frx":00AE
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
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "PO Generated By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   495
         TabIndex        =   76
         Top             =   3600
         Width           =   1485
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7830
         TabIndex        =   75
         Top             =   3600
         Width           =   1110
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Vouchers Generated"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   8010
         TabIndex        =   32
         Top             =   2205
         Width           =   900
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Recoveries"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1035
         TabIndex        =   30
         Top             =   2205
         Width           =   945
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded Seat"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   585
         TabIndex        =   25
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8190
         TabIndex        =   28
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation Reason"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   540
         TabIndex        =   23
         Top             =   1395
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8400
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   990
         TabIndex        =   13
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Payee"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1485
         TabIndex        =   11
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Source of Fund"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7695
         TabIndex        =   21
         Top             =   1215
         Width           =   1290
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   19
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8010
         TabIndex        =   17
         Top             =   675
         Width           =   990
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3510
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Order No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   810
         TabIndex        =   3
         Top             =   405
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14235
      TabIndex        =   60
      Top             =   8280
      Width           =   14235
      Begin VB.CommandButton cmdCancelApprovedlDemandNo 
         Caption         =   "Cancel Approved Demand No"
         Height          =   465
         Left            =   1935
         TabIndex        =   94
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdCancelRequest 
         Caption         =   "Cancel Pay Order &Request"
         Height          =   465
         Left            =   270
         TabIndex        =   61
         Top             =   0
         Width           =   1635
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&Lose"
         Height          =   465
         Left            =   8535
         TabIndex        =   64
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   465
         Left            =   7500
         TabIndex        =   63
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   465
         Left            =   6480
         TabIndex        =   62
         Top             =   0
         Width           =   1005
      End
   End
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   14160
      Top             =   9120
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14325
      TabIndex        =   0
      Top             =   0
      Width           =   14325
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payorder Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   630
         TabIndex        =   1
         Top             =   0
         Width           =   2685
      End
   End
   Begin VB.Frame fraNextLevel 
      BackColor       =   &H00EDF7F7&
      Caption         =   "Pay Order Cancel Request Details"
      Height          =   2970
      Left            =   4920
      TabIndex        =   38
      Top             =   5160
      Width           =   9315
      Begin VB.TextBox txtReqSeatCancel 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2040
         Width           =   1545
      End
      Begin VB.TextBox txtReqUserCancel 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   2040
         Width           =   2715
      End
      Begin VB.TextBox txtApprovedSec 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox txtApprovedAO 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   2400
         Width           =   2625
      End
      Begin VB.Frame fraInter 
         BackColor       =   &H00EDF7F7&
         Height          =   1455
         Left            =   360
         TabIndex        =   40
         Top             =   540
         Width           =   4455
         Begin VB.CheckBox chkRecoveriesPaidYesInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   285
            Left            =   3120
            TabIndex        =   48
            Top             =   630
            Width           =   600
         End
         Begin VB.CheckBox chkPaidToPartyYesInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   195
            Left            =   3120
            TabIndex        =   42
            Top             =   225
            Width           =   600
         End
         Begin VB.CheckBox chkPaidToPatyNoInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   195
            Left            =   3750
            TabIndex        =   43
            Top             =   225
            Width           =   600
         End
         Begin VB.CheckBox chkRecoveriesPaidNoInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   285
            Left            =   3750
            TabIndex        =   49
            Top             =   630
            Width           =   600
         End
         Begin VB.CheckBox chkChequeCancelledYesInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   285
            Left            =   3120
            TabIndex        =   45
            Top             =   1035
            Width           =   600
         End
         Begin VB.CheckBox chkChequeCancelledNoInter 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   285
            Left            =   3750
            TabIndex        =   46
            Top             =   1035
            Width           =   600
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Made and Cash/Cheque issued to Party"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   120
            TabIndex        =   41
            Top             =   225
            Width           =   3000
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Recoveries Paid"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1635
            TabIndex        =   47
            Top             =   675
            Width           =   1365
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque Cancelled"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   44
            Top             =   1080
            Width           =   1560
         End
      End
      Begin VB.Frame fraFinal 
         BackColor       =   &H00EDF7F7&
         Height          =   1455
         Left            =   4875
         TabIndex        =   51
         Top             =   540
         Width           =   4380
         Begin VB.CheckBox chkChequeCancelledNoApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   285
            Left            =   3750
            TabIndex        =   57
            Top             =   1110
            Width           =   600
         End
         Begin VB.CheckBox chkChequeCancelledYesApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   285
            Left            =   3120
            TabIndex        =   56
            Top             =   1110
            Width           =   600
         End
         Begin VB.CheckBox chkRecoveriesPaidNoApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   285
            Left            =   3750
            TabIndex        =   59
            Top             =   705
            Width           =   600
         End
         Begin VB.CheckBox chkPaidToPatyNoApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "No"
            Height          =   195
            Left            =   3720
            TabIndex        =   54
            Top             =   225
            Width           =   600
         End
         Begin VB.CheckBox chkPaidToPartyYesApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   195
            Left            =   3120
            TabIndex        =   53
            Top             =   225
            Width           =   600
         End
         Begin VB.CheckBox chkRecoveriesPaidYesApp 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Yes"
            Height          =   285
            Left            =   3120
            TabIndex        =   65
            Top             =   705
            Width           =   600
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque Cancelled"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   55
            Top             =   1155
            Width           =   1560
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Recoveries Paid"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1560
            TabIndex        =   58
            Top             =   750
            Width           =   1365
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00592525&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Made and Cash/Cheque issued to Party"
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2880
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   93
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   79
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Certifications"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3510
         TabIndex        =   70
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Level"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   50
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Intermediate Level"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   39
         Top             =   360
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmPayOrderCancellations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRequestID              As Double
Private mCancelStatus           As Integer    '   =
Private mPayorderStatus         As Integer
Private mPayorderRequested      As Variant
Private mPayorderRequestedSeat  As Variant
Private mPayorderApproved       As Variant
Private mPayorderApprovedSeat   As Variant
Private mDBPaidToParty          As Integer
Private mDBRecoveryRemitted     As Integer
Private mDBChequeCancelled      As Integer
Private mPaymentVoucherNo       As Variant
Private mDemandID               As Variant
Private mDemandNo               As Variant
Dim mPreviousYearMode           As Integer  '1 mPreviousYearMode set from Pending task,2 mPreviousYearMode set from Payorder cancel List
Dim mPendingTaskReqID           As Integer  'Pending task RequestID  -To update status
Dim mPendingTransactionDate     As Date


Private Sub chkChequeCancelledNoApp_Click()
    If chkChequeCancelledNoApp.Value = 1 Then
        chkChequeCancelledYesApp.Value = 0
    Else
        chkChequeCancelledYesApp.Value = 1
    End If
End Sub

Private Sub chkChequeCancelledNoInter_Click()
    If chkChequeCancelledNoInter.Value = 1 Then
        chkChequeCancelledYesInter.Value = 0
    Else
        chkChequeCancelledYesInter.Value = 1
    End If
End Sub

Private Sub chkChequeCancelledYesApp_Click()
    If chkChequeCancelledYesApp.Value = 1 Then
        chkChequeCancelledNoApp.Value = 0
    Else
        chkChequeCancelledNoApp.Value = 1
    End If
End Sub

Private Sub chkChequeCancelledYesInter_Click()
    If chkChequeCancelledYesInter.Value = 1 Then
        chkChequeCancelledNoInter.Value = 0
    Else
        chkChequeCancelledNoInter.Value = 1
    End If
End Sub

Private Sub chkPaidToPartyYesApp_Click()
    If chkPaidToPartyYesApp.Value = 1 Then
        chkPaidToPatyNoApp.Value = 0
    Else
        chkPaidToPatyNoApp.Value = 1
    End If
End Sub

Private Sub chkPaidToPartyYesInter_Click()
    If chkPaidToPartyYesInter.Value = 1 Then
        chkPaidToPatyNoInter.Value = 0
    Else
        chkPaidToPatyNoInter.Value = 1
    End If
End Sub

Private Sub chkPaidToPatyNoApp_Click()
    If chkPaidToPatyNoApp.Value = 1 Then
        chkPaidToPartyYesApp.Value = 0
    Else
        chkPaidToPartyYesApp.Value = 1
    End If
End Sub

Private Sub chkPaidToPatyNoInter_Click()
    If chkPaidToPatyNoInter.Value = 1 Then
        chkPaidToPartyYesInter.Value = 0
    Else
        chkPaidToPartyYesInter.Value = 1
    End If
End Sub

Private Sub chkRecoveriesPaidNoApp_Click()
    If chkRecoveriesPaidNoApp.Value = 1 Then
        chkRecoveriesPaidYesApp.Value = 0
    Else
        chkRecoveriesPaidYesApp.Value = 1
    End If
End Sub

Private Sub chkRecoveriesPaidNoInter_Click()
    If chkRecoveriesPaidNoInter.Value = 1 Then
        chkRecoveriesPaidYesInter.Value = 0
    Else
        chkRecoveriesPaidYesInter.Value = 1
    End If
End Sub

Private Sub chkRecoveriesPaidYesApp_Click()
    If chkRecoveriesPaidYesApp.Value = 1 Then
        chkRecoveriesPaidNoApp.Value = 0
    Else
        chkRecoveriesPaidNoApp.Value = 1
    End If
End Sub

    Private Sub chkRecoveriesPaidYesInter_Click()
        If chkRecoveriesPaidYesInter.Value = 1 Then
            chkRecoveriesPaidNoInter.Value = 0
        Else
            chkRecoveriesPaidNoInter.Value = 1
        End If
    End Sub

    Private Sub cmdCancel_Click()
        txtPayOrderNo.Tag = -1
        txtPayOrderNo.Text = ""
        Call InitForm
    End Sub

    Private Sub cmdCancelApprovedlDemandNo_Click()
        Dim mSql        As String
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mVoucherID  As Double
        Dim mPayOrderNo As String
        Dim objdb As New clsDB
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        'mCnn.BeginTrans
        mSql = " Select faVouchers.intVoucherID,faVouchers.intVoucherNo,intKeyID2 PayOrderNo From faVouchers Where numLinkKEyID in("
        mSql = mSql + vbNewLine + " Select faVouchers.intVoucherID intVoucherID From faReverseEntry Inner Join faPayOrder ON faPayOrder.vchPAyOrderNo=faReverseEntry.numDemandNo"
        mSql = mSql + vbNewLine + " Inner Join faVouchers On faPayOrder.vchPAyOrderNo=faVouchers.intKEyID2"
        mSql = mSql + vbNewLine + " Where faReverseEntry.tnyStatus=2 And tnyPaid=1 And faReverseEntry.tnyVoucherTypeID=50 And numLinkKeyId is Null And faVouchers.tnyVoucherTypeID in (40,20)"
        mSql = mSql + vbNewLine + " And faPayOrder.intPayOrderID=" & txtPayOrderNo.Tag & ")"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.BOF And Rec.EOF) Then
            While Not (Rec.EOF)
                If Rec!tnyVoucherTypeID = 40 Then
                     'On Error GoTo ErrRollBack:
                     mVoucherID = Rec!intVoucherID
                     mPayOrderNo = Rec!PayOrderNo
                     mSql = "Delete From faTransactionChild Where intTransactionID in (Select intTransactionID From faTransactions Where intVoucherID =" & mVoucherID & ")"
                     'mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faTransactions Where intVoucherID  =" & mVoucherID
                     'mCnn.Execute mSQL
                     mSql = "Delete From faVoucherAddress Where intVoucherID  =" & mVoucherID
                     'mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faVoucherChild Where intVoucherID  =" & mVoucherID
                     'mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faVouchers Where intVoucherID  =" & mVoucherID
                    ' mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faIDemandAddress Where numDemandID in (Select numDemandID From faIDemandTbl Where intKeyID2  =" & mPayOrderNo & ")"
                     'mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faIDemandChild Where numDemandID in (Select numDemandID From faIDemandTbl Where intKeyID2  =" & mPayOrderNo & ")"
                    ' mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Delete From faIDemandTbl Where intKeyID2  =" & mPayOrderNo
                     'mCnn.Execute mSQL
                     mSql = mSql + vbNewLine + "Update faReverseEntry set tnyPaid=0,tnyStatus=0,dtAuthorisationDateSec=Null,numAuthorisedBySec=Null,tnyRecoveryRemitted=Null,"
                     mSql = mSql + " numAuthorisedByAO=Null,dtAuthorisationDateAO=Null,tnyChequeCancelled=Null  Where numDemandNo  =" & mPayOrderNo
                     mSql = mSql + vbNewLine + " Update faPayOrder set tnyCancelled=0 Where intPayOrderID=" & txtPayOrderNo.Tag
                     objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                Else
                    mSql = "Update faReverseEntry set tnyPaid=0,tnyStatus=0,dtAuthorisationDateSec=Null,numAuthorisedBySec=Null,tnyRecoveryRemitted=Null,"
                    mSql = mSql + vbNewLine + " numAuthorisedByAO=Null,dtAuthorisationDateAO=Null,tnyChequeCancelled=Null  Where numDemandNo  =" & mPayOrderNo
                    mSql = mSql + vbNewLine + " Update faPayOrder set tnyCancelled=0 Where intPayOrderID=" & txtPayOrderNo.Tag
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                End If
                Rec.MoveNext
            Wend
            MsgBox "Approved Demand Cancelled Sucessfully", vbInformation
            cmdCancelApprovedlDemandNo.Visible = False
        Else
            mSql = " Select faVouchers.intVoucherID intVoucherID,numDemandNo From faReverseEntry Inner Join faPayOrder ON faPayOrder.vchPAyOrderNo=faReverseEntry.numDemandNo"
            mSql = mSql + vbNewLine + " Inner Join faVouchers On faPayOrder.vchPAyOrderNo=faVouchers.intKEyID2"
            mSql = mSql + vbNewLine + " Where faReverseEntry.tnyStatus=2 And tnyPaid=1 And faReverseEntry.tnyVoucherTypeID=50 And numLinkKeyId is Null And faVouchers.tnyVoucherTypeID in (40,20)"
            mSql = mSql + vbNewLine + " And faPayOrder.intPayOrderID=" & txtPayOrderNo.Tag
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.BOF And Rec.EOF) Then
                mPayOrderNo = Rec!numDemandNo
                mVoucherID = Rec!intVoucherID
                mSql = "Delete From faIDemandAddress Where numDemandID in (Select numDemandID From faIDemandTbl Where intKeyID2  =" & mPayOrderNo & ")"
                mSql = mSql + vbNewLine + "Delete From faIDemandChild Where numDemandID in (Select numDemandID From faIDemandTbl Where intKeyID2  =" & mPayOrderNo & ")"
                mSql = mSql + vbNewLine + "Delete From faIDemandTbl Where intKeyID2  =" & mPayOrderNo
                mSql = mSql + vbNewLine + "Update faReverseEntry set tnyPaid=0,tnyStatus=0,dtAuthorisationDateSec=Null,numAuthorisedBySec=Null,tnyRecoveryRemitted=Null,"
                mSql = mSql + " numAuthorisedByAO=Null,dtAuthorisationDateAO=Null,tnyChequeCancelled=Null  Where numDemandNo  =" & mPayOrderNo
                mSql = mSql + vbNewLine + " Update faPayOrder set tnyCancelled=0 Where intPayOrderID=" & txtPayOrderNo.Tag
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                MsgBox "Approved Demand Cancelled Sucessfully", vbInformation
                cmdCancelApprovedlDemandNo.Visible = False
            End If
        End If
'        mCnn.CommitTrans
'        Exit Sub
'
'ErrRollBack:
'
'        MsgBox "Cancel Failed ", vbInformation
'        mCnn.RollbackTrans
    End Sub

    Private Sub cmdCancelRequest_Click()
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If mCancelStatus = 0 Then
            If MsgBox("Do you really want to Remove this Cancel Request", vbYesNo) = vbYes Then
                mCnn.Execute "Update faReverseEntry Set tnyStatus = 4 Where intRequestID = " & mRequestID
                MsgBox "Removed Cancellation Request", vbInformation
            End If
        ElseIf mCancelStatus = 4 Then
            MsgBox "This Request is Already Cancelled"
        Else
            MsgBox "Invalid payorder Status"
        End If
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub InitForm()
        txtPayorderDate.Text = ""
        txtPaymentType.Tag = -1
        txtPaymentType.Text = ""
        txtPayee.Text = ""
        txtNetAmount.Text = ""
        txtStatus.Tag = -1
        txtStatus.Text = ""
        txtFunctionary.Tag = -1
        txtFunctionary.Text = ""
        txtFunction.Tag = -1
        txtFunction.Text = ""
        txtSourceofFund.Tag = -1
        txtSourceofFund.Text = ""
        
        cmbCancelReasons.ListIndex = -1
        txtForwardedSeat.Tag = -1
        txtForwardedSeat.Text = ""
        txtRemarks.Text = ""
        
        vsGridRecoveries.Rows = 1
        vsGridRecoveries.Rows = 5
        vsGriidVouchers.Rows = 1
        vsGriidVouchers.Rows = 5
        
        txtReqSeat.Tag = -1
        txtReqSeat.Text = ""
        txtReqUser.Tag = -1
        txtReqUser.Text = ""
        txtAppSeat.Tag = -1
        txtAppSeat.Text = ""
        txtAppUser.Tag = -1
        txtAppUser.Text = ""
        
        txtDemandNo.Tag = -1
        txtDemandNo.Text = ""
        
        txtCancelStatus.Tag = -1
        txtCancelStatus.Text = ""
        
        txtProceddingsNo.Tag = -1
        txtProceddingsNo.Text = ""
        txtProceedingsDate.Text = ""
        txtProceedingsRemarks.Text = ""
        
        chkPaidToPartyYesInter.Value = 0
        chkPaidToPatyNoInter.Value = 0
        chkChequeCancelledYesInter.Value = 0
        chkChequeCancelledNoInter.Value = 0
        chkRecoveriesPaidYesInter.Value = 0
        chkRecoveriesPaidNoInter.Value = 0
        chkPaidToPartyYesApp.Value = 0
        chkPaidToPatyNoApp.Value = 0
        chkChequeCancelledYesApp.Value = 0
        chkChequeCancelledNoApp.Value = 0
        chkRecoveriesPaidYesApp.Value = 0
        chkRecoveriesPaidNoApp.Value = 0
        
        txtReqSeatCancel.Tag = -1
        txtReqSeatCancel.Text = ""
        txtReqUserCancel.Tag = -1
        txtReqUserCancel.Text = ""
        txtApprovedAO.Tag = -1
        txtApprovedAO.Text = ""
        txtApprovedSec.Tag = -1
        txtApprovedSec.Text = ""
        
        mRequestID = -1
        mCancelStatus = -1
        mPayorderStatus = 0
        mPayorderRequested = -1
        mPayorderRequestedSeat = -1
        mPayorderApproved = -1
        mPayorderApprovedSeat = -1
        mDBPaidToParty = -1
        mDBRecoveryRemitted = -1
        mDBChequeCancelled = -1
        mPaymentVoucherNo = -1
        mDemandID = -1
        mDemandNo = -1
        
        vsGridRecoveries.Rows = 1
        vsGridRecoveries.Rows = 5
        vsGriidVouchers.Rows = 1
        vsGriidVouchers.Rows = 5
    End Sub

    Private Sub cmdForwardedSeat_Click()
        Dim mSql    As String
        gbSearchID = -1
        gbSearchStr = ""
        mSql = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (3,5,6) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        frmSearchSeat.SQLString = mSql
        frmSearchSeat.Show vbModal
        If gbSearchID > 0 Then
            txtForwardedSeat.Tag = gbSearchID
            txtForwardedSeat.Text = gbSearchStr
        End If
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

Private Sub cmdPayorderSearch_Click()
    frmSearchPaymentOrder.Staus = 1
'    frmSearchPaymentOrder.chkListToApprove.value = 1
    frmSearchPaymentOrder.chkListToApprove.Visible = True
    frmSearchPaymentOrder.Show vbModal
    If gbSearchID > 0 Then
        txtPayOrderNo.Tag = gbSearchID
        txtPayOrderNo.Text = gbSearchStr
        gbSearchID = -1
        gbSearchStr = ""
        Call txtPayOrderNo_LostFocus
    End If
End Sub

Private Sub cmdProceedings_Click()
    gbSearchID = -1
    gbSearchStr = ""
    frmProceedings.chkEdit.Value = 0
    frmProceedings.Module = 70
    frmProceedings.Show vbModal
    If gbSearchID > 0 Then
        Dim objProceedings As New clsProceedings
        With objProceedings
            .ProceedingsID = gbSearchID
            .getProceedingsByID
            If .Used > 0 Then
                MsgBox "This Proceedings already used", vbInformation
                .ProceedingsID = -1
            Else
                txtProceddingsNo.Tag = .ProceedingsID
                txtProceddingsNo.Text = .ProceedingsNo
                txtProceedingsDate.Text = CheckDateInMMM(.ProceedingsDate)
                txtProceedingsRemarks.Text = .Remarks
            End If
        End With
    End If
    gbSearchID = -1
    gbSearchStr = ""
End Sub

Private Sub cmdSearchFunction_Click()
    gbSearchStr = ""
    gbSearchID = -1
    frmSearchFunction.Show vbModal
    If gbSearchID > 0 Then
        txtFunction.Tag = gbSearchID
        txtFunction.Text = gbSearchStr
    End If
    gbSearchStr = ""
    gbSearchID = -1
End Sub

Private Sub cmdSearchFunctionary_Click()
    gbSearchStr = ""
    gbSearchID = -1
    frmSearchFunctionary.Show vbModal
    If gbSearchID > 0 Then
        txtFunctionary.Tag = gbSearchID
        txtFunctionary.Text = gbSearchStr
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub

Private Sub cmdSearchSourceofFund_Click()
    gbSearchStr = ""
    gbSearchID = -1
    
    frmSearchMasters.Connection = enuSourceString.Saankhya
    frmSearchMasters.QrySP = Qyery
    frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
    frmSearchMasters.Show vbModal
    If gbSearchID <> -1 Then
        txtSourceofFund.Text = gbSearchStr
        txtSourceofFund.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub

Private Sub cmdSearchTransactionType_Click()
    gbSearchStr = ""
    gbSearchID = -1
    
    frmSearchTransactionType.ModeOfTransaction = 2
    frmSearchTransactionType.Show vbModal
    If gbSearchID > 0 Then
        txtPaymentType.Tag = gbSearchID
        txtPaymentType.Text = gbSearchStr
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub



'Private Sub Form_Activate()
'    Me.Top = 0
'    Me.Left = 0
'End Sub

Private Sub Form_Load()
    Call InitForm
    PopulateList cmbCancelReasons, "Select vchReason,intReasonID From faReasons Where intType = 70", , True, True, True
    If mPreviousYearMode = 1 Then
        GetPendingTaskDetails
    End If
    If val(txtPayOrderNo.Text) > 0 Then
        Call FillPayOrder(txtPayOrderNo.Text)
    End If
End Sub
    Private Sub ReverseSulekhaExpenseDetails(ByVal mVoucherID As Double)
        Dim mSql    As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim objAcc  As New clsAccounts
        Dim arrIn   As Variant
        If objdb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
            arrIn = Array(mVoucherID)
            objdb.ExecuteSP "ExpenseDetails_U", arrIn, , , mCnn
        End If
        
    End Sub
Private Sub FillPayOrder(ByVal mValPaymentOrderNo As Double)
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim objAcc As New clsAccounts
  
    
    
    mPayorderStatus = 0
    mCancelStatus = -1
    mDemandID = -1
    mDemandNo = ""
    txtDemandNo.Tag = -1
    txtDemandNo.Text = ""
    fraNextLevel.Enabled = False
    FraProceedings.Enabled = True
    
    mDBPaidToParty = 0
    mDBRecoveryRemitted = 0
    mDBChequeCancelled = 0
    cmdSave.Enabled = True
    fraInter.Enabled = False
    fraFinal.Enabled = False
    vsGridRecoveries.Editable = flexEDNone
    '-------------------------------'
    ' Validations
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Sub
    End If
    '-------------------------------'
    Call InitForm
    '-----------------------------------------'
    '               Demand Details            '
    
    mSql = "SELECT numDemandID,vchDemandNo FROM faIDemandTBL Where tnyStatus <> 3 And intKeyID2 = " & mValPaymentOrderNo
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        mDemandID = Rec!numDemandID
        mDemandNo = Rec!vchDemandNo
        txtDemandNo.Tag = mDemandID
        txtDemandNo.Text = mDemandNo
    End If
    Rec.Close
    '------------------------------------------------------------------------'
    '                       Payorder Cancel Details                          '
    mSql = "Select  faReverseEntry.intRequestID,faReverseEntry.intReasonID,faReasons.vchReason,faReasons.intType,faReasons.intCategory,numForwardedSeatID,dbo.fnGetSeat(numForwardedSeatID) ForwardedSeat," & vbNewLine
    mSql = mSql + "     faReverseEntry.vchRemarks CancelRemarks,faProceedings.intProceedingsID,faProceedings.vchProceedingsNo,faProceedings.dtProceedingsDate,faProceedings.vchRemarks,faReverseEntry.tnyStatus," & vbNewLine
    mSql = mSql + "     faReverseEntry.tnyPaid,faReverseEntry.tnyRecoveryRemitted,faReverseEntry.tnyChequeCancelled,numRequestedSeatID,dbo.fnGetSeat(numRequestedSeatID) RequestedSeat,numRequestedUserID,dbo.fnGetUser(numRequestedUserID) RequestedUser,numAuthorisedByAO,dbo.fnGetUser(numAuthorisedByAO) ApprovedAO, dbo.fnGetUser(numAuthorisedBySec) ApprovedSec " & vbNewLine
    mSql = mSql + "From faReverseEntry" & vbNewLine
    mSql = mSql + "Inner Join faReasons On faReverseEntry.intReasonID = faReasons.intReasonID" & vbNewLine
    mSql = mSql + "Left Join faProceedings On faReverseEntry.numDemandNo = faProceedings.intVoucherNo" & vbNewLine
    mSql = mSql + "Where tnyStatus <> 3 And numDemandNo = " & mValPaymentOrderNo
    Rec.Open mSql, mCnn
    cmbCancelReasons.ListIndex = -1
    If Not (Rec.EOF And Rec.BOF) Then
        mRequestID = IIf(IsNull(Rec!intRequestID), -1, Rec!intRequestID)
        txtRemarks.Text = IIf(IsNull(Rec!CancelRemarks), "", Rec!CancelRemarks)
        cmbCancelReasons.Text = IIf(IsNull(Rec!vchReason), "", Rec!vchReason)
        txtForwardedSeat.Tag = IIf(IsNull(Rec!numForwardedSeatID), -1, Rec!numForwardedSeatID)
        txtForwardedSeat.Text = IIf(IsNull(Rec!ForwardedSeat), "", Rec!ForwardedSeat)
'        txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
        mCancelStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        txtCancelStatus.Tag = IIf(IsNull(Rec!tnyStatus), -1, Rec!tnyStatus)
        If IIf(IsNull(Rec!tnyStatus), -1, Rec!tnyStatus) = 0 Then
            txtCancelStatus.Text = "Requested Only"
           ' cmdSave.Caption = "First Level Approve"
        ElseIf Rec!tnyStatus = 1 Then
            txtCancelStatus.Text = "First Level Approved"
            'cmdSave.Caption = "Final Approve"
            'cmdSave.Caption = "First Level Approve"
        ElseIf Rec!tnyStatus = 2 Then
            cmdSave.Enabled = False
            txtCancelStatus.Text = "Final Level Approved"
            'cmdSave.Caption = "Final Approve"
        Else
            txtCancelStatus.Text = "Request Rejected"
        End If
        
        txtProceddingsNo.Tag = IIf(IsNull(Rec!intProceedingsID), -1, Rec!intProceedingsID)
        txtProceddingsNo.Text = IIf(IsNull(Rec!vchProceedingsNo), "", Rec!vchProceedingsNo)
        txtProceedingsDate.Text = IIf(IsNull(Rec!dtProceedingsDate), "", Rec!dtProceedingsDate)
        If Trim(txtProceedingsDate.Text) <> "" Then
            txtProceedingsDate.Text = CheckDateInMMM(txtProceedingsDate.Text)
        End If
        txtProceedingsRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
        txtReqSeatCancel.Tag = IIf(IsNull(Rec!numRequestedSeatID), -1, Rec!numRequestedSeatID)
        txtReqSeatCancel.Text = IIf(IsNull(Rec!RequestedSeat), "", Rec!RequestedSeat)
        txtReqUserCancel.Tag = IIf(IsNull(Rec!numRequestedUserID), -1, Rec!numRequestedUserID)
        txtReqUserCancel.Text = IIf(IsNull(Rec!RequestedUser), "", Rec!RequestedUser)
        txtApprovedAO.Tag = IIf(IsNull(Rec!numAuthorisedByAO), -1, Rec!numAuthorisedByAO)
        txtApprovedAO.Text = IIf(IsNull(Rec!ApprovedAO), "", Rec!ApprovedAO)
        txtApprovedSec.Text = IIf(IsNull(Rec!ApprovedSec), "", Rec!ApprovedSec)
        
        mDBPaidToParty = IIf(IsNull(Rec!tnyPaid), 0, Rec!tnyPaid)
        mDBRecoveryRemitted = IIf(IsNull(Rec!tnyRecoveryRemitted), 0, Rec!tnyRecoveryRemitted)
        mDBChequeCancelled = IIf(IsNull(Rec!tnyChequeCancelled), 0, Rec!tnyChequeCancelled)
        '-----------------------------------------------------------'
        '                   Check Box Certifications                '
        chkPaidToPartyYesInter.Value = mDBPaidToParty
        chkChequeCancelledYesInter.Value = mDBChequeCancelled
        chkRecoveriesPaidYesInter.Value = mDBRecoveryRemitted
        If mCancelStatus = 2 Then
            chkPaidToPartyYesApp.Value = mDBPaidToParty
            chkChequeCancelledYesApp.Value = mDBChequeCancelled
            chkRecoveriesPaidYesApp.Value = mDBRecoveryRemitted
        End If
        '-----------------------------------------------------------'
    End If
    Rec.Close
    If mCancelStatus = 0 Or mCancelStatus = 1 Then
        fraNextLevel.Enabled = True
    End If
    '------------------------------------------------------------------------'
    '-----------------------------------------'
    '           Payorder Details              '
    mSql = "Select  faPayorder.intPayOrderID,faPayorder.vchPayOrderNo,faPayorder.dtPayOrderDate,faPayorder.dtDueDate,faPayorder.intFunctionaryID,dbo.fnGetFunctionary(faPayorder.intFunctionaryID) Functionary,faPayorder.intFunctionID,dbo.fnGetFunction(faPayorder.intFunctionID) [Function],faPayorder.intTransactionTypeID,dbo.fnGetTransactionType(intTransactionTypeID)TransactionType,faPayorder.vchBillNo,faPayorder.numBillAmount,faPayorder.dtBillDate,faPayorder.intInstrumentTypeID, " & vbNewLine
    mSql = mSql + "faPayorder.intCashOrBankHeadID,faPayorder.vchDescription,faPayorder.vchTitle,faPayorder.intSubLedgerTypeID,faPayorder.intPayToSubLedgerID,faPayorder.intSubsidiaryCashBookID,faPayorder.intImplementingOfficerID,faPayorder.numProjectNo,faPayorder.intStockRegisterID," & vbNewLine
    mSql = mSql + "faPayorder.vchStockRefNo,faPayorder.intAssetTypeID,faPayorder.intAssetID,faPayorder.numFwdSeatID,faPayorder.intLocalBodyID,faPayorder.intZonalID,faPayorder.intFinancialYearID,faPayorder.numUserID,faPayorder.numSeatID,faPayorder.numApprovingOfficerID," & vbNewLine
    mSql = mSql + "faPayorder.numApprovingSeatID,faPayorder.dtApprovingDate,faPayorder.intVoucherID,faPayorder.intVoucherNo,faPayorder.dtVoucherDate,faPayorder.tnyStatus,faPayorder.intKeyID,faPayorder.numKeyID,faPayorder.dtKeyDate,faReverseEntryTrChild.intRequestID," & vbNewLine
    mSql = mSql + "faPayorder.tnyCancelled , faPayorder.intAppID, faPayorder.intModuleID, faPayorder.intSourceOfFundID,dbo.fnGetSourceOfFund(intSourceOfFundID)SourceOfFund, faPayorder.intAllotmentID, faPayorder.intAgreementID, faPayorder.tnyCategoryID, faPayorder.tnySectorID, faPayorder.tnyIsFinalBill, faPayOrderChild.numAmount,faPayOrderAddress.vchName," & vbNewLine
    mSql = mSql + "dbo.fnGetUser(faPayorder.numUserID) [User],dbo.fnGetSeat(faPayorder.numSeatID)Seat,dbo.fnGetUser(faPayorder.numApprovingOfficerID)Approver,dbo.fnGetSeat(faPayorder.numApprovingSeatID)ApproverSeat,faPayOrderChild.intAccountHeadID,faPayOrderChild.vchAccountHeadCode,faPayOrderChild.tnyCategoryFlag" & vbNewLine
    mSql = mSql + "From faPayorder" & vbNewLine
    mSql = mSql + "Inner Join faPayOrderChild On faPayorder.intPayOrderID = faPayOrderChild.intPayOrderID" & vbNewLine
    mSql = mSql + "Inner Join faPayOrderAddress On faPayorder.intPayOrderID = faPayOrderAddress.intPayOrderID" & vbNewLine
    mSql = mSql + "Left Join faReverseEntryTrChild On faReverseEntryTrChild.intRequestID = " & mRequestID & " And faPayOrderChild.intAccountHeadID = faReverseEntryTrChild.intAccountHeadID" & vbNewLine
    mSql = mSql + "Where faPayorder.vchPayorderNo = " & mValPaymentOrderNo
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        txtPayOrderNo.Tag = IIf(IsNull(Rec!intPayOrderID), -1, Rec!intPayOrderID)
        txtPayOrderNo.Text = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
        txtPayorderDate.Text = IIf(IsNull(Rec!dtPayOrderDate), "", Rec!dtPayOrderDate)
        If Trim(txtPayorderDate.Text) <> "" Then
            txtPayorderDate.Text = CheckDateInMMM(txtPayorderDate.Text)
        End If
        txtPaymentType.Tag = IIf(IsNull(Rec!intTransactionTypeID), -1, Rec!intTransactionTypeID)
        txtPaymentType.Text = IIf(IsNull(Rec!TransactionType), "", Rec!TransactionType)
        txtPayee.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
        txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), -1, Rec!intFunctionaryID)
        txtFunctionary.Text = IIf(IsNull(Rec!Functionary), "", Rec!Functionary)
        txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), -1, Rec!intFunctionID)
        txtFunction.Text = IIf(IsNull(Rec!Function), "", Rec!Function)
        txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceOfFundID), -1, Rec!intSourceOfFundID)
        txtSourceofFund.Text = IIf(IsNull(Rec!SourceOfFund), "", Rec!SourceOfFund)
'        txtRemarks.Text = Rec!vchDescription ' This is Cancellation Remark
        txtReqSeat.Tag = IIf(IsNull(Rec!numSeatID), -1, Rec!numSeatID)
        mPayorderRequestedSeat = IIf(IsNull(Rec!numSeatID), -1, Rec!numSeatID)
        txtReqSeat.Text = IIf(IsNull(Rec!Seat), "", Rec!Seat)
        mPayorderRequested = IIf(IsNull(Rec!numUserID), -1, Rec!numUserID)
        txtReqUser.Tag = IIf(IsNull(Rec!numUserID), -1, Rec!numUserID)
        txtReqUser.Text = IIf(IsNull(Rec!User), "", Rec!User)
        txtAppSeat.Tag = IIf(IsNull(Rec!numApprovingSeatID), -1, Rec!numApprovingSeatID)
        txtAppSeat.Text = IIf(IsNull(Rec!ApproverSeat), "", Rec!ApproverSeat)
        mPayorderApproved = IIf(IsNull(Rec!numApprovingOfficerID), -1, Rec!numApprovingOfficerID)
        mPayorderApprovedSeat = IIf(IsNull(Rec!numApprovingSeatID), -1, Rec!numApprovingSeatID)
        txtAppUser.Tag = IIf(IsNull(Rec!numApprovingOfficerID), -1, Rec!numApprovingOfficerID)
        txtAppUser.Text = IIf(IsNull(Rec!Approver), "", Rec!Approver)
        mPayorderStatus = IIf(IsNull(Rec!tnyStatus), 0, Rec!tnyStatus)
        vsGridRecoveries.Rows = 1
        While Not Rec.EOF
            If Rec!tnyCategoryFlag = 3 Then     ' Checking Net
                txtNetAmount.Text = Format(IIf(IsNull(Rec!numAmount), 0, Rec!numAmount), "0.00")
            End If
            If Rec!tnyCategoryFlag = 2 Then     ' Checking Recoveries
                If val(Rec!vchAccountHeadCode) >= 350200100 And val(Rec!vchAccountHeadCode) <= 350309900 Then '// Recovery Heads
                    With vsGridRecoveries
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID)
                        .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                        objAcc.SetAccountID (IIf(IsNull(Rec!intAccountHeadID), -1, Rec!intAccountHeadID))
                        .TextMatrix(.Rows - 1, 2) = objAcc.AccountHead
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rec!numAmount), 0, Rec!numAmount)
                        If IsNull(Rec!intRequestID) = False Then
                            .Cell(flexcpChecked, .Rows - 1, 4) = 1
                        Else
                            .Cell(flexcpChecked, .Rows - 1, 4) = 0
                        End If
                    End With
                End If
            End If
            Rec.MoveNext
        Wend
    End If
    Rec.Close
    '-----------------------------------------------------------'
    '                       Voucher Details                     '
    
    mPaymentVoucherNo = -1
    vsGriidVouchers.Rows = 1
    mSql = "Select faVouchers.intVoucherID,faVouchers.tnyVoucherTypeID, faVouchers.intVoucherNo,faVouchers.intInstrumentTypeID,decProjectID From faVouchers "
    mSql = mSql + "Left JOIN faVoucherSub ON faVouchers.intVoucherID = faVoucherSub.intVoucherID "
    mSql = mSql + " Where intKeyID2 = " & mValPaymentOrderNo
    Rec.CursorLocation = adUseClient
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        
        '
        'NOTE:- IF Project Related Vouhcer then CHECK CONNECTION TO SULEKHA
        
        Rec.Find "tnyVoucherTypeID = 20 "
'        Commented on 31/may/2018
'        If Not IsNull(Rec!decProjectID) Then
'            Dim mCnnSulekha As New ADODB.Connection
'            If Not (objDB.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
'            MsgBox "Connection To Plan [Sulekha] Module not found", vbCritical
'            Call InitForm
'            Exit Sub
'            End If
'        End If
        'END OF CHECKING CONNECTION TO SULEKHA
        
        Rec.MoveFirst
        While Not Rec.EOF
            With vsGriidVouchers
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
                .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rec!tnyVoucherTypeID), -1, Rec!tnyVoucherTypeID)
                .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                If Rec!tnyVoucherTypeID <> 20 And mPayorderStatus < 2 Then
                    mPayorderStatus = 1
                Else
                    mPaymentVoucherNo = IIf(IsNull(Rec!intVoucherNo), -1, Rec!intVoucherNo)
                    mPayorderStatus = 2
                End If
            End With
            Rec.MoveNext
        Wend
        
    End If
    Rec.Close
    
    '''---------------------------------------'
    '''    Added On 23.11.12 By Anisha C
    '''---------------------------------------'
    Rec.Open "Select intFinancialYearID From faVouchers Where tnyVoucherTypeID=20 And intKeyID2 = " & mValPaymentOrderNo, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        If mPreviousYearMode = 1 Then
            If Rec!intFinancialYearID <> gbFinancialYearID - 1 Then
                MsgBox "Payment voucher FinacialYear is not matching", vbInformation
            End If
        Else
            If Rec!intFinancialYearID <> gbFinancialYearID Then
                MsgBox "Payment voucher FinacialYear is not matching", vbInformation
            End If
        End If
    End If
    Rec.Close
    '''---------------------------------------'
    
    '-----------------------------------------'
    '             Payorder Status             '
    txtStatus.Tag = mPayorderStatus
    If mPayorderStatus = 0 Then
        txtStatus.Text = "Not Approved"
    ElseIf mPayorderStatus = 1 Then
        txtStatus.Text = "Approved"
    Else
        txtStatus.Text = "Approved and Paid"
    End If
    If mPayorderStatus = 2 Then
        If mCancelStatus = 0 And gbSeatID <> txtReqSeatCancel.Tag Then
            fraNextLevel.Enabled = True
            fraInter.Enabled = True
            vsGridRecoveries.Editable = flexEDKbdMouse
        ElseIf mCancelStatus = 1 Then
            fraNextLevel.Enabled = True
            fraFinal.Enabled = True
            vsGridRecoveries.Editable = flexEDKbdMouse
        End If
    End If
End Sub
Private Sub GetPendingTaskDetails()
    Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
            'Dim mSourceID   As Variant
            'Dim mProjectId  As Variant
            'Dim mCategoryID As Integer
            'Dim mSubSectorID As Variant
            'Dim objProj     As New clsProject
            'Dim objProFund  As New clsProjectFund
            'Dim mCol        As Collection
            'Dim mRow        As Integer
        Dim mTaskID     As Integer
        'Dim objTr As New clsTransactionType
        'Dim mTrTypeID As Integer
        
        
        'On Error GoTo Err
        If mPendingTaskReqID > 0 Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSql = "Select * from faPendingTaskRequest Where intRequestID= " & mPendingTaskReqID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mPendingTransactionDate = Rec!dtTransactionDate
                    
                End If
                Rec.Close
            End If
        End If
End Sub

Private Function GetSeatGroup(ByVal numSeatID As Variant) As Integer
    Dim mSql As String
    Dim Rec                 As New ADODB.Recordset
    Dim objdb               As New clsDB
    Dim mCn               As New ADODB.Connection
    Dim mSeatGroup      As Integer
    If objdb.CreateNewConnection(mCn, enuSourceString.Saankhya) Then
        mSql = "Select intGroupID,* From faSeats Where numSeatID=" & numSeatID
        Rec.Open mSql, mCn
        If Not (Rec.EOF Or Rec.BOF) Then
            mSeatGroup = Rec!intGroupID
        Else
            mSeatGroup = 0
        End If
        Rec.Close
    End If
    mCn.Close
    GetSeatGroup = mSeatGroup
End Function
Private Sub cmdSave_Click()
    Dim mSql                As String
    Dim mCnn                As New ADODB.Connection
    Dim Rec                 As New ADODB.Recordset
    Dim objdb               As New clsDB
    Dim mArrayInput         As Variant
    Dim mArrayOut           As Variant
    
    Dim mPaidToParty        As Integer
    Dim mRecoveryRemitted   As Integer
    Dim mChequeCancelled    As Integer
    
    Dim mPayorderID         As Variant
    Dim mPayOrderNo         As Variant
    Dim mCancelReason       As Integer
    Dim mForwardedSeat      As Variant
    Dim mBoolCancel         As Boolean
    Dim mReverseID          As Double
    Dim mVoucherID          As Double ''Used to pass voucherid to udate cancel flag in sulekha database added on 26-12-12
    Dim mTrnDate            As Date
    Dim mCurFinYear         As Integer
    Dim mSeatgroupID        As Integer
    '----------------------------------------------------------------------------------'
    
    mVoucherID = 0
    
    ' Validations
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Sub
    End If
    
    If val(txtPayOrderNo.Tag) < 1 Then
        MsgBox "Please Enter Payorder Number", vbInformation
        Exit Sub
    End If
    If cmbCancelReasons.ListIndex < 1 Then
        MsgBox "Please Select a Reason", vbInformation
        Exit Sub
    End If
    If val(txtForwardedSeat.Tag) < 1 Then
        MsgBox "Please select Forwarded Seat", vbInformation
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------'
    mPaidToParty = 0
    mRecoveryRemitted = 0
    mChequeCancelled = 0
    
    mBoolCancel = False
    '-------------------------------'
    '   Checking Cetifications      '
    If mPayorderStatus = 2 Then
        If mCancelStatus = 0 Or gbUserID = txtApprovedAO.Tag Then
            If chkPaidToPartyYesInter.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesInter.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesInter.Value = 1 Then
                mChequeCancelled = 1
            End If
        ElseIf mCancelStatus = 1 Then
            If chkPaidToPartyYesApp.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesApp.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesApp.Value = 1 Then
                mChequeCancelled = 1
            End If
        End If
    End If
    '-------------------------------'
    
    mSeatgroupID = GetSeatGroup(mPayorderRequestedSeat)
    
    
    mPayorderID = txtPayOrderNo.Tag
    mPayOrderNo = Trim(txtPayOrderNo.Text)
    mCancelReason = cmbCancelReasons.ItemData(cmbCancelReasons.ListIndex)
    mForwardedSeat = txtForwardedSeat.Tag
    If mPreviousYearMode = 1 Then
        mTrnDate = mPendingTransactionDate
        mCurFinYear = gbFinancialYearID - 1
    Else
        mTrnDate = gbTransactionDate
        mCurFinYear = gbFinancialYearID
    End If
    
    
    
StartCase:
    Select Case mCancelStatus   '///status From reverse Entry table tnystatus
        Case -1     '// Not Requested  (For New PayOrder cancel Request)
            
            'NOTE: Check Proceedings Details ' Added by Aiby on 13-Nov-2011
            If Not (val(txtProceddingsNo.Tag) > 0) Then
                MsgBox "Please enter the Proceedings Number", vbInformation
                txtProceddingsNo.SetFocus
                Exit Sub
            End If
            cmdSave.Enabled = False ''' Added On 3 Aug 2015 By Anisha To avoid receipt doubbling
            Select Case mPayorderStatus
                Case 0
                    If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                        'If gbUserID = mPayorderRequested Then      ' Changed to Seat Validation
                        'If gbSeatID = mPayorderRequestedSeat Then ' Changed for panchayat (Forward)
                        If gbSeatID = mPayorderRequestedSeat And gbLBPanchayat = 1 Then
                            '// Code for Save Payorder cancellations
                            '// Table faReverse Entry
                            mArrayInput = Array(-1, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                            objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                            '-----------------------------------------------------------'
                            '                       Proceedings Updation                '
                            Call UpdateProceedings(0)
                            '-----------------------------------------------------------'
                        ElseIf gbSeatID = mPayorderRequestedSeat And gbLBPanchayat = 1 Or mSeatgroupID = gbSeatGroupAccountSectionClerk Then   ' Changed to Seat Validation
                            '// Code for Save Payorder cancellations
                            '// Table faReverse Entry
                            mArrayInput = Array(-1, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                            objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                            '-----------------------------------------------------------'
                            '                       Proceedings Updation                '
                            Call UpdateProceedings(0)
                            '-----------------------------------------------------------'
                        Else
                            MsgBox "This Payorder Generated at " & txtReqSeat.Text & vbNewLine & "Please Request PO Cancel through " & txtReqSeat.Text, vbInformation
                            Exit Sub
                        End If
                    Else
                        MsgBox "An Accounts Clerk can Apply for Cancellation", vbInformation
                        Exit Sub
                    End If
                Case 5, 3 'For Panchayat Forwarded PO
                    If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                        If gbSeatID = mPayorderRequestedSeat And gbLBPanchayat = 1 Then
                            '// Code for Save Payorder cancellations
                            '// Table faReverse Entry
                            mArrayInput = Array(-1, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                            objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                            '-----------------------------------------------------------'
                            '                       Proceedings Updation                '
                            Call UpdateProceedings(0)
                            '-----------------------------------------------------------'
                        Else
                            MsgBox "This Payorder Generated at " & txtReqSeat.Text & vbNewLine & "Please Request PO Cancel through " & txtReqSeat.Text, vbInformation
                            Exit Sub
                        End If
                    End If
                Case 1, 2  '//PayOrder Approved And Paid =2,PayOrder Approved=1
                    If gbLBType = 4 Then '// Corporation
                        If gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                            'If gbUserID = mPayorderApproved Then       ' Changed to Seat Validation
                            If gbSeatID = mPayorderApprovedSeat Then   ' Changed to Seat Validation
                                '// Save in Pay order Details
                                '// Table faReverseEntry
                                mArrayInput = Array(-1, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                mPaymentVoucherNo, Null, Null, Null, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                                '-----------------------------------------------------------'
                                '                       Proceedings Updation                '
                                Call UpdateProceedings(0)
                                '-----------------------------------------------------------'
                            Else
                                MsgBox "This Payorder Approved By " & txtAppUser.Text, vbInformation
                                Exit Sub
                            End If
                        Else
                            MsgBox "An Accounts Supdt can only Apply for Cancellation", vbInformation
                            Exit Sub
                        End If
                    Else                 '// Panchayats or Municipalities
                        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                            'If gbUserID = mPayorderRequested Then      ' Changed to Seat Validation
                            If gbSeatID = mPayorderRequestedSeat Then   ' Changed to Seat Validation
                                '// Code for Save Payorder cancellations
                                '// Table faReverse Entry
                                mArrayInput = Array(-1, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                mPaymentVoucherNo, Null, Null, Null, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                                '-----------------------------------------------------------'
                                '                       Proceedings Updation                '
                                Call UpdateProceedings(0)
                                'NOTE:- PENDING TASK
                                If mPreviousYearMode = 1 Then
                                    mSql = "Update faPendingTaskRequest SET tnyStatus = 8 Where intRequestID=" & mPendingTaskReqID & "  "
                                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                                End If
                                '-----------------------------------------------------------'
                            Else
                                'MsgBox "This Payorder Requested By " & txtReqUser.Text, vbInformation
                                MsgBox "This Payorder Generated at " & txtReqSeat.Text & vbNewLine & "Please Request PO Cancel through" & txtReqSeat.Text, vbInformation
                                Exit Sub
                            End If
                        Else
                            MsgBox "An Accounts Clerk can only Apply for Cancellation", vbInformation
                            Exit Sub
                        End If
                    End If
            End Select
            
        Case 0    '// Requested But Not Approved (Payment order Cancel Request)
            Select Case mPayorderStatus
                Case 0, 5, 3 '// PayOrder Status (Pay Order  not approved)
                    If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then '// Edit Request Stage
                        'If gbUserID = mPayorderRequested Then          ' Changed to Seat Validation
                        If gbSeatID = mPayorderRequestedSeat Then   ' Changed to Seat Validation
                            '// Code for Save Payorder cancellations
                            '// Table faReverse Entry
                            mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                0, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                            objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                            '-----------------------------------------------------------'
                            '                       Proceedings Updation                '
                            Call UpdateProceedings(0)
                            '-----------------------------------------------------------'
                        Else
                            MsgBox "this Payorder Requested By " & txtReqUser.Text, vbInformation
                            Exit Sub
                        End If
                    Else    '// Approval Stage
                        If gbLBType = 4 Then
                            If gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                                '// Edit Cancel Payorder
                                mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), mPayorderRequested, txtReqSeat.Tag, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                2, gbUserID, gbTransactionDate, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                                '// Cancel Payorder
                                mSql = "Update faPayOrder Set tnyCancelled = 1 Where vchPayorderNo = '" & mPayOrderNo & "'"
                                mCnn.Execute mSql
                                '-----------------------------------------------------------'
                                '                       Proceedings Used                    '
                                Call UpdateProceedings(1)
                                '-----------------------------------------------------------'
                            Else
                                MsgBox "An Accounts Supdt Can Cancel the Payorder", vbInformation
                                Exit Sub
                            End If
                        Else       '// Panchayats or Municipalities
                            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                                '// Edit Cancel Payorder
                                mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), mPayorderRequested, txtReqSeat.Tag, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                2, gbUserID, gbTransactionDate, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                                '// Cancel Payorder
                                mSql = "Update faPayOrder Set tnyCancelled = 1 Where vchPayorderNo = '" & mPayOrderNo & "'"
                                mCnn.Execute mSql
                                '-----------------------------------------------------------'
                                '                       Proceedings Used                    '
                                Call UpdateProceedings(1)
                                '-----------------------------------------------------------'
                            Else
                                MsgBox "An Accounts Officer Can Cancel the Payorder", vbInformation
                                Exit Sub
                            End If
                        End If
                    End If
                Case 1 '// PayOrder APPROVED
                
                    If mPreviousYearMode = 1 Then
                        Call GetPendingTaskDetails
                        If mPendingTransactionDate >= DateAdd("yyyy", -1, gbStartingDate) And mPendingTransactionDate <= DateAdd("yyyy", -1, gbEndingDate) Then
                            mTrnDate = mPendingTransactionDate
                            mCurFinYear = gbFinancialYearID - 1
                        End If
                    End If
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then '// Approval Stage
                        '// Edit Payorder Cancellation Request  Final Level, tnyStatus = 2
                        mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                            Trim(txtRemarks.Text), mPayorderRequested, txtReqSeatCancel.Tag, _
                                            Null, Null, mForwardedSeat, mCurFinYear, _
                                            2, gbUserID, gbTransactionDate, mPayOrderNo, mPayorderID, _
                                            Null, Null, Null, Null, Null, Null)
                        objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                        '-----------------------------------------------------------'
                        '                       Proceedings Used                    '
                        Call UpdateProceedings(1)
                        
                        '-----------------------------------------------------------'
                        '// Reverse Entry
                        
                        ' NOTE: SQL QUERY CHANGED BY AIBY on 06-May-2013
                        mSql = "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo,dtDate, faVouchers.intFinancialYearID,dtStartingDate, dtEndingDate  From faVouchers"
                        mSql = mSql + " Inner join faFinancialYear ON faFinancialYear.intFinancialYearID = faVouchers.intFinancialYearID"
                        mSql = mSql + " Where intKeyID2 = " & mPayOrderNo
                        
                        Rec.Open mSql, mCnn
                        'Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPayOrderNo, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            While Not (Rec.EOF)
                            
                                If Not (mTrnDate >= Rec!dtStartingDate And mTrnDate <= Rec!dtEndingDate) Then
                                    mTrnDate = Rec!dtDate
                                End If
                            
                                mArrayInput = Array(Rec!intVoucherID, mTrnDate)
                                objdb.ExecuteSP "spSaveReverseVouchers", mArrayInput, mArrayOut, , mCnn
                                If IsArray(mArrayOut) Then
                                    mVoucherID = Rec!intVoucherID
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,intExternalModuleID=70, tnyReversed = 1  Where intVoucherID = " & mArrayOut(0, 0)
                                    mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=70, tnyReversed = 1 Where intVoucherID=" & mArrayOut(0, 0)
                                End If
                                If Rec!tnyVoucherTypeID = 40 Then
                                    'mVoucherID = Rec!intVoucherID
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                    mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                    
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                End If
                                If Rec!tnyVoucherTypeID = 10 Then
                                    'mVoucherID = Rec!intVoucherID
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                    mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                    
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                End If
                                
                                If Rec!tnyVoucherTypeID = 20 Then
                                    mVoucherID = Rec!intVoucherID
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                    mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                    
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                End If
                                Rec.MoveNext
                            Wend
                        End If
                        Rec.Close
                        '// Cancel Payorder
                        mSql = "Update faPayOrder Set tnyCancelled = 1 ,intAllotmentID=Null Where vchPayorderNo = '" & mPayOrderNo & "'"
                        mCnn.Execute mSql
                        If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                            If mVoucherID > 0 Then
                                'ReverseSulekhaExpenseDetails (mVoucherID)
                            End If
                        Else
                            
                        End If
                        
                    Else '// New Request Stage
                        If gbLBType = 4 Then '// Corporation
                            If gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                                'If gbUserID = mPayorderApproved Then      ' Changed to Seat Validation
                                If gbSeatID = mPayorderApprovedSeat Then   ' Changed to Seat Validation
                                    '// Save in Pay order Details
                                    '// Table faReverseEntry
                                    mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                1, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                                    objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                                    mRequestID = mArrayOut(0, 0)
                                    'Call SaveReverseChild(mRequestID)
                                Else
                                    MsgBox "Payorder Approved By " & txtAppUser.Text, vbInformation
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Accounts Supdt. Can Apply for Cancellation", vbInformation
                                Exit Sub
                            End If
                        Else                 '// Panchayats or Municipalities
                            If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                                'If gbUserID = mPayorderRequested Then      ' Changed to Seat Validation
                                If gbSeatID = mPayorderRequestedSeat Then   ' Changed to Seat Validation
                                    '// Code for Save Payorder cancellations
                                    '// Table faReverse Entry
                                    mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                Trim(txtRemarks.Text), gbUserID, gbSeatID, _
                                                Null, Null, mForwardedSeat, mCurFinYear, _
                                                1, Null, Null, mPayOrderNo, mPayorderID, _
                                                Null, Null, Null, Null, Null, Null)
                                    objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                                    mRequestID = mArrayOut(0, 0)
                                    'Call SaveReverseChild(mRequestID)
                                Else
                                    MsgBox "Payorder Requested By " & txtReqUser.Text, vbInformation
                                    Exit Sub
                                End If
                            Else
                                MsgBox "Accounts clerk Can Apply for Cancellation", vbInformation
                                Exit Sub
                            End If
                        End If
                    End If
                Case 2
                    If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
                    '-------- For Panchayat---------------------------------
                    '--------- Modified On 1/6/11 By Anisha C
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            If mDBPaidToParty < 1 Then
                                'Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPayOrderNo, mCnn
                                'If Not (Rec.EOF And Rec.BOF) Then
                                '
                                'While Not (Rec.EOF)
                                
                                ' NOTE: SQL QUERY CHANGED BY AIBY on 06-May-2013
                                mSql = "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo,dtDate, faVouchers.intFinancialYearID,dtStartingDate, dtEndingDate  From faVouchers"
                                mSql = mSql + " Inner join faFinancialYear ON faFinancialYear.intFinancialYearID = faVouchers.intFinancialYearID"
                                mSql = mSql + " Where intKeyID2 = " & mPayOrderNo
                                
                                Rec.Open mSql, mCnn
                                'Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPayOrderNo, mCnn
                                If Not (Rec.EOF And Rec.BOF) Then
                                    While Not (Rec.EOF)
                                    
                                        If Not (mTrnDate >= Rec!dtStartingDate And mTrnDate <= Rec!dtEndingDate) Then
                                            mTrnDate = Rec!dtDate
                                        End If
                                    
                                        mArrayInput = Array(Rec!intVoucherID, mTrnDate)
                                        objdb.ExecuteSP "spSaveReverseVouchers", mArrayInput, mArrayOut, , mCnn
                                        If IsArray(mArrayOut) Then
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,intExternalModuleID=70, tnyReversed = 1 Where intVoucherID = " & mArrayOut(0, 0)
                                            mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=70, tnyReversed = 1 Where intVoucherID=" & mArrayOut(0, 0)
                                            '**********************************************************************************************************************
                                                Call UpdateVoucherIndex(Rec!intVoucherID)    'ADDED BY MINU FOR UPDATE tnyChangeFag IN faVoucherIndex
                                            '**********************************************************************************************************************
                                             'Call CancelVoucherNewACRMode(mPayOrderNo) 'TO CANCEL AUTO GENERATED RECEIPTS FROM NEW ACR MODE
                                        End If
                                        If Rec!tnyVoucherTypeID = 40 Then
                                            'mVoucherID = Rec!intVoucherID
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                            mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                            
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                        End If
                                        If Rec!tnyVoucherTypeID = 10 Then
                                            'mVoucherID = Rec!intVoucherID
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                            mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                            
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                        End If
                                
                                        If Rec!tnyVoucherTypeID = 20 Then
                                            mVoucherID = Rec!intVoucherID
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                            mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                            
                                            mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                                        End If
                                        Rec.MoveNext
                                    Wend
                                End If
                                Rec.Close
                                mBoolCancel = True
                            Else
                                 mBoolCancel = PaidToPartyProcess '// The Process involving Reversal and Demand Generation
                            End If
                            If mBoolCancel Then
                                '// Cancel Payorder
                                mSql = "Update faPayOrder Set tnyCancelled = 1,intAllotmentID=Null Where vchPayorderNo = '" & mPayOrderNo & "'"
                                mCnn.Execute mSql
                                If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                                    If mVoucherID > 0 Then
                                        'ReverseSulekhaExpenseDetails (mVoucherID)
                                    End If
                                Else
                                    
                                End If
                                '// Reverse Entry
                                mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                    Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                                    txtApprovedAO.Tag, mTrnDate, mForwardedSeat, gbFinancialYearID, _
                                                    2, gbUserID, gbTransactionDate, mPayOrderNo, mPayorderID, _
                                                    Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                                '-----------------------------------------------------------'
                                '                       Proceedings Used                    '
                                Call UpdateProceedings(1)
                                '-----------------------------------------------------------'
                            End If
                        Else
                            If gbSeatID = txtReqSeatCancel.Tag Then
                                    '// Edit Payorder Cancellation Request Intermediate Level, tnyStatus = 1
                                mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                        Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                                        gbUserID, mTrnDate, mForwardedSeat, mCurFinYear, _
                                                        1, Null, Null, mPayOrderNo, mPayorderID, _
                                                        Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                                objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                                mRequestID = mArrayOut(0, 0)
                                Call SaveReverseChild(mRequestID)
                                '-----------------------------------------------------------'
                                '                       Proceedings Updation                '
                                Call UpdateProceedings(0)
                                '-----------------------------------------------------------'
                                MsgBox "Updated Successfully", vbInformation
                            Else
                                MsgBox "Not a Valid Requested User ", vbInformation
                                Exit Sub
                            End If
                        End If
            
                    Else  '''--- For Muncipalities And Corporations
                        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                            '// Edit Payorder Cancellation Request Intermediate Level, tnyStatus = 1
                            mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                    Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                                    gbUserID, mTrnDate, mForwardedSeat, mCurFinYear, _
                                                    1, Null, Null, mPayOrderNo, mPayorderID, _
                                                    Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                            objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                            mRequestID = mArrayOut(0, 0)
                            Call SaveReverseChild(mRequestID)
                        Else
                            If gbLBType = 4 Then '// Corporation
                                If gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                                    'If gbUserID = mPayorderApproved Then      ' Changed to Seat Validation
                                    If gbSeatID = mPayorderApprovedSeat Then   ' Changed to Seat Validation
                                        '// Save in Pay order Details
                                        '// Table faReverseEntry
                                        mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                    Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                                    Null, Null, mForwardedSeat, mCurFinYear, _
                                                    0, Null, Null, mPayOrderNo, mPayorderID, _
                                                    Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                                        objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                                        mRequestID = mArrayOut(0, 0)
                                        'Call SaveReverseChild(mRequestID)
                                        '-----------------------------------------------------------'
                                        '                       Proceedings Updation                '
                                        Call UpdateProceedings(0)
                                        '-----------------------------------------------------------'
                                    Else
                                        MsgBox "Payorder Approved By " & txtAppUser.Text, vbInformation
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Only Account Supdt. Can Apply for Cancellation", vbInformation
                                    Exit Sub
                                End If
                            Else                 '// Panchayats or Municipalities
                                If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
                                    'If gbUserID = mPayorderRequested Then      ' Changed to Seat Validation
                                    If gbSeatID = mPayorderRequestedSeat Then   ' Changed to Seat Validation
                                        '// Code for Save Payorder cancellations
                                        '// Table faReverse Entry
                                        mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                                    Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                                    Null, Null, mForwardedSeat, mCurFinYear, _
                                                    0, Null, Null, mPayOrderNo, mPayorderID, _
                                                    Null, Null, Null, Null, Null, Null)
                                        objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                                        mRequestID = mArrayOut(0, 0)
                                        '-----------------------------------------------------------'
                                        '                       Proceedings Updation                '
                                        Call UpdateProceedings(0)
                                        '-----------------------------------------------------------'
                                        'Call SaveReverseChild(mRequestID)
                                    Else
                                        MsgBox "Payorder Requested by " & txtReqUser.Text, vbInformation
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox "Acccounts Clerk can Apply for Cancellation", vbInformation
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
            End Select
        Case 1      '// Requested and Approved First Level
            
            'PANCHAYATH
            If gbSeatGroupID = gbSeatGroupSecretary Or ((gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5) And gbSeatGroupID = gbSeatGroupAccountsOfficer) Then
                If mDBPaidToParty < 1 Then
                            'Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPayOrderNo, mCnn
                            'If Not (Rec.EOF And Rec.BOF) Then
                            'While Not (Rec.EOF)
                        ' NOTE: SQL QUERY CHANGED BY AIBY on 06-May-2013
                        mSql = "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo,dtDate, faVouchers.intFinancialYearID,dtStartingDate, dtEndingDate  From faVouchers"
                        mSql = mSql + " Inner join faFinancialYear ON faFinancialYear.intFinancialYearID = faVouchers.intFinancialYearID"
                        mSql = mSql + " Where intKeyID2 = " & mPayOrderNo
                        
                        Rec.Open mSql, mCnn
                        'Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & mPayOrderNo, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                        While Not (Rec.EOF)
                        
                            If Not (mTrnDate >= Rec!dtStartingDate And mTrnDate <= Rec!dtEndingDate) Then
                                mTrnDate = Rec!dtDate
                            End If
                        
                            mArrayInput = Array(Rec!intVoucherID, mTrnDate)
                            objdb.ExecuteSP "spSaveReverseVouchers", mArrayInput, mArrayOut, , mCnn
                            If IsArray(mArrayOut) Then
                                'mVoucherID = Rec!intVoucherID
                                mCnn.Execute "Update faVouchers Set tnysync=Null,intExternalModuleID=70, tnyReversed = 1 Where intVoucherID = " & mArrayOut(0, 0)
                                mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=70 , tnyReversed = 1  Where intVoucherID=" & mArrayOut(0, 0)
                            End If
                            If Rec!tnyVoucherTypeID = 40 Then
                                'mVoucherID = Rec!intVoucherID
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                            End If
                            If Rec!tnyVoucherTypeID = 10 Then
                                'mVoucherID = Rec!intVoucherID
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                            End If
                            If Rec!tnyVoucherTypeID = 20 Then
                                mVoucherID = Rec!intVoucherID
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks =  vchRemarks + 'Cancelled to " & CStr(mArrayOut(1, 0)) & "'Where intVoucherID = " & Rec!intVoucherID
                                mCnn.Execute "Update faTransactions Set tnysync=Null,tnyReversed=1 Where intVoucherID = " & Rec!intVoucherID
                                
                                mCnn.Execute "Update faVouchers Set tnysync=Null,tnyReconciled = 3,tnyReversed=1,numTockenID = Null,dtRealisationDate = '" & Format(gbTransactionDate, "dd/MMM/YYYY") & "',vchRemarks = vchRemarks + 'Cancelled From " & Rec!intVoucherNo & "'Where intVoucherID = " & mArrayOut(0, 0)
                            End If
                            Rec.MoveNext
                        Wend
                    End If
                    Rec.Close
                    mBoolCancel = True
                Else
                     mBoolCancel = PaidToPartyProcess '// The Process involving Reversal and Demand Generation
                End If
                If mBoolCancel Then
                    '// Cancel Payorder
                    mSql = "Update faPayOrder Set tnyCancelled = 1,intAllotmentID=Null Where vchPayorderNo = '" & mPayOrderNo & "'"
                    mCnn.Execute mSql
                    If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                        If mVoucherID > 0 Then
                            'ReverseSulekhaExpenseDetails (mVoucherID)
                        End If
                    Else
                        
                    End If
                    '// Reverse Entry
                    mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                        Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                        txtApprovedAO.Tag, gbTransactionDate, mForwardedSeat, mCurFinYear, _
                                        2, gbUserID, gbTransactionDate, mPayOrderNo, mPayorderID, _
                                        Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                    objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, , , mCnn
                    '-----------------------------------------------------------'
                    '                       Proceedings Used                    '
                    Call UpdateProceedings(1)
                    '-----------------------------------------------------------'
                End If
            Else ' CASE  1: FIRST LEVEL APPROVAL :: MUNICIPALITY
                If gbUserID = txtApprovedAO.Tag Then
                        '// Edit Payorder Cancellation Request Intermediate Level, tnyStatus = 1
                    mArrayInput = Array(mRequestID, mTrnDate, 70, 50, mCancelReason, _
                                            Trim(txtRemarks.Text), txtReqUserCancel.Tag, txtReqSeatCancel.Tag, _
                                            gbUserID, gbTransactionDate, mForwardedSeat, mCurFinYear, _
                                            1, Null, Null, mPayOrderNo, mPayorderID, _
                                            Null, mPaidToParty, mRecoveryRemitted, mChequeCancelled, Null, Null)
                    objdb.ExecuteSP "spSaveReverseEntry", mArrayInput, mArrayOut, , mCnn
                    mRequestID = mArrayOut(0, 0)
                    Call SaveReverseChild(mRequestID)
                    '-----------------------------------------------------------'
                    '                       Proceedings Updation                '
                    Call UpdateProceedings(0)
                    '-----------------------------------------------------------'
                Else
                    MsgBox "Invalid User ", vbInformation
                End If
            End If
        Case 2      '// Approved in Final Level
            
            
        Case 4      '// Rejected Request
            If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                mCancelStatus = 0
                GoTo StartCase
            Else
                MsgBox "The Payorder Cancel Request is Cancelled, To do active Edit the Cancel request", vbInformation
                Exit Sub
            End If
        Case Else
            MsgBox "The Payorder Cancel Status is invalid", vbInformation
    End Select
    
    'NOTE:- PENDING TASK
    
    If Trim(txtPayOrderNo.Text) <> "" Then
        Call FillPayOrder(Trim(txtPayOrderNo.Text))
        cmdSave.Enabled = False
    End If
    If mPreviousYearMode = 1 Then
         mCnn.Execute "Update faPendingTaskRequest Set tnyStatus = 8 Where intRequestID = " & mPendingTaskReqID
    End If
End Sub

    Private Sub UpdateProceedings(mUsedFlag As Integer)
        Dim mSql                As String
        Dim mCnn                As New ADODB.Connection
        Dim objdb               As New clsDB
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If txtProceddingsNo.Tag > 0 Then
            mSql = "Update faProceedings Set intVoucherID = Null, intVoucherNo = Null Where intModuleID = 70 And isNull(tnyUsed,0) <> 1 And intVoucherNo = " & txtPayOrderNo.Text & " And intProceedingsID <>  " & val(txtProceddingsNo.Tag) & ";" & vbNewLine
            mSql = mSql + "Update faProceedings Set intVoucherID = " & txtPayOrderNo.Tag & ", intVoucherNo = " & txtPayOrderNo.Text & ",tnyUsed = " & mUsedFlag & " Where intModuleID = 70 And intProceedingsID = " & val(txtProceddingsNo.Tag) & ";"
            mCnn.Execute mSql
        End If
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
        mPreviousYearMode = 0
        mPendingTaskReqID = 0
    End Sub

    Public Sub txtPayOrderNo_LostFocus()
        
        If val(Trim(txtPayOrderNo.Text)) > 0 Then
            If gbSeatGroupID = gbSeatGroupSecretary Then
                CheckNonReceiptedDemandNo (val(Trim(txtPayOrderNo.Text)))
            End If
            If mPreviousYearMode = 1 Then
                GetPendingTaskDetails
            End If
            'Call CheckTransferCredit(val(Trim(txtPayOrderNo.Text)))
            Call FillPayOrder(val(Trim(txtPayOrderNo.Text)))
        End If
        
    End Sub
    Private Sub CheckTransferCredit(ByVal mPayOrderNo As Double)
        Dim mSql                As String
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim objdb               As New clsDB
        Dim mSqlChild           As String
        Dim RecChild            As New ADODB.Recordset
        Dim mAccountHeadID      As Variant
        Dim mPayOrderDate       As Date
        Dim mPayOrderVrID       As Long
        Dim mVoucherID          As Long
        
        mSql = " SELECT faPayOrder.intVoucherID ,ISNULL(numProjectNo,0) numProjectID,intKeyID1 intAccountHeadID,CONVERT(varchar,dtPayOrderDate,103) dtPayOrderDate"
        mSql = mSql + " From faPayOrder"
        mSql = mSql + " INNER JOIN faVouchers ON faVouchers.intVoucherID=faPayOrder.intVoucherID"
        mSql = mSql + " WHERE vchPayOrderNo = " & mPayOrderNo
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "The Connection to Saankhya not Present", vbCritical
            Exit Sub
        End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!numProjectID <> 0 Then
                mAccountHeadID = Rec!intAccountHeadID
                mPayOrderDate = Rec!dtPayOrderDate
                mPayOrderVrID = Rec!intVoucherID
                If gbLBPanchayat = 1 Then
                    mSqlChild = "SELECT intAccountHeadID,intKeyID1,faVouchers.intVoucherID,CONVERT(varchar,dtDate,103) dtDate FROM faVouchers "
                    mSqlChild = mSqlChild + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                    mSqlChild = mSqlChild + " WHERE intTransactionTypeID=4010"
                    
                Else
                    mSqlChild = "SELECT intAccountHeadID,intKeyID1,faVouchers.intVoucherID,CONVERT(varchar,dtDate,103) dtDate  FROM faVouchers "
                    mSqlChild = mSqlChild + " INNER JOIN faVoucherChild ON faVoucherChild.intVoucherID=faVouchers.intVoucherID"
                    mSqlChild = mSqlChild + " WHERE intTransactionTypeID=4006"
                End If
                RecChild.Open mSqlChild, mCnn
                If Not (RecChild.EOF And RecChild.BOF) Then
                    If mPayOrderVrID > RecChild!intVoucherID And CDate(mPayOrderDate) >= CDate(RecChild!dtDate) Then
                        Exit Sub
                    Else
                        While Not RecChild.EOF
                            If RecChild!intAccountHeadID = mAccountHeadID Or RecChild!intKeyID1 = mAccountHeadID Then 'intKeyID1
                                
                                MsgBox "Transfer Credit Already Done.No Cancellation is possible", vbInformation
                                txtPayOrderNo.Text = ""
                                txtPayOrderNo.Tag = ""
                                Exit Sub
                            End If
                            RecChild.MoveNext
                        Wend
                    End If
                End If
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    
    End Sub
    Private Sub CheckNonReceiptedDemandNo(ByVal mPayOrderNo As Double)
        Dim mSql                As String
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim RecRecoveries       As New ADODB.Recordset
        Dim objdb               As New clsDB
        
        mSql = mSql + " Select faIDemandTBL.numDemandID From faReverseEntry"
        mSql = mSql + " Inner Join faIDemandTBL ON faIDemandTBL.intKeyID2=faReverseEntry.numDemandNo"
        mSql = mSql + " Where faReverseEntry.tnyStatus = 2 And tnyPaid = 1 And intVoucherID Is Null And tnyVoucherTypeID = 50"
        mSql = mSql + " And faReverseEntry.numDemandNo=" & mPayOrderNo
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "The Connction to Saankhya not Present", vbCritical
            Exit Sub
        End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            cmdCancelApprovedlDemandNo.Visible = True
        Else
            cmdCancelApprovedlDemandNo.Visible = False
        End If
    End Sub
Private Sub PaidToPartyProcessTest()
    Dim mSql                As String
    Dim mCnn                As New ADODB.Connection
    Dim Rec                 As New ADODB.Recordset
    Dim RecRecoveries       As New ADODB.Recordset
    Dim objdb               As New clsDB
    Dim mArrayIn            As Variant
    Dim mArrayOut           As Variant
    
    Dim mPaidToParty As Integer
    Dim mRecoveryRemitted As Integer
    Dim mChequeCancelled As Integer
    Dim mBoolStatus As Boolean
    Dim mBoolRecovery As Boolean
    Dim mRecoveryAmount As Double
    
    mPaidToParty = 0
    mRecoveryRemitted = 0
    mChequeCancelled = 0
    mBoolStatus = False
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Sub
    End If
    
    '-------------------------------'
    '   Checking Cetifications      '
    '-------------------------------'
    If mPayorderStatus = 2 Then
        If mCancelStatus = 0 Then
            If chkPaidToPartyYesInter.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesInter.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesInter.Value = 1 Then
                mChequeCancelled = 1
            End If
        ElseIf mCancelStatus = 1 Then
            If chkPaidToPartyYesApp.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesApp.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesApp.Value = 1 Then
                mChequeCancelled = 1
            End If
        End If
    End If
    
    If mDBPaidToParty = mPaidToParty Then
        If mDBRecoveryRemitted = mRecoveryRemitted Then
            If mDBChequeCancelled = mChequeCancelled Then
                mBoolStatus = True
            End If
        End If
    End If
    If mBoolStatus Then
        Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where intKeyID2 = " & txtPayOrderNo.Text, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Not (Rec.EOF)
                mRecoveryAmount = 0
                RecRecoveries.Open "Select faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,faVoucherChild.fltAmount From faVoucherChild" & vbNewLine & _
                                    "Inner Join faAccountHeads On faVoucherChild.intAccountHeadID = faAccountHeads.intAccountHeadID" & vbNewLine & _
                                    "Where intVoucherID = " & Rec!intVoucherID, mCnn
                mBoolRecovery = False
                If Not (RecRecoveries.EOF And RecRecoveries.BOF) Then
                    While Not RecRecoveries.EOF
                        If val(RecRecoveries!vchAccountHeadCode) >= 350200100 And val(RecRecoveries!vchAccountHeadCode) <= 350309900 Then '// Recovery Heads
                            mRecoveryAmount = mRecoveryAmount + RecRecoveries!fltAmount
                            mBoolRecovery = True
                        End If
                        RecRecoveries.MoveNext
                    Wend
                End If
                RecRecoveries.Close
                If cmbCancelReasons.ItemData(cmbCancelReasons.ListIndex) = 700 Then
                    If mDBPaidToParty = 1 Then
                        If mDBRecoveryRemitted = 1 Then
                            If Rec!tnyVoucherTypeID <> 20 And mBoolRecovery = False Then
                                mArrayIn = Array(Rec!intVoucherID, gbTransactionDate)
                                objdb.ExecuteSP "spSaveReverseVouchers", mArrayIn, mArrayOut, , mCnn
                                If IsArray(mArrayOut) Then
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,intExternalModuleID=70 , tnyReversed=1  Where intVoucherID = " & mArrayOut(0, 0)
                                    mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=70, tnyReversed=1  Where intVoucherID=" & mArrayOut(0, 0)
                                End If
                                '' Call Save Demand Saving With Amount + mRecoveryAmount
                                Call DemandGeneration(mRecoveryAmount)
                            End If
                        Else
                            If Rec!tnyVoucherTypeID <> 20 Then
                                mArrayIn = Array(Rec!intVoucherID, gbTransactionDate)
                                objdb.ExecuteSP "spSaveReverseVouchers", mArrayIn, mArrayOut, , mCnn
                                If IsArray(mArrayOut) Then
                                    mCnn.Execute "Update faVouchers Set tnysync=Null,intExternalModuleID=70, tnyReversed = 1 Where intVoucherID = " & mArrayOut(0, 0)
                                    mCnn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=70, tnyReversed = 1 Where intVoucherID=" & mArrayOut(0, 0)
                                End If
                                '' Call Save Demand Saving With Amount
                                Call DemandGeneration(0)
                            End If
                        End If
                    End If
                End If
                Rec.MoveNext
            Wend
        End If
        Rec.Close
    Else
        MsgBox "Please Check Paid to Party, Recovery Remitted, Cheque Cancelled etc..", vbInformation
    End If
End Sub

Private Function PaidToPartyProcess() As Boolean
    Dim mSql                As String
    Dim mCnn                As New ADODB.Connection
    Dim Rec                 As New ADODB.Recordset
    Dim RecRecoveries       As New ADODB.Recordset
    Dim objdb               As New clsDB
    Dim mArrayInput            As Variant
    Dim mArrayOut           As Variant
    Dim mLoop               As Integer
    
    Dim mPaidToParty        As Integer
    Dim mRecoveryRemitted   As Integer
    Dim mChequeCancelled    As Integer
    Dim mBoolStatus         As Boolean
    Dim mBoolRecovery       As Boolean
    Dim mRecoveryAmount     As Double
    Dim mOtherAmount        As Double
    Dim mRecoveryHeads      As String
    Dim mVoucherID          As Variant
    Dim mTransactionID      As Variant
    
    
    mPaidToParty = 0
    mRecoveryRemitted = 0
    mChequeCancelled = 0
    mBoolStatus = False
    mVoucherID = -1
    mTransactionID = -1
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Function
    End If
    
    '-------------------------------'
    '   Checking Cetifications      '
    If mPayorderStatus = 2 Then
        If mCancelStatus = 0 Then
            If chkPaidToPartyYesInter.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesInter.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesInter.Value = 1 Then
                mChequeCancelled = 1
            End If
        ElseIf mCancelStatus = 1 Then
            If chkPaidToPartyYesApp.Value = 1 Then
                mPaidToParty = 1
            End If
            If chkRecoveriesPaidYesApp.Value = 1 Then
                mRecoveryRemitted = 1
            End If
            If chkChequeCancelledYesApp.Value = 1 Then
                mChequeCancelled = 1
            End If
        End If
    End If
    
    If mDBPaidToParty = mPaidToParty Then
        If mDBRecoveryRemitted = mRecoveryRemitted Then
            If mDBChequeCancelled = mChequeCancelled Then
                mBoolStatus = True
            End If
        End If
    End If
    mRecoveryHeads = "(-1"
    mRecoveryAmount = 0
    mOtherAmount = 0
    If mBoolStatus Then
        For mLoop = 1 To vsGridRecoveries.Rows - 1
            If vsGridRecoveries.Cell(flexcpChecked, mLoop, 4) = 1 Then
                mRecoveryAmount = mRecoveryAmount + val(vsGridRecoveries.TextMatrix(mLoop, 3))
                mRecoveryHeads = mRecoveryHeads + "," + vsGridRecoveries.TextMatrix(mLoop, 0)
            Else
                mOtherAmount = mOtherAmount + val(vsGridRecoveries.TextMatrix(mLoop, 3))
            End If
        Next mLoop
        mRecoveryHeads = mRecoveryHeads + ")"
        '-------------------------------------------------------------------------------------------'
        '' Call Demand Generation With Recovery Amount
        Call DemandGeneration(mRecoveryAmount)
        If mOtherAmount <> 0 Then
                
                '** ----------------------- **'
                '**                         **'
                '**  BLOCKED BY AIBY        **'
                '**  ON 12th August 2011    **'
                '**                         **'
                '** ----------------------- **'
                
                ''''''''            '' Call Reverse JV With mOtherAmount and Unchecked Recovery Heads
                ''''''''            '--------------------'
                ''''''''            '' Reversing the JV
                ''''''''            Rec.Open "Select Distinct intVoucherID,tnyVoucherTypeID,intVoucherNo From faVouchers Where tnyVoucherTypeID = 40 And intKeyID2 = " & txtPayOrderNo.Text, mCnn
                ''''''''            If Not (Rec.EOF And Rec.BOF) Then
                ''''''''                While Not Rec.EOF
                ''''''''                    mVoucherID = Rec!intVoucherID
                ''''''''                    mArrayInput = Array(mVoucherID, gbTransactionDate)
                ''''''''                    objDB.ExecuteSP "spSaveReverseVouchers", mArrayInput, mArrayOut, , mCnn
                ''''''''
                ''''''''                    Rec.MoveNext
                ''''''''                Wend
                ''''''''            End If
                ''''''''            Rec.Close
                ''''''''            '---------------------'
                ''''''''            If mDBRecoveryRemitted = 1 Then
                ''''''''                '----------Updating faVouchers fltAmount with other Amount--------'
                ''''''''                mCnn.Execute "Update faVouchers Set fltAmount = " & mOtherAmount & " Where intVoucherID = " & mVoucherID
                ''''''''                '----------Deleting Recovery Heads Which is Already Paid----------'
                ''''''''                mCnn.Execute "Delete From faVoucherChild Where intVoucherID = " & mVoucherID & " And intAccountHeadID in " & mRecoveryHeads
                ''''''''                '----------------Deleting From faTransactionChild-----------------'
                ''''''''                Rec.Open "Select intTransactionID From faTransactions Where intVoucherID = " & mVoucherID, mCnn
                ''''''''                If Not (Rec.EOF And Rec.BOF) Then
                ''''''''                    mTransactionID = Rec!intTransactionID
                ''''''''                End If
                ''''''''                Rec.Close
                ''''''''                mCnn.Execute "Delete From faTransactionChild Where intTransactionID = " & mTransactionID & " And intAccountHeadID in " & mRecoveryHeads
                ''''''''                '-----Updating TransactionChild Amount With unRecovered Amount-----'
                ''''''''                mCnn.Execute "Update faTransactionChild Set fltAmount = " & mOtherAmount & " Where intByAccountHeadID is Null And intTransactionID = " & mTransactionID
                ''''''''            End If
                '** ----------------------- **'
                '**  END OF BLOCKED BY AIBY **'
                '** ----------------------- **'
                
        End If
        PaidToPartyProcess = True
        '-------------------------------------------------------------------------------------------'
    Else
        'Returning
        PaidToPartyProcess = False
        MsgBox "Please Verify Certifications", vbInformation
    End If
End Function

Private Sub SaveReverseChild(mValRequestID As Variant)
    Dim mSql                As String
    Dim mLoop               As Integer
    Dim mCnn                As New ADODB.Connection
    Dim Rec                 As New ADODB.Recordset
    Dim objdb               As New clsDB
    Dim mArray              As Variant
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Sub
    End If
    mCnn.Execute "Delete From faReverseEntryTrChild Where intRequestID = " & mValRequestID
    
    For mLoop = 1 To vsGridRecoveries.Rows - 1
        If vsGridRecoveries.Cell(flexcpChecked, mLoop, 4) = 1 Then
            mArray = Array(mValRequestID, val(vsGridRecoveries.TextMatrix(mLoop, 0)), val(vsGridRecoveries.TextMatrix(mLoop, 3)))
            objdb.ExecuteSP "spSaveReverseEntryTrChild", mArray, , , mCnn
        End If
    Next mLoop
End Sub

Private Sub DemandGeneration(ByVal mRecoveryAmt As Double)
    Dim mSql        As String
    Dim mCnn        As New ADODB.Connection
    Dim Rec         As New ADODB.Recordset
    Dim objdb       As New clsDB
    Dim objAcc      As New clsAccounts
    Dim arrInput    As Variant
    Dim arrOutPut   As Variant
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
        MsgBox "The Connction to Saankhya not Present", vbCritical
        Exit Sub
    End If
    If mPaymentVoucherNo > 0 Then
        mSql = "SELECT     faVouchers.intVoucherID, faVouchers.intTransactionTypeID, faVouchers.tnyVoucherTypeID, faVouchers.intVoucherNo, faVouchers.dtDate," & vbNewLine
        mSql = mSql + "faVouchers.fltAmount AS VoucherAmount,faVouchers.vchDescription, faVouchers.intInstrumentTypeID, faVouchers.vchInstrumentNo, faVouchers.dtInstrumentDate," & vbNewLine
        mSql = mSql + "faVouchers.vchDescription, faVouchers.intUserID, faVouchers.intCounterID, faVouchers.intKeyID1, faVouchers.intKeyID2, faVouchers.tnyCancelFlag," & vbNewLine
        mSql = mSql + "faVouchers.intFundID , faVoucherChild.intSlNo, faVoucherChild.intAccountHeadID, faVoucherChild.tnyDebitOrCredit, faVoucherChild.fltAmount" & vbNewLine
        mSql = mSql + "FROM       faVouchers INNER JOIN" & vbNewLine
        mSql = mSql + "faVoucherChild ON faVouchers.intVoucherID = faVoucherChild.intVoucherID" & vbNewLine
        mSql = mSql + "WHERE tnyVoucherTypeID = 20 AND intKeyID2 = " & txtPayOrderNo.Text
    End If
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        Dim mUDemand As uDemand
        ' Saving to Demand TBL
        With mUDemand
            .intLBID = gbLocalBodyID
            .tnyExtAppID = AppID.Saankhya
            .tnyExtModuleID = 70
            .tnyDemandType = 70
            .intTransactionTypeID = txtPaymentType.Tag
            .intYearID = gbFinancialYearID
            .tnyPeriodID = Null
            .dtDemandDate = IIf(IsDate(txtPayorderDate.Text), txtPayorderDate.Text, gbTransactionDate)
            .numSubLedgerID = Null
            .intKeyID = Rec!intKeyID1
            .intKeyID2 = txtPayOrderNo.Text
            .vchRemarks = "Auto Generated Demand:" & vbNewLine & IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            .tnyStatus = 0
            .intVoucherID = Null
            .dtVoucherDate = Null
            .tnyArrearFlag = Null
            .dtExpiryDate = gbTransactionDate
            .numDemandID = IIf(mDemandID < 1, Null, mDemandID)
            .intFinancialYearID = gbFinancialYearID
            .numSeatID = gbSeatID
            .intSectionID = gbSectionID
            .numUserID = gbUserID
            .numCounterID = gbCounterID
            .vchAdminNote = Null
            .vchDemandNo = IIf(mDemandNo < 1, Null, mDemandNo)
            .numZoneID = Null
            .intWardNo = Null
            .intDoorNo = Null
            .vchDoorNo2 = Null
            .numForwardedSeatID = Null
            .dtDueDate = gbTransactionDate
            .intInstrumentTypeID = Rec!intInstrumentTypeID
            .vchInstrumentNo = Rec!vchInstrumentNo
            .dtInstrumentDate = Rec!dtInstrumentDate
            .vchDrawnFrom = Null
            .vchDrawnPlace = Null
            .tnyAccrualType = Null
            .numLocationID = gbLocationID
            .intFunctionID = 6
            .intFunctionaryID = 4
            .intSourceFundID = 4
            .dtTransactionDate = gbTransactionDate
            .intDemandMode = 0
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
                        .numDemandID, _
                        .intFinancialYearID, _
                        .numSeatID, _
                        .intSectionID, _
                        .numUserID, _
                        .numCounterID, .vchAdminNote, .vchDemandNo, .numZoneID, .intWardNo, .intDoorNo, .vchDoorNo2, .numForwardedSeatID, .dtDueDate, _
                        .intInstrumentTypeID, .vchInstrumentNo, .dtInstrumentDate, .vchDrawnFrom, .vchDrawnPlace, .tnyAccrualType, _
                        .numLocationID, .intFunctionID, .intFunctionaryID, .intSourceFundID, .dtTransactionDate, .intDemandMode)
        End With
        objdb.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn
        mDemandID = arrOutPut(0, 0)
        mDemandNo = arrOutPut(1, 0)
        txtDemandNo.Tag = mDemandID
        txtDemandNo.Text = mDemandNo
        
        '// Assuming only one Row in Foucher Child
        mCnn.Execute "Delete From faIDemandChild Where numDemandID=" & mDemandID
        ' Demand Saving To Child
        Dim mUDemandChild As uDemandChild
        With mUDemandChild
            .numDemandID = mDemandID
            .intLBID = gbLocalBodyID
            .tnySlNo = 1
            objAcc.SetAccounts (Rec!intAccountHeadID)
            .intAccountHeadID = Rec!intAccountHeadID
            .vchAccountHeadCode = objAcc.AccountCode
            .fltAmount = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount) + mRecoveryAmt
            .vchRemarks = ""
            .tnyStatus = 0
            .dtOnDate = gbTransactionDate
            .intYearID = gbFinancialYearID
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
        ' Demand Save To Demand Address
    
        arrInput = Array(mDemandID, _
                        gbLocalBodyID, _
                        Null, Null, Null, _
                        Null, Null, Null, _
                        Null, Null, Null, _
                        Null, Null, Null, _
                        Null, Null, Null, Null, IIf(mDemandID < 1, 0, 1))
            
        objdb.ExecuteSP "spSaveIDemandAddress", arrInput, , , mCnn
    End If
    Rec.Close
End Sub


Private Sub vsGriidVouchers_DblClick()

    Dim mCnn            As New ADODB.Connection
    Dim objdb           As New clsDB
    Dim Rec             As New ADODB.Recordset
    Dim RecAccHeads     As New ADODB.Recordset
    Dim mSql            As String
    Dim mSqlAccHeads    As String
    Dim mRowCount       As Double
        If vsGriidVouchers.Rows < 2 Then Exit Sub
        If val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 1)) = 20 Then
            frmIntegratedPayments.ViewMode = 1 ''Added on 16.07.2013 by anisha
            Call frmIntegratedPayments.DisplayVoucherDetails(val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 2)))
            frmIntegratedPayments.cmdNew.Enabled = False
            frmIntegratedPayments.cmdSave.Enabled = False
            'Unload Me
        ElseIf val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 1)) = 40 Then
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            mSql = "Select *,faTransactionChild.tinDebitOrCreditFlag  From faVouchers"
            mSql = mSql + " Left Join faTransactions On faTransactions.intVoucherId = faVouchers.intVoucherId"
            mSql = mSql + " Left Join faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID "
            mSql = mSql + " Left Join faTransactionType On faVouchers.intTransactionTypeID = faTransactionType.intTransactionTypeID"
            mSql = mSql + " Left Join faFunctions On fatransactions.intFunctionId = faFunctions.intFunctionId"
            mSql = mSql + " Left Join faFunctionaries On faTransactions.intFunctionaryId = faFunctionaries.intFunctionaryId"
            mSql = mSql + " Left Join faFunds On faFunds.intFundId = faTransactions.intFundId"
            mSql = mSql + " Left Join faFields On faTransactions.intFieldID = faFields.intFieldID"
            mSql = mSql + " Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID"
            mSql = mSql + " Left Join faInstrumentTypes On faVouchers.intInstrumentTypeID = faInstrumentTypes.intInstrumentTypeID"
            mSql = mSql + " Left Join faAccountHeads On faVouchers.intKeyID1 = faAccountHeads.intAccountHeadID"
            mSql = mSql + " Left Join faBanks On faVouchers.intKeyID1 = faBanks.intAccountHeadID"
            mSql = mSql + " Where faVouchers.intVoucherNo = " & val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 2))
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                frmJournalEntry.txtVoucherNo.Tag = IIf(IsNull(Rec.Fields(0)), "", Rec.Fields(0)) 'intVocherID
                frmJournalEntry.txtReference.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
                frmJournalEntry.txtReference.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
                frmJournalEntry.txtDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                frmJournalEntry.txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
                frmJournalEntry.txtFund.Tag = IIf(IsNull(Rec.Fields(34)), "", Rec.Fields(34)) 'intFundID
                frmJournalEntry.txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                frmJournalEntry.txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                frmJournalEntry.txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                frmJournalEntry.txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                frmJournalEntry.txtField.Text = IIf(IsNull(Rec!vchField), "", Rec!vchField)
                frmJournalEntry.txtField.Tag = IIf(IsNull(Rec!intFieldID), "", Rec!intFieldID)
                If Not IsNull(Rec!tinDebitOrCreditFlag) Then
                    If (Rec!tinDebitOrCreditFlag) = 0 Then
                        frmJournalEntry.optDebit.Value = False
                        frmJournalEntry.optCredit.Value = True
                    Else
                        frmJournalEntry.optDebit.Value = True
                        frmJournalEntry.optCredit.Value = False
                    End If
                End If
                frmJournalEntry.txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                frmJournalEntry.txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                frmJournalEntry.txtAccountHead.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
                frmJournalEntry.txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                mSqlAccHeads = "Select * From faTransactionChild"
                mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & val(frmJournalEntry.txtReference.Tag)
                mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
                RecAccHeads.Open mSqlAccHeads, mCnn
                mRowCount = 1
                While Not Rec.EOF
                    While Not RecAccHeads.EOF
                        frmJournalEntry.vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                        frmJournalEntry.vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                        frmJournalEntry.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                        frmJournalEntry.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        frmJournalEntry.vsGrid.Rows = frmJournalEntry.vsGrid.Rows + 1
                        mRowCount = mRowCount + 1
                        RecAccHeads.MoveNext
                    Wend
                    Rec.MoveNext
                Wend
                RecAccHeads.Close
            End If
            Rec.Close
            frmJournalEntry.cmdNew.Enabled = False
            frmJournalEntry.cmdSave.Enabled = False
            'Unload Me
        ElseIf val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 1)) = 30 Then
            frmContraEntry.ListContraDemandOrVoucher (val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 2)))
            frmContraEntry.cmdNew.Enabled = False
            frmContraEntry.cmdSave.Enabled = False
        ElseIf val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 1)) = 10 Then
            Call DisplayReceiptDetails(val(vsGriidVouchers.TextMatrix(vsGriidVouchers.Row, 2)))
            
        End If
        Me.Hide
End Sub
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
            mSql = mSql + " Where intVoucherNo=" & mVoucherNo
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Not IsNull(Rec!tnyCancelFlag) Then
                    If Rec!tnyCancelFlag = 1 Then
                        frmReceipt.lblMessage.Visible = True
                        frmReceipt.lblMessage.Caption = "This is a Cancelled Receipt"
                        frmReceipt.Timer1.Enabled = True
                    End If
                End If
                frmReceipt.txtReceiptNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
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
                        frmReceipt.vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchperiodicity), "", RecAccHeads!vchperiodicity)
                        '--------------------------------------------------------
                        mArrearFlag = IIf(IsNull(RecAccHeads!tnyArrearFlag), "", RecAccHeads!tnyArrearFlag)
                        If mArrearFlag = 0 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        ElseIf mArrearFlag = 1 Then
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        Else
                            frmReceipt.vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                        End If
                        frmReceipt.vsGrid.Rows = frmReceipt.vsGrid.Rows + 1
                        mRowCount = mRowCount + 1
                        RecAccHeads.MoveNext
                    Wend
                    Rec.MoveNext
                Wend
                RecAccHeads.Close
                Call frmSearchReceipts.Calculate
            End If
            mCnn.Close
    
   End Sub
'     Private Sub GetPendingTaskDetails()
'        Dim mSQL As String
'        Dim objDB       As New clsDB
'        Dim mCnn        As New ADODB.Connection
'        Dim Rec         As New ADODB.Recordset
'
'        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
'            mSQL = "Select * From faPendingTaskRequest Where intRequestID=" & mPendingTaskReqID
'            Set Rec = objDB.ExecuteSP(mSQL, , , , mCnn, adCmdText)
'            If Not (Rec.EOF Or Rec.BOF) Then
'                mPendingTransactionDate = IIf(IsDate(Rec!dtTransactionDate), Rec!dtTransactionDate, DateAdd("yyyy", -1, gbEndingDate))
'            End If
'            Rec.Close
'        End If
'    End Sub
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property
    Public Property Let PendingTaskReqID(mData As Integer)
        mPendingTaskReqID = mData
    End Property
    
    Public Property Let PreviousYearTransactionDate(mData As Date)
        mPendingTransactionDate = mData
    End Property
'''    Public Sub CancelVoucherNewACRMode(mPayOrderNo As Variant)   '******TO CANCEL THE AUTOGENERATED RECEIPTS FROM NEW ACR MODE PAYMENTS
'''        Dim mCnn   As New ADODB.Connection
'''        Dim Rec    As New ADODB.Recordset
'''        Dim objDB  As New clsDB
'''        Dim mSQL   As String
'''
'''        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
'''
'''        mSQL = " Update faVouchers SET tnyStatus=4,tnyCancelFlag=1 WHERE intKeyID2=" & mPayOrderNo & " AND intExternalModuleID=1"
'''        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
'''
'''        mSQL = " UPDATE faTransactions SET tnyStatus=4 WHERE intVoucherID IN ("
'''        mSQL = mSQL + " SELECT intVoucherID FROM faVouchers  WHERE intKeyID2=" & mPayOrderNo & " AND intExternalModuleID=1  )"
'''        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
'''        mCnn.Close
'''
'''    End Sub


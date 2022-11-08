VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDemandInterface 
   BackColor       =   &H00DAF2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " D e m a n d    I n t e r f a c e"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   1635
   ClientWidth     =   11835
   Icon            =   "frmDemandInterface.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTransactionDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4950
      TabIndex        =   89
      Text            =   "txtTransactionDate"
      Top             =   720
      Width           =   1395
   End
   Begin VB.TextBox txtAccountCode 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   990
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox cmbMode 
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
      TabIndex        =   86
      Top             =   45
      Width           =   3165
   End
   Begin VB.TextBox txtTransactionType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   705
      Width           =   2835
   End
   Begin VB.CommandButton cmdSearchTransactionType 
      BackColor       =   &H00F8FFF9&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   690
      Width           =   315
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   405
      Left            =   7245
      TabIndex        =   44
      Top             =   6570
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtFunction 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   1530
      Width           =   2370
   End
   Begin VB.TextBox txtFunctionary 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   1530
      Width           =   2370
   End
   Begin VB.TextBox txtSourceOfFund 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8715
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   1530
      Width           =   2370
   End
   Begin VB.CheckBox chkTag 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Caption         =   "Address Grouping"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   6660
      Width           =   1650
   End
   Begin VB.TextBox txtReference 
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
      Left            =   8250
      TabIndex        =   67
      Top             =   780
      Width           =   3435
   End
   Begin VB.ComboBox cmbOutDoorStaff 
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
      Height          =   330
      Left            =   8250
      Style           =   2  'Dropdown List
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   435
      Width           =   3465
   End
   Begin VB.CheckBox chkSkipPrinting 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAF2F2&
      Caption         =   "Skip Printing"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   10470
      TabIndex        =   51
      Top             =   6660
      Width           =   1245
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   405
      Left            =   4095
      TabIndex        =   41
      Top             =   6570
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   5145
      TabIndex        =   42
      Top             =   6570
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CanceL"
      Height          =   405
      Left            =   6195
      TabIndex        =   43
      Top             =   6570
      Width           =   1005
   End
   Begin VB.TextBox txtDemandNo 
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
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7395
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "<New>"
      Top             =   75
      Width           =   1635
   End
   Begin VB.TextBox txtDemandDate 
      Alignment       =   2  'Center
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
      Left            =   10050
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAF2F2&
      Height          =   5265
      Left            =   0
      TabIndex        =   45
      Top             =   1305
      Width           =   11850
      Begin VB.CheckBox chkRoundOff 
         Caption         =   "Check1"
         Height          =   195
         Left            =   10455
         TabIndex        =   79
         Top             =   3240
         Width           =   210
      End
      Begin VB.CommandButton cmdSourceOfFund 
         BackColor       =   &H00F8FFF9&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11115
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cmdFunctionary 
         BackColor       =   &H00F8FFF9&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton cmdFunction 
         BackColor       =   &H00F8FFF9&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7065
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   225
         Width           =   315
      End
      Begin VB.ComboBox cmbSeat 
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
         Left            =   9030
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4650
         Width           =   2430
      End
      Begin VB.TextBox txtAdminNote 
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
         Left            =   7995
         TabIndex        =   40
         Top             =   4080
         Width           =   3450
      End
      Begin VB.TextBox txtRemarks 
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
         Left            =   7995
         TabIndex        =   38
         Top             =   3540
         Width           =   3450
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9690
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   2895
         Width           =   1755
      End
      Begin VB.TextBox txtCurrentAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2580
         Width           =   1755
      End
      Begin VB.TextBox txtArrearAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7965
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2580
         Width           =   1725
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   1815
         Left            =   60
         TabIndex        =   9
         Top             =   675
         Width           =   11745
         _cx             =   20717
         _cy             =   3201
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDemandInterface.frx":1CCA
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00DAF2F2&
         Height          =   2775
         Left            =   60
         TabIndex        =   50
         Top             =   2460
         Width           =   7800
         Begin VB.Frame fraInstrument 
            BackColor       =   &H00DAF2F2&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   60
            TabIndex        =   56
            Top             =   1470
            Width           =   2850
            Begin VB.TextBox txtDrawnPlace 
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
               Left            =   1005
               MaxLength       =   50
               TabIndex        =   60
               Top             =   930
               Width           =   1800
            End
            Begin VB.TextBox txtDrawnFrom 
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
               Left            =   1005
               MaxLength       =   50
               TabIndex        =   59
               Top             =   615
               Width           =   1800
            End
            Begin VB.TextBox txtInstrumentNo 
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
               Left            =   1005
               MaxLength       =   15
               TabIndex        =   57
               Top             =   0
               Width           =   1800
            End
            Begin VB.TextBox txtInstrumentDate 
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
               Left            =   1005
               MaxLength       =   12
               TabIndex        =   58
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label lblDrawnPlace 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Drawn Place"
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
               Left            =   45
               TabIndex        =   69
               Top             =   960
               Width           =   930
            End
            Begin VB.Label lblDrawnFrom 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Drawn From"
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
               Left            =   90
               TabIndex        =   63
               Top             =   660
               Width           =   900
            End
            Begin VB.Label lblInstDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inst.Date"
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
               Left            =   360
               TabIndex        =   62
               Top             =   345
               Width           =   645
            End
            Begin VB.Label lblInstNo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inst. No"
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
               Left            =   450
               TabIndex        =   61
               Top             =   45
               Width           =   540
            End
         End
         Begin VB.ComboBox cmbInstrumentType 
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
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1125
            Width           =   1800
         End
         Begin VB.TextBox txtPhone 
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
            Left            =   3885
            MaxLength       =   14
            TabIndex        =   36
            Top             =   2115
            Width           =   2010
         End
         Begin VB.TextBox txtPin 
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
            Left            =   6180
            MaxLength       =   6
            TabIndex        =   34
            Top             =   1800
            Width           =   915
         End
         Begin VB.TextBox txtPost 
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
            Left            =   3885
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1800
            Width           =   2025
         End
         Begin VB.TextBox txtInitial4 
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
            Left            =   7380
            MaxLength       =   1
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtInitial3 
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
            Left            =   7065
            MaxLength       =   1
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtInitial2 
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
            Left            =   6750
            MaxLength       =   1
            TabIndex        =   20
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtInitial1 
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
            Left            =   6435
            MaxLength       =   1
            TabIndex        =   19
            Top             =   210
            Width           =   315
         End
         Begin VB.TextBox txtMainPlace 
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
            Left            =   3885
            MaxLength       =   100
            TabIndex        =   30
            Top             =   1485
            Width           =   3210
         End
         Begin VB.TextBox txtLocalPlace 
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
            Left            =   3885
            MaxLength       =   100
            TabIndex        =   28
            Top             =   1170
            Width           =   3210
         End
         Begin VB.TextBox txtStreet 
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
            Left            =   3885
            MaxLength       =   100
            TabIndex        =   26
            Top             =   855
            Width           =   3210
         End
         Begin VB.TextBox txtHouseName 
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
            Left            =   3885
            MaxLength       =   100
            TabIndex        =   24
            Top             =   540
            Width           =   3210
         End
         Begin VB.TextBox txtName 
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
            Left            =   3885
            MaxLength       =   100
            TabIndex        =   18
            Top             =   210
            Width           =   2535
         End
         Begin VB.ComboBox cmbZone 
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
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   135
            Width           =   1800
         End
         Begin VB.TextBox txtDoorNo2 
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
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   16
            Top             =   795
            Width           =   690
         End
         Begin VB.TextBox txtDoorNo1 
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
            Left            =   1065
            MaxLength       =   5
            TabIndex        =   15
            Top             =   795
            Width           =   1110
         End
         Begin VB.TextBox txtWardNo 
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
            Left            =   1065
            MaxLength       =   3
            TabIndex        =   13
            Top             =   480
            Width           =   1800
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instrument"
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
            Left            =   300
            TabIndex        =   53
            Top             =   1185
            Width           =   750
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
            Left            =   3165
            TabIndex        =   35
            Top             =   2175
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
            Left            =   5940
            TabIndex        =   33
            Top             =   1845
            Width           =   630
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
            Left            =   3525
            TabIndex        =   31
            Top             =   1845
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
            Left            =   3090
            TabIndex        =   29
            Top             =   1530
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
            Left            =   3030
            TabIndex        =   27
            Top             =   1215
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
            Left            =   3420
            TabIndex        =   25
            Top             =   915
            Width           =   435
         End
         Begin VB.Label Label10 
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
            Left            =   2895
            TabIndex        =   23
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   3450
            TabIndex        =   17
            Top             =   255
            Width           =   405
         End
         Begin VB.Label Label8 
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
            Left            =   660
            TabIndex        =   10
            Top             =   195
            Width           =   375
         End
         Begin VB.Label Label7 
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
            Left            =   450
            TabIndex        =   14
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ward No"
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
            Left            =   405
            TabIndex        =   12
            Top             =   510
            Width           =   630
         End
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
         Left            =   10710
         TabIndex        =   80
         Top             =   3210
         Width           =   705
      End
      Begin VB.Label lblSourceOfFund 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Of Fund"
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
         Left            =   7515
         TabIndex        =   75
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblFunctionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
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
         Left            =   165
         TabIndex        =   74
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
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
         Left            =   4050
         TabIndex        =   73
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded To"
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
         Left            =   7935
         TabIndex        =   55
         Top             =   4710
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrative Notes( if any)"
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
         Left            =   7965
         TabIndex        =   39
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   7965
         TabIndex        =   37
         Top             =   3330
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8790
         TabIndex        =   49
         Top             =   2940
         Width           =   840
      End
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
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.ComboBox cmbSections 
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   3165
   End
   Begin VB.TextBox txtAccountHead 
      Height          =   285
      Left            =   2655
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   990
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdAcHead 
      Caption         =   ".."
      Height          =   285
      Left            =   5625
      TabIndex        =   82
      Top             =   990
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label23 
      Caption         =   "Transaction Date"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   4950
      TabIndex        =   87
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode Of Collection"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Head"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   585
      TabIndex        =   83
      Top             =   1035
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   7200
      TabIndex        =   66
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label lblCombo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out Door Collection Staff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   6435
      TabIndex        =   65
      Top             =   480
      Width           =   1785
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demand No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   6525
      TabIndex        =   7
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   9090
      TabIndex        =   5
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   180
      TabIndex        =   3
      Top             =   765
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   990
      TabIndex        =   1
      Top             =   400
      Width           =   540
   End
End
Attribute VB_Name = "frmDemandInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
    
    '*****************************************************************************************'
    '* Application ID           : 115                                                        *'
    '* Application Name         : Saankhya Double Entry                                      *'
    '* Screen id                : Receipts                                                   *'
    '* Version No               : Ver 1.0.0                                                  *'
    '* Form Designed By         : Aiby                                                       *'
    '* Created on               :                                                            *'
    '* Coding By                : Aiby                                                       *'
    '* Modified By              : Vinod                                                      *'
    '* Coded on                 : 24-Aug-2008                                                *'
    '* Reviewed By              : Prasanth Krishna                                           *'
    '* Reviewed on              : 02-Sep-2008                                                *'
    '* Purpose                  : Common Demand Generation/Interface to integrate other      *'
    '*                            Local Masters and Demand Generation Forms and Calculations *'
    '*                                                                                       *'
    '* Name of Database         : DB_Finance, DB_iSaankhyaMasters                            *'
    '* Name of Table(s)         : faiDemandTbl, faiDemandChild, faiDemandAddress             *'
    '* Look up Table(s)         :                                                            *'
    '* DSN                      : dsnFA ( UserName=FAUser; PWD=FAUser )                      *'
    '*                                                                                       *'
    '*=======================================================================================*'
    '* | Number  | Modification Date |   Modified By         |   Name of function/Variable   *'
    '* |---------|-------------------|-----------------------|-------------------------------*'
    '* |         |                   |                       |                               *'
    '* |         |                   |                       |                               *'
    '*=======================================================================================*'
    ' Notes :-
    '
    '*****************************************************************************************'
    
    Dim mNewFlag            As Boolean
    Dim mEditFlag           As Boolean
    Dim mSeatPrefix         As String
    Dim mWardNo             As Variant
    Dim mName               As Variant
    Dim mInit1              As Variant
    Dim mInit2              As Variant
    Dim mInit3              As Variant
    Dim mInit4              As Variant
    Dim mHouse              As Variant
    Dim mStreet             As Variant
    Dim mLocalPlace         As Variant
    Dim mMainPlace          As Variant
    Dim mPost               As Variant
    Dim mPin                As Variant
    Dim mPhone              As Variant
    Dim mZonalCollection    As Variant
    Dim mRoundOffDecimalPlace As Boolean
    Dim mDemandNo           As Variant
    Dim mReverse            As Variant      'Property Variable for Reverse. Set as 1 From ReverseRequestForm
    Dim mProfTaxInstTypeMode As Integer     ' 1= Trader 2=Employees 3=Self Drawing Officer
    
    'Dim mPendingTask        As Integer  '1 Pending Task in Previous Year
    'Dim mPendingTaskReqID   As Integer  'Pending Task RequestID for Demand
    
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearRequestID As Integer
    
    Dim mMonthID As Integer   '''' Change on 26 August 2014 By Sabeen
    Dim mMonthName As Variant '''' Change on 26 August 2014 By Sabeen

    
    Private Sub SetVariables(mClearFlag As Boolean)
    On Error GoTo err
        If mClearFlag Then
        
            mWardNo = Null
            mName = Null
            mInit1 = Null
            mInit2 = Null
            mInit3 = Null
            mInit4 = Null
            
            mHouse = Null
            mStreet = Null
            mLocalPlace = Null
            mMainPlace = Null
            mPost = Null
            mPin = Null
            mPhone = Null
        Else
            mWardNo = txtwardno.Text
            mName = txtName.Text
            mInit1 = txtInitial1.Text
            mInit2 = txtInitial2.Text
            mInit3 = txtInitial3.Text
            mInit4 = txtInitial4.Text
            
            mHouse = Trim(txtHouseName.Text)
            mStreet = Trim(txtStreet.Text)
            mLocalPlace = Trim(txtLocalPlace.Text)
            mMainPlace = Trim(txtMainPlace.Text)
            mPost = Trim(txtPost.Text)
            mPin = Trim(txtPin)
            mPhone = Trim(txtPhone)
        End If
        Exit Sub
err:
            MsgBox err.Description
    End Sub
    
    Private Sub GetVariables(Optional mClearFlag = False)
        On Error GoTo err
        If mClearFlag = False Then
            txtwardno.Text = mWardNo
            txtName.Text = mName
            txtInitial1.Text = mInit1
            txtInitial2.Text = mInit2
            txtInitial3.Text = mInit3
            txtInitial4.Text = mInit4
            
            txtHouseName.Text = mHouse
            txtStreet.Text = mStreet
            txtLocalPlace.Text = mLocalPlace
            txtMainPlace.Text = mMainPlace
            txtPost.Text = mPost
            txtPin.Text = mPin
            txtPhone.Text = mPhone
        Else
            txtwardno.Text = ""
            txtwardno.Tag = ""
            txtName.Text = ""
            txtName.Tag = ""
            txtInitial1.Text = ""
            txtInitial2.Text = ""
            txtInitial3.Text = ""
            txtInitial4.Text = ""
            
            txtHouseName.Text = ""
            txtStreet.Text = ""
            txtLocalPlace.Text = ""
            txtMainPlace.Text = ""
            txtPost.Text = ""
            txtPin.Text = ""
            txtPhone.Text = ""
        End If
         Exit Sub
err:
            MsgBox err.Description
    End Sub
    
        Private Sub CheckBudgetDetails(mTransactionTypeID As Variant)
            Dim mCnn    As New ADODB.Connection
            Dim Rec     As New ADODB.Recordset
            Dim objdb   As New clsDB
            Dim mSql    As String
            
            '*********************************************************************************************'
            '  Procedure to fill the Function, Functionary & Source of Fund according to Transaction Type '
            '*********************************************************************************************'
            On Error GoTo err
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            
            mSql = "Select * From faTransactionType"
            mSql = mSql + " Left Join faFunctions On faTransactionType.intFunctionID = faFunctions.intFunctionID"
            mSql = mSql + " Left Join faFunctionaries On faTransactionType.intFunctionaryID = faFunctionaries.intFunctionaryID"
            mSql = mSql + " Left Join suSourceOfFund On faTransactionType.intSourceFundID = suSourceOfFund.intSourceFundID"
            mSql = mSql + " Where intTransactionTypeID = " & mTransactionTypeID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                txtSourceofFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
            End If
            Rec.Close
            Exit Sub
err:
            MsgBox err.Description
    End Sub

    Private Sub FetchDemand(vchDemandNo As String)
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mArrIn      As Variant
        Dim mRowCount   As Integer
        Dim mArrearFlag As Variant
        
        '*********************************************************************************************'
        '       Procedure to refill the details of Demand in the Case of Zonal Collection             '
        '*********************************************************************************************'
        On Error GoTo err
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            mArrIn = Array(vchDemandNo)
            Set Rec = objdb.ExecuteSP("spGetDemandDetails", mArrIn, , , mCnn, adCmdStoredProc)
            cmbSections.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
            'cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
            
            txtDemandNo.Text = vchDemandNo
            txtDemandDate.Text = IIf(IsNull(Rec!dtDemandDate), "", Rec!dtDemandDate)
            txtDemandNo.Tag = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
            
            'If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeOutDoor Then
            If txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                cmbOutDoorStaff.Enabled = True
                If IsNull(Rec!vchUserName) = False Then
                    cmbOutDoorStaff.Text = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                End If
            End If
            If IsNull(Rec!chvZoneNameEnglish) = False Then
                cmbOutDoorStaff.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
            End If
            txtwardno.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
            If IsNull(Rec!vchInstrumentType) = False Then
                cmbInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                cmbInstrumentType.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
            End If
            If cmbInstrumentType.Tag <> "" Then
                If cmbInstrumentType.Tag <> 1 Then
                    lblInstNo.Visible = True
                    txtInstrumentNo.Visible = True
                    txtInstrumentNo.Enabled = True
                    lblInstDate.Visible = True
                    txtInstrumentDate.Visible = True
                    txtInstrumentDate.Enabled = True
                    lblDrawnFrom.Visible = True
                    txtDrawnFrom.Visible = True
                    txtDrawnFrom.Enabled = True
                    lblDrawnPlace.Visible = True
                    txtDrawnPlace.Visible = True
                    txtDrawnPlace.Enabled = True
                    
                    txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtInstrumentDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                    txtDrawnFrom.Text = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
                    txtDrawnPlace.Text = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
                End If
            End If
            txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            txtInitial1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
            txtInitial2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
            txtInitial3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
            txtInitial4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
            txtHouseName.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
            txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
            txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            txtPost.Text = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
            txtPin.Text = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
            txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                        
            txtRemarks.Text = IIf(IsNull(Rec.Fields(12).Value), "", Rec.Fields(12).Value)
            txtAdminNote.Text = IIf(IsNull(Rec!vchAdminNote), "", Rec!vchAdminNote)
            If IsNull(Rec!numForwardedSeatID) = False Then
                cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
            End If
            
            mRowCount = 1
            While Not Rec.EOF
                vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), "", Rec!tnyArrearFlag)
                If mArrearFlag = 0 Or mArrearFlag = "" Then
                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                End If
                If mArrearFlag = 1 Then
                    vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                End If
                vsGrid.Rows = vsGrid.Rows + 1
                mRowCount = mRowCount + 1
                Rec.MoveNext
            Wend
            Call Calculate
        Else
            MsgBox "Connection To Finance does not exit, Please contact your System Administrator", vbInformation
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub DispalyDemand(vchDemandNo As String)
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim mSql            As String
        Dim mDemandNo       As String
        Dim mArrearFlag     As Variant
        Dim mRowCount       As Integer
        Dim mSeatID         As Variant
        Dim mVoucherID      As Variant
        Dim mStatus         As Variant
        Dim mCancelFlag     As Variant
        
        '*********************************************************************************************'
        '                  Procedure to refill the details of Demand                                  '
        '*********************************************************************************************'

        On Error GoTo err
        mEditFlag = True
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mDemandNo = vchDemandNo
        Call FormInitialize
        If mDemandNo <> "" Then
            mSql = "Select numSeatID,intVoucherID From faIDemandTBL Where vchDemandNo='" & mDemandNo & "'"
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mSeatID = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
                mVoucherID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
            End If
            Rec.Close
            
            If mVoucherID <> "" Then
                mSql = "Select tnyStatus,tnyCancelFlag From faVouchers"
                mSql = mSql + " Where intVoucherID='" & mVoucherID & "'"
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                    mCancelFlag = IIf(IsNull(Rec!tnyCancelFlag), "", Rec!tnyCancelFlag)
                End If
                Rec.Close
                If mStatus <> "" Then
                    If mStatus = 4 Or mCancelFlag = 1 Then
                        mSql = "Update  faIDemandTBL"
                        mSql = mSql + " Set tnyStatus=0"
                        mSql = mSql + " Where intVoucherID='" & mVoucherID & "'"
                        mCnn.Execute mSql
                    End If
                End If
            End If
            If mSeatID <> "" Then
                If Trim(mSeatID) = Trim(gbSeatID) Or gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                    mSql = "Select *,faIDemandTbl.tnyStatus As Status,faIDemandTbl.numUserID As UserID,faIDemandTBL.numSeatID As SeatID From faIDemandTBL"
                    mSql = mSql + " Inner Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
                    mSql = mSql + " Inner Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
                    mSql = mSql + " Left Join faInstrumentTypes On faIDemandTBL.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
                    mSql = mSql + " Inner Join faSection On faIDemandTBL.intSectionID=faSection.intSectionID"
                    mSql = mSql + " Inner Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
                    mSql = mSql + " Inner Join faAccountHeads On faIDemandChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                    'mSQL = mSQL + " Left Join DB_Sanchayalite..snPDE_ODStaff On faIDemandTBL.intKeyID2=DB_Sanchayalite..snPDE_ODStaff.numUserID"
                    mSql = mSql + " Left Join faUser On faIDemandTBL.intKeyID2 = faUser.numUserId And tnyOutDoorStaffs = 1"
                    mSql = mSql + " Left Join faFunctions on faIDemandTBL.intFunctionID = faFunctions.intFunctionID"
                    mSql = mSql + " Left Join faFunctionaries on faIDemandTBL.intFunctionaryID = faFunctionaries.intFunctionaryID"
                    mSql = mSql + " Left Join suSourceOfFund On faIDemandTBl.intSourceFundID = suSourceOfFund.intSourceFundID"
                    mSql = mSql + " Left Join DB_Masters..GM_Zone On faIDemandTBL.numZoneID=DB_Masters..GM_Zone.numZoneID"
                    'mSql = mSql + " Left Join DB_Masters..GL_Seats On faIDemandTBL.numForwardedSeatID=DB_Masters..GL_Seats.numSeatID"
                    mSql = mSql + " Left Join faSeats On faIDemandTBL.numForwardedSeatID=faSeats.numSeatID"
                    mSql = mSql + " Where faIDemandTBL.vchDemandNo='" & mDemandNo & "'"
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        If Rec!Status <> 1 Then
                            If Rec!Status <> 9 Then
                                cmbSections.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
                                'cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                                txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                                txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                                cmdSearchTransactionType.Enabled = False
                                
                                txtDemandNo.Text = mDemandNo
                                txtDemandNo.Tag = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                                txtDemandDate.Text = IIf(IsNull(Rec!dtDemandDate), "", Rec!dtDemandDate)
                                txtDemandDate.Tag = IIf(IsNull(Rec!SeatID), "", Rec!SeatID) 'Demand Generated SeatID
                                txtReference.Tag = IIf(IsNull(Rec!UserID), "", Rec!UserID)  'Demand Generated UserID
                                
                                
                                If IsNull(Rec!intDemandMode) Then
                                    If txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                                        cmbMode.Text = "Out Door Collection"
                                    Else
                                        cmbMode.Text = "Direct"
                                    End If
                                Else
                                    cmbMode.ListIndex = IIf(IsNull(Rec!intDemandMode), "", Rec!intDemandMode)
                                End If
                                
                                If cmbMode.ItemData(cmbMode.ListIndex) Then
                                    
                                End If
                                
                                
                                If cmbMode.ItemData(cmbMode.ListIndex) = 3 Then
                                'If txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                                    cmbOutDoorStaff.Enabled = True
                                    Call FillOutDoorStaffs
                                    If IsNull(Rec!vchUserName) = False Then
                                        cmbOutDoorStaff.Text = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                                    End If
                                'ElseIf (cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeBFundSSSFund Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeMoneyOrderReturns) Then
                                ElseIf txtTransactionType.Tag = gbTransactionTypeBFundSSSFund Or txtTransactionType.Tag = gbTransactionTypeMoneyOrderReturns Then
                                    If (Rec!Status = 0) Then
                                        MsgBox "Can't edit this Demand (Already approved)"
                                        FormInitialize
                                        Exit Sub
                                    End If
                                End If
                                'End If
                                
                                If IsNull(Rec!dtTransactionDate) Then
                                Else
                                   txtTransactionDate.Text = Format(Rec!dtTransactionDate, "dd-mmm-yyyy")
                                End If
                                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                                txtSourceofFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                                txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
                                
                                If IsNull(Rec!chvZoneNameEnglish) = False Then
                                    cmbOutDoorStaff.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
                                End If
                                txtwardno.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                                txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                                txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                                If IsNull(Rec!vchInstrumentType) = False Then
                                    cmbInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                                    cmbInstrumentType.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                                End If
                                If cmbInstrumentType.Tag <> "" Then
                                    If cmbInstrumentType.Tag <> 1 Then
                                        lblInstNo.Visible = True
                                        txtInstrumentNo.Visible = True
                                        txtInstrumentNo.Enabled = True
                                        lblInstDate.Visible = True
                                        txtInstrumentDate.Visible = True
                                        txtInstrumentDate.Enabled = True
                                        lblDrawnFrom.Visible = True
                                        txtDrawnFrom.Visible = True
                                        txtDrawnFrom.Enabled = True
                                        lblDrawnPlace.Visible = True
                                        txtDrawnPlace.Visible = True
                                        txtDrawnPlace.Enabled = True
                                        
                                        txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                                        txtInstrumentDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                                        txtDrawnFrom.Text = IIf(IsNull(Rec!vchDrawnFrom), "", Rec!vchDrawnFrom)
                                        txtDrawnPlace.Text = IIf(IsNull(Rec!vchDrawnPlace), "", Rec!vchDrawnPlace)
                                    End If
                                End If
                                txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                                txtInitial1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
                                txtInitial2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
                                txtInitial3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
                                txtInitial4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
                                txtHouseName.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                                txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                                txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                                txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                                txtPost.Text = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
                                txtPin.Text = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
                                txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                                            
                                txtRemarks.Text = IIf(IsNull(Rec.Fields(12).Value), "", Rec.Fields(12).Value)
                                txtAdminNote.Text = IIf(IsNull(Rec!vchAdminNote), "", Rec!vchAdminNote)
                                If IsNull(Rec!numForwardedSeatID) = False Then
                                    cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                                End If
                                
                                mRowCount = 1
                                While Not Rec.EOF
                                    vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                                    vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                                    vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                                    vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                                    mArrearFlag = IIf(IsNull(Rec!tnyArrearFlag), "", Rec!tnyArrearFlag)
                                    If mArrearFlag = 0 Or mArrearFlag = "" Then
                                        vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                                    End If
                                    If mArrearFlag = 1 Then
                                        vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                                    End If
                                    vsGrid.Rows = vsGrid.Rows + 1
                                    mRowCount = mRowCount + 1
                                    Rec.MoveNext
                                Wend
                                Call Calculate
                            Else
                                MsgBox "Can't edit this Demand (Demand Cancelled)", vbCritical
                                Exit Sub
                            End If
                        Else
                            MsgBox "Can't edit the Demand (Receipt Issued)", vbCritical
                            Exit Sub
                        End If
                    Else
                        MsgBox "Demand Number doesn't exists !!", vbInformation
                    End If
                    Rec.Close
                Else
                    MsgBox "You are not authorized to edit this Demand!", vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "Demand Number doesn't exists !!", vbInformation
                Exit Sub
            End If
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    
    Private Sub FineCalculation()
        Dim mCount As Long
        Dim mAccountHeadCode As String
        Dim mPeriodId As Integer
        Dim mYearID As Integer
        Dim mMonthID As Integer
        Dim mNoOfMonths As Integer
        Dim mDate As Date
        Dim mTransactionTypeID As Integer
        Dim mAmtCurrent As Double
        Dim mFine As Double
        Dim mSql As String
        
        
        ' Find Arrear AccountHead Code as per transactiontype
        'mTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
        mTransactionTypeID = txtTransactionType.Tag
        Select Case mTransactionTypeID
            Case gbTransactionTypeProfTaxTrade
                mAccountHeadCode = "431190102"
        End Select
        
        
        For mCount = 1 To vsGrid.Rows - 1
            'Find Matching Arrear Head Code/Year/Period
            If val(vsGrid.TextMatrix(mCount, 0)) = mAccountHeadCode Then
                
                'Year
                If vsGrid.Cell(flexcpValue, mCount, 2) > 0 Then
                    mYearID = vsGrid.Cell(flexcpValue, mCount, 2)
                End If
                
                'Period
                If vsGrid.Cell(flexcpValue, mCount, 3) > 0 Then
                    mPeriodId = vsGrid.Cell(flexcpValue, mCount, 3)
                End If
                
                'Month
                If mPeriodId = 1 Then
                    mMonthID = 9
                Else
                    mMonthID = 3
                End If
                
                'Find Year as per Period
                If mPeriodId = 2 Then mYearID = mYearID + 1
                
                'Find Fine Date
                mDate = DateSerial(mYearID, mMonthID, 1)
                
                'Find No of Months from the Transaction Date
                mMonthID = DateDiff("m", mDate, gbTransactionDate) + 1
                
                'Calculate the Fine
                If mYearID <= gbFinancialYearID Then
                    If mYearID = gbFinancialYearID And mPeriodId = gbCurrentPeriodID Then
                        mFine = mFine + 0
                    Else
                        mFine = mFine + val(vsGrid.TextMatrix(mCount, 4)) * mMonthID / 100
                    End If
                End If
            End If
        Next
        If mFine > 0 Then
            For mCount = 1 To vsGrid.Rows - 1
                If val(vsGrid.TextMatrix(mCount, 0)) = gbAcHeadCodePenalInterest Then
                    Exit For
                End If
            Next
            If mCount <= vsGrid.Rows - 1 Then
                vsGrid.TextMatrix(mCount, 5) = Format(mFine, "0.00")
            Else
                mSql = gbAcHeadCodePenalInterest & vbTab & "Penal Interest" & vbTab & vbTab & vbTab & vbTab & Format(mFine, "0.00")
                vsGrid.AddItem mSql, mCount
            End If
        End If
        
    End Sub
    
    Public Sub Calculate()
        Dim mAmtArrear As Double
        Dim mAmtCurrent As Double
        Dim mCount As Long
        For mCount = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mCount, 4)) Then
                mAmtArrear = mAmtArrear + val(vsGrid.Cell(flexcpText, mCount, 4))
            Else
                mAmtCurrent = mAmtCurrent + val(vsGrid.Cell(flexcpText, mCount, 5))
            End If
        Next
        txtArrearAmt.Text = Format(mAmtArrear, "0.00")
        txtCurrentAmt.Text = Format(mAmtCurrent, "0.00")
        txtGrandTotal.Text = Format(val(txtArrearAmt) + val(txtCurrentAmt), "0.00")
        
        'Call FineCalculation
        
       
    End Sub
    
    Private Sub ValuesForHiddenColumns()
        Dim mYearID As Integer
        If vsGrid.Row = 0 Then Exit Sub
        
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 2)) Then
            vsGrid.TextMatrix(vsGrid.Row, 7) = vsGrid.TextMatrix(vsGrid.Row, 2)
        Else
             vsGrid.TextMatrix(vsGrid.Row, 7) = ""
             vsGrid.TextMatrix(vsGrid.Row, 2) = ""
        End If
        
        If vsGrid.TextMatrix(vsGrid.Row, 3) = "" Then  'Period
            vsGrid.TextMatrix(vsGrid.Row, 8) = 3
        Else
            vsGrid.TextMatrix(vsGrid.Row, 8) = vsGrid.TextMatrix(vsGrid.Row, 3)
        End If
        
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 7)) Then
            If vsGrid.TextMatrix(vsGrid.Row, 7) < gbFinancialYearID Then  ' Arrear Flag
                vsGrid.TextMatrix(vsGrid.Row, 9) = 1
            Else
                vsGrid.TextMatrix(vsGrid.Row, 9) = 0
            End If
        Else
            vsGrid.TextMatrix(vsGrid.Row, 9) = 0
        End If
        
        If val(vsGrid.TextMatrix(vsGrid.Row, 4)) > 0 Then   'Arrear Amount
            vsGrid.TextMatrix(vsGrid.Row, 11) = val(vsGrid.TextMatrix(vsGrid.Row, 4))
        Else                                          'Current Amount
            vsGrid.TextMatrix(vsGrid.Row, 11) = val(vsGrid.TextMatrix(vsGrid.Row, 5))
        End If
        
    End Sub
    
    Private Sub FormInitialize()
        vsGrid.Clear 1, 0
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.Value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        
        cmdSearchTransactionType.Enabled = True
        cmbInstrumentType.Enabled = True
        
        txtTransactionType.Text = ""
        txtTransactionType.Tag = ""
        On Error Resume Next
        cmbSections.Text = GetSetting("iSaankhyaMasters", "DemandGenerator", "Section", "")
        cmbZone.Text = GetSetting("iSaankhyaMasters", "DemandGenerator", "Zone", "")
        chkSkipPrinting.Value = GetSetting("iSaankhyaMasters", "DemandGenerator", "SkipPrint", "0")
        
        mNewFlag = True
        cmdSave.Enabled = True
        cmbInstrumentType.Text = "Cash"
        cmbOutDoorStaff.Enabled = False
        On Error GoTo 0
        Call FinancialYearSetForPEndingTask
        If mPreviousYearMode = 1 Then
            'txtDemandDate.Text = DdMmmYy(gbTransactionDate)
            cmdNew.Enabled = False
        Else
            txtDemandDate.Text = DdMmmYy(gbTransactionDate)
        End If
        
        
        
        cmbMode.Text = "Direct"
        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
            cmbSections.Text = "Panchayat Office"
        End If
        On Error GoTo 0
             
    End Sub
    Private Sub FinancialYearSetForPEndingTask()
        Dim mSql    As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim Trndate     As Date
        Dim mTrnYear    As Integer
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            If mPreviousYearMode = 1 Then
                mSql = "Select * From faPendingTaskRequest Where intRequestID=" & mPreviousYearRequestID
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.EOF Or Rec.BOF) Then
                        Trndate = Rec!dtTransactionDate
                        mTrnYear = gbFinancialYearID - 1
                        'gbTransactionDate = Trndate
                        'gbFinancialYearID = mTrnYear
                        If Not IsNull(Rec!intTransactionTypeID) Then
                            txtTransactionType.Text = FindMaster("faTransactionType", "vchTransactionType", "intTransactionTypeID", Rec!intTransactionTypeID)
                            txtTransactionType.Tag = Rec!intTransactionTypeID
                            cmdSearchTransactionType.Enabled = False
                            txtTransactionType_GotFocus
                        End If
                        If IsDate(Rec!dtTransactionDate) Then
                            txtTransactionDate.Text = DdMmmYy(Rec!dtTransactionDate)
                            txtDemandDate.Text = txtTransactionDate.Text
                        End If
                        
                        If Not IsNull(Rec!intInstrumentTypeID) Then
                            If Rec!intInstrumentTypeID > 0 Then
                                cmbInstrumentType.Text = FindMaster("faInstrumentTypes", "vchInstrumentType", "intInstrumentTypeID", Rec!intInstrumentTypeID)
                            End If
                            cmbInstrumentType.Enabled = False
                        End If
                        
                        If IsNumeric(Rec!fltAmount) Then
                            txtGrandTotal.Tag = Rec!fltAmount
                        Else
                            txtGrandTotal.Tag = "0"
                        End If
                    End If
                    Rec.Close
            Else
                mSql = "Select *,GetDate() as TrnDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    mTrnYear = Rec!intFinancialYear
                    Trndate = Rec!Trndate
                    'gbTransactionDate = Trndate
                    'gbFinancialYearID = mTrnYear
                End If
            End If
        End If
    End Sub
    
    Private Sub FillInstruments()
        Dim mSql As String
        If mReverse Then
            mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumenttypeID<>6"
        Else
            mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes "
        End If
        Call PopulateList(cmbInstrumentType, mSql, "Cash", True, True, True)
    End Sub
    Private Sub FillOutDoorStaffs()
        On Error GoTo last
        Dim mSql As String
        mSql = "SELECT     vchEmpName, numUserID FROM GM_User WHERE (tnyOutDoorStaffs = 1) And intLBID = " & gbLocalBodyID
        Call PopulateList(cmbOutDoorStaff, mSql, , True, True, True, enuSourceString.DBMaster)
'        mSQL = mSQL + "SELECT chvEmployeeName, numUserID From snPDE_ODStaff Where intLBID = " & gbLocalBodyID
'        Call PopulateList(cmbOutDoorStaff, mSQL, , True, True, True, SanchayaLite)
        Exit Sub
last:
        MsgBox "Contact Administrator & Add Out Door Staafs through Admin Module-->Saankhya-->OutdoorStaff", vbInformation
    End Sub
    Private Sub FillZoneInSub()
        Call PopulateList(cmbOutDoorStaff, "Select chvZoneNameEnglish, numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", , True, True, True, enuSourceString.DBMaster)
    End Sub
    Private Sub FillZone()
       Call PopulateList(cmbZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Where intLBID = " & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, enuSourceString.DBMaster)
    End Sub
    Private Sub FillSections()
        Dim mSql As String
        If gbLBType = 3 Or gbLBType = 4 Then
            mSql = "Select vchSectionName, intSectionID From faSection Order By vchSectionName"
        Else
            mSql = "Select vchSectionName, intSectionID From faSection WHERE intSectionID > 99 Order By vchSectionName"
        End If
        Call PopulateList(cmbSections, mSql, , True, , True)
     End Sub
    Private Sub FillSeats()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim mQuery As String
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
'        mQuery = "Select Left(Convert( VarChar(10),numSeatID), 7) As Prefix From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
'        Rec.Open mQuery, mCnn
'        If Not (Rec.EOF And Rec.BOF) Then
'            mSeatPrefix = IIf(IsNull(Rec!Prefix), "", Rec!Prefix)
'        End If
'        Rec.Close
        mQuery = "Select intLocationID From faLBSettings Where intLBID = " & gbLocalBodyID & " And tnyLBTypeID = " & gbLBType
        Rec.Open mQuery, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mSeatPrefix = IIf(IsNull(Rec!intLocationID), "", Rec!intLocationID)
        End If
        Rec.Close
        
        mSql = "Select chvSeatTitle, Right(Convert( VarChar(10),numSeatID), 3) From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " And intGroupID IN(" & gbSeatGroupAccountsOfficer & "," & gbSeatGroupAccountsSuperintended & ") Order By chvSeatTitle"
        'mSQL = "Select chvSeatTitle, numSeatID From GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        Call PopulateList(cmbSeat, mSql, , True, True, True, enuSourceString.DBMaster)
        
    End Sub
'    Private Sub FillTransactionTypes()
'        Dim mSql As String
'        Dim mSectionID As Long
'        If cmbSections.ListIndex > 0 Then
'            mSectionID = cmbSections.ItemData(cmbSections.ListIndex)
''            If mSectionID = 99 Then
''                mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 Order By vchTransactionType"
''            Else
''                mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 And intSectionID = " & mSectionID & " Order By vchTransactionType"
''            End If
'        '**********************************************'
'        '   Modified By Poornima On 07-June-2010
'        '**********************************************'
'
'            If mSectionID = 99 Then
'                mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
'                mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
'                mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
'                mSql = mSql + " Where (faTransactionType.intGroupID = 10) And faSectionWiseTransactionTypes.tnyList = 1"
'                mSql = mSql + " And isNull(tnyHidden,0)=0"
'                mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
'            Else
'                mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
'                mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
'                mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
'                mSql = mSql + " Where (faTransactionType.intGroupID = 10)And faSectionWiseTransactionTypes.tnyList =1 And  faSectionWiseTransactionTypes.intSectionID =  " & mSectionID
'                mSql = mSql + " And isNull(tnyHidden,0)=0"
'                mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
'            End If
'            If mSql = "" Then
'                If mReverse = 1 Then
'                    mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
'                    mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
'                    mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
'                    mSql = mSql + " Where (faTransactionType.intGroupID = 10)"
'                   '' mSQL = mSQL + " And isNull(tnyHidden,0)=0"
'                    mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
'                End If
'            End If
'            Call PopulateList(cmbTransactionType, mSql, , True, , True)
'        Else
'            cmbTransactionType.Clear
'            cmbTransactionType.AddItem ""
'        End If
'    End Sub
    Private Sub FillGridYear()
        Dim mLoop As Integer
        Dim mItem As String
        mItem = "#0; "
        For mLoop = gbFinancialYearID + 1 To 1970 Step -1
            mItem = mItem & "|#" & mLoop & ";" & CStr(mLoop) & "-" & CStr(mLoop + 1)
        Next
        vsGrid.ColComboList(2) = mItem
        
        'mItem = "#0; "
        'mItem = mItem & "|#" & 1 & "; First Half"
        'mItem = mItem & "|#" & 2 & "; Second Half"
        'mItem = mItem & "|#" & 3 & "; Full Year"
        'vsGrid.ColComboList(3) = mItem
        
         'Note:- Filling Month
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        objdb.SetConnection mCnn
        mSql = "Select * From faPeriodicity"
        mItem = "#0; "
        Rec.Open mSql, mCnn, adOpenForwardOnly, adLockReadOnly
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                mItem = mItem & "|#" & Rec!intPeriodicityID & "; " & Rec!vchperiodicity
                Rec.MoveNext
            Wend
        End If
        vsGrid.ColComboList(3) = mItem
        
    End Sub
    Private Sub FillAccountHeads()
        Call gFillVSGrid(vsGrid, 1, "spGetAccHead4Receipts", enuSourceString.Saankhya)
    End Sub

Private Sub chkRoundOff_Click()
    mRoundOffDecimalPlace = chkRoundOff.Value
End Sub

Private Sub chkTag_Click()
    If chkTag.Value = 0 Then
        Call GetVariables(True)
    End If
End Sub

    Private Sub cmbInstrumentType_Click()
        Dim mInstrumentTypeID As Long
        If cmbInstrumentType.ListIndex > -1 Then
            mInstrumentTypeID = cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
            If mInstrumentTypeID <> gbInstrumentCash Then
                fraInstrument.Visible = True
                txtInstrumentNo.Enabled = True
                txtInstrumentDate.Enabled = True
                txtDrawnFrom.Enabled = True
                txtDrawnPlace.Enabled = True
            Else
                fraInstrument.Visible = False
                txtInstrumentNo.Text = ""
                txtInstrumentDate.Text = ""
                txtDrawnFrom.Text = ""
                txtDrawnPlace.Text = ""
            End If
        End If
    End Sub
    Private Sub cmbInstrumentType_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
Private Sub cmbMode_Click() ' Modified  on 05/07/2011
     ' Call SetVariables(True)
     
     If cmbMode.ItemData(cmbMode.ListIndex) = 1 Then
                If mPreviousYearMode <> 1 Then
                    txtTransactionDate.Enabled = False
                    txtTransactionDate.Text = ""
                    ' Call FillOutDoorStaffs
                End If
     End If
     If cmbMode.ItemData(cmbMode.ListIndex) = 3 Then 'gbTransactionTypeOutDoor Then
                Call FillOutDoorStaffs
                lblCombo.Caption = "Out Door Collection Staff"
                cmbOutDoorStaff.Enabled = True
                lblCombo.Enabled = True
                lblCombo.Left = 6450
                lblCombo.Top = 480
                If mPreviousYearMode <> 1 Then
                    txtTransactionDate.Enabled = True
                    txtTransactionDate.Text = ""
                End If
                'ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeZonalCollection Then
    ElseIf cmbMode.ItemData(cmbMode.ListIndex) = 2 Then 'gbTransactionTypeZonalCollection Then
                'cmbTransactionType.Tag = gbTransactionTypeZonalCollection
                txtTransactionType.Tag = gbTransactionTypeZonalCollection
                Call FillZoneInSub
                lblCombo.Caption = "Zonal Office"
                lblCombo.Left = 7260
                lblCombo.Top = 485
                cmbOutDoorStaff.Enabled = True
                lblCombo.Enabled = True
                If mPreviousYearMode <> 1 Then
                    txtTransactionDate.Enabled = True
                    txtTransactionDate.Text = ""
                End If
    ElseIf cmbMode.ItemData(cmbMode.ListIndex) = 3 Then 'gbTransactionTypeZonalCollection Then
              '  cmbTransactionType.Tag = gbTransactionTypeZonalCollection
              '  txtTransactionType.Tag = gbTransactionTypeZonalCollection
              '  Call FillZoneInSub
              '  lblCombo.Caption = "Zonal Office"
              '  lblCombo.Left = 7260
              '  lblCombo.Top = 485
              '  cmbOutDoorStaff.Enabled = True
              '  lblCombo.Enabled = True
                If mPreviousYearMode <> 1 Then
                    txtTransactionDate.Enabled = True
                    txtTransactionDate.Text = ""
                End If
    ElseIf cmbMode.ItemData(cmbMode.ListIndex) = 5 Then
        If mPreviousYearMode <> 1 Then
            txtTransactionDate.Enabled = True
            txtTransactionDate.Text = ""
        End If
    Else
            lblCombo.Enabled = False
            cmbOutDoorStaff.Enabled = False
                
   End If
End Sub


    Private Sub cmbOutDoorStaff_Click()
        If cmbOutDoorStaff.ListIndex > 0 Then
            cmbOutDoorStaff.Tag = cmbOutDoorStaff.ItemData(cmbOutDoorStaff.ListIndex)
            txtName.Text = cmbOutDoorStaff.Text
        Else
            cmbOutDoorStaff.Tag = ""
        End If
    End Sub

    Private Sub cmbSections_Click()
        'Call FillTransactionTypes
    End Sub
'    Private Sub checkProfTaxEmpOrTraders()
'        Dim objDb   As New clsDb
'        Dim mCnn    As New ADODB.Connection
'        Dim Rec     As New ADODB.Recordset
'        Dim mSql    As Variant
'
'        objDb.SetConnection mCnn
'        mSql = " Select tnyLinkProfessionTaxEmployee from faConfig"
'        Rec.Open mSql, mCnn
'        If Rec!tnyLinkProfessionTaxEmployee = 1 Then
'            frmSearchProfTaxInstitutions.Show vbModal
'        End If
'        Rec.Close
'    End Sub


'    Private Sub cmbTransactionType_Click()
'        'On Error GoTo err
'        '-----------------For Clearing the Budget Details-----------------------'
'        txtFunctionary.Text = ""
'        txtFunctionary.Tag = ""
'        txtFunction.Text = ""
'        txtFunction.Tag = ""
'        txtSourceOfFund.Text = ""
'        txtSourceOfFund.Tag = ""
'        '------------------------------------------------------------------------'
'        If cmbTransactionType.ListIndex > 0 Then
'            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeOutDoor Then
'                Call FillOutDoorStaffs
'                lblCombo.Caption = "Out Door Collection Staff"
'                cmbOutDoorStaff.Enabled = True
'                lblCombo.Enabled = True
'                lblCombo.Left = 5565
'                lblCombo.Top = 480
'            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeZonalCollection Then
'                cmbTransactionType.Tag = gbTransactionTypeZonalCollection
'                Call FillZoneInSub
'                lblCombo.Caption = "Zonal Office"
'                lblCombo.Left = 6465
'                lblCombo.Top = 485
'                cmbOutDoorStaff.Enabled = True
'                lblCombo.Enabled = True
'            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxEmp Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxTrade Then
'                'If gbLinkWithProfTaxEmp Then
'                '    frmSearchProfTaxInstitutions.Show vbModal
'                'End If
'            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeBFundSSSFund Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeMoneyOrderReturns Then
'                If gbLBType = 3 Then
'                    Dim mCnn    As New ADODB.Connection
'                    Dim objDb   As New clsDB
'                    Dim Rec     As New ADODB.Recordset
'                    Dim mSql    As String
'                    Dim i       As Integer
'
'                    If (objDb.CreateNewConnection(mCnn, enuSourceString.DBMaster)) Then
'                        mSql = "Select Right(Convert(Varchar(20),numSeatID),3) As SeatID,chvSeatTitle From GL_Seats Where intGroupID = " & gbSeatGroupAccountsOfficer
'                        Rec.Open mSql, mCnn
'                        If Not (Rec.EOF And Rec.BOF) Then
'                            cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
'                            For i = 0 To cmbSeat.ListCount - 1
'                                If (cmbSeat.List(i) = Rec!chvSeatTitle) Then
'                                    cmbSeat.ListIndex = i
'                                End If
'                            Next
'                            'cmbSeat.ItemData(cmbSeat.ListIndex) = IIf(IsNull(Rec!SeatID), "", Rec!SeatID)
'                        End If
'                        Rec.Close
'                    Else
'                        MsgBox "Connection To Master does not exit, Please contact your System Administrator", vbInformation
'                        Exit Sub
'                    End If
'                End If
'            Else
'                lblCombo.Caption = "Out Door Collection Staff"
'                cmbOutDoorStaff.Enabled = False
'                lblCombo.Enabled = False
'                lblCombo.Left = 5565
'                lblCombo.Top = 480
'            End If
'            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) <> 0 Then
'                Call CheckBudgetDetails(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
'            End If
'        End If
'        Exit Sub
'err:
'        MsgBox err.Description
'    End Sub
'    Private Sub cmbTransactionType_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
'    End Sub
'    Private Sub cmbTransactionType_LostFocus()
'        If cmbTransactionType.ListIndex > -1 Then
'            If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxEmp Then
'                If gbLinkWithProfTaxEmp Then
'                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 2
'                    frmSearchProfTaxInstitutions.Show vbModal
'                End If
''            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxTrade Then
''                If gbLinkWithProfTaxEmp Then
''                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 1
''                    frmSearchProfTaxInstitutions.Show vbModal
''                End If
'            End If
'        End If
'    End Sub

'    Private Sub cmbTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
'        Select Case val(cmbTransactionType.Tag)
'            Case Is = gbTransactionTypeProfTaxEmp Or gbTransactionTypeProfTaxTrade
'                If gbLinkWithProfTaxEmp Then
'                    frmSearchProfTaxInstitutions.Show vbModal
'                End If
'        End Select
'        If val(cmbTransactionType.Tag) = gbTransactionTypeProfTaxEmp Or gbTransactionTypeProfTaxTrade Then
'            If gbLinkWithProfTaxEmp Then
'                frmSearchProfTaxInstitutions.Show vbModal
'            End If
'        End If
'    End Sub
'    Private Sub cmbTransactionType_LostFocus()
'        If cmbTransactionType.ListIndex > 0 Then
'            Select Case val(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
'                Case Is = gbTransactionTypeProfTaxEmp Or gbTransactionTypeProfTaxTrade
'                    If gbLinkWithProfTaxEmp Then
'                        frmSearchProfTaxInstitutions.Show vbModal
'                    End If
'            End Select
'        End If
'    End Sub

    Private Sub cmbZone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    
    Private Sub cmdAcHead_Click()
        gbSearchStr = ""
        gbSearchID = -1
        Dim mSql As String
            If cmbInstrumentType.ListIndex > 0 Then
                Select Case cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
                    Case 1 '[Cash]
                     mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.tinHiddenFlag = 0 AND  faAccountHeads.intGroupID =1 "
                    Case Else
                     mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.tinHiddenFlag = 0 AND faAccountHeads.intGroupID =2"
                End Select
                frmSearchAccountHeads.VoucherMode = 300
                frmSearchAccountHeads.SQLString = mSql
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.Show vbModal
                txtAccountCode.SetFocus
            Else
                MsgBox "Please select an Instrument", vbInformation
                cmbInstrumentType.Enabled = True
                cmbInstrumentType.SetFocus
            End If
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdFunction_Click()
        gbSearchStr = ""
        gbSearchID = -1
        On Error GoTo err:
        frmSearchFunction.Show vbModal
        If Not gbSearchStr = "" Then
            txtFunction.Text = gbSearchStr
            txtFunction.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFunctionary_Click()
        gbSearchStr = ""
        gbSearchID = -1
        On Error GoTo err:
            frmSearchFunctionary.Show vbModal
            If Not gbSearchStr = "" Then
                txtFunctionary.Text = gbSearchStr
                txtFunctionary.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdNew_Click()
        If chkTag.Value = 1 Then
            Call SetVariables(False)
        Else
            Call SetVariables(True)
        End If
        
        Call FormInitialize
        DemandNo = ""
        cmdSave.Caption = "Save"
        cmdSearchTransactionType.Enabled = True
        If chkTag.Value = 1 Then
            Call GetVariables
        End If
    End Sub
    


'''    Private Sub cmdReject_Click() 'ADDED BY MINU FOR REJECTIONS
'''        frmReject.Mode = 1
'''        frmReject.RequestTypeID = txtDemandNo.Tag
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''    End Sub

    Private Sub cmdSave_Click()
        Dim intLBID             As Variant
        Dim tnyExtAppID         As Variant
        Dim tnyExtModuleID      As Variant
        Dim tnyDemandType       As Variant
        Dim intTransactionTypeID As Variant
        Dim intYearID           As Variant
        Dim tnyPeriodID         As Variant
        Dim dtDemandDate        As Variant
        Dim numSubLedgerID      As Variant
        Dim intKeyID            As Variant
        Dim intKeyID2           As Variant
        Dim vchRemarks          As Variant
        Dim tnyStatus           As Variant
        Dim intVoucherID        As Variant
        Dim dtVoucherDate       As Variant
        Dim tnyArrearFlag       As Variant
        Dim dtExpiryDate        As Variant
        Dim numDemandID         As Variant
        Dim vchDemandNo         As Variant
        Dim intSectionID        As Variant
        Dim numSeatID           As Variant
        Dim numUserID           As Variant
        Dim vchAdminNote        As Variant
        Dim intAccountHeadID    As Variant
        Dim vchAccountHeadCode  As Variant
        Dim fltAmount           As Variant
        Dim dtOnDate            As Variant
        Dim numZoneID           As Variant
        Dim intWardNo           As Variant
        Dim intDoorNo           As Variant
        Dim vchDoorNo2          As Variant
        Dim numForwardedSeatID  As Variant
        Dim ForwardedSeatID     As String
        Dim dtDueDate           As Variant
        Dim dtTransactionDate   As Variant  ' Added on 05/07/2011
        Dim intDemandMode       As Integer  ' Added on 07/07/2011
        Dim numRefSubID         As Double   ' Added On 12 Jul 2011
        Dim intInstrumentTypeID As Variant
        Dim vchInstrumentNo     As Variant
        Dim dtInstrumentDate    As Variant
        Dim vchDrawnFrom        As Variant
        Dim vchDrawnPlace       As Variant
        Dim tnyAccrualType      As Variant
        
        
        Dim vchName_6           As String
        Dim vchInit1_7          As String
        Dim vchInit2_8          As String
        Dim vchInit3_9          As String
        Dim vchInit4_10         As String
        Dim vchHouseName_11     As String
        Dim vchStreet_12        As String
        Dim vchLocalPlace_13    As String
        Dim vchMainPlace_14     As String
        Dim vchPost_15          As String
        Dim vchPin_16           As String
        Dim vchPhone_17         As String
        Dim mServerDate         As String
        
        Dim mLoop               As Variant
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        
        Dim objdb               As New clsDB
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim mFinYearID          As Integer
        
        '*********************************************************************************************'
        '                           Procedure to Save the Demand                                      '
        '*********************************************************************************************'
        '--------------------------------------------------------------'
        ' S e c t i o n                                                '
        '--------------------------------------------------------------'
        If cmbSections.ListIndex < 1 Then
            MsgBox "Choose a Section, Please!", vbInformation
            cmbSections.SetFocus
            Exit Sub
        Else
            intSectionID = cmbSections.ItemData(cmbSections.ListIndex)
        End If
        
        '--------------------------------------------------------------'
        ' T r a n s a c t i o n  T y p e                               '
        '--------------------------------------------------------------'
        'If cmbTransactionType.ListIndex <= 0 Then
        If val(txtTransactionType.Tag) <= 0 Then
            MsgBox "Choose the type of Transaction!", vbInformation
'            cmbTransactionType.SetFocus
            cmdSearchTransactionType.SetFocus
            Exit Sub
        Else
            'intTransactionTypeID = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
            intTransactionTypeID = txtTransactionType.Tag
        End If
        
        '--------------------------------------------------------------'
        ' Check whether Budget Centre and Source of Fund is defined    '                               '
        '--------------------------------------------------------------'
        If intTransactionTypeID <> "" Then
            If txtFunctionary.Text = "" Then
                MsgBox "Please select the Functionary !", vbInformation
                cmdFunctionary.SetFocus
                Exit Sub
            End If
            If txtFunction.Text = "" Then
                MsgBox "Please select the Function !", vbInformation
                'cmdFunction.SetFocus
                Exit Sub
            End If
            If txtSourceofFund.Text = "" Then
                MsgBox "Please select the Source of Fund !", vbInformation
                cmdSourceOfFund.Enabled = True
                cmdSourceOfFund.SetFocus
                Exit Sub
            End If
        End If
        '-------Added by Sunil Babu on 05-07-2011-----------------------
        '-------Modified by Anisha on 07-07-2011------------------------
        If cmbMode.ItemData(cmbMode.ListIndex) < 1 Then
            MsgBox "Please Select A Mode of Collection", vbInformation
            Exit Sub
        Else
            intDemandMode = cmbMode.ItemData(cmbMode.ListIndex)
            If cmbMode.ItemData(cmbMode.ListIndex) <> 1 Then
                If txtTransactionDate.Text = "" Then
                    MsgBox "Please Enter the Transaction date", vbInformation
                    Exit Sub
                Else
                    If CDate(txtTransactionDate.Text) > gbTransactionDate Then
                        MsgBox "Please Enter Valid Date", vbApplicationModal
                        Exit Sub
                    Else
                        dtTransactionDate = txtTransactionDate.Text
                    End If
                End If
            Else
                dtTransactionDate = DdMmmYy(gbTransactionDate)
            End If
            
        End If
        If mZonalCollection = 1 Then
            intDemandMode = 9
        End If
        '-------------------------------------------------------------------'
        ' Mode: Direct Transaction Type : Property Tax::: Validate Door No  '
        '-------------------------------------------------------------------'
            If cmbMode.ItemData(cmbMode.ListIndex) = 1 Then
                If val(txtTransactionType.Tag) = gbTransactionTypePTax Then
                    If val(txtwardno.Text) = 0 Then
                        MsgBox "Please enter the Ward No", vbInformation
                        txtwardno.SetFocus
                        Exit Sub
                    End If
                    If val(txtDoorNo1.Text) = 0 Then
                        MsgBox "Please enter the Door No", vbInformation
                        txtDoorNo1.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        '-------------------------------------------------------------------'
        
        '--------------------------------------------------------------'
        ' Check whether any Account Head is selected or not            '
        '--------------------------------------------------------------'
        If val(txtGrandTotal) <= 0 Then
            MsgBox "Choose an Account head and an Amount"
            Exit Sub
        End If
        If Trim(txtName.Text) = "" And Trim(txtHouseName) = "" Then
            MsgBox "Enter the Name and condtact info, if any..!", vbInformation
            txtName.SetFocus
            Exit Sub
        End If
        '--------------------------------------------------------------'
        ' F o r w a r d e d   t o   S e a t                            '
        '--------------------------------------------------------------'
        If cmbSeat.ListIndex > 0 Then
            numForwardedSeatID = cmbSeat.ItemData(cmbSeat.ListIndex)
            ForwardedSeatID = Right("000" + CStr(numForwardedSeatID), 3)
            numForwardedSeatID = mSeatPrefix + ForwardedSeatID
        Else
            numForwardedSeatID = Null
        End If
        
        'NOTE:-  ADDED ON 06-May-2013 ' PREVIOUS YEAR TRANSACTIONS
        If mPreviousYearMode = 1 Then
            If val(txtGrandTotal.Text) <> val(txtGrandTotal.Tag) Then
                MsgBox "Requested Amount for Demand is Rs." & val(txtGrandTotal.Tag)
                Exit Sub
            End If
        End If
        
        
        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandTBL          **'
        '---------------------------------------------------------'
        '
        '@intLBID            int,
        '@tnyExtAppID        As Integer
        '@tnyExtModuleID     As Integer
        '@tnyDemandType      As Integer
        '@intTransactionTypeID   smallint,
        '@intYearID          smallint,
        '@tnyPeriodID        tinyint,
        '@dtDemandDate       smalldatetime,
        '@numSubLedgerID     Numeric,
        '@intKeyID           Numeric,
        '@intKeyID2          Numeric,
        '@vchRemarks         varChar(100),
        '@tnyStatus          TinyInt,
        '@intVoucherID       Int ,
        '@dtVoucherDate      SmallDateTime,
        '@tnyArrearFlag      TinyInt = 0,
        '@dtExpiryDate       SmallDateTime = Null,
        '@numDemandID        Numeric = Null Output
        '
        '@numSeatID          Numeric     = Null,
        '@intSectionID       Int         = Null,
        '@numUserID          Numeric     = Null,
        '@numCounterID       Numeric     = Null,
        '@vchAdminNote       varChar(100)    = Null,
        '@vchDemandNo        varChar(20)     = Null,
        '@numZoneID          Numeric     = Null,
        '@intWardNo          Int         = Null,
        '@intDoorNo         int         = Null,
        '@vchDoorNo2        varChar(10)     = Null
        '@intForwaredSeat   Numeric         = Null
        '@dtDueDate         SmallDateTime = Null
        
        '----------------------------------------------------------------------------------'
        ' Added by Vinod on 26-Mar-2011                                                  '
        '----------------------------------------------------------------------------------'
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                Set Rec = mCnn.Execute("Select GetDate()")
                If IsDate(Rec.Fields(0)) Then
                    mServerDate = DdMmmYy(Rec.Fields(0))
                Else
                    MsgBox "Didn't able to Access Server Date", vbInformation
                    Exit Sub
                End If
                Rec.Close
                Set mCnn = Nothing
            End If
        '----------------------------------------------------------------------------------'
        
        intLBID = gbLocalBodyID
        tnyExtAppID = AppID.Saankhya
        If mReverse = 1 Then
            ' For Reverse Entry  Added On 18/8/2010
            tnyExtModuleID = 55
            If val(txtAccountHead.Tag) > 0 Then
                intKeyID = val(txtAccountHead.Tag)
            End If
            If cmdAcHead.Enabled Then
                If val(txtAccountHead.Tag) > 0 Then
                    intKeyID = val(txtAccountHead.Tag)
                Else
                    MsgBox "Please Select Account Head Code (Bank /Cash)", vbApplicationModal
                    Exit Sub
                End If
                '''''''------------------------------------------------------------------
                '''''''Modified On 21/03/2011 By Anisha For Reverse Mode
                '''''''------------------------------------------------------------------
                
                If val(txtAccountHead.Tag) = gbAcHeadIDCash Then
                   If cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex) <> 1 Then
                        MsgBox "Please Select correct Instrument", vbApplicationModal
                        cmbInstrumentType.SetFocus
                        Exit Sub
                   End If
                Else
                   If cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex) = 1 Then
                        MsgBox "Please Select correct Instrument", vbApplicationModal
                        cmbInstrumentType.SetFocus
                        Exit Sub
                   End If
                End If
                
            End If
        Else
            tnyExtModuleID = 99
            intKeyID = Null
        End If
        
        ' ---------------------------------------------------------------------------------- '
        ' NOTE :- PREVIOUS YEAR TRANSACTION - DEMAND BY REQUEST                              '
        ' ---------------------------------------------------------------------------------- '
        If mPreviousYearMode = 1 Then
            mFinYearID = gbFinancialYearID - 1
            dtTransactionDate = DdMmmYy(txtTransactionDate.Text)
        Else
            mFinYearID = gbFinancialYearID
            If Not IsDate(dtTransactionDate) Then
                dtTransactionDate = DdMmmYy(gbTransactionDate) 'gbTransactionDate
            End If
        End If
        
        
        
        ' ---------------------------------------------------------------------------------- '
        
        
        
        
        tnyDemandType = 10
        dtDemandDate = IIf(IsDate(txtDemandDate), txtDemandDate, dtTransactionDate) ' gbTransactionDate) ' Changed on 06-May-2013
        numSubLedgerID = Null
        
        If val(cmbOutDoorStaff.Tag) > 0 Then
            intKeyID2 = val(cmbOutDoorStaff.Tag)
        End If
        vchRemarks = Trim(txtRemarks)
        numSeatID = gbSeatID
        numUserID = gbUserID
        
        'If Not (cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeBFundSSSFund Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeMoneyOrderReturns) Then
        
        ' The following 2 Transaction Types BFund and Money Order Returns needs approval from higher authority, if the transaction is not in Current Date
        If Not (txtTransactionType.Tag = gbTransactionTypeBFundSSSFund Or txtTransactionType.Tag = gbTransactionTypeMoneyOrderReturns) Then
            tnyStatus = 0
        Else
            If gbTransactionDate <> mServerDate Then
                If Not (cmbSeat.ListIndex > 0) Then
                    MsgBox "Please select the Forwarded Seat", vbInformation
                    cmbSeat.SetFocus
                    Exit Sub
                End If
                If gbSeatGroupID <> gbSeatGroupAccountsOfficer And gbSeatGroupID <> gbSeatGroupAccountsSuperintended Then
                    tnyStatus = 8
                Else
                    If mDemandNo <> "" Then
                        numSeatID = val(txtDemandDate.Tag) 'Demand Generated SeatID
                        numUserID = val(txtReference.Tag)  'Demand Generated UserID
                    End If
                    tnyStatus = 0
                End If
            Else
                tnyStatus = 0
            End If
        End If
        
        intVoucherID = Null
        dtVoucherDate = Null
        tnyArrearFlag = Null
        
        
        '************MODIFIED BY Sabeen**************************************
        ''''        If mPreviousYearMode = 1 Then
        ''''            dtExpiryDate = dtTransactionDate 'gbTransactionDate 'Changed on 06-May-2013
        ''''        Else
        ''''            dtExpiryDate = gbTransactionDate 'Changed on 06-May-2013
        ''''        End If

        mMonthID = Month(dtTransactionDate)
        mMonthName = MonthName(mMonthID)
''''        Select Case mMonthID
''''            Case 1, 3, 5, 7, 8, 10, 12:
''''                dtExpiryDate = " 31 / " + CStr(mMonthName) + " / " + CStr(gbFinancialYearID)
''''            Case 2:
''''                dtExpiryDate = "28/" + CStr(mMonthName) + " / " + CStr(gbFinancialYearID)
''''            Case 4, 6, 9, 11:
''''                dtExpiryDate = "30/" + CStr(mMonthName) + " / " + CStr(gbFinancialYearID)
''''        End Select
        
        If mMonthID < 4 Then
            dtExpiryDate = "1/" + CStr(MonthName(mMonthID + 1)) + " / " + CStr(gbFinancialYearID + 1)
            dtExpiryDate = DdMmmYy(DateAdd("d", -1, dtExpiryDate))
        Else
            If mMonthID = 12 Then
                dtExpiryDate = "1/" + CStr("january") + " / " + CStr(gbFinancialYearID + 1)
                dtExpiryDate = DdMmmYy(DateAdd("d", -1, dtExpiryDate))
      
            Else
                dtExpiryDate = "1/" + CStr(MonthName(mMonthID + 1)) + " / " + CStr(gbFinancialYearID)
                dtExpiryDate = DdMmmYy(DateAdd("d", -1, dtExpiryDate))
            End If
        End If
        
'        Select Case mMonthID
'            Case 1, 3:
'                dtExpiryDate = " 31 / " + CStr(mMonthName) + " / " + CStr(gbFinancialYearID + 1)
'            Case 5, 7, 8, 10, 12:
'                dtExpiryDate = " 31 / " + CStr(mMonthName) + " / " + CStr(gbFinancialYearID)
'            Case 2:
'                dtExpiryDate = "28/" + CStr(mMonthName) + " / " + CStr(gbFinancialYearID + 1)
'            Case 4, 6, 9, 11:
'                dtExpiryDate = "30/" + CStr(mMonthName) + " / " + CStr(gbFinancialYearID)
'        End Select
        
        
        
        
      '*******************************************************************
        
        'intSectionID = gbSectionID
        dtDueDate = dtTransactionDate 'gbTransactionDate 'Changed on 06-May-2013
        
        vchAdminNote = Trim(txtAdminNote)
        vchDemandNo = Null
'        If (cmbOutDoorStaff.ListIndex > 0) Then
'            numZoneID = cmbOutDoorStaff.ItemData(cmbOutDoorStaff.ListIndex)
'        Else
'            numZoneID = Null
'        End If
            
        If cmbMode.ItemData(cmbMode.ListIndex) = 2 Then  'To Fetch Zone
            If cmbOutDoorStaff.ListIndex < 0 Then
                 MsgBox "Please Select Zonal Office", vbInformation
                 Exit Sub
            Else
                numZoneID = cmbOutDoorStaff.ItemData(cmbOutDoorStaff.ListIndex)
            End If
        ElseIf cmbMode.ItemData(cmbMode.ListIndex) = 3 Then   'To Fetch OutDoorStaff
            If cmbOutDoorStaff.ListIndex < 0 Then
                 MsgBox "Please Select OutDoorStaff", vbInformation
                 Exit Sub
            Else
                intKeyID2 = cmbOutDoorStaff.ItemData(cmbOutDoorStaff.ListIndex)
            End If
        End If
        
        intWardNo = IIf(val(txtwardno) > 0, val(txtwardno), Null)
        intDoorNo = IIf(val(txtDoorNo1) > 0, val(txtDoorNo1), Null)
        vchDoorNo2 = IIf(Len(Trim(txtDoorNo2)), Trim(txtDoorNo2), Null)
        
        intInstrumentTypeID = cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
        If intInstrumentTypeID <> 1 Then
            If Trim(txtInstrumentNo.Text) = "" Then
                MsgBox "Enter the Instrument No !", vbInformation
                txtInstrumentNo.SetFocus
                Exit Sub
            End If
            If Not IsDate(txtInstrumentDate) Then
                MsgBox "Please enter the Instrument Date!", vbInformation
                txtInstrumentDate.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtDrawnFrom)) = 0 Then
                MsgBox "Please enter the Bank Or Treasury !", vbInformation
                txtDrawnFrom.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtDrawnPlace)) = 0 Then
                MsgBox "Please enter the Bank/Treasury's Place!", vbInformation
                txtDrawnPlace.SetFocus
                Exit Sub
            End If
        End If
        
        vchInstrumentNo = Trim(txtInstrumentNo.Text)
        If IsDate(txtInstrumentDate) Then
            dtInstrumentDate = txtInstrumentDate
        End If
        vchDrawnFrom = Trim(txtDrawnFrom)
        vchDrawnPlace = Trim(txtDrawnPlace)
        If txtDemandNo.Tag <> "" Then
            numDemandID = Trim(txtDemandNo.Tag)
        Else
            numDemandID = ""
        End If
        '-------- Transaction date---------------                   Added on 05/07/2011
'        If IsNull(dtTransactionDate.Text) Then
'            dtTransactionDate = Null
'        Else
'            dtTransactionDate = dttranDate.value
'        End If
'
        arrInput = Array(intLBID, tnyExtAppID, tnyExtModuleID, _
        tnyDemandType, intTransactionTypeID, _
        intYearID, tnyPeriodID, dtDemandDate, _
        numSubLedgerID, _
        intKeyID, _
        intKeyID2, _
        vchRemarks, _
        tnyStatus, _
        intVoucherID, _
        dtVoucherDate, _
        tnyArrearFlag, _
        dtExpiryDate, _
        IIf(numDemandID = "", Null, numDemandID), _
        mFinYearID, _
        numSeatID, _
        intSectionID, _
        numUserID, _
        gbCounterID, _
        vchAdminNote, _
        vchDemandNo, _
        numZoneID, _
        intWardNo, _
        intDoorNo, _
        vchDoorNo2, _
        numForwardedSeatID, dtDueDate, intInstrumentTypeID, vchInstrumentNo, dtInstrumentDate, vchDrawnFrom, vchDrawnPlace, Null, gbLocationID, txtFunctionary.Tag, txtFunction.Tag, txtSourceofFund.Tag, dtTransactionDate, intDemandMode)
        
        If Not objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            MsgBox "Didn't able to establish a connection with Database server!", vbInformation
            Exit Sub
        End If
        
        'mCnn.BeginTrans
        'On Error GoTo ErrRollBack
        objdb.ExecuteSP "spSaveIDemandTBL", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        
        If IsArray(arrOutPut) Then
            numDemandID = arrOutPut(0, 0)
            vchDemandNo = arrOutPut(1, 0)
            txtDemandNo = vchDemandNo
        Else
            txtDemandNo = ""
            MsgBox "Didn't able to Generate Demand ID!", vbInformation
            GoTo ErrRollBack:
        End If
        
        
        '---------------------------------------------------------
        
        
        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandChild        **'
        '---------------------------------------------------------'
        '@numDemandID_1      [numeric]   ,
        '@intLBID_2      [int]       ,
        '@tnySlNo_3      [tinyint]   ,
        '@intAccountHeadID_4     [int]       ,
        '@vchAccountHeadCode_5   [varchar](20)   ,
        '@fltAmount_6        [float]     ,
        '@vchRemarks_7       [varchar](100)  ,
        '@tnyStatus_8        [tinyint]   = 0 ,
        '@dtOnDate_9     [SmallDateTime] = Null
        '
        '@intYearID      [Integer] = Null,
        '@tnyPeriodID        [TinyInt] = Null,
        '@tnyArrearFlag      [TinyInt] = Null
        '---------------------------------------------------------'
        
        Dim objAc As New clsAccounts
        Dim mSql As String 'added by sunil
        Dim intTransacriontypeID As Variant  ' added by sunil
        mCnn.Execute "Delete From faIDemandChild Where numDemandID=" & numDemandID
        
        For mLoop = 1 To vsGrid.Rows - 1
            fltAmount = 0
            objAc.SetAccountCode (vsGrid.TextMatrix(mLoop, 0))
            If objAc.AccountHeadID > -1 Then
                intAccountHeadID = objAc.AccountHeadID
                vchAccountHeadCode = objAc.AccountCode
                If val(vsGrid.TextMatrix(mLoop, 4)) > 0 Then
                    tnyArrearFlag = 1
                    fltAmount = val(vsGrid.TextMatrix(mLoop, 4))
                End If
                If val(vsGrid.TextMatrix(mLoop, 5)) Then
                    tnyArrearFlag = 0
                    fltAmount = val(vsGrid.TextMatrix(mLoop, 5))
                End If
                If fltAmount <= 0 Then
                    'Amount is Zero
                    GoTo Skip
                End If
                If val(vsGrid.TextMatrix(mLoop, 7)) > 0 Then
                    intYearID = val(vsGrid.TextMatrix(mLoop, 7))
                Else
                    intYearID = Null
                End If
                tnyPeriodID = val(vsGrid.TextMatrix(mLoop, 8))
                '-------Added by Sunil----------
'                If mZonalCollection = 1 Then
'                    tnyStatus = 9
'                Else
                '======To identify Advance amount added by sunil=======
                If val(vsGrid.TextMatrix(mLoop, 14)) = 1 Then
                    tnyStatus = 10
                Else
                     tnyStatus = 0
                End If
               ' End If
               If mZonalCollection = 1 Then
                  intTransacriontypeID = val(vsGrid.TextMatrix(mLoop, 13)) ' Added on 5-08-2011 TransactionTypeId For Integrated Zonal Collection
                  '--------------------------------
                '  tnyStatus = 0
                  dtOnDate = dtTransactionDate ' gbTransactionDate ' Changed on 06-May-2013
                  vchRemarks = ""
                  
                  arrInput = Array(numDemandID, _
                  intLBID, _
                  mLoop, _
                  intAccountHeadID, _
                  vchAccountHeadCode, _
                  fltAmount, _
                  vchRemarks, _
                  tnyStatus, _
                  dtOnDate, _
                  intYearID, _
                  tnyPeriodID, _
                  tnyArrearFlag, _
                  intTransacriontypeID _
                  )
            Else
                arrInput = Array(numDemandID, _
                  intLBID, _
                  mLoop, _
                  intAccountHeadID, _
                  vchAccountHeadCode, _
                  fltAmount, _
                  vchRemarks, _
                  tnyStatus, _
                  dtOnDate, _
                  intYearID, _
                  tnyPeriodID, _
                  tnyArrearFlag _
                  )
            End If
                objdb.ExecuteSP "spSaveIDemandChild", arrInput, , , mCnn, adCmdStoredProc
            End If
Skip:
        Next
        
        '---------------------------------------------------------'
        '**         Input Variable to spSaveIDemandChild        **'
        '---------------------------------------------------------'
        
        mCnn.Execute "Delete From faIDemandAddress Where numDemandID=" & numDemandID
        
        vchName_6 = Trim(txtName)
        vchInit1_7 = Trim(txtInitial1)
        vchInit2_8 = Trim(txtInitial2)
        vchInit3_9 = Trim(txtInitial3)
        vchInit4_10 = Trim(txtInitial4)
        vchHouseName_11 = Trim(txtHouseName)
        vchStreet_12 = Trim(txtStreet)
        vchLocalPlace_13 = Trim(txtLocalPlace)
        vchMainPlace_14 = Trim(txtMainPlace)
        vchPost_15 = Trim(txtPost)
        vchPin_16 = Trim(txtPin)
        vchPhone_17 = Trim(txtPhone)
        
        arrInput = Array(numDemandID, _
        gbLocalBodyID, _
        numZoneID, _
        intWardNo, _
        intDoorNo, _
        vchDoorNo2, _
        vchName_6, _
        vchInit1_7, vchInit2_8, vchInit3_9, vchInit4_10, _
        vchHouseName_11, _
        vchStreet_12, _
        vchLocalPlace_13, _
        vchMainPlace_14, _
        vchPost_15, _
        vchPin_16, _
        vchPhone_17)
        
        objdb.ExecuteSP "spSaveIDemandAddress", arrInput, , , mCnn, adCmdStoredProc
        On Error GoTo 0
        mNewFlag = False
        cmdSave.Enabled = False
                
        If intTransactionTypeID = gbTransactionTypeProfTaxTradeAccrual Then
            Call AccrualJournalByDemandID(numDemandID)
        End If
        '----------For Reverse------------------------------------
        If mReverse = 1 Then
            If numDemandID <> "" Then
                frmReverseRequest.mDemandNo = numDemandID
                frmReverseRequest.mRevDemand = True
            Else
                frmReverseRequest.mDemandNo = ""
                frmReverseRequest.mRevDemand = True
            End If
        End If
        '---------------------------------------------------------
        ' NOTE:- PREVIOUS YEAR : TRANSACTIONS
        
        If mPreviousYearMode Then
        If mPreviousYearRequestID > 0 Then
            'Dim mSql As String
            mSql = "Update faPendingTaskRequest SET  numDemandID = " & numDemandID & ", tnyStatus = 8 Where intRequestID = " & mPreviousYearRequestID & "  "
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        End If
        End If
        
        
        If chkSkipPrinting.Value = 0 Then
            PrintDemandSlip (numDemandID)
        End If
        Exit Sub
ErrRollBack:
        MsgBox "Unexpected Error! RollBacking!", vbInformation
        mCnn.RollbackTrans
        
    End Sub

'Private Sub cmdSend_Click()
'    Dim objDb As New clsDB
'    Dim mCnn As New ADODB.Connection
'    Dim mCnnSvr As New ADODB.Connection
'    Dim Rec As New ADODB.Recordset
'    Dim RecSvr As New ADODB.Recordset
'    Dim mSql As String
'    Dim mDemandID As Double
'
'    'Note:-Inserting into Demand Table
'    mSql = "Select * From faIDemandTbl Where intTransactionTypeID = " & gbTransactionTypeZonalCollection & " AND dtDemandDate = '" & txtDemandDate & "'"
'    objDb.SetConnection mCnn
'    Rec.CursorLocation = adUseClient
'    Rec.Open mSql, mCnn, adOpenForwardOnly, adLockBatchOptimistic, adCmdText
'    If Not (Rec.BOF And Rec.EOF) Then
'        objDb.CreateNewConnection mCnnSvr, SaankhyaHO
'        mCnnSvr.BeginTrans
'        On Error GoTo ErrRollBack:
'        RecSvr.CursorLocation = adUseServer
'        RecSvr.Open "faIDemandTbl", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
'        RecSvr.AddNew
'        mDemandID = Rec!numDemandID
'
'        RecSvr!numDemandID = Rec!numDemandID
'        RecSvr!intLBID = Rec!intLBID
'        RecSvr!tnyExtAppID = Rec!tnyExtAppID
'        RecSvr!tnyExtModuleID = Rec!tnyExtModuleID
'        RecSvr!tnyDemandType = Rec!tnyDemandType
'        RecSvr!intTransactionTypeID = Rec!intTransactionTypeID
'        RecSvr!intYearID = Rec!intYearID
'        RecSvr!tnyPeriodID = Rec!tnyPeriodID
'        RecSvr!dtDemandDate = Rec!dtDemandDate
'        RecSvr!numSubLedgerID = Rec!numSubLedgerID
'        RecSvr!intKeyID = Rec!intKeyID
'        RecSvr!intKeyID2 = Rec!intKeyID2
'        RecSvr!vchRemarks = Rec!vchRemarks
'        RecSvr!tnyStatus = 0
'        RecSvr!tnyArrearFlag = Rec!tnyArrearFlag
'        'RecSvr!intVoucherID = Rec!intVoucherID
'        'RecSvr!dtVoucherDate = Rec!dtVoucherDate
'        RecSvr!dtExpiryDate = Rec!dtExpiryDate
'        RecSvr!intFinancialYearID = Rec!intFinancialYearID
'        RecSvr!numSeatID = Rec!numSeatID
'        RecSvr!intSectionID = Rec!intSectionID
'        RecSvr!numUserID = Rec!numUserID
'        RecSvr!numCounterID = Rec!numCounterID
'        RecSvr!vchAdminNote = Rec!vchAdminNote
'        RecSvr!vchDemandNo = Rec!vchDemandNo
'        RecSvr!numZoneID = Rec!numZoneID
'        RecSvr!intWardNo = Rec!intWardNo
'        RecSvr!intDoorNo = Rec!intDoorNo
'        RecSvr!vchDoorNo2 = Rec!vchDoorNo2
'        RecSvr!numForwardedSeatID = Rec!numForwardedSeatID
'        RecSvr!intInstrumentTypeID = Rec!intInstrumentTypeID
'        RecSvr!vchInstrumentNo = Rec!vchInstrumentNo
'        RecSvr!dtInstrumentDate = Rec!dtInstrumentDate
'        RecSvr!vchDrawnFrom = Rec!vchDrawnFrom
'        RecSvr!vchDrawnPlace = Rec!vchDrawnPlace
'        RecSvr!dtDueDate = Rec!dtDueDate
'        RecSvr!tnyAccrualType = Rec!tnyAccrualType
'        RecSvr!numLocationID = Rec!numLocationID
'        RecSvr.Update
'
'        Rec.Close
'        RecSvr.Close
'
'        'Note:-Inserting into DemandChild Table
'        mSql = "Select * From faIDemandChild Where numDemandID = " & mDemandID
'        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
'        If Not (Rec.BOF And Rec.EOF) Then
'            RecSvr.CursorLocation = adUseServer
'            RecSvr.Open "faIDemandChild", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
'            While Not Rec.EOF
'                RecSvr.AddNew
'                RecSvr!numDemandID = Rec!numDemandID
'                RecSvr!intLBID = Rec!intLBID
'                RecSvr!tnySlNo = Rec!tnySlNo
'                RecSvr!intAccountHeadID = Rec!intAccountHeadID
'                RecSvr!vchAccountHeadCode = Rec!vchAccountHeadCode
'                RecSvr!fltAmount = Rec!fltAmount
'                RecSvr!intYearID = Rec!intYearID
'                RecSvr!tnyPeriodID = Rec!tnyPeriodID
'                RecSvr!tnyArrearFlag = Rec!tnyArrearFlag
'                RecSvr!vchRemarks = Rec!vchRemarks
'                RecSvr!tnyStatus = Rec!tnyStatus
'                RecSvr!dtOnDate = Rec!dtOnDate
'                RecSvr!snyRate = Rec!snyRate
'                'RecSvr!intVoucherID = Rec!intVoucherID
'                'RecSvr!dtVoucherDate = Rec!dtVoucherDate
'                RecSvr.Update
'                Rec.MoveNext
'            Wend
'        End If
'        Rec.Close
'        RecSvr.Close
'
'        'Note:- Inserting into DemandAddress Table
'        mSql = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
'        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
'        If Not (Rec.BOF And Rec.EOF) Then
'            RecSvr.CursorLocation = adUseServer
'            RecSvr.Open "faIDemandAddress", mCnnSvr, adOpenDynamic, adLockOptimistic, adCmdTable
'            RecSvr.AddNew
'
'            RecSvr!numDemandID = Rec!numDemandID
'            RecSvr!numZoneID = Rec!numZoneID
'            RecSvr!intWardNo = Rec!intWardNo
'            RecSvr!intDoorNo = Rec!intDoorNo
'            RecSvr!vchDoorNo2 = Rec!vchDoorNo2
'            RecSvr!vchName = Rec!vchName
'            RecSvr!vchInit1 = Rec!vchInit1
'            RecSvr!vchInit2 = Rec!vchInit2
'            RecSvr!vchInit3 = Rec!vchInit3
'            RecSvr!vchInit4 = Rec!vchInit4
'            RecSvr!vchHouseName = Rec!vchHouseName
'            RecSvr!vchStreet = Rec!vchStreet
'            RecSvr!vchLocalPlace = Rec!vchLocalPlace
'            RecSvr!vchMainPlace = Rec!vchMainPlace
'            RecSvr!vchPost = Rec!vchPost
'            RecSvr!vchPin = Rec!vchPin
'            RecSvr!vchPhone = Rec!vchPhone
'
'            RecSvr.Update
'            RecSvr.Close
'            Rec.Close
'        End If
'        mCnnSvr.CommitTrans
'        mCnn.Close
'    End If
'    Exit Sub
'
'ErrRollBack:
'    mCnnSvr.RollbackTrans
'    mCnnSvr.Close
'
'End Sub

'Private Sub Command1_Click()
'    Call FetchCollectionDetails("6-Aug-09")
'End Sub

Private Sub Command2_Click()
    
End Sub

    Private Sub cmdSearchTransactionType_Click()
        On Error GoTo err:
         
        Dim mSectionID As Long
        Dim mSql        As String
        Dim mModID As Integer
        mModID = cmbMode.ItemData(cmbMode.ListIndex)
        If frmDemandInterface.cmbSections.ListIndex > 0 Then
            mSectionID = frmDemandInterface.cmbSections.ItemData(frmDemandInterface.cmbSections.ListIndex)
'            If mSectionID = 99 Then
'                mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 Order By vchTransactionType"
'            Else
'                mSql = "Select vchTransactionType, intTransactionTypeID From faTransactionType Where intGroupID =10 And intSectionID = " & mSectionID & " Order By vchTransactionType"
'            End If
        '**********************************************'
        '   Modified By Poornima On 07-June-2010
        '**********************************************'
                   
            If mSectionID = 99 Then
                mSql = "SELECT DISTINCT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
                mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSql = mSql + " Where (faTransactionType.intGroupID = 10) And faSectionWiseTransactionTypes.tnyList = 1"
                mSql = mSql + " And isNull(tnyHidden,0)=0 and faTransactionType.inttransactionTypeID NOT IN (9996,9997,9998)"
                mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
            Else
                mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
                mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSql = mSql + " Where (faTransactionType.intGroupID = 10)And faSectionWiseTransactionTypes.tnyList =1 And  faSectionWiseTransactionTypes.intSectionID =  " & mSectionID
                mSql = mSql + " And isNull(tnyHidden,0)=0 and faTransactionType.inttransactionTypeID NOT IN (9996,9997,9998)"
                mSql = mSql + " And faSectionWiseTransactionTypes.intSectionID = " & mSectionID
                mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
            End If
            If mSql = "" Then
                If mReverse = 1 Then
                    mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                    mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
                    mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                    mSql = mSql + " Where (faTransactionType.intGroupID = 10) and faTransactionType.inttransactionTypeID NOT IN (9996,9997,9998)"
                   '' mSQL = mSQL + " And isNull(tnyHidden,0)=0"
                    mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
                End If
            End If
                    'Call PopulateList(cmbTransactionType, mSql, , True, , True)
        '        Else
        '            cmbTransactionType.Clear
        '            cmbTransactionType.AddItem ""
                'End If
        Else
        
                mSql = "SELECT DISTINCT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
                mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
                mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSql = mSql + " Where (faTransactionType.intGroupID = 10) And faSectionWiseTransactionTypes.tnyList = 1"
                mSql = mSql + " And isNull(tnyHidden,0)=0 and faTransactionType.inttransactionTypeID NOT IN (9996,9997,9998)"
                mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
        End If
         
        frmSearchTransactionType.ModeOfTransaction = 1
        frmSearchTransactionType.StrQuery = mSql
        gbSearchStr = ""
        gbSearchID = -1
        frmSearchTransactionType.Show vbModal
        If Not gbSearchStr = "" Then
            txtTransactionType.Text = gbSearchStr
            txtTransactionType.Tag = gbSearchID
            txtTransactionType.SetFocus
        Else
            txtTransactionType.Text = ""
            txtTransactionType.Tag = ""
            txtTransactionType.SetFocus
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSourceOfFund_Click()
        On Error GoTo err
        gbSearchStr = ""
        gbSearchID = -1
        frmSearchMasters.SQLQry = "Select * From suSourceOfFund"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If Not gbSearchStr = "" Then
            txtSourceofFund.Text = gbSearchStr
            txtSourceofFund.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    
    
    Private Sub Form_Activate()
       ' Me.Left = 0
        'Me.Top = 0
        'Call FillOutDoorStaffs
        
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        'If cmbTransactionType.ListIndex > 0 Then
        If val(txtTransactionType.Tag) > 0 Then
            'If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypePTax Then
            If txtTransactionType.Tag = gbTransactionTypePTax Then
                If KeyCode = vbKeyF8 Then
                    frmPropertyTaxCalculator.DemandMode = True
                    frmPropertyTaxCalculator.cmdCopyToReceipt.Caption = "Copy to Demand"
                    frmPropertyTaxCalculator.Show vbModal
                    frmPropertyTaxCalculator.DemandMode = False
                End If
            End If
        End If
    End Sub
    Private Sub FillMode()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
       
 
        'cmbMode.AddItem " "
        'cmbMode.ItemData(cmbMode.NewIndex) = 0
'        cmbMode.AddItem "Direct"
'        cmbMode.ItemData(cmbMode.NewIndex) = 1
'        cmbMode.AddItem "Zonal Office Collection"
'        cmbMode.ItemData(cmbMode.NewIndex) = 2
'        cmbMode.AddItem "Out Door Collection"
'        cmbMode.ItemData(cmbMode.NewIndex) = 3
'        cmbMode.AddItem "Friends JanasevanaKendram Collections"
'        cmbMode.ItemData(cmbMode.NewIndex) = 4


        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select vchDemandMode,intDemandModeID From faDemandMode"
        Call PopulateList(cmbMode, mSql, , True, True, True, enuSourceString.Saankhya)
        
    End Sub

    Private Sub GetPreviousYearsTransaction()
         Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        
        Dim mTaskID     As Integer
        Dim objTr As New clsTransactionType
        Dim mTrTypeID As Integer
        
        Dim objInst As New clsInstruments
        
        'On Error GoTo Err
        If mPreviousYearRequestID > 0 Then
            If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSql = "SELECT * FROM faPendingTaskRequest WHERE intRequestID = " & mPreviousYearRequestID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    
                    'MODE OF COLLECTION
                    cmbMode.Text = "Direct"
                    
                    'INSTRUMENT TYPE
                    objInst.SetInstrumentType (Rec!intInstrumentTypeID)
                    If objInst.InstrumentTypeID > 0 Then
                        cmbInstrumentType.Text = objInst.InstrumentType
                    End If
                    
                    'DEMAND DATE
                    txtDemandDate.Text = DdMmmYy(Rec!dtTransactionDate)
                    
                    'TRANSACTION TYPE
                    mTrTypeID = IIf(IsNull(Rec!intTransactionTypeID), -1, Rec!intTransactionTypeID)
                    objTr.SetTransactionType (mTrTypeID)
                    If objTr.TransactionTypeID > 0 Then
                        txtTransactionType.Text = objTr.TransactionType
                        txtTransactionType.Tag = objTr.TransactionTypeID
                    End If
                    
                    
                End If
                Rec.Close
            End If
        End If
        Exit Sub

err:

    MsgBox err
    
    End Sub
    
    Private Sub Form_Load()
        vsGrid.ColComboList(0) = "|..."
        Call FillMode
        Call FillInstruments
        
        Call FillGridYear
        Call FillZone
        Call FillSections
        Call FillSeats
        Call FormInitialize
        
        If mZonalCollection = 1 Then
            cmbMode.ListIndex = 2
            cmbMode.Enabled = False
            Call FetchCollectionDetails(CDate(frmListOfZonalDailyCollection.vsGrid.TextMatrix(frmListOfZonalDailyCollection.vsGrid.Row, 0)))
        End If
        
        If mPreviousYearMode = 1 Then
            
        End If
        
        If mDemandNo <> "" Then
            Call DispalyDemand(CStr(mDemandNo))
        End If
        
        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
            cmbSections.Text = "Panchayat Office"
        End If
        
    End Sub
    

    Private Sub Form_Unload(Cancel As Integer)
        '  Saving In Registory  '
        SaveSetting "iSaankhyaMasters", "DemandGenerator", "Section", CStr(cmbSections.Text)
        SaveSetting "iSaankhyaMasters", "DemandGenerator", "Zone", CStr(cmbZone.Text)
        SaveSetting "iSaankhyaMasters", "DemandGenerator", "SkipPrint", CStr(chkSkipPrinting.Value)
        mPreviousYearMode = 0
        Call FinancialYearSetForPEndingTask
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

    Private Sub txtAdminNote_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    
    Private Sub txtDemandNo_LostFocus()
        If txtDemandNo.Text <> "" Then
            If gbFetchDemandFromHO = 1 Then
                Call FetchDemand(txtDemandNo.Text)
            Else
                Call DispalyDemand(txtDemandNo.Text)
            End If
        End If
    End Sub

    Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtDoorNo1_LostFocus()
        If val(txtDoorNo1) > 0 Then
            txtDoorNo1.Text = Format(val(txtDoorNo1), "#0")
        Else
            txtDoorNo1.Text = ""
        End If
    End Sub
    Private Sub txtDoorNo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtDrawnFrom_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtDrawnFrom_LostFocus()
        If txtDrawnFrom.Text <> "" Then
            txtDrawnFrom.Text = FormatIntoProperCase(txtDrawnFrom.Text)
        End If
    End Sub

    Private Sub txtDrawnPlace_LostFocus()
        If txtDrawnPlace.Text <> "" Then
            txtDrawnPlace.Text = FormatIntoProperCase(txtDrawnPlace.Text)
        End If
    End Sub

    Private Sub txtHouseName_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtHouseName_LostFocus()
        If txtHouseName.Text <> "" Then
            txtHouseName.Text = FormatIntoProperCase(txtHouseName.Text)
        End If
    End Sub

    Private Sub txtInitial1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtInitial1_LostFocus()
        txtInitial1.Text = UCase(txtInitial1)
    End Sub
    Private Sub txtInitial2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtInitial2_LostFocus()
        txtInitial2.Text = UCase(txtInitial2)
    End Sub
    Private Sub txtInitial3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtInitial3_LostFocus()
        txtInitial3.Text = UCase(txtInitial3)
    End Sub
    Private Sub txtInitial4_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtInitial4_LostFocus()
        txtInitial4.Text = UCase(txtInitial4)
    End Sub
    
    Private Sub txtInstrumentDate_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtInstrumentDate_LostFocus()
        txtInstrumentDate = Trim(txtInstrumentDate)
        If Len(txtInstrumentDate) Then
            txtInstrumentDate.Text = CheckDateInMMM(txtInstrumentDate.Text)
            '------------------------------------------'
            '   Added To Validate the Cheque Date      '
            '------------------------------------------'
            Dim mInstrumentTypeID As Integer
            Dim mDt As Date
            mDt = txtInstrumentDate
            If cmbInstrumentType.ListIndex > -1 Then
                mInstrumentTypeID = cmbInstrumentType.ItemData(cmbInstrumentType.ListIndex)
            Else
                mInstrumentTypeID = 0
            End If
            
            If mInstrumentTypeID = 5 Or mInstrumentTypeID = 4 Then
                If mDt > gbTransactionDate Then
                    MsgBox "Post dated cheques will not accepted!", vbInformation
                    txtInstrumentDate.Text = DdMmmYy(gbTransactionDate)
                    txtInstrumentDate.SetFocus
                End If
                If mPreviousYearMode <> 1 Then
                    If mDt < DateAdd("d", -180, gbTransactionDate) Then
                        MsgBox "Upto Six Months Validity Cheques Can Only Accept", vbInformation
                        txtInstrumentDate.Text = DdMmmYy(gbTransactionDate)
                        txtInstrumentDate.SetFocus
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub txtInstrumentNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtLocalPlace_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtLocalPlace_LostFocus()
        If txtLocalPlace.Text <> "" Then
            txtLocalPlace.Text = FormatIntoProperCase(txtLocalPlace.Text)
        End If
    End Sub

    Private Sub txtMainPlace_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtMainPlace_LostFocus()
        If txtMainPlace.Text <> "" Then
            txtMainPlace.Text = FormatIntoProperCase(txtMainPlace.Text)
        End If
    End Sub

    Private Sub txtName_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtName_LostFocus()
        If txtName.Text <> "" Then
            txtName.Text = FormatIntoProperCase(txtName.Text)
        End If
    End Sub

    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = Asc("-")) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 13 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPost_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtPost_LostFocus()
        If txtPost.Text <> "" Then
            txtPost.Text = FormatIntoProperCase(txtPost.Text)
        End If
    End Sub

    Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtRemarks_LostFocus()
        If txtRemarks.Text <> "" Then
            txtRemarks.Text = FormatIntoProperCase(txtRemarks.Text)
        End If
    End Sub

    Private Sub txtStreet_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub

    Private Sub txtStreet_LostFocus()
        If txtStreet.Text <> "" Then
            txtStreet.Text = FormatIntoProperCase(txtStreet.Text)
        End If
    End Sub

Private Sub txtTransactionDate_LostFocus()
  txtTransactionDate.Text = CheckDateInMMM(txtTransactionDate.Text)
End Sub

'Private Sub txtTransactionDate_Validate(Cancel As Boolean)
'  If Me.txtTransactionDate.Text = "" Then
'        Exit Sub
'    End If
'     If IsNumeric(Me.txtTransactionDate.Text) Then
'        Me.txtTransactionDate.Text = Format(Me.txtTransactionDate.Text, "##/##/##")
'    End If
'    'Checks to make sure it's a valid set of numerics
'    If InStr(Me.txtTransactionDate.Text, "/") Then
'            Me.txtTransactionDate.Text = Format(Me.txtTransactionDate.Text, "dd/MM/yyyy")
'            Else: GoTo Error1
'    End If
'        'Keeps users from entering extra digits
'    If Len(Me.txtTransactionDate.Text) <> 8 Then
'        GoTo Error1
'    End If
'
'    'Makes sure the Month is not greater then 12
'    If Left(Me.txtTransactionDate.Text, 2) > 12 Then
'        GoTo Error1
'    End If
'
'    'Makes sure the Day is not greater then 31
'    If mID(Me.txtTransactionDate.Text, 4, 2) > 31 Then
'        GoTo Error1
'    End If
'
'    'Exits sub if no errors were encountered
'  Exit Sub
'
'    'Keeps user in Text field, highlights text, and displays
'    'a MsgBox telling them the valid date format.
'Error1:
'    Cancel = True
'    txtTransactionDate.SelStart = 0
'    txtTransactionDate.SelLength = Len(txtTransactionDate.Text)
'   ' temp = MsgBox("Please enter date in MMDDYY Format.", vbOKCancel + vbExclamation + vbDefaultButton1 + vbApplicationModal, Error)
'   MsgBox ("Error")
'End Sub

    Private Sub txtTransactionType_GotFocus()
        'On Error GoTo err
        '-----------------For Clearing the Budget Details-----------------------'
        txtFunctionary.Text = ""
        txtFunctionary.Tag = ""
        txtFunction.Text = ""
        txtFunction.Tag = ""
        txtSourceofFund.Text = ""
        txtSourceofFund.Tag = ""
        '------------------------------------------------------------------------'
        'If cmbTransactionType.ListIndex > 0 Then
        If val(txtTransactionType.Tag) > 0 Then
            'If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeOutDoor Then
            If cmbMode.ItemData(cmbMode.ListIndex) = 3 Then
                Call FillOutDoorStaffs
                lblCombo.Caption = "Out Door Collection Staff"
                cmbOutDoorStaff.Enabled = True
                lblCombo.Enabled = True
                lblCombo.Left = 6450
                lblCombo.Top = 480
            'ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeZonalCollection Then
            ElseIf cmbMode.ItemData(cmbMode.ListIndex) = 2 Then
                'cmbTransactionType.Tag = gbTransactionTypeZonalCollection
                'txtTransactionType.Tag = gbTransactionTypeZonalCollection
                Call FillZoneInSub
                lblCombo.Caption = "Zonal Office"
                lblCombo.Left = 7260
                lblCombo.Top = 485
                cmbOutDoorStaff.Enabled = True
                lblCombo.Enabled = True
            'ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxEmp Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxTrade Then
                'If gbLinkWithProfTaxEmp Then
                '    frmSearchProfTaxInstitutions.Show vbModal
                'End If
            'ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeBFundSSSFund Or cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeMoneyOrderReturns Then
            ElseIf txtTransactionType.Tag = gbTransactionTypeBFundSSSFund Or txtTransactionType.Tag = gbTransactionTypeMoneyOrderReturns Then
                If gbLBType = 3 Then
                    Dim mCnn    As New ADODB.Connection
                    Dim objdb   As New clsDB
                    Dim Rec     As New ADODB.Recordset
                    Dim mSql    As String
                    Dim i       As Integer
                    
                    If (objdb.CreateNewConnection(mCnn, enuSourceString.DBMaster)) Then
                        mSql = "Select Right(Convert(Varchar(20),numSeatID),3) As SeatID,chvSeatTitle From GL_Seats Where intGroupID = " & gbSeatGroupAccountsOfficer
                        Rec.Open mSql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                            For i = 0 To cmbSeat.ListCount - 1
                                If (cmbSeat.List(i) = Rec!chvSeatTitle) Then
                                    cmbSeat.ListIndex = i
                                End If
                            Next
                            'cmbSeat.ItemData(cmbSeat.ListIndex) = IIf(IsNull(Rec!SeatID), "", Rec!SeatID)
                        End If
                        Rec.Close
                    Else
                        MsgBox "Connection To Master does not exit, Please contact your System Administrator", vbInformation
                        Exit Sub
                    End If
                End If
            Else
                lblCombo.Caption = "Out Door Collection Staff"
                cmbOutDoorStaff.Enabled = False
                lblCombo.Enabled = False
                lblCombo.Left = 5565
                lblCombo.Top = 480
            End If
            'If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) <> 0 Then
            If txtTransactionType.Tag > 0 Then
            '    Call CheckBudgetDetails(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
                Call CheckBudgetDetails(txtTransactionType.Tag)
            End If
            'End If
        End If
        'End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub txtTransactionType_KeyDown(KeyCode As Integer, Shift As Integer)
'        If KeyCode = vbKeyDelete Then
'            txtTransactionType.Text = ""
'            txtTransactionType.Tag = ""
'        End If
    End Sub

    Private Sub txtTransactionType_LostFocus()
        'If cmbTransactionType.ListIndex > -1 Then
        If val(txtTransactionType.Tag) > 0 Then
         '   If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxEmp Then
            If txtTransactionType.Tag = gbTransactionTypeProfTaxEmp Then
                If gbLinkWithProfTaxEmp Then
                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 2
                    frmSearchProfTaxInstitutions.Show vbModal
                End If
'            ElseIf cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeProfTaxTrade Then
'                If gbLinkWithProfTaxEmp Then
'                    frmSearchProfTaxInstitutions.ProfTaxInstTypeMode = 1
'                    frmSearchProfTaxInstitutions.Show vbModal
'                End If
            End If
        End If
        Dim mCnt    As Integer
        If mReverse = 1 Then
            If vsGrid.TextMatrix(1, 0) <> "" Then
                For mCnt = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mCnt, 0) <> "" Then
                        If FunTrTypeAccHeadValidate(val(txtTransactionType.Tag), vsGrid.TextMatrix(mCnt, 0)) = False Then
                            MsgBox "AccountHead is not Defined in this TransactionType"
                            Exit Sub
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then KeyAscii = 0: Call PressTabKey
    End Sub
    Private Sub txtWardNo_LostFocus()
        If val(txtwardno) > 0 Then
            txtwardno.Text = Format(val(txtwardno), "#0")
        Else
            txtwardno.Text = ""
        End If
    End Sub
    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGrid.Row > 1 Then
            If vsGrid.TextMatrix(vsGrid.Row - 1, 0) = "" Or _
               (val(vsGrid.TextMatrix(vsGrid.Row - 1, 4)) <= 0 And _
               val(vsGrid.TextMatrix(vsGrid.Row - 1, 5)) <= 0) Then
               Cancel = True
               Exit Sub
            End If
        End If
        
        If Col = 4 Or Col = 5 Then
            If Trim(vsGrid.TextMatrix(Row, 0)) = "" Then
                Cancel = True
            End If
        End If
        
        If Len(gbSearchStr) Then
            Dim objAccHead As New clsAccounts
            objAccHead.SetAccountCode (Token(gbSearchStr, " "))
            If objAccHead.AccountHeadID > 0 Then
                vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                vsGrid.TextMatrix(Row, 1) = objAccHead.AccountHead
                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
            End If
            vsGrid.Col = vsGrid.Col + 2
            vsGrid.Redraw = flexRDDirect
            gbSearchStr = ""
            gbSearchID = -1
        End If
        
    End Sub
    Private Sub vsGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
        If OldRow >= vsGrid.Rows - 1 Then
            vsGrid.Rows = vsGrid.Rows + 5
        End If
    End Sub
    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        'cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
        If vsGrid.Row <= 9 Then
            Dim mSql As String
            Dim mIndex As Long
            'If cmbTransactionType.ListIndex > -1 Then
            If val(txtTransactionType.Tag) > 0 Then
                'mIndex = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
                mIndex = txtTransactionType.Tag
            End If
            'End If
            If mIndex > 0 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                mSql = mSql + " Where intTransactionTypeID = " & mIndex & " And tinHiddenFlag = 0 And faAccountHeads.intGroupID is Null Order By faTransactionTypeChild.intOrder"
                frmSearchAccountHeads.SQLString = mSql '"Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            Else
                If gbLBPanchayat = 1 Then
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null and intMinorAccountHeadID<>220 Order By faAccountHeads.vchAccountHeadCode"
            
            Else
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.cmdSearch.Enabled = False
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null and intMinorAccountHeadID<>248 Order By faAccountHeads.vchAccountHeadCode"
            End If
                'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null Order By faAccountHeads.vchAccountHeadCode"
            
            End If
            frmSearchAccountHeads.VoucherMode = 100
            'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
        Else
            MsgBox "Can't print more than 9 rows in this Receipt", vbInformation
        End If
    End Sub
    Private Sub vsGrid_CellChanged(ByVal Row As Long, ByVal Col As Long)
        Dim objAccHead As clsAccounts
        Dim mAmt As Double
        If vsGrid.Row > 0 Then
            If Col = 1 And vsGrid.ComboIndex > -1 Then
                Set objAccHead = New clsAccounts
                If objAccHead.FindAccountByHead(Trim(vsGrid.ComboItem)) Then
                vsGrid.TextMatrix(Row, 0) = objAccHead.AccountCode
                vsGrid.TextMatrix(Row, 6) = objAccHead.AccountHeadID
                End If
            ElseIf Col = 4 Then
                If mRoundOffDecimalPlace Then
                    vsGrid.TextMatrix(Row, 4) = Format(val(vsGrid.TextMatrix(Row, 4)), "#0")
                Else
                    vsGrid.TextMatrix(Row, 4) = Format(val(vsGrid.TextMatrix(Row, 4)), "0.00")
                End If
                mAmt = Format(val(vsGrid.TextMatrix(Row, 4)), "0.00")
                If (mAmt - Int(mAmt)) > 0 Then
                    mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
                End If
                vsGrid.TextMatrix(Row, 4) = Format(mAmt, "0.00")
                
                If val(vsGrid.TextMatrix(Row, 4)) > 0 Then
                    vsGrid.TextMatrix(Row, 5) = ""
                End If
                Call Calculate
            ElseIf Col = 5 Then
                If mRoundOffDecimalPlace Then
                    vsGrid.TextMatrix(Row, 5) = Format(val(vsGrid.TextMatrix(Row, 5)), "#0")
                Else
                    vsGrid.TextMatrix(Row, 5) = Format(val(vsGrid.TextMatrix(Row, 5)), "0.00")
                End If
                
                mAmt = Format(val(vsGrid.TextMatrix(Row, 5)), "0.00")
                If (mAmt - Int(mAmt)) > 0 Then
                    mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
                End If
                vsGrid.TextMatrix(Row, 5) = Format(mAmt, "0.00")
                
                If val(vsGrid.TextMatrix(Row, 5)) > 0 Then
                    vsGrid.TextMatrix(Row, 4) = ""
                End If
                Call Calculate
            ElseIf Col = 6 Then
                If vsGrid.TextMatrix(Row, 0) = gbAcHeadCodeAdvanceDandO Then
                    Dim mLoop As Integer
                    Dim mItem As String
                    mItem = "#0; "
                    For mLoop = gbFinancialYearID + 5 To 1970 Step -1
                        mItem = mItem & "|#" & mLoop & ";" & CStr(mLoop) & "-" & CStr(mLoop + 1)
                    Next
                    vsGrid.ColComboList(2) = mItem
                End If
            End If
            Call ValuesForHiddenColumns
        End If
    End Sub
    Private Sub vsGrid_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
        If Col = 0 Then
            If KeyCode >= Asc("0") And KeyCode <= Asc("9") Or KeyCode = vbKeyBack Then
            ElseIf KeyCode = Asc(vbTab) Or KeyCode = 13 Then
                gbSearchStr = vsGrid.Cell(flexcpText, Row, Col)
                vsGrid.Cell(flexcpText, Row, Col) = ""
                Call vsGrid_BeforeEdit(Row, Col, False)
            Else
                KeyCode = 0
            End If
        End If
    End Sub

    Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'       -------------------------------------------
'        Modified On 18/02/2011 By Anisha
        If vsGrid.Col = 0 Then
            KeyAscii = 0
        End If
        If vsGrid.Col = 4 Or vsGrid.Col = 5 Then
            vsGrid.EditMaxLength = 15
        End If
'      ------------------------------------------
        If vsGrid.Row > 9 Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub vsGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        ''----------------------------------------------------------------------------'
        '' Selection and Deselection of Demands in Grid only permits for 3 demands    '
        '' 3 Demands = 6 Rows in Receipt in the case of Property Tax                                                                                                                                                                                                                                                                                                                                 '
        '' Selection must be periodicity Order                                        '
        ''----------------------------------------------------------------------------'
        'Dim mLoop As Long
        'If Row > 0 Then
        '    If vsGrid.Cell(flexcpChecked, Row, Col) = 2 Then
        '        If mNumberOfSelections < 3 Then
        '            If Row = 1 Or vsGrid.Cell(flexcpChecked, Row - 1, Col) = vbChecked Then
        '                vsGrid.Cell(flexcpChecked, Row, Col) = vbChecked
        '                mNumberOfSelections = mNumberOfSelections + 1 'IIf(Row Mod 2 = 0, 1, 0)
        '            Else
        '                Cancel = True
        '            End If
        '        Else
        '            Cancel = True
        '        End If
        '    Else ' Already  Checked
        '        If vsGrid.Cell(flexcpChecked, Row - 1, Col) = 1 Then
        '        For mLoop = Row To vsGrid.Rows - 1
        '            If vsGrid.TextMatrix(Row, 10) <> vsGrid.TextMatrix(mLoop, 10) Then
        '                If vsGrid.Cell(flexcpChecked, mLoop, 12) = vbChecked Then
        '                    Cancel = True
        '                End If
        '                mNumberOfSelections = mNumberOfSelections - 1
        '                Exit For
        '            End If
        '        Next mLoop
        '        Else
        '            Cancel = True
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub PrintDemandSlip(mDemandID As Variant)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim RecChild As New ADODB.Recordset
        Dim RecAddress As New ADODB.Recordset
        Dim arrInput As Variant
        Dim objTranType As New clsTransactionType
        Dim mSql As String
        Dim mTotalAmt As Double
        '*********************************************************************************************'
        '                  Procedure to Print the Demand Slip                                         '
        '*********************************************************************************************'
        
        arrInput = Array(mDemandID)
        objdb.SetConnection mCnn
        
        mSql = "        Select faIDemandTbl.*, vchSectionName From faIDemandTbl Inner Join "
        mSql = mSql + " faSection On faSection.intSectionID = faIDemandTbl.intSectionID"
        mSql = mSql + " Where numDemandID = " & mDemandID
        
        Rec.Open mSql, mCnn, adOpenStatic, adLockOptimistic
        If Not (Rec.EOF And Rec.BOF) Then
        'Call FileInitialize
        If chkSkipPrinting.Value = 0 Then
            'Call FileInitialize
            Call PrinterInit
                On Error Resume Next
                Print #gbFileNO,
                Print #gbFileNO, Style(gbTitle1, True, True)
                Print #gbFileNO, Style("  Demand Slip", True, True)
                Print #gbFileNO,
                Print #gbFileNO, "Demand No:"; Rec!vchDemandNo
                Print #gbFileNO, "Demand Date : "; DdMmmYy(Rec!dtDemandDate)
                
                objTranType.SetTransactionType Rec!intTransactionTypeID
                Print #gbFileNO, Rec!vchSectionName
                
                If objTranType.TransactionTypeID > 0 Then
                    Print #gbFileNO, objTranType.TransactionType
                Else
                    Print #gbFileNO, "Transaction Type : Unknown < Please Contact System Administrator"
                End If
                
                '----------------------------------------------------------'
                ' iDemandChild recordset is only required here if One have
                ' to Print the head wise details.
                '----------------------------------------------------------'
                mSql = "Select * From faIDemandChild Where numDemandID = " & mDemandID
                RecChild.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                While Not RecChild.EOF
                    mTotalAmt = mTotalAmt + RecChild!fltAmount
                    RecChild.MoveNext
                Wend
                
                Print #gbFileNO, "Amount : " & mTotalAmt
                mSql = "Select * From faIDemandAddress Where numDemandID = " & mDemandID
                RecAddress.Open mSql, mCnn, adOpenStatic, adLockOptimistic
                If Not (RecAddress.BOF And RecAddress.EOF) Then
                    Print #gbFileNO, "Ward    : " & RecAddress!intWardNo; "       ";
                    Print #gbFileNO, "Door No : " & RecAddress!intDoorNo & IIf(Len(RecAddress!vchDoorNo2), "/" & RecAddress!vchDoorNo2, "")
                    Print #gbFileNO, "Name    : " & RecAddress!vchName
                    Print #gbFileNO, "Phone   : " & RecAddress!vchPhone
                End If
                
                Print #gbFileNO,
                Print #gbFileNO, Rec!vchRemarks
                Print #gbFileNO,
                Print #gbFileNO, "Prepared By " & gbUserName & "     Seat Name " & gbSeatName
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                Print #gbFileNO,
                
            Close #gbFileNO
            'Shell "Print " & gbFileName
            'ShellPad
            End If
            Rec.Close
        End If 'chkSkipPrint
        Set mCnn = Nothing
        Set objdb = Nothing
        
    End Sub
    
    Public Sub AccrualJournalByDemandID(mDemandID As Variant)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim RecIDemand As New ADODB.Recordset
        Dim mSql As String
        Dim RecTran As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mTotalAmount As Variant
                
        Dim intTransactionID                As Variant
        Dim intLocalBodyID                  As Variant
        Dim intFinancialYearID              As Variant
        Dim dtTransactionDate               As Variant
        Dim intExternalApplicationID        As Variant
        Dim intExternalApplicationModuleID  As Variant
        Dim intFunctionID                   As Variant
        Dim intFunctionaryID                As Variant
        Dim intFieldID                      As Variant
        Dim intFundID                       As Variant
        Dim intBudgetCentreID               As Variant
        Dim vchNarration                    As Variant
        Dim intTransactionTypeID            As Variant
        Dim intProcessID                    As Variant
        Dim vchGroup                        As Variant
        Dim intGroupID                      As Variant
        Dim intKeyID                        As Variant
        Dim numSubLedgerID                  As Variant
        Dim numUserID                       As Variant
        Dim intVoucherNo                    As Variant
                
        'Dim intTransactionID                As Variant
        Dim intSerialNo                     As Variant
        Dim intAccountHeadID                As Variant
        Dim fltAmount                       As Variant
        Dim tinDebitOrCreditFlag            As Variant
        Dim intByAccountHeadID              As Variant
        'Dim vchNarration                    As Variant
        'Dim intFundID                       As Variant
      
        
        mSql = " Select * From faIDemandTbl Inner Join "
        mSql = mSql + " faIDemandChild ON faIDemandChild.numDemandID = faIDemandTbl.numDemandID Inner Join"
        mSql = mSql + " faTransactionType On faTransactionType.intTransactionTypeID = faIDemandTbl.intTransactionTypeID"
        mSql = mSql + " Where faIDemandTbl.numDemandID = " & mDemandID
        
        objdb.SetConnection mCnn
        RecIDemand.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
        If RecIDemand.BOF And RecIDemand.EOF Then
            MsgBox "There is no Demand Found for Proceed", vbInformation
            Exit Sub
        End If
        intTransactionID = -1
        intLocalBodyID = gbLocalBodyID
        intFinancialYearID = gbFinancialYearID
        dtTransactionDate = RecIDemand!dtDueDate
        intExternalApplicationID = 115
        intExternalApplicationModuleID = 0
        intFunctionID = RecIDemand!intFunctionID
        intFunctionaryID = RecIDemand!intFunctionaryID
        intFieldID = Null
        intFundID = gbFundID
        intBudgetCentreID = Null
        vchNarration = RecIDemand!vchRemarks
        intTransactionTypeID = RecIDemand!intTransactionTypeID
        intProcessID = Null
        vchGroup = "JV"
        intGroupID = 40
        intKeyID = Null
        numSubLedgerID = RecIDemand!numDemandID
        numUserID = gbUserID
        intVoucherNo = Null
        
        arrInput = Array( _
        intTransactionID, _
        intLocalBodyID, _
        intFinancialYearID, _
        dtTransactionDate, _
        intExternalApplicationID, _
        intExternalApplicationModuleID, _
        intFunctionID, _
        intFunctionaryID, _
        intFieldID, _
        intFundID, _
        intBudgetCentreID, _
        vchNarration, _
        intTransactionTypeID, _
        intProcessID, _
        vchGroup, _
        intGroupID, _
        intKeyID, _
        numSubLedgerID, _
        numUserID, _
        intVoucherNo)
        
        objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut
        If IsArray(arrOutPut) Then
            intTransactionID = arrOutPut(0, 0)
        End If
        intSerialNo = 1
        
        mSql = " Select * From faTransactionType INNER JOIN "
        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intTransactionTypeID = faTransactionType.intTransactionTypeID "
        mSql = mSql + " Where faTransactionType.intTransactionTypeID = " & RecIDemand!intTransactionTypeID
        mSql = mSql + " Order By intOrder"
        RecTran.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        If Not (RecTran.BOF And RecTran.EOF) Then
            intByAccountHeadID = RecTran!intAccountHeadID
            While Not RecIDemand.EOF
                RecTran.MoveFirst
                While Not RecTran.EOF
                    Debug.Print RecTran!intOrder
                    If RecTran!intAccountHeadID = RecIDemand!intAccountHeadID Then
                            intSerialNo = intSerialNo + 1
                            intAccountHeadID = RecIDemand!intAccountHeadID
                            fltAmount = RecIDemand!fltAmount
                            tinDebitOrCreditFlag = RecTran!tinDebitOrCredit
                            'intByAccountHeadID
                            vchNarration = Null
                            intFundID = gbFundID
                            mTotalAmount = mTotalAmount + RecIDemand!fltAmount
                            
                            arrInput = Array( _
                            intTransactionID, _
                            intSerialNo, _
                            intAccountHeadID, _
                            fltAmount, _
                            tinDebitOrCreditFlag, _
                            intByAccountHeadID, _
                            vchNarration, _
                            intFundID)
                            
                            objdb.ExecuteSP "spSaveTransactionChild", arrInput
                            GoTo SkipLoop:
                    End If
                    RecTran.MoveNext
                Wend
SkipLoop:
                RecIDemand.MoveNext
            Wend
                
            RecTran.MoveFirst
            intSerialNo = 1
            intAccountHeadID = RecTran!intAccountHeadID
            fltAmount = mTotalAmount
            tinDebitOrCreditFlag = RecTran!tinDebitOrCredit
            intByAccountHeadID = Null
            'vchNarration = RecIDemand!vchNarration
            intFundID = gbFundID
            
            arrInput = Array( _
            intTransactionID, _
            intSerialNo, _
            intAccountHeadID, _
            fltAmount, _
            tinDebitOrCreditFlag, _
            intByAccountHeadID, _
            vchNarration, _
            intFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput
            
        End If 'If Not (RecTran.BOF And RecTran.EOF) Then
        'mSQL = "Update faIDemandTbl Set tnyStatus = 1 Where numDemandID = " & mDemandID
        'mCnn.Execute mSQL
        RecIDemand.Close
        RecTran.Close
    End Sub

    Private Sub FetchCollectionDetails(mDate As Date)
        Dim objdb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSql        As String
        Dim mDt         As Date
        Dim arrInput    As Variant
        Dim i           As Integer
        Dim mTranType   As Integer
        Dim mAdvFlag    As Boolean
        Dim mAdvAmount  As Double
        Dim mCount      As Integer
        Dim objAcc  As New clsAccounts
        Dim mFlag As Integer
        '*********************************************************************************************'
        '                  Procedure to fetch the Collection details from a Zonal Office              '
        '*********************************************************************************************'
        mDt = mDate
        mTranType = 0
        mAdvAmount = 0
        mAdvFlag = False
        mSql = "Select * From faIDemandTbl Where intTransactionTypeID = " & gbTransactionTypeZonalCollection
        mSql = mSql + " And dtDemandDate = '" & DdMmmYy(mDt) & "'"
        mSql = mSql + " And tnyStatus <> 9"
        objdb.SetConnection mCnn
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not (Rec.BOF And Rec.EOF) Then
'            MsgBox "Demand Already Generated on this Date!", vbInformation
'            Exit Sub
            'Added By sunil-
            txtDemandNo.Text = Rec!vchDemandNo
            txtInstrumentNo.Text = Rec!vchInstrumentNo
            txtDemandNo.Locked = True
            frmDemandInterface.cmdSave.Enabled = False
            frmDemandInterface.cmdNew.Enabled = False
            vsGrid.Enabled = False
            mFlag = 1
        End If
        Rec.Close
        
        cmbSections.Text = "Janasevana Kendram"
        txtDemandDate.Text = DdMmmYy(mDt)
'        For i = 0 To cmbTransactionType.ListCount - 1
'            If cmbTransactionType.ItemData(i) = 9997 Then
'                mTranType = 1
'            End If
'        Next
'        If mTranType = 0 Then
'            cmbTransactionType.AddItem ("Zonal Office Collection")
'            cmbTransactionType.ItemData(cmbTransactionType.NewIndex) = 9997
'        End If
        
'        cmbTransactionType.Text = "Zonal Office Collection"
        txtTransactionType.Text = "Zonal Office Collection"
        txtTransactionType.Tag = 9997
        Call txtTransactionType_GotFocus
              
        cmbOutDoorStaff.Text = gbLocation
        cmbOutDoorStaff.Tag = gbLocationID
        
        cmbInstrumentType.Text = "Directly Debited To Bank"
        cmbInstrumentType.Tag = InstrumentType.DirectlyDebited
        
        txtDrawnFrom.Text = gbRemittingBank 'ReadIniFile(gbSaankhyaINI, "Receipt", "RemittingBank")
        txtDrawnPlace.Text = gbRemittingPlaceOfBank 'ReadIniFile(gbSaankhyaINI, "Receipt", "RemittingPlaceOfBank")
        
        txtDemandDate.Enabled = False
        'cmbTransactionType.Enabled = False
        cmdSearchTransactionType.Enabled = False
        cmbOutDoorStaff.Enabled = False
        cmbInstrumentType.Enabled = False
        txtName.Enabled = False
        txtDrawnFrom.Enabled = IIf(Len(Trim(txtDrawnFrom)), False, True)
        txtDrawnPlace.Enabled = IIf(Len(Trim(txtDrawnPlace)), False, True)
        ' vsGrid.Clear
        'vsGrid.Rows = 1
        vsGrid.Clear 1, 1
        vsGrid.Rows = 1
        vsGrid.Rows = 1
    '   vsGrid.Row = 50
        Dim mRow As Variant
      '  mRow = 1
        arrInput = Array(mDt)
        objdb.SetConnection mCnn
        Set Rec = objdb.ExecuteSP("spHeadWiseConsolidationZone", arrInput, , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            While Not Rec.EOF
                If Rec!fltAmount <> 0 Then
                    vsGrid.Rows = vsGrid.Rows + 1
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchAccountHeadCode
                    objAcc.SetAccountCode (Rec!vchAccountHeadCode)
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = Rec!vchAccountHead
                    If InStr(1, Rec!vchAccountHead, "Arrear") Then
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = Format(Rec!fltAmount, "0.00")
                    Else
                        vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = Format(Rec!fltAmount, "0.00")
                    End If
                    
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = objAcc.AccountHeadID
'                    If Rec!vchAccountHeadCode = gbAcHeadCodeAdvancePTax Then
'                        mAdvFlag = True
'                        mAdvAmount = Rec!fltAmount
'                        vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = "0.00"          ' Advance Displayed as 0.00
'                    End If
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 13) = Rec!intTransactionTypeID
                    vsGrid.TextMatrix(vsGrid.Rows - 1, 14) = 0
                      
                End If
                Rec.MoveNext
            Wend
        End If
        Rec.Close
        '=========================================
        Dim nn As Variant
        Dim mRowCount As Variant
        Dim rr As Variant
        nn = vsGrid.Rows
        For rr = 1 To nn - 1
            If vsGrid.TextMatrix(rr, 1) = "" Then
                Exit For
            End If
        Next rr
        mRowCount = rr - 1
        '=======================================================
        
        
        
        
        ''      Advance Deduction       '' Modified By sunil on 05-08-2011 for Transactiontype wise Details
        mSql = "Select faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,Sum(faTransactionChild.fltAmount)[fltAmount],faTransactionType.intTransactionTypeID as intTransTypeID From faTransactionChild " & _
                "Inner Join faTransactions On faTransactions.intTransactionID = faTransactionChild.intTransactionID " & _
                "Inner Join faTransactionType on faTransactionType.intTransactionTypeID=faTransactions.intTransactionTypeID " & _
                "Inner Join faVouchers On faTransactions.intVoucherID = faVouchers.intVoucherID " & _
                "inner Join faAccountHeads on faAccountHeads.intAccountHeadID=faTransactionChild.intAccountHeadID " & _
                "Where intByAccountHeadID = 1157 And tnyCancelFlag = 0 And intInstrumentTypeID = 1 And dtDate  = '" & Format(mDt, "dd/MMM/yyyy") & "' Group By faAccountHeads.intAccountHeadID,faTransactionType.intTransactionTypeID,faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            While Not Rec.EOF
                For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mCount, 13) = Rec!intTranstypeID Then
                    If vsGrid.TextMatrix(mCount, 6) = Rec!intAccountHeadID Then
                        If val(vsGrid.TextMatrix(mCount, 4)) = 0 Then
                            vsGrid.TextMatrix(mCount, 5) = val(vsGrid.TextMatrix(mCount, 5)) - Rec!fltAmount
                        Else
                            vsGrid.TextMatrix(mCount, 4) = val(vsGrid.TextMatrix(mCount, 4)) - Rec!fltAmount
                        End If
                        '=========================For Advance Head ===================
                        
                                 vsGrid.Rows = vsGrid.Rows + 1
                                 vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = Rec!vchAccountHeadCode
                              '   objAcc.SetAccountCode (Rec!vchAccountHeadCode)
                                 vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = Rec!vchAccountHead
                                 vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = Format(Rec!fltAmount, "0.00")
                                 vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = Rec!intAccountHeadID
                                 vsGrid.TextMatrix(vsGrid.Rows - 1, 13) = Rec!intTranstypeID
                                vsGrid.TextMatrix(vsGrid.Rows - 1, 14) = 1 'To Identify Advance Amount
                                 'vsGrid.Cell(flexcpCustomFormat, mLoop, 0, , 8)
                                 vsGrid.Cell(flexcpBackColor, vsGrid.Rows - 1, 0, , 8) = &HC0FFC0
                                 
                   End If
                End If
                mRowCount = mRowCount + 1
                Next
               Rec.MoveNext
            Wend
        End If
        If mAdvAmount <> 0 Then
            cmdSave.Enabled = False
        End If
        Rec.Close
        Call Calculate
  '      Me.Mode = 0  '
         vsGrid.Editable = flexEDNone
         
         '======Lock =======
         If mFlag = 1 Then
                cmbSections.Enabled = False
                txtDemandDate.Enabled = False
                txtTransactionType.Enabled = False
                cmbOutDoorStaff.Enabled = False
                cmbInstrumentType.Enabled = False
                txtDrawnFrom.Enabled = False
                txtDrawnPlace.Enabled = False
                txtDemandDate.Enabled = False
                cmdSearchTransactionType.Enabled = False
                cmbOutDoorStaff.Enabled = False
                cmbInstrumentType.Enabled = False
                txtName.Enabled = False
                txtInitial1.Enabled = False
                txtInitial2.Enabled = False
                txtInitial3.Enabled = False
                txtInitial4.Enabled = False
                txtDrawnFrom.Enabled = IIf(Len(Trim(txtDrawnFrom)), False, True)
                txtDrawnPlace.Enabled = IIf(Len(Trim(txtDrawnPlace)), False, True)
                cmbZone.Enabled = False
                txtwardno.Enabled = False
                txtDoorNo1.Enabled = False
                txtDoorNo2.Enabled = False
                cmbInstrumentType.Enabled = False
                txtInstrumentNo.Enabled = False
                txtInstrumentDate.Enabled = False
                txtHouseName.Enabled = False
                txtStreet.Enabled = False
                txtLocalPlace.Enabled = False
                txtMainPlace.Enabled = False
                txtPost.Enabled = False
                txtPin.Enabled = False
                chkRoundOff.Enabled = False
                txtPhone.Enabled = False
                txtArrearAmt.Enabled = False
                txtCurrentAmt.Enabled = False
                txtGrandTotal.Enabled = False
                txtRemarks.Enabled = False
                txtAdminNote.Enabled = False
                cmbSeat.Enabled = False
                chkSkipPrinting.Enabled = False
                chkTag.Enabled = False
                txtFunction.Enabled = False
                txtFunctionary.Enabled = False
                txtSourceofFund.Enabled = False
                txtReference.Enabled = False
           End If
        '===========================================
         
         
         
    End Sub
   
    Public Property Let Mode(mMode As Integer)
        mZonalCollection = mMode
    End Property

    Public Property Let DemandNo(mData As Variant)
        mDemandNo = mData
    End Property
    
'     Public Property Let ProfTaxInstTypeMode(mData As Boolean)
'        mProfTaxInstTypeMode = mData
'    End Property
'
'    Public Property Get ProfTaxInstTypeMode() As Boolean
'        ProfTaxInstTypeMode = mProfTaxInstTypeMode
'    End Property
'
''---------------------=======================================--------------------------
''-----------------Reverse Entry Demand------------------------------------------
''                 Codded By anisha
''                 Dated:22/10/2010
''--------------------------------------------------------------------------------------
''---------------------***************************************--------------------------
    
    Public Sub ReverseDemandDetails(ByVal mVoucherID As Variant, ByVal mType As Integer)
        'Fill Demand Details Using Voucher No
        'mType=1 Demand No, 2=VoucherID
        'For Reverse Entry
        
            Dim objdb       As New clsDB
            Dim objAc       As New clsAccounts
            Dim Rec         As New ADODB.Recordset
            Dim RecChild    As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim objUser     As New clsUser
            Dim mSql        As String
            Dim mRowCount   As Integer
            Dim mPeriodId   As Integer
            Dim mYearID     As Integer
            Dim mArrearFlag As Integer
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            If mType = 1 Then
                mSql = "Select vchDemandNo From faIDemandTBL Where numDemandID=" & mVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    txtDemandNo.Text = IIf(IsNull(Rec!vchDemandNo), "", Rec!vchDemandNo)
                    Call txtDemandNo_LostFocus
                End If
            Else
                mSql = "Select * From faVouchers"
                mSql = mSql + " Inner Join faVoucherChild On faVouchers.intVoucherID=faVoucherChild.intVoucherID"
                mSql = mSql + " Inner Join faVoucherAddress On faVouchers.intVoucherID=faVoucherAddress.intVoucherID"
                mSql = mSql + " Inner Join faTransactions On faTransactions.intVoucherID=faVouchers.intVoucherID"
                mSql = mSql + " Inner Join faTransactionType On faVouchers.intTransactionTypeID=faTransactionType.intTransactionTypeID"
                mSql = mSql + " Inner Join faSection On faTransactionType.intSectionID=faSection.intSectionID"
                mSql = mSql + " Inner Join faInstrumentTypes On faVouchers.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
                mSql = mSql + " Inner Join faAccountHeads On faVouchers.intKeyID1=faAccountHeads.intAccountHeadID"
                mSql = mSql + " Left Join faFunctionaries On faFunctionaries.intFunctionaryID=faTransactions.intFunctionaryID"
                mSql = mSql + " Left Join faFunctions On faFunctions.intFunctionID=faTransactions.intFunctionID"
                mSql = mSql + " Left Join faFunds On faFunds.intFundID=faTransactions.intFundID"
                mSql = mSql + " Left Join DB_Masters..GM_Zone On faVouchers.numZoneID=DB_Masters..GM_Zone.numZoneID"
                mSql = mSql + " Where faVouchers.intVoucherID=" & mVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                        If gbLBType = 1 Or gbLBType = 2 Or gbLBType = 5 Then
                            cmbSections.Text = "Panchayat Office"
                        Else
                            cmbSections.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
                        End If
    '                    If cmbTransactionType.ListIndex = -1 Then
    '                        mSql = ""
    '                        mSql = "SELECT faTransactionType.vchTransactionType, faSectionWiseTransactionTypes.intTransactionTypeID "
    '                        mSql = mSql + " FROM faSectionWiseTransactionTypes INNER JOIN "
    '                        mSql = mSql + " faTransactionType ON faSectionWiseTransactionTypes.intTransactionTypeID = faTransactionType.intTransactionTypeID "
    '                        mSql = mSql + " Where (faTransactionType.intGroupID = 10)"
    '                        '' mSQL = mSQL + " And isNull(tnyHidden,0)=0"
    '                        mSql = mSql + " ORDER BY faTransactionType.vchTransactionType"
    '                        Call PopulateList(cmbTransactionType, mSql, , True, True, True)
    '                        cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
    '                    Else
    '                        cmbTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
    '                    End If
                        
                        txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                        txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                        If mType <> 1 Then
                            If txtTransactionType.Tag = gbTransactionTypeZonalCollection Then
                                cmbMode.ListIndex = 2
                            ElseIf txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                                cmbMode.ListIndex = 3
                            ElseIf txtTransactionType.Tag = gbTransactionTypeFriendsCollection Then
                                cmbMode.ListIndex = 4
                            Else
                                cmbMode.ListIndex = 1
                            End If
                        End If
                        'If cmbTransactionType.ItemData(cmbTransactionType.ListIndex) = gbTransactionTypeOutDoor Then
                        If txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                            cmbOutDoorStaff.Enabled = True
                            objUser.SetUser (Rec!numUserID)
                            Call FillOutDoorStaffs
                            If IsNull(objUser.UserName) = False Then
                                cmbOutDoorStaff.Text = IIf(IsNull(objUser.UserName), "", objUser.UserName)
                            End If
                        End If
                        'End If
                        
        '                If IsNull(Rec!chvZoneNameEnglish) = False Then
        '                    .cmbOutDoorStaff.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
        '                End If
                        txtwardno.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
                        txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                        txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
                        If IsNull(Rec!vchInstrumentType) = False Then
                            cmbInstrumentType.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                            cmbInstrumentType.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                        End If
'                        If cmdAcHead.Enabled Then
                        objAc.SetAccountID (IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1))
                        txtAccountHead.Text = objAc.AccountHead
                        txtAccountHead.Tag = objAc.AccountHeadID
                        txtAccountCode.Text = objAc.AccountCode
                            
'                        End If
                        If cmbInstrumentType.Tag <> "" Then
                            If cmbInstrumentType.Tag <> 1 Then
                                lblInstNo.Visible = True
                                txtInstrumentNo.Visible = True
                                txtInstrumentNo.Enabled = True
                                lblInstDate.Visible = True
                                txtInstrumentDate.Visible = True
                                txtInstrumentDate.Enabled = True
                                lblDrawnFrom.Visible = True
                                txtDrawnFrom.Visible = True
                                txtDrawnFrom.Enabled = True
                                lblDrawnPlace.Visible = True
                                txtDrawnPlace.Visible = True
                                txtDrawnPlace.Enabled = True
                                txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                                txtInstrumentDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
                                txtDrawnFrom.Text = IIf(IsNull(Rec!vchBank), "", Rec!vchBank)
                                txtDrawnPlace.Text = IIf(IsNull(Rec!vchBankPlace), "", Rec!vchBankPlace)
                            End If
                        End If
                        txtFunction.Text = IIf(IsNull(Rec!vchFunction), "Accounts", Rec!vchFunction)
                        txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), 6, Rec!intFunctionID)
                        txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "Accounts Department", Rec!vchFunctionary)
                        txtFunctionary.Tag = IIf(IsNull(Rec!vchFunctionary), 4, Rec!vchFunctionary)
                        txtSourceofFund.Text = IIf(IsNull(Rec!vchFund), "General Fund", Rec!vchFund)
                        txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceFundID), 1, Rec!intSourceFundID)
                        
                        txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                        txtInitial1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
                        txtInitial2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
                        txtInitial3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
                        txtInitial4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
                        txtHouseName.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                        '.txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                        txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                        txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                        '.txtPost.Text = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
                        '.txtPin.Text = IIf(IsNull(Rec!vchPin), "", Rec!vchPin)
                        txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                                    
                        txtRemarks.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                        txtAdminNote.Text = IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
        '                If IsNull(Rec!numForwardedSeatID) = False Then
        '                    .cmbSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
        '                End If
                        mSql = ""
                        mSql = "Select * From faVoucherChild"
                        mSql = mSql + " Inner Join faAccountHeads On faVoucherChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
                        mSql = mSql + " Left Join faPeriodicity On faPeriodicity.intPeriodicityID=faVoucherChild.tnyPeriodID"
                        mSql = mSql + " Where intVoucherID=" & mVoucherID
                        RecChild.Open mSql, mCnn
                        mRowCount = 1
                        While Not Rec.EOF
                            While Not RecChild.EOF
                                vsGrid.TextMatrix(mRowCount, 0) = IIf(IsNull(RecChild!vchAccountHeadCode), "", RecChild!vchAccountHeadCode)
                                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecChild!vchAccountHead), "", RecChild!vchAccountHead)
                                mPeriodId = IIf(IsNull(RecChild!tnyPeriodID), "", RecChild!tnyPeriodID)
                                mYearID = IIf(IsNull(RecChild!intYearID), 0, RecChild!intYearID)
                                If mYearID <> 0 Then
                                    vsGrid.TextMatrix(mRowCount, 2) = mYearID & "-" & mYearID + 1
                                End If
                                'vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchPeriodicity), "", Rec!vchPeriodicity)
                                mArrearFlag = IIf(IsNull(RecChild!tnyArrearFlag), "", RecChild!tnyArrearFlag)
                                If mArrearFlag = 0 Then
                                    vsGrid.TextMatrix(mRowCount, 5) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
                                End If
                                If mArrearFlag = 1 Then
                                    vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecChild!fltAmount), "", RecChild!fltAmount)
                                End If
                                
    '                            If vsGrid.TextMatrix(mRowCount, 0) = gbAcHeadCodeAdvanceBuilding _
    '                            Then
    '                                vsGrid.TextMatrix(mRowCount, 14) = 1
    '                            End If
                                vsGrid.Rows = vsGrid.Rows + 1
                                mRowCount = mRowCount + 1
                                RecChild.MoveNext
                            Wend
                            Rec.MoveNext
                        Wend
                        RecChild.Close
                End If
            End If
        End Sub
        
        Private Function FunTrTypeAccHeadValidate(ByVal mTypeID As Integer, ByVal AccCode As String) As Boolean
        'To validate AccountHead and TransactionType
           Dim mSql     As String
           Dim objdb    As New clsDB
           Dim Rec      As New ADODB.Recordset
           Dim mCnn     As New ADODB.Connection
           Dim mFlag    As Boolean
           mFlag = False
           objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
           mSql = "Select vchAccountHeadCode From faTransactionTypeChild Where intTransactionTypeID=" & mTypeID
           Rec.Open mSql, mCnn
           
           If Not (Rec.EOF And Rec.BOF) Then
                Rec.Close
                mSql = mSql + " And vchAccountHeadCode=" & AccCode
                Rec.Open mSql, mCnn
                If (Rec.EOF And Rec.BOF) Then
                    mFlag = False
                Else
                    mFlag = True
                End If
                FunTrTypeAccHeadValidate = mFlag
           Else
                FunTrTypeAccHeadValidate = True
           End If
        End Function
        
        Public Property Let Reverse(mData As Integer)
            mReverse = mData
        End Property
        Public Property Let PreviousYearMode(mData As Integer)
            mPreviousYearMode = mData
        End Property
        Public Property Let PendingTaskReqID(mData As Integer)
            mPreviousYearRequestID = mData
        End Property

        

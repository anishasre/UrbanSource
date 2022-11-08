VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDemand 
   BackColor       =   &H00DAF2F2&
   Caption         =   "Demand"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   Icon            =   "frmDemand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11820
   Begin VB.TextBox txtTransactionDate 
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
      Height          =   255
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   900
      Width           =   1635
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   45
      Width           =   3840
   End
   Begin VB.TextBox txtOutDoorStaff 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   435
      Width           =   4290
   End
   Begin VB.TextBox txtTransactionType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   615
      Width           =   3840
   End
   Begin VB.TextBox txtSection 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   315
      Width           =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAF2F2&
      Height          =   5505
      Left            =   -30
      TabIndex        =   6
      Top             =   1110
      Width           =   11850
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2130
         Left            =   60
         TabIndex        =   12
         Top             =   570
         Width           =   11745
         _cx             =   20717
         _cy             =   3757
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDemand.frx":1CCA
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
      Begin VB.TextBox txtForwardedSeat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7995
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   5025
         Width           =   1770
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   195
         Width           =   2640
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   195
         Width           =   2640
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   195
         Width           =   2640
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DAF2F2&
         Height          =   2850
         Left            =   60
         TabIndex        =   13
         Top             =   2655
         Width           =   7800
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
            Height          =   315
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   1125
            Width           =   1785
         End
         Begin VB.TextBox txtZone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   135
            Width           =   1800
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
            Left            =   1065
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   37
            Top             =   465
            Width           =   1800
         End
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
            Left            =   1065
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   36
            Top             =   795
            Width           =   1110
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
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   35
            Top             =   795
            Width           =   690
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
            Height          =   300
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   34
            Top             =   225
            Width           =   2535
         End
         Begin VB.TextBox txtHouseName 
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
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   33
            Top             =   540
            Width           =   3210
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
            Height          =   300
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   32
            Top             =   855
            Width           =   3210
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
            Height          =   300
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   31
            Top             =   1170
            Width           =   3210
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
            Height          =   300
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   30
            Top             =   1485
            Width           =   3210
         End
         Begin VB.TextBox txtInitial1 
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
            Left            =   6435
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   29
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtInitial2 
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
            Left            =   6765
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   28
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtInitial3 
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
            Left            =   7095
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
         End
         Begin VB.TextBox txtInitial4 
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
            Left            =   7425
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   225
            Width           =   315
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
            Height          =   300
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1800
            Width           =   2025
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
            Height          =   300
            Left            =   6180
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   24
            Top             =   1800
            Width           =   915
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
            Height          =   315
            Left            =   3885
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   23
            Top             =   2115
            Width           =   2010
         End
         Begin VB.Frame fraInstrument 
            BackColor       =   &H00DAF2F2&
            BorderStyle     =   0  'None
            Height          =   1305
            Left            =   60
            TabIndex        =   14
            Top             =   1455
            Width           =   2850
            Begin VB.TextBox txtInstrumentNo 
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
               Left            =   1005
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   17
               Top             =   0
               Width           =   1800
            End
            Begin VB.TextBox txtInstrumentDate 
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
               Left            =   1005
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   18
               Top             =   330
               Width           =   1800
            End
            Begin VB.TextBox txtDrawnFrom 
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
               Left            =   1005
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   16
               Top             =   660
               Width           =   1800
            End
            Begin VB.TextBox txtDrawnPlace 
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
               Left            =   1005
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   15
               Top             =   990
               Width           =   1800
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
               TabIndex        =   22
               Top             =   75
               Width           =   540
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
               TabIndex        =   21
               Top             =   375
               Width           =   645
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
               TabIndex        =   20
               Top             =   690
               Width           =   900
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
               TabIndex        =   19
               Top             =   990
               Width           =   930
            End
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
            TabIndex        =   49
            Top             =   510
            Width           =   630
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
            TabIndex        =   48
            Top             =   840
            Width           =   585
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
            TabIndex        =   47
            Top             =   195
            Width           =   375
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
            TabIndex        =   46
            Top             =   255
            Width           =   405
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
            TabIndex        =   45
            Top             =   585
            Width           =   960
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
            TabIndex        =   44
            Top             =   915
            Width           =   435
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
            TabIndex        =   43
            Top             =   1215
            Width           =   825
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
            TabIndex        =   42
            Top             =   1530
            Width           =   765
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
            TabIndex        =   41
            Top             =   1845
            Width           =   315
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
            TabIndex        =   40
            Top             =   1845
            Width           =   630
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
            TabIndex        =   39
            Top             =   2175
            Width           =   690
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
            TabIndex        =   38
            Top             =   1185
            Width           =   750
         End
      End
      Begin VB.TextBox txtArrearAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2775
         Width           =   1725
      End
      Begin VB.TextBox txtCurrentAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9690
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2775
         Width           =   1755
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9690
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3075
         Width           =   1755
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   450
         Left            =   7995
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3645
         Width           =   3450
      End
      Begin VB.TextBox txtAdminNote 
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
         Height          =   450
         Left            =   7995
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4335
         Width           =   3450
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
         TabIndex        =   56
         Top             =   3135
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         TabIndex        =   55
         Top             =   3435
         Width           =   630
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
         TabIndex        =   54
         Top             =   4095
         Width           =   2055
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
         Left            =   7965
         TabIndex        =   53
         Top             =   4815
         Width           =   1035
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
         Left            =   4215
         TabIndex        =   52
         Top             =   225
         Width           =   615
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
         Left            =   330
         TabIndex        =   51
         Top             =   225
         Width           =   855
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
         Left            =   7680
         TabIndex        =   50
         Top             =   225
         Width           =   1155
      End
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
      Left            =   10035
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   1635
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
      Left            =   7380
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "<New>"
      Top             =   105
      Width           =   1635
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
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   765
      Width           =   4290
   End
   Begin VB.TextBox txtSourceOfFund 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8700
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   2370
   End
   Begin VB.TextBox txtFunctionary 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   2370
   End
   Begin VB.TextBox txtFunction 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1110
      Width           =   2370
   End
   Begin VB.Label lblTrDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transactio Date"
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
      Left            =   360
      TabIndex        =   75
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode of Collection"
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
      Left            =   225
      TabIndex        =   73
      Top             =   60
      Width           =   1320
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
      Left            =   975
      TabIndex        =   62
      Top             =   330
      Width           =   540
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
      Left            =   240
      TabIndex        =   61
      Top             =   660
      Width           =   1260
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
      Left            =   9075
      TabIndex        =   60
      Top             =   150
      Width           =   960
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
      Left            =   6510
      TabIndex        =   59
      Top             =   150
      Width           =   825
   End
   Begin VB.Label lblCombo 
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
      Left            =   5550
      TabIndex        =   58
      Top             =   510
      Width           =   1785
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
      Left            =   6330
      TabIndex        =   57
      Top             =   825
      Width           =   1005
   End
End
Attribute VB_Name = "frmDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mDemandNo As Variant
    
    '*********************************************************************************************'
    '                  Form for just viewing the Demand through Demand Register                   '
    '*********************************************************************************************'
    Private Sub FormInitialize()
        vsGrid.Clear 1, 0
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
    End Sub
    
        Private Sub DispalyDemand(vchDemandNo As String)
        Dim objDB           As New clsDB
        Dim mcnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim mSQL            As String
        Dim mDemandNo       As String
        Dim mArrearFlag     As Variant
        Dim mRowCount       As Integer
        Dim mSeatID         As Variant
        Dim mVoucherID      As Variant
        Dim mStatus         As Variant
        Dim mCancelFlag     As Variant
        
        '*********************************************************************************************'
        '                  Procedure to Fill all the details of a particular demand                   '
        '*********************************************************************************************'
        On Error GoTo err
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        mDemandNo = vchDemandNo
            
        mSQL = "Select *,faIDemandTbl.intTransactionTypeID[TransactionTypeID],faIDemandTbl.intSectionID[SectionID],faIDemandTbl.tnyStatus As Status,faIDemandTbl.numUserID As UserID,faIDemandTBL.numSeatID As SeatID From faIDemandTBL"
        mSQL = mSQL + " Inner Join faIDemandChild On faIDemandTBL.numDemandID=faIDemandChild.numDemandID"
        mSQL = mSQL + " Inner Join faIDemandAddress On faIDemandTBL.numDemandID=faIDemandAddress.numDemandID"
        mSQL = mSQL + " Left Join faInstrumentTypes On faIDemandTBL.intInstrumentTypeID=faInstrumentTypes.intInstrumentTypeID"
        mSQL = mSQL + " Left Join faSection On faIDemandTBL.intSectionID=faSection.intSectionID"
        mSQL = mSQL + " Inner Join faTransactionType On faIDemandTBL.intTransactionTypeID=faTransactionType.intTransactionTypeID"
        mSQL = mSQL + " Inner Join faAccountHeads On faIDemandChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
        'mSQL = mSQL + " Left Join DB_Sanchayalite..snPDE_ODStaff On faIDemandTBL.intKeyID2=DB_Sanchayalite..snPDE_ODStaff.numUserID"
        mSQL = mSQL + " Left Join faUser On faIDemandTBL.intKeyID2 = faUser.numUserId And tnyOutDoorStaffs = 0"
        mSQL = mSQL + " Left Join faFunctions on faIDemandTBL.intFunctionID = faFunctions.intFunctionID"
        mSQL = mSQL + " Left Join faFunctionaries on faIDemandTBL.intFunctionaryID = faFunctionaries.intFunctionaryID"
        mSQL = mSQL + " Left Join suSourceOfFund On faIDemandTBl.intSourceFundID = suSourceOfFund.intSourceFundID"
        mSQL = mSQL + " Left Join DB_Masters..GM_Zone On faIDemandTBL.numZoneID=DB_Masters..GM_Zone.numZoneID"
        mSQL = mSQL + " Left Join DB_Masters..GL_Seats On faIDemandTBL.numForwardedSeatID=DB_Masters..GL_Seats.numSeatID"
        '----Added on 7.Jul.2011 By Anisha
        mSQL = mSQL + " Left Join faDemandMode On faDemandMode.intDemandModeID=faIDemandTbl.intDemandMode"
        mSQL = mSQL + " Where faIDemandTBL.vchDemandNo='" & mDemandNo & "'"
        mSQL = mSQL + " Order By faIDemandChild.tnyArrearFlag Desc,faAccountHeads.vchAccountHeadCode"
        Rec.Open mSQL, mcnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtSection.Text = IIf(IsNull(Rec!vchSectionName), "", Rec!vchSectionName)
            txtSection.Tag = IIf(IsNull(Rec!SectionID), "", Rec!SectionID)
            txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
            txtTransactionType.Tag = IIf(IsNull(Rec!TransactionTypeID), "", Rec!TransactionTypeID)
            txtDemandNo.Text = mDemandNo
            txtDemandNo.Tag = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
            txtDemandDate.Text = IIf(IsNull(Rec!dtDemandDate), "", Rec!dtDemandDate)
            txtDemandDate.Tag = IIf(IsNull(Rec!SeatID), "", Rec!SeatID) 'Demand Generated SeatID
            txtReference.Tag = IIf(IsNull(Rec!UserID), "", Rec!UserID)  'Demand Generated UserID
            If IsNull(Rec!intDemandMode) Then
                txtMode.Text = "Direct"
                txtMode.Tag = 1
            Else
                txtMode.Text = Rec!vchDemandMode
                txtMode.Tag = Rec!intDemandMode
            End If
            If IsNull(Rec!dtTransactionDate) Then
                lblTrDate.Visible = False
                txtTransactionDate.Visible = False
            Else
                txtTransactionDate.Text = Format(Rec!dtTransactionDate, "dd-mmm-yyyy")
            End If
            If txtTransactionType.Tag = gbTransactionTypeOutDoor Then
                If IsNull(Rec!vchUserName) = False Then
                    txtOutDoorStaff.Text = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                End If
            End If
            
            txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
            txtSourceOfFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            txtSourceOfFund.Tag = IIf(IsNull(Rec!intSourceFundID), "", Rec!intSourceFundID)
            
            If IsNull(Rec!chvZoneNameEnglish) = False Then
                txtZone.Text = IIf(IsNull(Rec!chvZoneNameEnglish), "", Rec!chvZoneNameEnglish)
            End If
            txtWardNo.Text = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
            txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
            txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
            If IsNull(Rec!vchInstrumentType) = False Then
                txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
                txtInstrument.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
            End If
            If txtInstrument.Tag <> "" Then
                If txtInstrument.Tag <> 1 Then
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
                        
            txtRemarks.Text = IIf(IsNull(Rec.Fields(12).value), "", Rec.Fields(12).value)
            txtAdminNote.Text = IIf(IsNull(Rec!vchAdminNote), "", Rec!vchAdminNote)
            If IsNull(Rec!numForwardedSeatID) = False Then
                txtForwardedSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
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
            MsgBox "Demand Number doesn't exists !!", vbInformation
        End If
        Rec.Close
        Exit Sub
err:
        MsgBox err.Description
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
        txtGrandTotal.Text = Format(mAmtArrear + mAmtCurrent, "0.00")
        'txtRoundOff.Text = Format(RoundOffAdjustment(val(txtTotal)), "0.00")
        'txtTotal.Text = Format(val(txtTotal) + val(txtRoundOff) - val(txtAdvance), "0.00")
    End Sub
  
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.Width = 11940
        Me.Height = 7155
    End Sub

    Private Sub Form_Load()
        FormInitialize
        If DemandNo <> "" Then
            DispalyDemand (DemandNo)
        End If
    End Sub
    
    Public Property Let DemandNo(mVal As String)
        mDemandNo = mVal
    End Property
    
    Public Property Get DemandNo() As String
        DemandNo = mDemandNo
    End Property

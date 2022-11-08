VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUSevanaInward 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Sevana Inward Details"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15465
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameReceipt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Receipt Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      TabIndex        =   59
      Top             =   6240
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame frameReceiptSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   9165
         Begin VB.CommandButton cmdReceiptSearch 
            Appearance      =   0  'Flat
            Caption         =   "Search"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8040
            TabIndex        =   73
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox txtReceiptSearch 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   1995
         End
         Begin VB.OptionButton optPayedReceipt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Payed Receipt"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optInterruptReceipt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inetrrupt Receipt"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   70
            Top             =   240
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Receipt No"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   4770
            TabIndex        =   75
            Top             =   285
            Width           =   885
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Receipt Mode"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   74
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   12795
         Begin VB.TextBox txtNoOfCertificate 
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
            Height          =   315
            Left            =   8310
            TabIndex        =   63
            Top             =   150
            Width           =   765
         End
         Begin VB.CommandButton cmdCopy 
            Appearance      =   0  'Flat
            Caption         =   "Copy to Receipt"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7005
            TabIndex        =   62
            Top             =   3015
            Width           =   1965
         End
         Begin VB.TextBox txtNoofYears 
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
            Height          =   285
            Left            =   5715
            TabIndex        =   61
            Top             =   180
            Width           =   735
         End
         Begin VSFlex8LCtl.VSFlexGrid vsGrid 
            Height          =   2370
            Left            =   0
            TabIndex        =   64
            Top             =   600
            Width           =   8850
            _cx             =   15610
            _cy             =   4180
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
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmUSevanaInward.frx":0000
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount :"
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
            Left            =   4650
            TabIndex        =   68
            Top             =   3105
            Width           =   1185
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   5925
            TabIndex        =   67
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No of Certificates"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6660
            TabIndex        =   66
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label lblNoofYears 
            BackStyle       =   0  'Transparent
            Caption         =   "No of Years"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4590
            TabIndex        =   65
            Top             =   210
            Width           =   1095
         End
      End
   End
   Begin VB.Frame frameSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox txtNoCopeis 
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
         Height          =   285
         Left            =   1920
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboLanguage 
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
         Height          =   315
         Left            =   5760
         TabIndex        =   55
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame frameRegister 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Details in Register"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   8895
         Begin VB.ComboBox cboRelationship 
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
            Height          =   315
            Left            =   1440
            TabIndex        =   49
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox txtEnglishname 
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
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtMalayalamname 
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
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtRegNo 
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
            Height          =   285
            Left            =   5640
            TabIndex        =   46
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtBookNo 
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
            Height          =   285
            Left            =   7920
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdGetName 
            Appearance      =   0  'Flat
            Caption         =   "Get Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6120
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5625
            TabIndex        =   43
            Top             =   720
            Width           =   1395
         End
         Begin VB.CommandButton cmdClear 
            Appearance      =   0  'Flat
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7080
            TabIndex        =   42
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Relationship"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "English"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Malayalam"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Reg No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   51
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Book No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6960
            TabIndex        =   50
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No of Copies"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Language"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   57
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame frameReceiptexe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Receipt Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox txtReceiptBookNo 
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
         Height          =   285
         Left            =   5400
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtReceiptAmount 
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
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtReceiptNo 
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
         Height          =   285
         Left            =   5400
         TabIndex        =   32
         Top             =   840
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPReceiptDate 
         Height          =   300
         Left            =   1680
         TabIndex        =   35
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64880641
         CurrentDate     =   40038
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Book No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Receipt No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame frameSevana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sevana Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.ComboBox cboSubType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   6015
      End
      Begin VB.ComboBox cboHospitals 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   13
         Top             =   1230
         Width           =   6525
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1680
         Width           =   2805
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Cance&L"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtSubTypeID 
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
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkZonal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Zonal Office"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CheckBox chkInsideCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inside Country"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1770
         TabIndex        =   7
         Top             =   270
         Width           =   1575
      End
      Begin VB.CheckBox chkOutsideCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Outside Country"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   270
         Width           =   1755
      End
      Begin VB.TextBox txtwardno 
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbDepartment 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2325
      End
      Begin VB.ComboBox cmbSeat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2400
         Width           =   1545
      End
      Begin VB.ComboBox cmbSeatID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   465
      End
      Begin MSComCtl2.DTPicker dtpEventDate 
         Height          =   375
         Left            =   7260
         TabIndex        =   5
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64880641
         CurrentDate     =   40631
      End
      Begin MSComCtl2.DTPicker DTPApplDate 
         Height          =   300
         Left            =   1200
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64880641
         CurrentDate     =   40037
      End
      Begin MSComCtl2.DTPicker dtpDeliveryDate1 
         Height          =   360
         Left            =   6000
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64880641
         CurrentDate     =   40544
      End
      Begin VB.Label lblMandatory 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   135
         Index           =   2
         Left            =   6750
         TabIndex        =   30
         Top             =   330
         Width           =   135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Application Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblMandatory 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Event Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5820
         TabIndex        =   26
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sub Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   750
         Width           =   1635
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hospitals"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000005&
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4680
         TabIndex        =   22
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9000
         TabIndex        =   21
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label lblusername 
         BackColor       =   &H80000005&
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7320
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10455
      Left            =   9480
      TabIndex        =   76
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   18441
      _Version        =   393216
      MousePointer    =   5
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Checklist"
      TabPicture(0)   =   "frmUSevanaInward.frx":0218
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grvCheckList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VSFlex8LCtl.VSFlexGrid grvCheckList 
         Height          =   6015
         Left            =   0
         TabIndex        =   77
         Top             =   240
         Width           =   5775
         _cx             =   10186
         _cy             =   10610
         Appearance      =   0
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUSevanaInward.frx":0234
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
   End
End
Attribute VB_Name = "frmUSevanaInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Option Explicit
    
    Dim tnyType             As Variant
    Public SevanaTypeID     As Variant
    Public SevanaKioskID    As Variant
    Dim CommMarriageFee     As Variant
    Dim objdb               As New clsDB
    '-------------------------------------'
    Dim intTransactionTypeID As Integer
    
    
    Dim AmtNoofCert         As Double
    Dim SearchAmt           As Double
    Dim mTempTotal          As Double
    Dim AmtNoofExtraCert    As Double
    Dim mReceiptNo          As Variant
    Dim mReceiptAmt         As Double
    Dim mCurRow As Variant
      
    Private Sub Calculate()
        Dim mLoop As Integer
        Dim mTotalAmt As Double
        
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 0) = 1 Then
                mTotalAmt = mTotalAmt + val(vsGrid.TextMatrix(mLoop, 7))
            End If
        Next
        lblTotal.Caption = Format(mTotalAmt, "0.00")
    End Sub
    
    Public Sub ShowFrames()
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        
        Dim Rec As New ADODB.Recordset
        vsGrid.Clear  '''added on 17 jul 2015 By soumya V S
        mReceiptNo = 0
      '  mReceiptAmt = 0
        If cboSubType.ListIndex = -1 Then Exit Sub
        txtSubTypeID.Text = cboSubType.ItemData(cboSubType.ListIndex)
        'changed by soumya V S on 14.05.14
        If txtSubTypeID.Text = 93 Or txtSubTypeID.Text = 80 Or txtSubTypeID.Text = 81 Or txtSubTypeID.Text = 82 Or txtSubTypeID.Text = 83 Or txtSubTypeID.Text = 91 Or txtSubTypeID.Text = 92 Or txtSubTypeID.Text = 101 Or txtSubTypeID.Text = 102 Then
            MsgBox "This subtype is blocked", vbInformation
            txtSubTypeID.Text = ""
            cboSubType.ListIndex = 0
        Else
             'If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Then
             If txtSubTypeID.Text = "2" Then
                 Label2.Caption = "Arrival Date"
             Else
                 Label2.Caption = "Application Date"
             End If
             
             mSql = "Select tnyType,tnyToSeat from mSubjectSevanaSubType where intid= " & txtSubTypeID
             If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
                 MsgBox "Connection not present", vbDefaultButton1
                 Exit Sub
             End If
            
             Rec.Open mSql, mCnn
             If Not (Rec.BOF Or Rec.EOF) Then
                 SevanaTypeID = Rec!tnyType
                 SevanaKioskID = Rec!tnyToSeat
             End If
             Rec.Close
             If SevanaTypeID = 1 Then
                If chkZonal.Value = 1 Then
                    frameReceipt.Visible = False
                    'frameSearch.Visible = True
                    frameSearch.Top = 2500
                    Me.Left = 2250
                    Me.Top = 2000
                    'Me.Height = 2500
                    Me.Height = 2800
                    cmdOK.Enabled = True
                Else
                    frameReceipt.Visible = True
                    Me.Left = 2250
                    Me.Top = 2000
                    'Me.Height = 6250
                    Me.Height = 9720
                    frameReceipt.Top = 2800
                    cmdOK.Enabled = False
                    Call FillGrid
                End If
              ElseIf SevanaTypeID = 2 Then
'                If txtSubTypeID.Text = 110 Or txtSubTypeID.Text = 8 Or chkZonal.value = 1 Then
                 If txtSubTypeID.Text = 110 Or chkZonal.Value = 1 Or txtSubTypeID.Text = 8 Or txtSubTypeID.Text = 122 Or txtSubTypeID.Text = 115 Then    'Ranjitha 09/10
                    frameReceipt.Visible = False
                    frameSearch.Visible = True
                    frameSearch.Top = 2500
                    Me.Left = 2250
                    'Me.Top = 2000
                    Me.Top = 1500
                    Me.Height = 5750
                    txtNoOfCertificate.Text = 1
                    cmdOK.Enabled = True
                Else
                '
'                    'add Vipin on 26/05/2012
'                    If (txtSubTypeID.Text = 111) Then
'                    frameReceipt.Visible = False
'                    frameSearch.Visible = True
'                    frameReceipt.Top = 5550
'                    Me.Left = 2250
'                    'Me.Top = 2000
'                    Me.Top = 1000
'                    Me.Height = 6700
'                    'Me.Height = 9720
'                    Call FillGrid
'                    cmdOK.Enabled = True
'
'                    Else
                    frameReceipt.Visible = True
                    frameSearch.Visible = True
                    frameReceipt.Top = 5550
                    Me.Left = 2250
                    'Me.Top = 2000
                    Me.Top = 1000
                    'Me.Height = 8700
                    Me.Height = 9720
                    Call FillGrid
                    cmdOK.Enabled = False
                    If InwardMode <> 0 Then
                        Me.Top = 250
                        Me.Width = 9825
                        Me.Height = 10845
                    Else
                        Frame1.Top = 160
                    End If
                End If
             Else
                 frameReceipt.Visible = False
                 frameSearch.Visible = False
                 cmdOK.Enabled = True
                Me.Left = 2250
                 Me.Top = 2000
                 'Me.Height = 2500
                 Me.Height = 3300
             End If
             
             If (mCnn.State = 1) Then
                mCnn.Close
             End If
        End If
    
    End Sub


Private Sub cboSubType_Change()
'cmbDepartment.ListIndex = 0
'cmbSeat.ListIndex = 0
'cmbSeatID.Text = ""

End Sub

    Private Sub cboSubType_Click()
           'changed by soumya V S
        
      
     
        Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    'CHNAGED
    Dim i As Integer
    Dim J As Integer
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
        
     Call ShowFrames
 

        EnableControls
        
        'If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Or txtSubTypeID.Text = "148" Then
    If txtSubTypeID.Text = "2" Then
        'Label2.Caption = "Arrival Date"
        Label18.Caption = "Arrival Date"
    Else
        'Label2.Caption = "Application Date"
        Label18.Caption = "Event Date"
    End If
        
        
        'changed by soumya V S on 14.05.04
        getSubjectDeliverydate (txtSubTypeID.Text)
    ReDim arrIn(1)
    arrIn(0) = txtSubTypeID.Text
    arrIn(1) = txtWardNo.Text
   Set Rec = objdb.ExecuteSP("SpSelectSubTypeSeatCoding", arrIn, , , mCnn, adCmdStoredProc)
    'Set Rec = objDB.ExecuteSP("Sp_SelectHolidayList", arrIn, , , mCnn, adCmdStoredProc)
    'changed by soumya V S
    'CHANGED
     If Not (Rec.EOF Or Rec.BOF) Then
     'Label22.Caption = "Seat:" + Rec!chvSeatName
     'Label23.Caption = "User:" + Rec!chvUserNameEng
     If InwardMode = 0 Then
     frmUSoochikaInward.txtuserid = Rec!numUserID
     frmUSoochikaInward.txtseatid = Rec!numSeatID
     End If
      For i = 0 To cmbDepartment.ListCount - 1
            If (cmbDepartment.ItemData(i) = Rec!intDeptID) Then
                cmbDepartment.ListIndex = i
               ' Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & Rec!intDeptID, , True, True, True, enuSourceString.SoochikaUnicode)
                'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & Rec!intDeptID, , True, True, True, enuSourceString.SoochikaUnicode)
                
                
                'LATEST 24NOV
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                For J = 0 To cmbSeat.ListCount - 1
                    If (cmbSeatID.List(J) = Rec!numSeatID) Then
                    'chnaged by soumya V S
                        cmbSeatID.ListIndex = J
                        cmbSeat.ListIndex = J
                    End If
                Next
            End If
        Next
     
     

If (cmbSeat.ListIndex > 0) Then
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""

    Rec.Close
    End If
    End If
    'NOV18
     Else
       'If (cmbDepartment.ListIndex <> 0) Then
       
        'lblusername.Caption = ""
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
       ' Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
   
   FillCombo
  ' End If

End If
       'changed by soumya VS
       grvCheckList.Rows = 2
       grvCheckList.Clear 1
        FillEnclosureGrid val(txtSubTypeID.Text)
 
      

    End Sub
    
    Private Sub FillEnclosureGrid(ByVal SubID As Integer)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim arrOut As Variant
    Dim i As Integer
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    
    ReDim arrIn(1)
    arrIn(0) = 1
    arrIn(1) = SubID
    Set Rec = objdb.ExecuteSP("Sp_SelectSubTypeEnclosure", arrIn, arrOut, , mCnn, adCmdStoredProc)
    If IsArray(arrOut) Then
        For i = 0 To UBound(arrOut, 2)
            If i > 0 Then
                   grvCheckList.Rows = grvCheckList.Rows + 1
            End If
            grvCheckList.TextMatrix(i + 1, 3) = arrOut(1, i)
            grvCheckList.TextMatrix(i + 1, 2) = arrOut(0, i)
        Next i
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub



    Private Sub chkZonal_Click()
        ShowFrames
    End Sub

Private Sub cmbDepartment_Click()
 If (cmbDepartment.ListIndex <> 0) Then
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex), , True, True, True, enuSourceString.SoochikaUnicode)
        
        'add  by vipin 21-09-2012
        'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
      'LATEST 24Nov

        Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)

    End If
End Sub

Private Sub cmbseat_Click()
'changed by soumya
cmbSeatID.ListIndex = cmbSeat.ListIndex
End Sub

Private Sub cmbSeatID_Click()
Dim arrIn As Variant
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
            Exit Sub
        End If
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""
End If
End Sub

    Private Sub cmdClear_Click()
        cboRelationship.ListIndex = -1
        txtRegNo.Text = ""
        txtBookNo.Text = ""
        txtEnglishname.Text = ""
        txtMalayalamname.Text = ""
        
        cboRelationship.Enabled = True
        txtRegNo.Enabled = True
        txtBookNo.Enabled = True
        txtEnglishname.Enabled = True
        txtMalayalamname.Enabled = True
    End Sub
 Private Sub EnableControls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is Buttons Then
            ctl.Enabled = True
        End If
        
    Next ctl
End Sub
Public Sub DisableControls()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = False
            'ctl.Locked = True
        ElseIf TypeOf ctl Is DTPicker Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is CheckBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is VSFlexGrid Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is Buttons Then
            ctl.Enabled = False
        End If
      
    Next ctl
End Sub

    Private Sub cmdClose_Click()
        If InwardMode = 0 Then
            frmUSoochikaInward.txtSubID.Text = ""
            frmUSoochikaInward.txtSubject.Text = ""
            gbSevanaMainTypeID = 0
        Else
            frmUSoochikaManualInward.txtSubID.Text = ""
            frmUSoochikaManualInward.txtSubject.Text = ""
        End If
        Unload Me
    End Sub

    Private Sub cmdCopy_Click()
        Dim flag
'           Dim mCnnSoochika As New ADODB.Connection
'            Dim mCnnSevana As New ADODB.Connection
'            Dim InwNo As Variant
'            flag = 1
'            If Validate <> 0 Then
'                    objdb.CreateNewConnection mCnnSoochika, enuSourceString.SoochikaUnicode
'                    mCnnSoochika.BeginTrans
'
'                    If InwardMode = 0 Then
'                        InwNo = frmUSoochikaInward.SaveSoochika(mCnnSoochika)
'                    Else
'                        InwNo = frmUSoochikaManualInward.SaveSoochika(mCnnSoochika)
'                    End If
'                    frmUSoochikaInward.SaveAttachment (InwNo)
'                    mCnnSoochika.CommitTrans
'               End If
        
        
        If cmdCopy.Caption <> "Save" Then
            Call frmReceiptsCounter.CheckInterruptReceiptRequestStatus
If frmReceiptsCounter.InterruptedMode = False And InwardMode = 1 Then
                MsgBox "You have no authority to take receipt ", vbInformation, "receipt"
                Exit Sub
        End If
            If CopyValidation Then
                Call copyToReceipt
            End If
        Else
            mReceiptNo = txtReceiptSearch.Text
            'Added by sunil on 14.08.2012
            If gbSevanaMainTypeID = 5 Then
                txtNoCopeis.Text = val(txtNoOfCertificate.Text)
            End If
            cmdOK_Click
        End If
    End Sub
    
    Private Function CopyValidation() As Boolean
        On Error GoTo err:
            CopyValidation = False
            Dim mCount As Integer
            Dim flag As Boolean
            
            For mCount = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpChecked, mCount, 0) = vbChecked Then
                    flag = True
                End If
            Next
            
            If flag = False Then
                MsgBox "Please Select the Amount", vbInformation
                vsGrid.SetFocus
                Exit Function
            End If
            
            If SevanaTypeID = 2 Then
                If Validate = 0 Then
                    Exit Function
                End If
            End If
            
            
         If (gbSevanaMainTypeID = 1) Then
               If Validate = 0 Then
               Exit Function
               End If
            End If
            
            CopyValidation = True
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Sub cmdGetName_Click()
        
        Dim con5 As New ADODB.Connection
        Dim rs5 As New ADODB.Recordset
        Dim rs6 As New ADODB.Recordset
        Dim Qry
        Dim Qry1
        
        If (objdb.CreateNewConnection(con5, enuSourceString.SevanaRegn) = False) Then
            MsgBox "Sevena Connection Failed", vbDefaultButton1
            Exit Sub
        End If
        
        
        If Trim(txtRegNo) <> "" And Trim(txtBookNo.Text) <> "" Then 'If registration no empty
         If gbSevanaMainTypeID = 1 Then        'Birth
            If cboRelationship.ListIndex <> 2 Then
                Qry = "SELECT     chvMalFather, chvEngFather ,chvMalChild,chvEngChild From tBirthRep " _
                & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
                frmUSevanaInward.cboRelationship.ListIndex = 1
             Else
                Qry = "SELECT     chvMalMother, chvEngMother ,chvMalChild,chvEngChild From tBirthRep " _
                & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
                frmUSevanaInward.cboRelationship.ListIndex = 2
             End If
             'Added by vipin on 19.07.2012
             If IIf(IsNull(txtBookNo.Text), "", txtBookNo.Text) <> "" And IIf(IsNull(txtRegNo.Text), "", txtRegNo.Text) <> "" Then
                    Qry1 = "set dateformat dmy select BirthDate from BIRTHSEARCHVIEW1 where "
                    Qry1 = Qry1 + "  chvRegnNo like '" & txtRegNo.Text & "%' "
                    Qry1 = Qry1 + " and bookno = '" & IIf(IsNull(txtBookNo.Text), "", txtBookNo.Text) & "' "
                    rs6.Open Qry1, con5
                    Dim Age As Integer
                    If Not (rs6.EOF And rs6.BOF) Then
                        If IsNull(rs6!BirthDate) Then
                            Exit Sub
                        Else
                            Dim BirthDate As Integer
                            BirthDate = Year((rs6!BirthDate))
                            Age = Year(gbTransactionDate) - BirthDate
                             If (Age >= 6 And frmUSevanaInward.txtSubTypeID = 111) Then
                                MsgBox "Age > 6..!! pet name correction not Allowd."
                                Exit Sub
                            End If
                        End If
                   End If
            End If
               
             
         ElseIf gbSevanaMainTypeID = 2 Then   'Death
             Qry = "SELECT     chvMalDeadName, chvEngDeadName From tDeathRep " _
             & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
             frmUSevanaInward.cboRelationship.ListIndex = 0
         ElseIf gbSevanaMainTypeID = 3 Then   'Still birth
             Qry = "SELECT     chvMalFather, chvEngFather From tStillBirthRep " _
             & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
             frmUSevanaInward.cboRelationship.ListIndex = 0
             '--------------------Commented by savitha on 24.01.2008-------
         ElseIf gbSevanaMainTypeID = 4 Then   'Marriage
             '-----------modified on 15/10/08 by nisha ninan
            If cboRelationship.ListIndex = 0 Then
                 Qry = "SELECT     tMarriageMal.chvGroom,tMarriageEng.chvGroom FROM tMarriageEng INNER JOIN tMarriageMal ON tMarriageEng.chvAckNo = tMarriageMal.chvAckNo " _
                 & "WHERE     (tMarriageEng.chvRegnNo = '" & Trim(txtRegNo) & "')"
                 frmUSevanaInward.cboRelationship.ListIndex = 0
             Else
                 Qry = "SELECT     tMarriageMal.chvGroom,tMarriageEng.chvGroom FROM tMarriageEng INNER JOIN tMarriageMal ON tMarriageEng.chvAckNo = tMarriageMal.chvAckNo " _
                 & "WHERE     (tMarriageEng.chvRegnNo = '" & Trim(txtRegNo) & "')"
                 frmUSevanaInward.cboRelationship.ListIndex = 0
'                 Qry = "SELECT     tMarriageMal.chvBride,tMarriageEng.chvBride FROM tMarriageEng INNER JOIN tMarriageMal ON tMarriageEng.chvAckNo = tMarriageMal.chvAckNo " _
'                 & "WHERE     (tMarriageEng.chvRegnNo = '" & Trim(txtRegNo) & "')"
'                 frmUSevanaInward.cboRelationship.ListIndex = 0
             End If
             '---------added by nisha on 10/10/08
             ElseIf gbSevanaMainTypeID = 5 Then 'Common Marriage
             If cboRelationship.ListIndex = 0 Then
              '.............Modified by savitha on 27.01.2009 for commonMarriage
        
        '            Qry = "select tMarriageMalayalam.chvHusName as MalHus,tMarriageEnglish.chvHusName from tMarriageEnglish inner join tMarriageMalayalam on tMarriageEnglish.chvAckNo=tMarriageMalayalam.chvAckNo " _
        '            & " where  (tMarriageEnglish.chvRegnNo='" & Trim(txtRegistrationno) & "')"
                 
                 Qry = "select chvHusName from tMarriageEnglish  where  chvRegnNo='" & Trim(txtRegNo) & "'"
                 
                 frmUSevanaInward.cboRelationship.ListIndex = 0
             Else
        '            Qry = "select tMarriageMalayalam.chvWfeName as MalWfe,tMarriageEnglish.chvWfeName from tMarriageEnglish inner join tMarriageMalayalam on tMarriageEnglish.chvAckNo=tMarriageMalayalam.chvAckNo " _
        '            & " where  (tMarriageEnglish.chvRegnNo='" & Trim(txtRegistrationno) & "')"
               '  Qry = "select chvWfeName from tMarriageEnglish  where  chvRegnNo='" & Trim(txtRegNo) & "'"
                Qry = "select chvHusName from tMarriageEnglish  where  chvRegnNo='" & Trim(txtRegNo) & "'"
                 frmUSevanaInward.cboRelationship.ListIndex = 0
             End If
            End If
          
        rs5.Open Qry, con5
            If rs5.EOF Then 'If the regno is invalid
               MsgBox "Invalid Reg. No"
               txtRegNo.Text = ""
               txtBookNo.Text = ""
               txtRegNo.SetFocus
            Else
            '.........Modified by savitha on 27.01.2009 for CommonMarriage
            
              If gbSevanaMainTypeID = 5 Then
              Dim rsMal As New ADODB.Recordset
               Dim SQL As String
                     If cboRelationship.ListIndex = 0 Then
                               
                             SQL = "select isnull(tMarriageMalayalam.chvHusName,'\ðInbn«nñ') from tMarriageMalayalam inner join tMarriageEnglish on  tMarriageEnglish.chvackno=tMarriageMalayalam.chvackno where  tMarriageEnglish.chvRegnNo='" & Trim(txtRegNo) & "'"
                             rsMal.Open SQL, con5
                             If rsMal.EOF = False Then
                                  txtMalayalamname.Text = rsMal(0)
                             Else
                                  txtMalayalamname.Text = "\ðInbn«nñ"
                             End If
                      Else
             
                         SQL = "select isnull(tMarriageMalayalam.chvWfeName,'\ðInbn«nñ' from tMarriageMalayalam inner join tMarriageEnglish on  tMarriageEnglish.chvackno=tMarriageMalayalam.chvackno where  tMarriageEnglish.chvRegnNo='" & Trim(txtRegNo) & "'"
                         rsMal.Open SQL, con5
                             If rsMal.EOF = False Then
                                  txtMalayalamname.Text = rsMal(0)
                             Else
                                  txtMalayalamname.Text = "\ðInbn«nñ"
                             End If
                     End If
                 txtEnglishname.Text = rs5(0)
                 txtMalayalamname.Enabled = False
                 txtEnglishname.Enabled = False
                 txtRegNo.Enabled = False
                 txtBookNo.Enabled = False
                 cboRelationship.Enabled = False
              Else
              '.................................................
              'Commented & Added on 31.07.2009 by Sreeja----start
        '          txtEngCertName.Text = rs5(1)
        '          txtMalCertName.Text = rs5(0)
                
                '''----------------------------------------------
                If (frmUSevanaInward.txtSubTypeID = 113 Or frmUSevanaInward.txtSubTypeID = 111 Or frmUSevanaInward.txtSubTypeID = 114 Or frmUSevanaInward.txtSubTypeID = 115 Or frmUSevanaInward.txtSubTypeID = 116 Or frmUSevanaInward.txtSubTypeID = 117 Or frmUSevanaInward.txtSubTypeID = 118 Or frmUSevanaInward.txtSubTypeID = 119) Then
                    If IsNull(rs5(2)) And IsNull(rs5(3)) Then
                        MsgBox "Name Not Given"
                        Exit Sub
                    End If
                End If
                
                If (frmUSevanaInward.txtSubTypeID = 74 Or frmUSevanaInward.txtSubTypeID = 75 Or frmUSevanaInward.txtSubTypeID = 8 Or frmUSevanaInward.txtSubTypeID = 20 Or frmUSevanaInward.txtSubTypeID = 59) Then
'               If (rs5(2)) <> "Not Given" Then
'               If IsNull(rs5(2)) Then
'                    MsgBox "Child Name Already Given"
'                  frmUSevanaInward.cboRelationship.ListIndex = -1
'                Exit Sub
'                End If
'
'                'If (rs5(2)) <> "\ðInbn«nñ" Then
'                 MsgBox "Child Name Already Given"
'                 frmUSevanaInward.cboRelationship.ListIndex = -1
'                Exit Sub
'                End If

               If Not IsNull(rs5(2)) Or Not IsNull(rs5(3)) Then
                  MsgBox "Child Name Already Given"
                  frmUSevanaInward.cboRelationship.ListIndex = -1
                Exit Sub
               End If
               
                End If

                
                '''----------------------------------------------
                
                 txtEnglishname.Text = IIf(IsNull(rs5(1)), "Not Given", rs5(1))
                 txtMalayalamname.Text = IIf(IsNull(rs5(0)), "\ðInbn«nñ", rs5(0))
              '--------------------------------------------end
                'Modified by Arun A on 6.5.2006 for disabling Editing
                txtMalayalamname.Enabled = False
                txtEnglishname.Enabled = False
                txtRegNo.Enabled = False
                txtBookNo.Enabled = False
                cboRelationship.Enabled = False
               End If
            End If
        
        
        Else
        '.......................Modified by savitha on 24.01.2009
            If gbSevanaMainTypeID = 4 Or gbSevanaMainTypeID = 5 Then
                 MsgBox " Please Enter the Registration Number And the book Number ? ", vbInformation
                 txtRegNo.SetFocus
                 txtEnglishname.Text = ""
                 txtMalayalamname.Text = ""
                
            Else
            MsgBox " Please Enter the Registration Number And the book Number ? ", vbInformation
            'MsgBox "Enter the Reg.No"
            txtRegNo.SetFocus
            txtEnglishname.Text = ""
            txtMalayalamname.Text = ""
            txtBookNo.Text = ""
            End If
            '.......................................................
        End If 'Exits if reg no empty
        '--------------------------------------------------------------------
         If (con5.State = 1) Then
            con5.Close
         End If
    End Sub

    Private Sub cmdOK_Click()
        On Error GoTo err:
            
            '----------------------------------------'
            '----------------------------------------'
            '''If SevanaTypeID = 1 Or SevanaTypeID = 2 Then
            '''    If chkZonal.Value = 0 And txtSubTypeID.Text <> 110 Then
            '''        Call cmdCopy_Click
            '''        Exit Sub
            '''    End If
            '''End If
            '----------------------------------------'
            '----------------------------------------'
                      
            
            Dim flag
            Dim mCnnSoochika As New ADODB.Connection
            Dim mCnnSevana As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim InwNo As Variant
            Dim mSql As String
            'changed by soumya v S 29Sep
            Dim arrIn As Variant
            Dim ss As Variant
        
            flag = 1
            If Validate <> 0 Then
                objdb.CreateNewConnection mCnnSoochika, enuSourceString.SoochikaUnicode
                mCnnSoochika.BeginTrans
                On Error GoTo ErroRollBack:
                    If InwardMode = 0 Then
                    'changed by soumya vS on 13.08
                    'CHANGED
                        frmUSoochikaInward.dtpDeliveryDate.Value = dtpDeliveryDate1.Value
                       ' mSQl = "SELECT  numCurrentUserID From tSeatDetails WHERE numSeatID=" & cmbSeatID.Text
                       ' Set Rec = mCnnSoochika.Execute(mSQl)
                        'frmUSoochikaInward.txtuserid.Text = Rec!numCurrentUserID
                        InwNo = frmUSoochikaInward.SaveSoochika(mCnnSoochika)
                        'changed by soumya v S 29Sep
                                            
                        'If (Label22.Caption <> "") Then
                       'ss = frmUSoochikaInward.updateseat(mCnnSoochika)
                     
                        'Else
                        'frmUSoochikaInward.Label22.Caption = ""
                        ''frmUSoochikaInward.Label23.Caption = ""
                        'frmUSoochikaInward.txtuserid.Text = ""
                        'frmUSoochikaInward.txtseatid.Text = ""
                           
                        'End If
                    Else
                        InwNo = frmUSoochikaManualInward.SaveSoochika(mCnnSoochika)
                    End If
                   
                    objdb.CreateNewConnection mCnnSevana, enuSourceString.SevanaRegn
                    mCnnSevana.BeginTrans
                    If InwardMode = 0 Then
                        Call frmUSoochikaInward.SaveSevana(InwNo, SevanaTypeID, SevanaKioskID, 0, 0, mCnnSevana)
                    Else
                        Call frmUSoochikaManualInward.SaveSevana(InwNo, SevanaTypeID, SevanaKioskID, mReceiptNo, mReceiptAmt, mCnnSevana)
                    End If
                    'added by soumya vs on status doubling issue in tb_inward On 14 Oct 2016
                     If (SevanaTypeID <> 1) And (SevanaTypeID <> 2) Then
                            SaveSevanaStatus mCnnSevana, Right(InwNo, 6), Year(Now()), Now(), gbSevanaMainTypeID, txtSubTypeID.Text
                     End If
                
                mCnnSoochika.CommitTrans
                mCnnSevana.CommitTrans
                If InwardMode = 0 Then
                    MsgBox "Inward is saved with inward no of : " & Right(InwNo, 6), vbInformation, "SOOCHIKA"
                  '  frmUSoochikaInward.Ack (frmUSoochikaInward.lSoochikaFeildID)
                  '***************
                            'changed by soumya V S
                            mSql = "SELECT tLBSettings.flgAttachment FROM tLBSettings"
                            Set Rec = mCnnSoochika.Execute(mSql)
                            If (Rec.Fields(0) = "1") Then
                            frmUSoochikaInward.SaveAttachment (InwNo)
                            End If
                            frmUSoochikaInward.ShowAckReport (InwNo)
                            Unload frmUSevanaInward
                            frmUSoochikaInward.DisableControls
                            frmUSoochikaInward.cmdNew.Enabled = True
                            frmUSoochikaInward.cmdSave.Enabled = False
                            frmUSoochikaInward.cmdNew.SetFocus
                  
                  '***************
                   
                End If
                Unload Me
                'MsgBox " HAppy New Year"
                If InwardMode = 0 Then
                    frmUSoochikaInward.DisableControls
                    frmUSoochikaInward.cmdNew.Enabled = True
                    frmUSoochikaInward.cmdSave.Enabled = False
                    frmUSoochikaInward.cmdNew.SetFocus
                Else
                    frmUSoochikaManualInward.DisableControls
                    frmUSoochikaManualInward.cmdNew.Enabled = True
                    frmUSoochikaManualInward.cmdSave.Enabled = False
                    frmUSoochikaManualInward.cmdNew.SetFocus
                End If
            End If
        Exit Sub
err:
        MsgBox (Error$)
        Exit Sub
ErroRollBack:
        MsgBox (Error$)
        If mCnnSoochika.State Then
            mCnnSoochika.RollbackTrans
        End If
        
        If mCnnSevana.State Then
            mCnnSevana.RollbackTrans
        End If
    End Sub
Private Sub SaveSevanaStatus(mCnn As ADODB.Connection, InwNo As Variant, Year As Variant, Dt As Variant, mID As Variant, SID As Variant)
    Dim Sevarr As Variant
    ReDim Sevarr(4)
    Sevarr(0) = InwNo
    Sevarr(1) = Year
    Sevarr(2) = Dt
    Sevarr(3) = mID
    Sevarr(4) = SID
   objdb.ExecuteSP "sp_insertinwardstatusSoochika", Sevarr, , , mCnn, adCmdStoredProc
End Sub
    Public Function Validate()
    'NoV18
    Dim strDate As Date
        Dim flag
        flag = 1
    
       'Nov18
           If (dtpDeliveryDate1.Value > 1) Then
           strDate = Format(Date, "dd/Mm/yyyy")
           If (dtpDeliveryDate1.Value < strDate) Then
           MsgBox "Delivery date should be greater than today !!!", vbInformation, "SOOCHIKA"
           flag = 0
           dtpDeliveryDate1.SetFocus
           GoTo last
           ElseIf (CheckHoliday(dtpDeliveryDate1.Value) = True) Then
           MsgBox "The selected delivery date is holiday !!!", vbInformation, "SOOCHIKA"
           flag = 0
           dtpDeliveryDate1.SetFocus
           GoTo last
           End If
        'CHNAGED
        ElseIf cmbDepartment.ListIndex = -1 Then
         flag = 0
         MsgBox "Please select the Department", vbInformation, "SOOCHIKA"
         cmbDepartment.SetFocus
         GoTo last
        ElseIf cmbSeat.ListIndex = -1 Then
         flag = 0
         MsgBox "Please select the Seat", vbInformation, "SOOCHIKA"
         cmbSeat.SetFocus
         GoTo last
        ElseIf txtSubTypeID.Text = "" Then
            flag = 0
            MsgBox "Enter SubType", vbDefaultButton1
            txtSubTypeID.SetFocus
            GoTo last
        ElseIf cboSubType.ListIndex < 0 Then
            flag = 0
            MsgBox "select subtype", vbInformation
            cboSubType.SetFocus
            GoTo last
        ElseIf DTPApplDate.Value = 0 Then
            flag = 0
            MsgBox "select the Application/Arrival Date", vbDefaultButton1
            DTPApplDate.SetFocus
            GoTo last
        End If
        If SevanaTypeID = 2 Then
            If (gbLBID <> 167) Then
                If txtSubTypeID.Text < 76 Or txtSubTypeID.Text > 79 Then      ' Modified on 27.03.2010 Demaded By Arun Adoor
                    If txtRegNo.Text = "" Then
                        flag = 0
                        MsgBox "Please Enter Registration Number", vbInformation
                        txtRegNo.SetFocus
                        GoTo last
                    End If
                    
                    If txtBookNo.Text = "" Then
                        flag = 0
                        MsgBox "Please Enter Book Number", vbInformation
                        txtBookNo.SetFocus
                        GoTo last
                    End If
                    
                    If txtEnglishname.Text = "" And txtMalayalamname.Text = "" Then
                        flag = 0
                        MsgBox "Please Click GetName to Search the Names", vbInformation
                        cmdGetName.SetFocus
                        GoTo last
                    End If
                    If cboRelationship.ListIndex < 0 Then
                        flag = 0
                        MsgBox "Select Relationship", vbInformation
                        cboRelationship.SetFocus
                        GoTo last
                    ElseIf txtEnglishname.Text = "" Then
                        flag = 0
                        MsgBox "Searching not successfull,pls make research"
                        txtEnglishname.SetFocus
                        GoTo last
                    Else
                        If txtSubTypeID.Text = 110 Then
                            If txtRemarks.Text = "" Then
                                flag = 0
                                MsgBox "Please enter remarks", vbInformation
                                txtRemarks.SetFocus
                                GoTo last
                            End If
                        End If
                    End If
                End If
            End If
        End If
          
'If (ValidateEnclosure() = False) Then
        'MsgBox "Please select any Enclosures", vbInformation, "SOOCHIKA"
         'flag = 0
        'SSTab1.Tab = 0
        'GoTo last
        
last:         Validate = flag

  ' End If
    End Function

Private Sub cmdReceiptSearch_Click()
    frmSearchVouchers.Show vbModal
    If gbSearchID <> -1 Then
        GetVoucherDetails (gbSearchID)
    End If
End Sub
 Public Function GetVoucherDetails(ByVal intVoucherID As Long) As Boolean
        On Error GoTo err:
            Dim mSql As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim objdb As New clsDB
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * from faVouchers Where intVoucherID = " & intVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtReceiptSearch.Tag = Rec!intVoucherID
                    txtReceiptSearch.Text = Rec!intVoucherNo
                    lblTotal.Caption = Rec!fltAmount
                    mReceiptAmt = CDbl(lblTotal.Caption)
                    mReceiptNo = txtReceiptSearch.Text
                
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function

    Private Sub cmdsearch_Click()
    
        '--------------------------------------------'
        
        If txtRegNo.Text <> "" And txtBookNo.Text <> "" Then
            cmdGetName_Click
            Exit Sub
        End If
        
        '--------------------------------------------'
    
    
        If gbSevanaMainTypeID = 1 Then
            frmSevanaBirthSearch.Show vbModal           'Birth Search
        ElseIf gbSevanaMainTypeID = 2 Then
            frmSevanadethsearch.Show vbModal            'Death Search
        ElseIf gbSevanaMainTypeID = 3 Then
            frmSevanaStillBirth.Show vbModal            'Still Birth Search
        ElseIf gbSevanaMainTypeID = 4 Then
            frmSevanaMarriageSearch.Show vbModal        'Marriage Search
        ElseIf gbSevanaMainTypeID = 5 Then
            frmSevanaCommonMarriageSearch.Show vbModal  'Common Marriage Search
        End If
        
    End Sub

Private Sub dtpEventDate_LostFocus()
'Added dy Syalima on 28/07/2018
    If txtSubTypeID.Text = 3 Then
        If CDate(dtpEventDate.Value) + 60 > Date Then
            MsgBox "Arrival date is less than 60 days.. Use Normal Registration, born outside the country...", vbInformation, "SOOCHIKA"
            txtSubTypeID.Text = 2
'        Else
'            MsgBox "Arrival date is greater than 60 days.. Use Delayed reporting, born outside the country (Within 1 year or after 1 year)...", vbInformation, "SOOCHIKA"
        End If
    End If
End Sub

    Private Sub Form_Activate()
        If gbSevanaMainTypeID = 0 Then
            Unload Me
        End If
       If txtSubTypeID.Text = "" Then
            'dtpEventDate.value = Null
            dtpEventDate.Value = Format(Now, "dd/MM/yyyy")
            Me.Left = 2250
            Me.Top = 2000
            'Me.Height = 2500
            Me.Height = 3300
        Else
        
            ShowFrames
        End If
    End Sub

    Public Sub Form_Load()
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
       
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection not present", vbDefaultButton1, "SOOCHIKA"
            Exit Sub
        End If
                
        DTPApplDate.Value = Date
        DTPReceiptDate.Value = Date
        PopulateList cboSubType, "Select TypeofSubRequest,intID from mSubjectSevanaSubtype where intsubTypeID='" & gbSevanaMainTypeID & "'", , , , True, enuSourceString.SoochikaUnicode
        If gbSevanaMainTypeID = 4 Or gbSevanaMainTypeID = 5 Then
            Label3.Visible = False
            cboHospitals.Visible = False
        Else
            Label3.Visible = True
            cboHospitals.Visible = True
            PopulateList cboHospitals, "CBOSelectHospital", , True, , True, enuSourceString.SevanaRegn
        End If
        PopulateList cboRelationship, "select chvdescription,intid from mCertificateOwners where intregtype=" & gbSevanaMainTypeID, , , , True, enuSourceString.SevanaRegn
        'NOV18
        'Call PopulateList(cmbDepartment, "SP_SelectDepartment 1", , True, True, True, enuSourceString.SoochikaUnicode)
        
        FillCombo
        cboLanguage.Clear
        cboLanguage.AddItem "Malayalam"
        cboLanguage.ItemData(cboLanguage.NewIndex) = 1
        cboLanguage.AddItem "English"
        cboLanguage.ItemData(cboLanguage.NewIndex) = 2
        cboLanguage.ListIndex = 1
        
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
        If InwardMode = 0 Then
            frameReceiptSearch.Visible = False
        End If
             grvCheckList.Rows = 2
       grvCheckList.Clear 1
      'CHNAGED
      'NOV18
   ' cmbDepartment.ListIndex = 0
    'cmbSeat.Clear
    'cmbSeatID.Clear

    End Sub

Private Sub FillCombo()
 Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant

    Dim i As Integer
    Dim J As Integer
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    Call PopulateList(cmbDepartment, "SP_SelectDepartment 1", , True, True, True, enuSourceString.SoochikaUnicode)
    'ReDim arrIn(1)
    'arrIn(0) = frmUSoochikaInward.txtSubID.Text
    'arrIn(1) = frmUSoochikaInward.txtwardno.Text
   'Set Rec = objDb.ExecuteSP("SpSelectSubjectSeatCoding", arrIn, , , mCnn, adCmdStoredProc)
   
     'If Not (Rec.EOF Or Rec.BOF) Then

     'frmUSoochikaInward.txtuserid = Rec!numUserID
     'frmUSoochikaInward.txtseatid = Rec!numSeatID
     'changed by soumya vs on Dec2
     If InwardMode = 0 Then
      For i = 0 To cmbDepartment.ListCount - 1
            If (cmbDepartment.ItemData(i) = frmUSoochikaInward.Text2.Text) Then
                cmbDepartment.ListIndex = i
                'LATEST 24Nov
                'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & frmUSoochikaInward.Text2.Text, , True, True, True, enuSourceString.SoochikaUnicode)
                'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & frmUSoochikaInward.Text2.Text, , True, True, True, enuSourceString.SoochikaUnicode)
                
                
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)

                For J = 0 To cmbSeat.ListCount - 1
                    If (cmbSeatID.List(J) = frmUSoochikaInward.Text1.Text) Then
                    'chnaged by soumya V S
                        cmbSeatID.ListIndex = J
                        cmbSeat.ListIndex = J
                    End If
                Next
            End If
        Next
     
     Else
     For i = 0 To cmbDepartment.ListCount - 1
            If (cmbDepartment.ItemData(i) = frmUSoochikaManualInward.Text2.Text) Then
                cmbDepartment.ListIndex = i
                'LATEST 24Nov
                'Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & frmUSoochikaInward.Text2.Text, , True, True, True, enuSourceString.SoochikaUnicode)
                'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & frmUSoochikaInward.Text2.Text, , True, True, True, enuSourceString.SoochikaUnicode)
                
                
                Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and tSeatDetails.numCurrentUserID is not null and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
                Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 99 and tUserdetails.tnySuiteActive=0 and tUserDetails.tnyActive=0 and  tSeatDetails.numCurrentUserID is not null  and tSeatDetails.intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)

                For J = 0 To cmbSeat.ListCount - 1
                    If (cmbSeatID.List(J) = frmUSoochikaManualInward.Text1.Text) Then
                    'chnaged by soumya V S
                        cmbSeatID.ListIndex = J
                        cmbSeat.ListIndex = J
                    End If
                Next
            End If
        Next
     End If

If (cmbSeat.ListIndex > 0) Then
ReDim arrIn(0)
arrIn(0) = cmbSeatID.Text
Set Rec = objdb.ExecuteSP("spSelectUser", arrIn, , , mCnn, adCmdStoredProc)
If Not (Rec.EOF Or Rec.BOF) Then
lblusername.Caption = Rec!chvUserNameEng
Else
lblusername.Caption = ""

    Rec.Close
    End If
    End If
    
    ' Else
       'If (cmbDepartment.ListIndex <> 0) Then
       
        'lblusername.Caption = ""
       ' Call PopulateList(cmbSeatID, "select numSeatID,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
        'Call PopulateList(cmbSeat, "select chvSeatname,chvSeatname from tSeatDetails where numCurrentUserID is not null and intDeptID=" & cmbDepartment.ItemData(cmbDepartment.ListIndex) & "order by chvSeatname", , True, True, True, enuSourceString.SoochikaUnicode)
   ' End If

'End If
       
     
End Sub

Private Sub optInterruptReceipt_Click()
    If optInterruptReceipt.Value = True Then
        txtReceiptSearch.Text = ""
        cmdReceiptSearch.Enabled = False
        cmdCopy.Caption = "Copy to Receipt"
    End If
End Sub

Private Sub optPayedReceipt_Click()
    If optPayedReceipt.Value = True Then
        cmdCopy.Caption = "Save"
        cmdReceiptSearch.Enabled = True
    End If
End Sub

    Private Sub txtBookNo_Change()
'        txtEnglishname.Text = ""
'        txtMalayalamname.Text = ""
    End Sub

    Private Sub txtBookNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub

    Private Sub txtNoCopeis_Change()
'        If txtReceiptAmount.Text <> "" And gbSevanaMainTypeID = 5 Then
'            txtReceiptAmount.Text = Val(CommMarriageFee) * Val(txtNoCopeis.Text)
'        End If
    End Sub
    
    Private Sub txtNoCopeis_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
        
    Private Sub txtNoOfCertificate_Change()
'        If vsGrid.TextMatrix(1, 6) = "" Then Exit Sub
'        If gbSevanaMainTypeID = 1 Or gbSevanaMainTypeID = 2 Or gbSevanaMainTypeID = 3 Or gbSevanaMainTypeID = 4 Then
'            If txtNoOfCertificate.Text <> "" Then
'                vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6) * val(txtNoOfCertificate.Text)
'            Else
'                vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6)
'            End If
'        ElseIf gbSevanaMainTypeID = 5 Then
'            If txtNoOfCertificate.Text <> "" Then
'                vsGrid.TextMatrix(6, 7) = vsGrid.TextMatrix(6, 6) * val(txtNoOfCertificate.Text)
'            Else
'                vsGrid.TextMatrix(6, 7) = vsGrid.TextMatrix(6, 6)
'            End If
'        End If

  '  Call CalculateAmount
  
    Dim intNoOfCer As Integer
    Dim Amount As Integer
    If txtNoOfCertificate.Text = "" Then
        intNoOfCer = 1
    Else
     intNoOfCer = val(txtNoOfCertificate.Text)
    End If
    If intNoOfCer <> 0 Then
       Amount = intNoOfCer * val(vsGrid.TextMatrix(1, 6))
       vsGrid.TextMatrix(1, 7) = Amount
'       If mCurRow = 3 Then
'        If val((vsGrid.TextMatrix(3, 6))) <> 0 Then
'             amount = val(vsGrid.TextMatrix(3, 6))
'             vsGrid.TextMatrix(3, 7) = amount
'        End If
'       End If
    End If
    Call Calculate
    End Sub
    
    Private Sub txtNoOfCertificate_ChangeoLD()
        If vsGrid.TextMatrix(1, 6) = "" Then Exit Sub
        If txtNoOfCertificate.Text <> "" Then
            vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6) * val(txtNoOfCertificate.Text)
        Else
            vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6)
        End If
    End Sub

    Private Sub txtNoofYears_Change()
'        If vsGrid.TextMatrix(2, 6) = "" Then Exit Sub
'        If txtNoofYears.Text <> "" Then
'            vsGrid.TextMatrix(2, 7) = vsGrid.TextMatrix(2, 6) * val(txtNoofYears.Text)
'        Else
'            vsGrid.TextMatrix(2, 7) = vsGrid.TextMatrix(2, 6)
'        End If
    CalculateAmount
    Call Calculate
    End Sub
    
    Private Sub txtReceiptAmount_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtReceiptBookNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    
Private Sub txtReceiptSearch_Change()
    If (txtReceiptSearch.Text <> "") Then
        cmdCopy.Caption = "Save"
    Else
        cmdCopy.Caption = "Copy to Receipt"
    End If
End Sub

    Private Sub txtRegNo_Change()
'        txtEnglishname.Text = ""
'        txtMalayalamname.Text = ""
    End Sub

    Private Sub txtSubTypeID_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub

    '''Public Sub GetCommonMarriageFee(ByVal SevanaSubID As Variant)
    '''    Select Case SevanaSubID
    '''        Case 89, 90
    '''            frmUSevanaInward.txtReceiptAmount = "10"
    '''        Case 91
    '''            frmUSevanaInward.txtReceiptAmount = "110"
    '''        Case 92
    '''            frmUSevanaInward.txtReceiptAmount = "260"
    '''        Case 93
    '''            frmUSevanaInward.txtReceiptAmount = "5"
    '''        Case 96
    '''            frmUSevanaInward.txtReceiptAmount = "100"
    '''        Case 98
    '''            frmUSevanaInward.txtReceiptAmount = "25"
    '''        Case 99
    '''            frmUSevanaInward.txtReceiptAmount = "15"
    '''        Case 100
    '''            frmUSevanaInward.txtReceiptAmount = "15"
    '''        Case 101
    '''            frmUSevanaInward.txtReceiptAmount = "115"
    '''        Case 102
    '''            frmUSevanaInward.txtReceiptAmount = "265"
    '''        Case 103
    '''            frmUSevanaInward.txtReceiptAmount = "25"
    '''        Case 104
    '''            frmUSevanaInward.txtReceiptAmount = "125"
    '''        Case 94, 95, 97
    '''            frmUSevanaInward.txtReceiptAmount = ""
    '''    End Select
    '''    CommMarriageFee = frmUSevanaInward.txtReceiptAmount.Text
    '''End Sub

Private Sub txtSubTypeID_LostFocus()
    Dim flag
    Dim i As Integer
    
    flag = 0
    If txtSubTypeID.Text <> "" Then
        For i = 0 To cboSubType.ListCount - 1
            If val(txtSubTypeID.Text) = cboSubType.ItemData(i) Then
                cboSubType.ListIndex = i
                flag = 1
            End If
        Next
        If flag <> 1 Then
            MsgBox "Item not found", vbDefaultButton1
        End If
    End If
    'If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Or txtSubTypeID.Text = "148" Then
    If txtSubTypeID.Text = "2" Then
        'Label2.Caption = "Arrival Date"
        Label18.Caption = "Arrival Date"
    Else
        'Label2.Caption = "Application Date"
        Label18.Caption = "Event Date"
    End If
    If flag = 1 Then
        Call ShowFrames
    End If
End Sub
Private Sub FillGrid()
    On Error GoTo err:
        Dim mSql As String
        
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mRowCount As Integer
        
        Dim mScheduleType As Integer
        Dim mScheduleSubID As Integer
        
        
        Dim mFunctionId As Integer
        Dim mFunctionaryID As Integer
        
        If gbSevanaMainTypeID = 1 Or gbSevanaMainTypeID = 3 Then 'Birth
            mScheduleType = 13
            mScheduleSubID = 1
            intTransactionTypeID = gbTransactionTypeBrith
            
            mFunctionId = 46
            mFunctionaryID = 7
        ElseIf gbSevanaMainTypeID = 2 Then     'Death
            mScheduleType = 13
            mScheduleSubID = 2
            intTransactionTypeID = gbTransactionTypeDeath
            
            mFunctionId = 46
            mFunctionaryID = 7
'        ElseIf gbSevanaMainTypeID = 4 Then     'Marriage
'            mScheduleType = 15
'            mScheduleSubID = 3
'            intTransactionTypeID = gbTransactionTypeMarriage
'
'            mFunctionID = 47
'            mFunctionaryID = 7

        ElseIf gbSevanaMainTypeID = 5 Then     'CmnMarriage
            mScheduleType = 14
            mScheduleSubID = 3
            intTransactionTypeID = gbTransactionTypeCmnMarriage
            
            mFunctionId = 47
            mFunctionaryID = 7
        ElseIf gbSevanaMainTypeID = 4 Then     'Marriage
            mScheduleType = 15
            mScheduleSubID = 3
            intTransactionTypeID = gbTransactionTypeMarriage
            
            mFunctionId = 47
            mFunctionaryID = 7
        End If
        
        txtNoOfCertificate.Text = 1
       ' txtNoofYears.Text = 1
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
'            If InwardMode = 0 Then
'                If frmUSoochikaInward.chkBPL.value = 1 Or frmUSoochikaInward.chkSCST.value = 1 Then
'                    mSql = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltSpecialRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
'                Else
'                    mSql = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltFixedRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
'                End If
'            Else
'                If frmUSoochikaManualInward.chkBPL.value = 1 Or frmUSoochikaManualInward.chkSCST.value = 1 Then
'                    mSql = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltSpecialRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
'                Else
'                    mSql = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltFixedRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
'                End If
'            End If
'
'            mSql = mSql + " FROM         smScheduleMasters INNER JOIN "
'            mSql = mSql + " smAttributes ON smScheduleMasters.intAttributeID = smAttributes.intAttributeID "
'            mSql = mSql + " WHERE     (smScheduleMasters.intScheduleID = " & mScheduleType & ") and smAttributes.tnyGroupID = " & mScheduleSubID & " and smAttributes.intAttributeID <> 154  and smAttributes.intAttributeID <> 158 order by smAttributes.intAttributeID"   'Modified on 06.04.2009 by Suby / Modified on 17.04.2009
'            Rec.Open mSql, mCnn
'            mRowCount = 1
'            vsGrid.Rows = 2
'            While Not Rec.EOF And Not Rec.BOF
'                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
'                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
'                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchAttributeTitle), "", Rec!vchAttributeTitle)
'                If (InwardMode = 0) Then
'                    If frmUSoochikaInward.chkBPL.value = 1 Or frmUSoochikaInward.chkSCST.value = 1 Then
'                        vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                        vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                    Else
'                        vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                        vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                    End If
'                Else
'                    If frmUSoochikaManualInward.chkBPL.value = 1 Or frmUSoochikaManualInward.chkSCST.value = 1 Then
'                        vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                        vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
'                    Else
'                        vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                        vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                        vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
'                    End If
'                End If
'                vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intAttributeID), "", Rec!intAttributeID) 'Added on 03.04.2009
'                mRowCount = mRowCount + 1
'                vsGrid.Rows = vsGrid.Rows + 1
'                Rec.MoveNext
'            Wend
'            'Added on 17.04.2009 by Suby---
''            If gbSevanaMainTypeID= 5 Then
''                vsGrid.RowHidden(2) = True
''            End If
'            '------------------------------
'            If Rec.State = 1 Then Rec.Close
'            mSql = "Select * from smAttributeSevanaMapping Where intSevanaSubTypeID = " & val(txtSubTypeID.Text)
'            Rec.Open mSql, mCnn
'            While Not (Rec.EOF Or Rec.BOF)
'                For mRowCount = 1 To vsGrid.Rows - 1
'                    If vsGrid.TextMatrix(mRowCount, 12) = Rec!intAttributeID Then
'                        vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked
'                        Call vsGrid_AfterEdit(mRowCount, 0)
'                    End If
'                Next
'                Rec.MoveNext
'            Wend

        'Modified by Sunil On 03-jul-2012

'        mSql = "Select faSubTypes.intSubTypeID,vchSubTypeTitle,intAccountHeadID,vchAccountHeadCode,fltRate,* from faSubTypes"
'        mSql = mSql + " Inner Join faSubtypeSchedule ON faSubTypes.intSubTypeID=faSubTypeSchedule.intSubtypeID"
'        mSql = mSql + " Where intSevanSubTypeID =" & val(txtSubTypeID.Text) & "   And intSubTransactionType = " & mScheduleSubID
        txtNoofYears.Text = ""
        txtNoOfCertificate.Text = ""
        If txtSubTypeID.Text = 60 Then
            txtNoofYears.Text = 1
            txtNoOfCertificate.Enabled = False
        End If
        txtNoofYears.Enabled = True
        mSql = "Select faSubTypes.intSubTypeID,vchAlias,intSlNo,tnyMultipleFlag,fltrate,intAccountHeadID,vchAccountHeadCode from faSubTypes"
        mSql = mSql + " Inner join faSubTypeSchedule on faSubTypes.intSubTypeID=faSubTypeSchedule.intSubTypeID"
        mSql = mSql + " Where intSevanSubTypeID =" & val(txtSubTypeID.Text) '& "   And intSubTransactionType = " & mScheduleSubID
        If gbSevanaMainTypeID = 5 Then
            If InwardMode = 0 Then
                 If frmUSoochikaInward.chkBPL.Value = 1 Or frmUSoochikaInward.chkSCST.Value = 1 Then
                    mSql = mSql + " And  intSubTypeCategoryID=2"
                Else
                    mSql = mSql + " And  intSubTypeCategoryID=1"
                End If
            End If
            '---Added on 14.08.2012 by Sunil
            If InwardMode = 1 Then
                 If frmUSoochikaManualInward.chkBPL.Value = 1 Or frmUSoochikaManualInward.chkSCST.Value = 1 Then
                 
                    mSql = mSql + " And  intSubTypeCategoryID=2"
                Else
                    mSql = mSql + " And  intSubTypeCategoryID=1"
                End If
            End If
        Else
                mSql = mSql + " And  intSubTypeCategoryID=1"
        End If
        
        Rec.Open mSql, mCnn
        mRowCount = 1
  
        While Not Rec.EOF And Not Rec.BOF
        mCurRow = mRowCount
        vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
        vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias)
        vsGrid.TextMatrix(mRowCount, 15) = IIf(IsNull(Rec!intSlNo), "", Rec!intSlNo)
        
        If Rec!tnyMultipleFlag = 1 And val(txtSubTypeID.Text) <> 114 And val(txtSubTypeID.Text) <> 119 And val(txtSubTypeID.Text) <> 124 Then
        
            If val(txtSubTypeID.Text) = 134 Or val(txtSubTypeID.Text) = 135 Or val(txtSubTypeID.Text) = 99 Or val(txtSubTypeID.Text) = 100 Or _
               val(txtSubTypeID.Text) = 136 Or val(txtSubTypeID.Text) = 137 Then
                 vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                 vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                 vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                 txtNoofYears.Enabled = False
            Else
                vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                If IIf(IsNull(Rec!intSlNo), 0, Rec!intSlNo) <> 3 Then
                    vsGrid.TextMatrix(mRowCount, 7) = 0
                Else
                    vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                End If
                vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
            End If
        Else
            If (InwardMode = 0) Then
                If frmUSoochikaInward.chkBPL.Value = 1 Or frmUSoochikaInward.chkSCST.Value = 1 Then
                            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                Else
                            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
               End If
            Else
               If frmUSoochikaManualInward.chkBPL.Value = 1 Or frmUSoochikaManualInward.chkSCST.Value = 1 Then
                            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                              vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
               Else
                            vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                            vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
                              vsGrid.TextMatrix(mRowCount, 16) = IIf(IsNull(Rec!fltRate), "", Rec!fltRate)
               End If
        End If
        End If
'''       vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intSubTypeID), "", Rec!intSubTypeID)
'''       mRowCount = mRowCount + 1
'''       vsGrid.Rows = vsGrid.Rows + 1
'''       Rec.MoveNext

       vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked
       vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intSubTypeID), "", Rec!intSubTypeID)
       mRowCount = mRowCount + 1
       vsGrid.Rows = vsGrid.Rows + 1
       Call CertificateorNot
       Rec.MoveNext
       Call Calculate
   Wend
        
        Else
            MsgBox "Connection To iSaankhyaMasters does not exist, Please Contact your System Administrator", vbInformation
        End If
    Exit Sub
err:
    MsgBox (Error$)
End Sub

Private Function copyToReceipt() As Boolean
    On Error GoTo err:
        Dim objTrType As New clsTransactionType
        Dim objAcc As New clsAccounts
        Dim mLoop As Integer
        Dim mRowCnt As Integer
        Dim mText As String
        Dim i As Integer
        
        Dim mTotal As Variant
        
        mRowCnt = 0
        For mLoop = 0 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 0) = vbChecked Then
                mRowCnt = mRowCnt + 1
                End If
        Next
        
        If gbSevanaMainTypeID = 5 Then
            txtNoCopeis.Text = val(txtNoOfCertificate.Text)
'            If vsGrid.Cell(flexcpChecked, 6, 0) = vbChecked Then
'                For i = 7 To 11
'                    If vsGrid.Cell(flexcpChecked, i, 0) = vbChecked Then
'                        txtNoCopeis.Text = val(txtNoCopeis.Text) + 1
'                    End If
'                Next
'            End If
        End If
        Me.Hide
        Load frmReceiptsCounter
        
        
        
        objTrType.SetTransactionType (intTransactionTypeID)
    
         frmReceiptsCounter.vsGrid.Editable = flexEDNone
        frmReceiptsCounter.SoochikaConnected = True
    
        frmReceiptsCounter.txtTransactionType.Text = objTrType.TransactionType
        frmReceiptsCounter.txtTransactionType.Tag = intTransactionTypeID
        frmReceiptsCounter.SubLedgerID = 9999   ' Test Value    '
'        frmReceiptsCounter.cmbZone.Text = gbnumZonalID
'        frmReceiptsCounter.cmbDZone.Text = gbnumZonalID '   Added   '
        'frmReceiptsCounter.txtWard.Text = frmUSoochikaInward.txtWardNo.Text
        If InwardMode = 0 Then
            frmReceiptsCounter.txtWardNo.Text = frmUSoochikaInward.txtWardNo.Text
            frmReceiptsCounter.txtWard.Tag = frmUSoochikaInward.txtWardNo.Text
            frmReceiptsCounter.txtHouseNo1.Text = frmUSoochikaInward.txtDoorNo1.Text
            frmReceiptsCounter.txtHouseNo2.Text = frmUSoochikaInward.txtDoorNo2.Text
            frmReceiptsCounter.txtDoorNo1.Text = frmUSoochikaInward.txtDoorNo1.Text
            frmReceiptsCounter.txtDoorNo2.Text = frmUSoochikaInward.txtDoorNo2.Text
            frmReceiptsCounter.txtName.Text = frmUSoochikaInward.txtApplicantName.Text
            frmReceiptsCounter.txtMainPlace.Text = frmUSoochikaInward.txtMainPlace.Text
            frmReceiptsCounter.txtLocalPlace.Text = frmUSoochikaInward.txtLocalPlace.Text
            'frmReceiptsCounter.txtHouseName.Text = frmUSoochikaInward.txtHouseName.Text
        Else
            frmReceiptsCounter.InterruptedModeSoochika = True
            frmReceiptsCounter.txtWardNo.Text = frmUSoochikaManualInward.txtWardNo.Text
            frmReceiptsCounter.txtWard.Tag = frmUSoochikaManualInward.txtWardNo.Text
            frmReceiptsCounter.txtHouseNo1.Text = frmUSoochikaManualInward.txtDoorNo1.Text
            frmReceiptsCounter.txtHouseNo2.Text = frmUSoochikaManualInward.txtDoorNo2.Text
            frmReceiptsCounter.txtDoorNo1.Text = frmUSoochikaManualInward.txtDoorNo1.Text
            frmReceiptsCounter.txtDoorNo2.Text = frmUSoochikaManualInward.txtDoorNo2.Text
            frmReceiptsCounter.txtName.Text = frmUSoochikaManualInward.txtApplicantName.Text
        End If
        mText = ""
        If txtRegNo.Text <> "" Then
            mText = "RegNo:" + CStr(txtRegNo.Text) + " "
        End If
        
        If txtBookNo.Text <> "" Then
            mText = mText + "BookNo:" + CStr(txtBookNo.Text)
        End If
        frmReceiptsCounter.txtDescription.Text = mText
        
        frmReceiptsCounter.vsGrid.Rows = mRowCnt + 1
        
        mRowCnt = 1
        mTotal = 0
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, mLoop, 0) = vbChecked Then
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 0) = vsGrid.TextMatrix(mLoop, 2)
                objAcc.SetAccountCode (vsGrid.TextMatrix(mLoop, 2))
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 0) = objAcc.AccountCode
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 1) = objAcc.AccountHead
                
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 5) = val(vsGrid.TextMatrix(mLoop, 7))
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 6) = objAcc.AccountHeadID
                
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 7) = gbFinancialYearID
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 8) = 1
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 9) = 0
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 10) = 0
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 11) = val(vsGrid.TextMatrix(mLoop, 7))
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 12) = 0
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 13) = 0
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 14) = 0
                frmReceiptsCounter.vsGrid.TextMatrix(mRowCnt, 15) = 0
                mRowCnt = mRowCnt + 1
                mTotal = mTotal + val(vsGrid.TextMatrix(mLoop, 7))
           End If
        Next
        frmReceiptsCounter.txtTotalCurrent.Text = mTotal
        frmReceiptsCounter.txtTotal.Text = mTotal
        frmReceiptsCounter.txtGrandTotal.Text = mTotal
        'CHNAGED
        
        If (gbSevanaMainTypeID = 0) Then
         
            If InwardMode = 0 Then
             
                frmReceiptsCounter.cmbSeat.Text = frmUSoochikaInward.cmbSeat.Text
            Else
                frmReceiptsCounter.cmbSeat.Text = frmUSoochikaManualInward.cmbSeat.Text
            End If
        Else
           frmReceiptsCounter.cmbSeat.Text = cmbSeat.Text
        End If
        If IsDate(frmUSoochikaInward.dtpDeliveryDate.Value) Then
        'changed
            frmReceiptsCounter.txtDescription.Text = frmReceiptsCounter.txtDescription.Text & "|Delivery Date :" & DdMmmYy(frmUSevanaInward.dtpDeliveryDate1.Value)
        End If
        
        frmReceiptsCounter.txtWardNo.Enabled = False
        frmReceiptsCounter.txtDoorNo1.Enabled = False
        frmReceiptsCounter.txtDoorNo2.Enabled = False
        frmReceiptsCounter.txtName.Enabled = False
        frmReceiptsCounter.txtInit1.Enabled = False
        frmReceiptsCounter.txtInit2.Enabled = False
        frmReceiptsCounter.txtInit3.Enabled = False
        frmReceiptsCounter.txtInit4.Enabled = False
        frmReceiptsCounter.txtHouse.Enabled = False
        frmReceiptsCounter.txtStreet.Enabled = False
        frmReceiptsCounter.txtLocalPlace.Enabled = False
        frmReceiptsCounter.txtMainPlace.Enabled = False
        frmReceiptsCounter.txtPost.Enabled = False
        frmReceiptsCounter.txtPin.Enabled = False
        frmReceiptsCounter.txtPhone.Enabled = False
        
   
        
        frmReceiptsCounter.Visible = True
        frmReceiptsCounter.ZOrder (0)
        If InwardMode = 0 Then
            frmUSoochikaInward.ZOrder (1)
        Else
            frmUSoochikaManualInward.ZOrder (1)
        End If
        
    Exit Function
err:
    MsgBox (Error$)
End Function

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo err:
        If gbSevanaMainTypeID = 1 Or gbSevanaMainTypeID = 2 Or gbSevanaMainTypeID = 3 Or gbSevanaMainTypeID = 4 Then
            If vsGrid.Cell(flexcpChecked, 1, 0) = vbChecked Then
                txtNoOfCertificate.Enabled = True
            Else
                If Row = 1 Then
                    txtNoOfCertificate.Text = 1
                End If
               ' txtNoOfCertificate.Enabled = False
            End If
            
            If vsGrid.Cell(flexcpChecked, 2, 0) = vbChecked Then
                txtNoofYears.Enabled = True
            Else
                If Row = 2 Then
                    txtNoOfCertificate.Text = 1
                End If
              '  txtNoofYears.Enabled = False
            End If
        ElseIf gbSevanaMainTypeID = 5 Then
'            If vsGrid.Cell(flexcpChecked, 6, 0) = vbChecked Then
'                txtNoOfCertificate.Enabled = True
'            Else
'                If Row = 6 Then
'                    txtNoOfCertificate.Text = 1
'                End If
'              '  txtNoOfCertificate.Enabled = False
'            End If
        End If
        If vsGrid.Col = 0 Then
            Call Calculate
        End If
    Exit Sub
err:
    MsgBox (Error$)
End Sub
Private Sub dtpEventDate_Change()
    If ((chkInsideCountry.Value = 1 Or chkOutsideCountry.Value = 1) And (IsNull(dtpEventDate.Value) = False)) Then
        If (CDate(dtpEventDate.Value) > Date) Then
            MsgBox "Event date should be less than today's date!!!", vbInformation, "SOOCHIKA"
            dtpEventDate.Value = Date
            Exit Sub
        Else
            DTPApplDate.Value = dtpEventDate.Value
            SetSubType
        End If
    End If
End Sub
Private Sub chkInsideCountry_Click()
    If (chkInsideCountry.Value = 1) Then
        chkOutsideCountry.Value = 0
        Label18.Caption = "Event Date"
        txtSubTypeID.Locked = True
        cboSubType.Locked = True
        dtpEventDate.Value = Date
    Else
        txtSubTypeID.Locked = False
        cboSubType.Locked = False
        dtpEventDate.Value = Null
    End If
End Sub
Private Sub chkOutsideCountry_Click()
    If chkOutsideCountry.Value = 1 Then
        chkInsideCountry.Value = 0
        Label18.Caption = "Arrival Date"
        txtSubTypeID.Locked = True
        cboSubType.Locked = True
        dtpEventDate.Value = Date
    Else
        txtSubTypeID.Locked = False
        cboSubType.Locked = False
        dtpEventDate.Value = Null
    End If
End Sub

Private Sub SetSubType()
    Dim days As Integer
    Dim J As Integer
    days = 0
    Dim i As Integer
    Dim PrvDate As Date
    PrvDate = DateAdd("D", 1, Date)
'            TimeSpan span = dtpAppDate.Value.Date.Subtract(dtpEventDate.Value.Date);
'            days = span.Days + 1;
    days = DateDiff("D", dtpEventDate.Value, Date)
    days = days + 1
    For i = 0 To 10
        If ValidateHoliday(PrvDate) = True Then
            days = days - 1
            PrvDate = DateAdd("D", -1, PrvDate)
        Else
            Exit For
        End If
    Next
            
            If gbSevanaMainTypeID = 1 Then         'Birth Registration
            'CHNAGED by soumya VS on 03Dec
                If (chkOutsideCountry.Value = 1) And (days <= 60) Then
            
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 2 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                
                ElseIf chkOutsideCountry.Value = 1 And days > 60 And (days <= 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 3 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                    
                    
                    ElseIf chkOutsideCountry.Value = 1 And (days > 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 148 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                    
                ElseIf days <= 21 Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 1 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 30) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 4 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 5 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days > 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 6 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                End If
            ElseIf (gbSevanaMainTypeID = 2) Then    'Death Registration
                If (days <= 21) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 22 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 30) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 23 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 24 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days > 365) Then
                   For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 25 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                End If
            ElseIf (gbSevanaMainTypeID = 3) Then    'Still Birth Registration
                If (days <= 21) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 37 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 30) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 38 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days <= 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 39 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days > 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 4 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                End If
            ElseIf (gbSevanaMainTypeID = 4) Then    'Marriage Registration
                If (days <= 30) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 48 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days > 30) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 50 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                End If
            ElseIf (gbSevanaMainTypeID = 5) Then   'Common Marriage Registration
                If (days <= 45) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 89 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (DateDiff("D", dtpEventDate.Value, CDate("28/02/2008")) <= 0) Then
                   For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 90 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next '<=28/02/2008
                ElseIf (days <= 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 91 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                ElseIf (days > 365) Then
                    For J = 1 To cboSubType.ListCount
                        If cboSubType.ItemData(J - 1) = 92 Then
                            cboSubType.ListIndex = J - 1
                        End If
                    Next
                End If
            End If
End Sub
 Private Function ValidateHoliday(ByVal HolidayDate As Date)
    ValidateHoliday = False
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim arrIn As Variant
    ReDim arrIn(0)
    Dim i As Integer
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Soochika connection Failure !!", "SOOCHIKA", vbInformation
        Exit Function
    End If
    arrIn(0) = Year(HolidayDate)
    Set Rec = objdb.ExecuteSP("Sp_SelectHolidayList", arrIn, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF Or Rec.BOF) Then
            For i = 0 To Rec.RecordCount
               If Rec.Fields(0) = HolidayDate Then
                    ValidateHoliday = True
                    Return
                End If
            Next
        End If
 End Function

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Col = 6 Then
      Cancel = True
   End If
   If Col = 7 Then
      Cancel = True
   End If
End Sub
Private Sub CalculateAmount()
    Dim Amount As Integer
'    Dim intNoOfCer As Integer
'    If txtNoOfCertificate.Text = "" Then
'        intNoOfCer = 1
'    Else
'     intNoOfCer = val(txtNoOfCertificate.Text)
'    End If
'Added by sunil

'MsgBox val(vsGrid.TextMatrix(1, 2))
'MsgBox val(vsGrid.TextMatrix(2, 2))
    If txtSubTypeID.Text = 60 Or txtSubTypeID.Text = 61 Or txtSubTypeID.Text = 62 Then
              
        If val(vsGrid.TextMatrix(1, 2)) <> 0 Then
           Amount = val(txtNoofYears.Text) * val(vsGrid.TextMatrix(1, 16))
           vsGrid.TextMatrix(1, 6) = val(vsGrid.TextMatrix(1, 16))
        End If
        vsGrid.TextMatrix(1, 7) = Amount
    Else
        
        If val(vsGrid.TextMatrix(2, 2)) <> 0 Then
           Amount = val(txtNoofYears.Text) * val(vsGrid.TextMatrix(2, 16))
           vsGrid.TextMatrix(2, 6) = val(vsGrid.TextMatrix(2, 16))
        End If
    '    amount = val(vsGrid.TextMatrix(1, 6)) * intNoOfCer + 2 * val(txtNoofYears.Text)
        vsGrid.TextMatrix(2, 7) = Amount
    End If
End Sub
Private Sub CertificateorNot()
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim arrIn As Variant
    Dim mSql As String
    ReDim arrIn(0)
    Dim i As Integer
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
        mSql = "Select faSubTypes.intSubTypeID,vchAlias,intSlNo,tnyMultipleFlag,fltrate,intAccountHeadID,vchAccountHeadCode from faSubTypes"
        mSql = mSql + " Inner join faSubTypeSchedule on faSubTypes.intSubTypeID=faSubTypeSchedule.intSubTypeID"
        mSql = mSql + " Where intSevanSubTypeID=" & txtSubTypeID.Text & " and (vchAlias like'%certificate%' Or vchAlias like '%extract%') "
        Rec.Open mSql, mCnn
        If Not Rec.EOF And Not Rec.BOF Then
            txtNoOfCertificate.Text = 1
        End If
    End If
End Sub

'changed by soumya V S on 14.05.14
Private Sub getSubjectDeliverydate(ByVal SubTypeID As Integer)  'Hol check
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
      Dim Rec1 As New ADODB.Recordset
      Dim Rec2 As New ADODB.Recordset
       'Dim Rec3 As New ADODB.Recordset
      Dim i As Integer
      
    Dim arrIn As Variant
    Dim deliveryDate As Variant
    Dim strDate As Date
    
    Dim tot
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    
    Rec.Open "select intPeriod from tSubjectDeliveryPeriod where numSubTypeID=" & SubTypeID, mCnn
    Rec1.Open "Set dateformat DMY SELECT COUNT(*) as cnt FROM mholiday WHERE dtDate>getdate() and dtDate<=dATEADD(DD,15,GETDATE())", mCnn
    Rec2.Open "select numSubTypeID from tSubjectDeliveryPeriod where numSubTypeID<>0  and numSubTypeID IS NOT NULL and  numSubTypeID=" & SubTypeID, mCnn
    
    strDate = Format(Now, "dd/MM/yyyy")
        deliveryDate = strDate
        
    If Not (Rec2.EOF Or Rec2.BOF) Then
        If IsNull(Rec2!numSubTypeID) Then
          GoTo aa
          Else
          
        
        
        For i = 1 To CInt(Rec!intPeriod)
        deliveryDate = DateAdd("d", 1, deliveryDate)
        While (CheckHoliday(deliveryDate) = True)
        deliveryDate = DateAdd("d", 1, deliveryDate)
        Wend
        Next i
                
        dtpDeliveryDate1.Value = deliveryDate
        
    'End If
aa: dtpDeliveryDate1.Value = deliveryDate
End If
End If
dtpDeliveryDate1.Value = deliveryDate
    
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub

Private Function CheckHoliday(ByVal dtDate As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim arrIn As Variant
    Dim mSql As String
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Function
    End If
    mSql = " select * from mHoliday where convert(datetime,dtdate,103)=convert(datetime,'" & dtDate & "',103)"
    Rec.Open mSql, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        CheckHoliday = True
    Else
        CheckHoliday = False
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function


Private Function ValidateEnclosure()

    ValidateEnclosure = True
  
        For i = 1 To grvCheckList.Rows - 1
            If (grvCheckList.TextMatrix(i, 1) = "-1") Then
                ValidateEnclosure = True
                Exit Function
            End If
        Next i

End Function

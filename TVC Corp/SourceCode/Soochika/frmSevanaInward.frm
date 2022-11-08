VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSevanaInward 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sevana Inward Details"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   3495
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Visible         =   0   'False
      Width           =   9495
      Begin VB.TextBox txtNoOfCertificate 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   8370
         TabIndex        =   46
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
         Left            =   7065
         TabIndex        =   45
         Top             =   3015
         Width           =   1965
      End
      Begin VB.TextBox txtNoofYears 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5760
         TabIndex        =   43
         Top             =   180
         Width           =   735
      End
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2370
         Left            =   180
         TabIndex        =   42
         Top             =   570
         Width           =   8970
         _cx             =   15822
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
         FormatString    =   $"frmSevanaInward.frx":0000
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
         Left            =   4710
         TabIndex        =   49
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
         Left            =   5985
         TabIndex        =   48
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
         Left            =   6720
         TabIndex        =   47
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
         Left            =   4650
         TabIndex        =   44
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame frameSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   9495
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
         Height          =   1815
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   8895
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
            Left            =   7110
            TabIndex        =   51
            Top             =   1035
            Width           =   1395
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
            Left            =   5595
            TabIndex        =   20
            Top             =   1035
            Width           =   1395
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
            Left            =   6000
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   1455
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
            TabIndex        =   16
            Top             =   360
            Width           =   735
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
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtMalayalamname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "ML-TTRevathi"
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
            TabIndex        =   18
            Top             =   1320
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
            TabIndex        =   17
            Top             =   840
            Width           =   3255
         End
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
            TabIndex        =   14
            Top             =   360
            Width           =   3255
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
            TabIndex        =   40
            Top             =   360
            Width           =   735
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
            TabIndex        =   39
            Top             =   360
            Width           =   615
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
            TabIndex        =   38
            Top             =   1320
            Width           =   1095
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
            TabIndex        =   37
            Top             =   840
            Width           =   1095
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
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
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
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
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
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   4200
         TabIndex        =   34
         Top             =   360
         Width           =   1455
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
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
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
      Left            =   9240
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
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
         TabIndex        =   11
         Top             =   840
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
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
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
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPReceiptDate 
         Height          =   300
         Left            =   1680
         TabIndex        =   8
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
         Format          =   60751873
         CurrentDate     =   40038
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
         TabIndex        =   26
         Top             =   840
         Width           =   1095
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
         TabIndex        =   25
         Top             =   840
         Width           =   855
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
         TabIndex        =   24
         Top             =   360
         Width           =   735
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
         TabIndex        =   23
         Top             =   360
         Width           =   975
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CheckBox chkZonal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Zonal Office"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6690
         TabIndex        =   50
         Top             =   1560
         Width           =   1605
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
         Left            =   1770
         TabIndex        =   1
         Top             =   360
         Width           =   495
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
         TabIndex        =   7
         Top             =   1305
         Width           =   975
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
         TabIndex        =   6
         Top             =   840
         Width           =   975
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
         Height          =   735
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1320
         Width           =   4725
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
         Left            =   4080
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPApplDate 
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60751873
         CurrentDate     =   40037
      End
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
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   5895
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
         TabIndex        =   32
         Top             =   360
         Width           =   855
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
         TabIndex        =   31
         Top             =   360
         Width           =   135
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
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1455
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
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   975
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
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Width           =   135
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
         Left            =   3240
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSevanaInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim MainSubTypeID As Variant
    Dim tnyType As Variant
    Dim SevanaTypeID    As Variant
    Dim SevanaKioskID   As Variant
    Dim CommMarriageFee As Variant
    '-------------------------------------'
    Dim intTransactionTypeID As Integer
    
    
    Dim AmtNoofCert As Double
    Dim SearchAmt As Double
    Dim mTempTotal As Double
    Dim AmtNoofExtraCert As Double
      
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
        Dim mSQL As String
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        If cboSubType.ListIndex = -1 Then Exit Sub
        txtSubTypeID.Text = cboSubType.ItemData(cboSubType.ListIndex)
        If txtSubTypeID.Text = 93 Or txtSubTypeID.Text = 80 Or txtSubTypeID.Text = 81 Or txtSubTypeID.Text = 82 Or txtSubTypeID.Text = 83 Then
            MsgBox "This subtype is blocked", vbInformation
            txtSubTypeID.Text = ""
            cboSubType.ListIndex = 0
        Else
'             If MainSubTypeID = 5 Then
'                 GetCommonMarriageFee (txtSubTypeID.Text)            ' For getting the common marriage fee
'                 txtReceiptAmount.Enabled = False
'                 If SevanaTypeID = 2 Then
'                     txtNoCopeis.Text = 1
'                 Else
'                     txtNoCopeis.Text = ""
'                 End If
'             Else
'                 txtReceiptAmount.Enabled = True
'             End If
             If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Then
                 Label2.Caption = "Arrival Date"
             Else
                 Label2.Caption = "Application Date"
             End If
             
             mSQL = "Select tnyType,tnyToSeat from TblSubjectSubType where intid= " & txtSubTypeID
             If (objDb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
                 MsgBox "Connection not present", vbDefaultButton1
                 Exit Sub
             End If
            
             Rec.Open mSQL, mCnn
             If Not (Rec.BOF Or Rec.EOF) Then
                 SevanaTypeID = Rec!tnyType
                 SevanaKioskID = Rec!tnyToSeat
                 If InwardMode = 0 Then
                    frmSoochikaInward.SevanaTypeID = SevanaTypeID
                    frmSoochikaInward.SevanaKioskID = SevanaKioskID
                Else
                    frmSoochikaManualInward.SevanaTypeID = SevanaTypeID
                    frmSoochikaManualInward.SevanaKioskID = SevanaKioskID
                End If
             End If
'             If chkZonal.Value = 1 Then
'                SevanaTypeID = 0
'                frmSoochikaInward.SevanaTypeID = SevanaTypeID
'             End If
             If SevanaTypeID = 1 Then
                If chkZonal.value = 1 Then
                    frameReceipt.Visible = False
                    'frameSearch.Visible = True
                    frameSearch.Top = 2500
                    Me.Left = 2250
                    Me.Top = 2000
                    'Me.Height = 2500
                    Me.Height = 2700
                    cmdOK.Enabled = True
                Else
                    frameReceipt.Visible = True
                    Me.Left = 2250
                    Me.Top = 2000
                    'Me.Height = 6250
                    Me.Height = 6500
                    frameReceipt.Top = 2500
                    cmdOK.Enabled = False
                    Call FillGrid
                End If
             ElseIf SevanaTypeID = 2 Then
                If txtSubTypeID.Text = 110 Or chkZonal.value = 1 Then
                    frameReceipt.Visible = False
                    frameSearch.Visible = True
                    frameSearch.Top = 2500
                    Me.Left = 2250
                    'Me.Top = 2000
                    Me.Top = 1500
                    Me.Height = 5550
                    txtNoOfCertificate.Text = 1
                    cmdOK.Enabled = True
                Else
                    frameReceipt.Visible = True
                    frameSearch.Visible = True
                    frameReceipt.Top = 5250
                    Me.Left = 2250
                    'Me.Top = 2000
                    Me.Top = 1500
                    'Me.Height = 8700
                    Me.Height = 9405
                    Call FillGrid
                    cmdOK.Enabled = False
                End If
             Else
                 frameReceipt.Visible = False
                 frameSearch.Visible = False
                 cmdOK.Enabled = True
                Me.Left = 2250
                 Me.Top = 2000
                 'Me.Height = 2500
                 Me.Height = 2700
             End If
        End If
    End Sub
    Private Sub cboSubType_Click()
        Call ShowFrames
    End Sub

    Private Sub chkZonal_Click()
        ShowFrames
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

    Private Sub cmdClose_Click()
        If InwardMode = 0 Then
            frmSoochikaInward.txtSubID.Text = ""
            frmSoochikaInward.txtSubject.Text = ""
        Else
            frmSoochikaManualInward.txtSubID.Text = ""
            frmSoochikaManualInward.txtSubject.Text = ""
        End If
        Unload Me
    End Sub

    Private Sub cmdCopy_Click()
        Call frmReceiptsCounter.CheckInterruptReceiptRequestStatus
        If frmReceiptsCounter.InterruptedMode = False And InwardMode = 1 Then
            MsgBox "You have no authority to take receipt ", vbInformation, "receipt"
            Exit Sub
        End If
        If CopyValidation Then
            Call copyToReceipt
        End If
    End Sub
    
    Private Function CopyValidation() As Boolean
        On Error GoTo Err:
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
            
            
            CopyValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub cmdGetName_Click()
        Dim objDb As New clsDB
        Dim con5 As New ADODB.Connection
        Dim rs5 As New ADODB.Recordset
        Dim Qry
        
        If (objDb.CreateNewConnection(con5, enuSourceString.SevanaRegn) = False) Then
            MsgBox "Sevena Connection Failed", vbDefaultButton1
            Exit Sub
        End If
        
        If Trim(txtRegNo) <> "" And Trim(txtBookNo.Text) <> "" Then 'If registration no empty
         If MainSubTypeID = 1 Then        'Birth
             Qry = "SELECT     chvMalFather, chvEngFather From tBirthRep " _
             & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
             frmSevanaInward.cboRelationship.ListIndex = 1
         ElseIf MainSubTypeID = 2 Then   'Death
             Qry = "SELECT     chvMalDeadName, chvEngDeadName From tDeathRep " _
             & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
             frmSevanaInward.cboRelationship.ListIndex = 0
         ElseIf MainSubTypeID = 3 Then   'Still birth
             Qry = "SELECT     chvMalFather, chvEngFather From tStillBirthRep " _
             & "WHERE     (chvRegnNo = '" & Trim(txtRegNo) & "') and chvbookno='" & Trim(txtBookNo.Text) & "'"
             frmSevanaInward.cboRelationship.ListIndex = 0
             '--------------------Commented by savitha on 24.01.2008-------
         ElseIf MainSubTypeID = 4 Then   'Marriage
             '-----------modified on 15/10/08 by nisha ninan
            If cboRelationship.ListIndex = 0 Then
                 Qry = "SELECT     tMarriageMal.chvGroom,tMarriageEng.chvGroom FROM tMarriageEng INNER JOIN tMarriageMal ON tMarriageEng.chvAckNo = tMarriageMal.chvAckNo " _
                 & "WHERE     (tMarriageEng.chvRegnNo = '" & Trim(txtRegNo) & "')"
                 frmSevanaInward.cboRelationship.ListIndex = 0
             Else
                 Qry = "SELECT     tMarriageMal.chvBride,tMarriageEng.chvBride FROM tMarriageEng INNER JOIN tMarriageMal ON tMarriageEng.chvAckNo = tMarriageMal.chvAckNo " _
                 & "WHERE     (tMarriageEng.chvRegnNo = '" & Trim(txtRegNo) & "')"
                 frmSevanaInward.cboRelationship.ListIndex = 1
             End If
             '---------added by nisha on 10/10/08
             ElseIf MainSubTypeID = 5 Then 'Common Marriage
             If cboRelationship.ListIndex = 0 Then
              '.............Modified by savitha on 27.01.2009 for commonMarriage
        
        '            Qry = "select tMarriageMalayalam.chvHusName as MalHus,tMarriageEnglish.chvHusName from tMarriageEnglish inner join tMarriageMalayalam on tMarriageEnglish.chvAckNo=tMarriageMalayalam.chvAckNo " _
        '            & " where  (tMarriageEnglish.chvRegnNo='" & Trim(txtRegistrationno) & "')"
                 
                 Qry = "select chvHusName from tMarriageEnglish  where  chvRegnNo='" & Trim(txtRegNo) & "'"
                 
                 frmSevanaInward.cboRelationship.ListIndex = 0
             Else
        '            Qry = "select tMarriageMalayalam.chvWfeName as MalWfe,tMarriageEnglish.chvWfeName from tMarriageEnglish inner join tMarriageMalayalam on tMarriageEnglish.chvAckNo=tMarriageMalayalam.chvAckNo " _
        '            & " where  (tMarriageEnglish.chvRegnNo='" & Trim(txtRegistrationno) & "')"
                 Qry = "select chvWfeName from tMarriageEnglish  where  chvRegnNo='" & Trim(txtRegNo) & "'"
                 frmSevanaInward.cboRelationship.ListIndex = 1
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
            
              If MainSubTypeID = 5 Then
              Dim rsMal As New ADODB.Recordset
               Dim SQL As String
                     If cboRelationship.ListIndex = 0 Then
                               
                             SQL = "select tMarriageMalayalam.chvHusName from tMarriageMalayalam inner join tMarriageEnglish on  tMarriageEnglish.chvackno=tMarriageMalayalam.chvackno where  tMarriageEnglish.chvRegnNo='" & Trim(txtRegNo) & "'"
                             rsMal.Open SQL, con5
                             If rsMal.EOF = False Then
                                  txtMalayalamname.Text = rsMal(0)
                             Else
                                  txtMalayalamname.Text = "\ðInbn«nñ"
                             End If
                      Else
             
                         SQL = "select tMarriageMalayalam.chvWfeName from tMarriageMalayalam inner join tMarriageEnglish on  tMarriageEnglish.chvackno=tMarriageMalayalam.chvackno where  tMarriageEnglish.chvRegnNo='" & Trim(txtRegNo) & "'"
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
            If MainSubTypeID = 4 Or MainSubTypeID = 5 Then
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
            
    End Sub

    Private Sub cmdOK_Click()
        On Error GoTo Err:
            
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
            Dim objDb As New clsDB
            Dim mCnnSoochika As New ADODB.Connection
            Dim mCnnSevana As New ADODB.Connection
            Dim InwNo As Double
            
            flag = 1
            If Validate <> 0 Then
                objDb.CreateNewConnection mCnnSoochika, enuSourceString.SOOCHIKA
                mCnnSoochika.BeginTrans
                On Error GoTo ErroRollBack:
                    If InwardMode = 0 Then
                        InwNo = frmSoochikaInward.SaveSoochika(mCnnSoochika)
                    Else
                        InwNo = frmSoochikaManualInward.SaveSoochika(mCnnSoochika)
                    End If
                    objDb.CreateNewConnection mCnnSevana, enuSourceString.SevanaRegn
                    mCnnSevana.BeginTrans
                    If InwardMode = 0 Then
                        Call frmSoochikaInward.SaveSevana(InwNo, 0, 0, mCnnSevana)
                    Else
                        Call frmSoochikaManualInward.SaveSevana(InwNo, 0, 0, mCnnSevana)
                    End If
                mCnnSoochika.CommitTrans
                mCnnSevana.CommitTrans
                If InwardMode = 0 Then
                    frmSoochikaInward.Ack (frmSoochikaInward.lSoochikaFeildID)
                End If
                Unload Me
                'MsgBox " HAppy New Year"
                If InwardMode = 0 Then
                    frmSoochikaInward.ClearDetails
                Else
                    frmSoochikaManualInward.ClearDetails
                End If
            End If
        Exit Sub
Err:
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
    Public Function Validate()
        Dim flag
        flag = 1
        If txtSubTypeID.Text = "" Then
            flag = 0
            MsgBox "Enter SubType", vbDefaultButton1
            txtSubTypeID.SetFocus
            GoTo last
        ElseIf cboSubType.ListIndex < 0 Then
            flag = 0
            MsgBox "select subtype", vbInformation
            cboSubType.SetFocus
            GoTo last
        ElseIf DTPApplDate.value = 0 Then
            flag = 0
            MsgBox "select the Application/Arrival Date", vbDefaultButton1
            DTPApplDate.SetFocus
            GoTo last
        End If
'        If SevanaTypeID = 1 Then
'            If txtReceiptBookNo.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt book no", vbInformation
'                txtReceiptBookNo.SetFocus
'                GoTo last
'            ElseIf txtReceiptNo.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt No", vbInformation
'                txtReceiptNo.SetFocus
'                GoTo last
'            ElseIf txtReceiptAmount.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt Amount"
'                txtReceiptAmount.SetFocus
'                GoTo last
'            End If
        If SevanaTypeID = 2 Then
'            If txtReceiptBookNo.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt book no", vbInformation
'                txtReceiptBookNo.SetFocus
'                GoTo last
'            ElseIf txtReceiptNo.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt No", vbInformation
'                txtReceiptNo.SetFocus
'                GoTo last
'            ElseIf txtReceiptAmount.Text = "" Then
'                flag = 0
'                MsgBox "Enter Receipt Amount"
'                txtReceiptAmount.SetFocus
'                GoTo last
'            ElseIf txtNoCopeis.Text = "" Then
'                flag = 0
'                MsgBox "Enter no of copies", vbInformation
'                txtNoCopeis.SetFocus
'                GoTo last
'            Else
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
last:         Validate = flag
    End Function
    Private Sub cmdSearch_Click()
    
        '--------------------------------------------'
        
        If txtRegNo.Text <> "" And txtBookNo.Text <> "" Then
            cmdGetName_Click
            Exit Sub
        End If
        
        '--------------------------------------------'
    
    
        If MainSubTypeID = 1 Then
            frmSevanaBirthSearch.Show vbModal 'Birth Search
        ElseIf MainSubTypeID = 2 Then
            frmSevanadethsearch.Show vbModal                   'Death Search
        ElseIf MainSubTypeID = 3 Then
            frmSevanaStillBirth.Show vbModal                   'Still Birth Search
        ElseIf MainSubTypeID = 4 Then
            frmSevanaMarriageSearch.Show vbModal               'Marriage Search
        ElseIf MainSubTypeID = 5 Then
            frmSevanaCommonMarriageSearch.Show vbModal         'Common Marriage Search
        End If
        
    End Sub

    Private Sub Form_Activate()
        If MainSubTypeID = 0 Then
            Unload Me
        End If
       If txtSubTypeID.Text = "" Then
            Me.Left = 2250
            Me.Top = 2000
            'Me.Height = 2500
            Me.Height = 2700
        Else
            ShowFrames
        End If
    End Sub

    Private Sub Form_Load()
        Dim mCnn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        MainSubTypeID = 0
        If (objDb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
            MsgBox "Connection not present", vbDefaultButton1
            Exit Sub
        End If
        
        If InwardMode = 0 Then
            mSQL = "Select isnull(intMainSubID,'0') as intMainSubID from TblSubjectcoding where intsubID= " & frmSoochikaInward.txtSubID
        Else
            mSQL = "Select isnull(intMainSubID,'0') as intMainSubID from TblSubjectcoding where intsubID= " & frmSoochikaManualInward.txtSubID
        End If
        Rec.Open mSQL, mCnn
        If Not (Rec.BOF Or Rec.EOF) Then
            MainSubTypeID = Rec!intMainSubID
            frmSoochikaInward.SevanaMainSubid = Rec!intMainSubID
        End If
        
        If MainSubTypeID = 0 Then
            Exit Sub
        End If
        
        DTPApplDate.value = Date
        DTPReceiptDate.value = Date
        PopulateList cboSubType, "Select TypeofSubRequest,intID from TblSubjectSubType where intsubTypeID='" & MainSubTypeID & "'", , , , True, enuSourceString.SOOCHIKA
        If MainSubTypeID = 4 Or MainSubTypeID = 5 Then
            Label3.Visible = False
            cboHospitals.Visible = False
        Else
            Label3.Visible = True
            cboHospitals.Visible = True
            PopulateList cboHospitals, "CBOSelectHospital", , True, , True, enuSourceString.SevanaRegn
        End If
        PopulateList cboRelationship, "select chvdescription,intid from mCertificateOwners where intregtype=" & MainSubTypeID, , , , True, enuSourceString.SevanaRegn
        
        cboLanguage.Clear
        cboLanguage.AddItem "Malayalam"
        cboLanguage.ItemData(cboLanguage.NewIndex) = 1
        cboLanguage.AddItem "English"
        cboLanguage.ItemData(cboLanguage.NewIndex) = 2
        cboLanguage.ListIndex = 1
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
'        If txtReceiptAmount.Text <> "" And MainSubTypeID = 5 Then
'            txtReceiptAmount.Text = Val(CommMarriageFee) * Val(txtNoCopeis.Text)
'        End If
    End Sub
    
    Private Sub txtNoCopeis_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
        
    Private Sub txtNoOfCertificate_Change()
        If vsGrid.TextMatrix(1, 6) = "" Then Exit Sub
        If MainSubTypeID = 1 Or MainSubTypeID = 2 Or MainSubTypeID = 3 Or MainSubTypeID = 4 Then
            If txtNoOfCertificate.Text <> "" Then
                vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6) * val(txtNoOfCertificate.Text)
            Else
                vsGrid.TextMatrix(1, 7) = vsGrid.TextMatrix(1, 6)
            End If
        ElseIf MainSubTypeID = 5 Then
            If txtNoOfCertificate.Text <> "" Then
                vsGrid.TextMatrix(6, 7) = vsGrid.TextMatrix(6, 6) * val(txtNoOfCertificate.Text)
            Else
                vsGrid.TextMatrix(6, 7) = vsGrid.TextMatrix(6, 6)
            End If
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
        If vsGrid.TextMatrix(2, 6) = "" Then Exit Sub
        If txtNoofYears.Text <> "" Then
            vsGrid.TextMatrix(2, 7) = vsGrid.TextMatrix(2, 6) * val(txtNoofYears.Text)
        Else
            vsGrid.TextMatrix(2, 7) = vsGrid.TextMatrix(2, 6)
        End If
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
    '''            frmSevanaInward.txtReceiptAmount = "10"
    '''        Case 91
    '''            frmSevanaInward.txtReceiptAmount = "110"
    '''        Case 92
    '''            frmSevanaInward.txtReceiptAmount = "260"
    '''        Case 93
    '''            frmSevanaInward.txtReceiptAmount = "5"
    '''        Case 96
    '''            frmSevanaInward.txtReceiptAmount = "100"
    '''        Case 98
    '''            frmSevanaInward.txtReceiptAmount = "25"
    '''        Case 99
    '''            frmSevanaInward.txtReceiptAmount = "15"
    '''        Case 100
    '''            frmSevanaInward.txtReceiptAmount = "15"
    '''        Case 101
    '''            frmSevanaInward.txtReceiptAmount = "115"
    '''        Case 102
    '''            frmSevanaInward.txtReceiptAmount = "265"
    '''        Case 103
    '''            frmSevanaInward.txtReceiptAmount = "25"
    '''        Case 104
    '''            frmSevanaInward.txtReceiptAmount = "125"
    '''        Case 94, 95, 97
    '''            frmSevanaInward.txtReceiptAmount = ""
    '''    End Select
    '''    CommMarriageFee = frmSevanaInward.txtReceiptAmount.Text
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
    If txtSubTypeID.Text = "2" Or txtSubTypeID.Text = "3" Then
        Label2.Caption = "Arrival Date"
    Else
        Label2.Caption = "Application Date"
    End If
    If flag = 1 Then
        Call ShowFrames
    End If
End Sub
Private Sub FillGrid()
    On Error GoTo Err:
        Dim mSQL As String
        Dim objDb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mRowCount As Integer
        
        Dim mScheduleType As Integer
        Dim mScheduleSubID As Integer
        
        
        Dim mFunctionID As Integer
        Dim mFunctionaryID As Integer
        
        
        If MainSubTypeID = 1 Or MainSubTypeID = 3 Then 'Birth
            mScheduleType = 13
            mScheduleSubID = 1
            intTransactionTypeID = gbTransactionTypeBrith
            
            mFunctionID = 46
            mFunctionaryID = 7
        ElseIf MainSubTypeID = 2 Then     'Death
            mScheduleType = 13
            mScheduleSubID = 2
            intTransactionTypeID = gbTransactionTypeDeath
            
            mFunctionID = 46
            mFunctionaryID = 7
'        ElseIf mainSubTypeID = 4 Then     'Marriage
'            mScheduleType = 15
'            mScheduleSubID = 3
'            intTransactionTypeID = gbTransactionTypeMarriage
'
'            mFunctionID = 47
'            mFunctionaryID = 7

        ElseIf MainSubTypeID = 5 Then     'CmnMarriage
            mScheduleType = 14
            mScheduleSubID = 3
            intTransactionTypeID = gbTransactionTypeCmnMarriage
            
            mFunctionID = 47
            mFunctionaryID = 7
        ElseIf MainSubTypeID = 4 Then     'Marriage
            mScheduleType = 15
            mScheduleSubID = 3
            intTransactionTypeID = gbTransactionTypeMarriage
            
            mFunctionID = 47
            mFunctionaryID = 7
        End If
        
        txtNoOfCertificate.Text = 1
        txtNoofYears.Text = 1
        
        If objDb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
            If InwardMode = 0 Then
                If frmSoochikaInward.chkBPL.value = 1 Or frmSoochikaInward.chkSCST.value = 1 Then
                    mSQL = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltSpecialRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
                Else
                    mSQL = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltFixedRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
                End If
            Else
                If frmSoochikaManualInward.chkBPL.value = 1 Or frmSoochikaManualInward.chkSCST.value = 1 Then
                    mSQL = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltSpecialRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
                Else
                    mSQL = "SELECT  distinct  smScheduleMasters.intScheduleID, smScheduleMasters.fltFixedRate, smAttributes.vchAccountHeadCode, smAttributes.vchAttributeTitle, smAttributes.intAccountHeadID,smAttributes.intAttributeID"
                End If
            End If
            
            mSQL = mSQL + " FROM         smScheduleMasters INNER JOIN "
            mSQL = mSQL + " smAttributes ON smScheduleMasters.intAttributeID = smAttributes.intAttributeID "
            mSQL = mSQL + " WHERE     (smScheduleMasters.intScheduleID = " & mScheduleType & ") and smAttributes.tnyGroupID = " & mScheduleSubID & " and smAttributes.intAttributeID <> 154  and smAttributes.intAttributeID <> 158 order by smAttributes.intAttributeID"   'Modified on 06.04.2009 by Suby / Modified on 17.04.2009
            Rec.Open mSQL, mCnn
            mRowCount = 1
            vsGrid.Rows = 2
            While Not Rec.EOF And Not Rec.BOF
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!vchAttributeTitle), "", Rec!vchAttributeTitle)
                If frmSoochikaInward.chkBPL.value = 1 Or frmSoochikaInward.chkSCST.value = 1 Then
                    vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
                    vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
                    vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltSpecialRate), "", Rec!fltSpecialRate)
                Else
                    vsGrid.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
                    vsGrid.TextMatrix(mRowCount, 7) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
                    vsGrid.TextMatrix(mRowCount, 11) = IIf(IsNull(Rec!fltFixedRate), "", Rec!fltFixedRate)
                End If
                vsGrid.TextMatrix(mRowCount, 12) = IIf(IsNull(Rec!intAttributeID), "", Rec!intAttributeID) 'Added on 03.04.2009
                mRowCount = mRowCount + 1
                vsGrid.Rows = vsGrid.Rows + 1
                Rec.MoveNext
            Wend
            'Added on 17.04.2009 by Suby---
'            If MainSubTypeID= 5 Then
'                vsGrid.RowHidden(2) = True
'            End If
            '------------------------------
            If Rec.State = 1 Then Rec.Close
            mSQL = "Select * from smAttributeSevanaMapping Where intSevanaSubTypeID = " & val(txtSubTypeID.Text)
            Rec.Open mSQL, mCnn
            While Not (Rec.EOF Or Rec.BOF)
                For mRowCount = 1 To vsGrid.Rows - 1
                    If vsGrid.TextMatrix(mRowCount, 12) = Rec!intAttributeID Then
                        vsGrid.Cell(flexcpChecked, mRowCount, 0) = vbChecked
                        Call vsGrid_AfterEdit(mRowCount, 0)
                    End If
                Next
                Rec.MoveNext
            Wend
        Else
            MsgBox "Connection To iSaankhyaMasters does not exist, Please Contact your System Administrator", vbInformation
        End If
    Exit Sub
Err:
    MsgBox (Error$)
End Sub

Private Function copyToReceipt() As Boolean
    On Error GoTo Err:
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
        
        If MainSubTypeID = 5 Then
            txtNoCopeis.Text = val(txtNoOfCertificate.Text)
            If vsGrid.Cell(flexcpChecked, 6, 0) = vbChecked Then
                For i = 7 To 11
                    If vsGrid.Cell(flexcpChecked, i, 0) = vbChecked Then
                        txtNoCopeis.Text = val(txtNoCopeis.Text) + 1
                    End If
                Next
            End If
        End If
        Me.Hide
        Load frmReceiptsCounter
        
        
        
        objTrType.SetTransactionType (intTransactionTypeID)
    
        frmReceiptsCounter.SoochikaConnected = True
    
        frmReceiptsCounter.txtTransactionType.Text = objTrType.TransactionType
        frmReceiptsCounter.txtTransactionType.Tag = intTransactionTypeID
        frmReceiptsCounter.SubLedgerID = 9999   ' Test Value    '
'        frmReceiptsCounter.cmbZone.Text = gbnumZonalID
'        frmReceiptsCounter.cmbDZone.Text = gbnumZonalID '   Added   '
        'frmReceiptsCounter.txtWard.Text = frmSoochikaInward.txtWardNo.Text
        If InwardMode = 0 Then
            frmReceiptsCounter.txtWardNo.Text = frmSoochikaInward.txtWardNo.Text
            frmReceiptsCounter.txtWard.Tag = frmSoochikaInward.txtWardNo.Text
            frmReceiptsCounter.txtHouseNo1.Text = frmSoochikaInward.txtDoorNo1.Text
            frmReceiptsCounter.txtHouseNo2.Text = frmSoochikaInward.txtDoorNo2.Text
            frmReceiptsCounter.txtDoorNo1.Text = frmSoochikaInward.txtDoorNo1.Text
            frmReceiptsCounter.txtDoorNo2.Text = frmSoochikaInward.txtDoorNo2.Text
            frmReceiptsCounter.txtName.Text = frmSoochikaInward.txtSender.Text
        Else
            frmReceiptsCounter.txtWardNo.Text = frmSoochikaManualInward.txtWardNo.Text
            frmReceiptsCounter.txtWard.Tag = frmSoochikaManualInward.txtWardNo.Text
            frmReceiptsCounter.txtHouseNo1.Text = frmSoochikaManualInward.txtDoorNo1.Text
            frmReceiptsCounter.txtHouseNo2.Text = frmSoochikaManualInward.txtDoorNo2.Text
            frmReceiptsCounter.txtDoorNo1.Text = frmSoochikaManualInward.txtDoorNo1.Text
            frmReceiptsCounter.txtDoorNo2.Text = frmSoochikaManualInward.txtDoorNo2.Text
            frmReceiptsCounter.txtName.Text = frmSoochikaManualInward.txtSender.Text
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
        
        'frmReceiptsCounter.cmbSeat.Text = frmUSoochikaInward.cmbSeat.Text
        'frmReceiptsCounter.txtReceiptNo.Text = "Delivery Date :" & DdMmmYy(frmUSoochikaInward.dtpDeliveryDate.value)
        
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
            frmSoochikaInward.ZOrder (1)
        Else
            frmSoochikaManualInward.ZOrder (1)
        End If
        
    Exit Function
Err:
    MsgBox (Error$)
End Function

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo Err:
        If MainSubTypeID = 1 Or MainSubTypeID = 2 Or MainSubTypeID = 3 Or MainSubTypeID = 4 Then
            If vsGrid.Cell(flexcpChecked, 1, 0) = vbChecked Then
                txtNoOfCertificate.Enabled = True
            Else
                If Row = 1 Then
                    txtNoOfCertificate.Text = 1
                End If
                txtNoOfCertificate.Enabled = False
            End If
            
            If vsGrid.Cell(flexcpChecked, 2, 0) = vbChecked Then
                txtNoofYears.Enabled = True
            Else
                If Row = 2 Then
                    txtNoOfCertificate.Text = 1
                End If
                txtNoofYears.Enabled = False
            End If
        ElseIf MainSubTypeID = 5 Then
            If vsGrid.Cell(flexcpChecked, 6, 0) = vbChecked Then
                txtNoOfCertificate.Enabled = True
            Else
                If Row = 6 Then
                    txtNoOfCertificate.Text = 1
                End If
                txtNoOfCertificate.Enabled = False
            End If
        End If
        If vsGrid.Col = 0 Then
            Call Calculate
        End If
    Exit Sub
Err:
    MsgBox (Error$)
End Sub


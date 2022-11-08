VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptFinancialFilterFields 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Financial Report Filter Fields"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10620
   Begin VB.Frame fmeReceiptPayment 
      BackColor       =   &H80000009&
      Caption         =   "Receipt && Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3450
      Left            =   4005
      TabIndex        =   53
      Top             =   2385
      Width           =   5820
      Begin VB.Frame fmeRPDiscrepancy 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   855
         Left            =   90
         TabIndex        =   99
         Top             =   2520
         Visible         =   0   'False
         Width           =   5655
         Begin VB.ComboBox cmbYear 
            Height          =   330
            ItemData        =   "frmRptFinancialFilterFields.frx":0000
            Left            =   3840
            List            =   "frmRptFinancialFilterFields.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   450
            Width           =   1215
         End
         Begin VB.ComboBox cmbMonth 
            Height          =   330
            ItemData        =   "frmRptFinancialFilterFields.frx":0004
            Left            =   1770
            List            =   "frmRptFinancialFilterFields.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   480
            Width           =   1875
         End
         Begin VB.CheckBox chkRPDescrepancy 
            BackColor       =   &H8000000E&
            Caption         =   "RP Discrepancy"
            Height          =   480
            Left            =   30
            TabIndex        =   100
            Top             =   -30
            Width           =   1815
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            Height          =   210
            Left            =   3840
            TabIndex        =   103
            Top             =   210
            Width           =   420
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            Height          =   315
            Left            =   1830
            TabIndex        =   102
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.TextBox txtRpFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1155
         TabIndex        =   54
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtRpTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   56
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdRPShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1950
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpRPFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpRPTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   84
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   83
         Top             =   1215
         Width           =   735
      End
   End
   Begin VB.Frame fmeJournalBook 
      BackColor       =   &H80000009&
      Caption         =   "Journal Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3465
      TabIndex        =   31
      Top             =   1485
      Width           =   5820
      Begin VB.TextBox txtJBFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtJBTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   34
         Top             =   1455
         Width           =   1455
      End
      Begin VB.CommandButton cmdJBShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2070
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpJBFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpJBTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   77
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   76
         Top             =   1215
         Width           =   735
      End
   End
   Begin VB.Frame fmeBalanceSheet 
      BackColor       =   &H80000009&
      Caption         =   "Balance Sheet"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3720
      TabIndex        =   43
      Top             =   1935
      Width           =   5820
      Begin VB.TextBox txtBSTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2295
         TabIndex        =   44
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdBSShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2070
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpBSTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "As On"
         Height          =   210
         Left            =   2295
         TabIndex        =   80
         Top             =   1215
         Width           =   555
      End
   End
   Begin VB.Frame fmeCashBook 
      BackColor       =   &H80000009&
      Caption         =   "Cash Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3060
      TabIndex        =   9
      Top             =   810
      Width           =   5820
      Begin VB.CheckBox chkCashBookSummary 
         BackColor       =   &H8000000E&
         Caption         =   "Get Cash Book Summary"
         Height          =   480
         Left            =   120
         TabIndex        =   98
         Top             =   2730
         Width           =   2175
      End
      Begin VB.CommandButton cmdCBShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtCBTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCBFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpCBFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpCBTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label lblAccountHead 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "450100100    Cash"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   270
         Left            =   135
         TabIndex        =   67
         Top             =   585
         Width           =   5580
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   65
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   64
         Top             =   1215
         Width           =   975
      End
   End
   Begin VB.Frame fmeBankBook 
      BackColor       =   &H80000009&
      Caption         =   "Bank Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3195
      TabIndex        =   15
      Top             =   1035
      Width           =   5820
      Begin VB.CommandButton cmdBBBrowse 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   630
         Width           =   360
      End
      Begin VB.TextBox txtBBAccountHead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   990
         Width           =   4200
      End
      Begin VB.TextBox txtBBAccountHeadCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtBBFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   18
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtBBTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   20
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton cmdBBShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2610
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpBBFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpBBTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         Height          =   210
         Left            =   90
         TabIndex        =   70
         Top             =   675
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   69
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   68
         Top             =   1755
         Width           =   735
      End
   End
   Begin VB.Frame fmeSubLedger 
      BackColor       =   &H80000009&
      Caption         =   "Subsidiary Ledger Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   4170
      TabIndex        =   86
      Top             =   2640
      Width           =   5820
      Begin VB.CommandButton cmdSearchSubLedger 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   630
         Width           =   360
      End
      Begin VB.TextBox txtSubLedger 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   990
         Width           =   4200
      End
      Begin VB.TextBox txtSubLedgerCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   90
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtSubLedgerFromDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   89
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtSubLedgerToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   88
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton cmdShowSubLedger 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Cancel          =   -1  'True
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2565
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpSubFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpSubTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         Height          =   210
         Left            =   90
         TabIndex        =   97
         Top             =   675
         Width           =   1290
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   96
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   95
         Top             =   1755
         Width           =   735
      End
   End
   Begin VB.Timer tmrLocalBody 
      Interval        =   100
      Left            =   9990
      Top             =   5220
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   990
      Width           =   1320
   End
   Begin VB.CommandButton cmdSearchFund 
      BackColor       =   &H8000000A&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtFund 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   3495
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   450
      Width           =   4485
   End
   Begin VB.Frame fmeMenu 
      BackColor       =   &H8000000E&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   5955
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2895
      Begin VB.CommandButton cmdSubLedger 
         BackColor       =   &H80000000&
         Caption         =   "SubLedger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   5400
         Width           =   2760
      End
      Begin VB.CommandButton cmdReceiptPayment 
         BackColor       =   &H80000000&
         Caption         =   "Receipt && Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4650
         Width           =   2760
      End
      Begin VB.CommandButton cmdIncomeExpenditure 
         BackColor       =   &H80000000&
         Caption         =   "Income && Expenditure"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   2760
      End
      Begin VB.CommandButton cmdBalanceSheet 
         BackColor       =   &H80000000&
         Caption         =   "Balance Sheet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3330
         Width           =   2760
      End
      Begin VB.CommandButton cmdTrialBalance 
         BackColor       =   &H80000000&
         Caption         =   "Trial Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2655
         Width           =   2760
      End
      Begin VB.CommandButton cmdLedgerBook 
         BackColor       =   &H80000000&
         Caption         =   "Ledger Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1365
         Width           =   2760
      End
      Begin VB.CommandButton cmdJounalBook 
         BackColor       =   &H80000000&
         Caption         =   "Journal Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1995
         Width           =   2760
      End
      Begin VB.CommandButton cmdBankBook 
         BackColor       =   &H80000000&
         Caption         =   "Bank Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   765
         Width           =   2760
      End
      Begin VB.CommandButton cmdCashBook 
         BackColor       =   &H80000000&
         Caption         =   "Cash Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   2760
      End
   End
   Begin VB.Frame fmeLedgerBook 
      BackColor       =   &H80000009&
      Caption         =   "Ledger Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3330
      TabIndex        =   23
      Top             =   1260
      Width           =   5820
      Begin VB.CommandButton cmdLBShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2610
         Width           =   1320
      End
      Begin VB.TextBox txtLBTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   28
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtLBFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   26
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtLBAccountHeadCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   24
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtLBAccountHead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   990
         Width           =   4200
      End
      Begin VB.CommandButton cmdLBBrowse 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   630
         Width           =   360
      End
      Begin MSComCtl2.DTPicker dtpLBFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpLBTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1980
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   75
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   74
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         Height          =   210
         Left            =   90
         TabIndex        =   73
         Top             =   675
         Width           =   1290
      End
   End
   Begin VB.Frame fmeTrialBalance 
      BackColor       =   &H80000009&
      Caption         =   "Trial Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3585
      TabIndex        =   37
      Top             =   1710
      Width           =   5820
      Begin VB.CommandButton cmdTBShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2070
         Width           =   1320
      End
      Begin VB.TextBox txtTBTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   40
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtTBFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1155
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpTBFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpTBTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   79
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   78
         Top             =   1215
         Width           =   975
      End
   End
   Begin VB.Frame fmeIncomeExpenditure 
      BackColor       =   &H80000009&
      Caption         =   "Income && Expenditure"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3300
      Left            =   3885
      TabIndex        =   47
      Top             =   2145
      Width           =   5820
      Begin VB.TextBox txtIEFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   48
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtIETo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   50
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdIEShow 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2070
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpIEFrom 
         Height          =   315
         Left            =   2655
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin MSComCtl2.DTPicker dtpIETo 
         Height          =   315
         Left            =   4500
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483638
         Format          =   61145089
         CurrentDate     =   39612
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   210
         Left            =   1170
         TabIndex        =   82
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   210
         Left            =   3015
         TabIndex        =   81
         Top             =   1215
         Width           =   735
      End
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---Date---"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   9630
      TabIndex        =   66
      Top             =   495
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   210
      Left            =   3000
      TabIndex        =   63
      Top             =   495
      Width           =   480
   End
   Begin VB.Label lblLocalBody 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--Local Body--"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2970
      TabIndex        =   60
      Top             =   45
      Width           =   7620
   End
End
Attribute VB_Name = "frmRptFinancialFilterFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mLoop As Integer
    Dim frmNewRpt As New frmRptViewer
    Dim arInput As Variant
    Dim frmNewViewer As New frmRptViewer
    
    Private Sub cmbMonth_Click()
          Dim mMonthIndex As Integer
        If chkRPDescrepancy.value = 1 Then
            '---------------------------------------------------------------------------------'
            'Note:- Finding Range of Dates According the month selected
            '---------------------------------------------------------------------------------'
            
            mMonthIndex = cmbMonth.ItemData(cmbMonth.ListIndex)
            If gbLBPanchayat <> 1 Then
            If mMonthIndex > 3 Then
                txtRpFrom.Text = CheckDateInMMM(DateSerial(gbFinancialYearID, mMonthIndex, 1))
            Else
                txtRpTo.Text = CheckDateInMMM(DateSerial(gbFinancialYearID + 1, mMonthIndex, 1))
            End If
            
            End If
            
            If mMonthIndex > 3 Then
                txtRpFrom.Text = CheckDateInMMM(DateSerial(gbFinancialYearID, mMonthIndex + 1, 1) - 1)
            Else
                txtRpTo.Text = CheckDateInMMM(DateSerial(gbFinancialYearID + 1, mMonthIndex + 1, 1) - 1)
            End If
        End If
    End Sub
    Private Sub fillMonthCombo()
        cmbMonth.AddItem "April"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 4
        cmbMonth.AddItem "May"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 5
        cmbMonth.AddItem "June"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 6
        cmbMonth.AddItem "July"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 7
        cmbMonth.AddItem "August"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 8
        cmbMonth.AddItem "September"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 9
        cmbMonth.AddItem "October"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 10
        cmbMonth.AddItem "November"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 11
        cmbMonth.AddItem "December"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 12
        cmbMonth.AddItem "January"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 1
        cmbMonth.AddItem "February"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 2
        cmbMonth.AddItem "March"
        cmbMonth.ItemData(cmbMonth.NewIndex) = 3
    End Sub
    Public Sub cmdBalanceSheet_Click()
        Call fmeBalanceSheet_Click
    End Sub

    Public Sub cmdBankBook_Click()
        Call fmeBankBook_Click
    End Sub

    Private Sub cmdBBBrowse_Click()
        Dim mSql As String
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = 2"
        
        mSql = " Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
        mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID"
        mSql = mSql + " WHERE faAccountHeads.intGroupID = 2 Order By vchAccountHeadCode"
                
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        txtBBAccountHeadCode.SetFocus
    End Sub

    Private Sub cmdBBShow_Click()
        arInput = Array(val(txtBBAccountHeadCode.Tag), CDate(txtBBFrom.Text), CDate(txtBBTo.Text))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptBankBook.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Private Sub cmdBSShow_Click()
        arInput = Array(CDate(txtBSTo.Text))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptB1Schedule.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
        Dim arrInput As Variant
        arrInput = Array(arInput(0), val(txtFund))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptBalanceSheetSchedule.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arrInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
        Dim frmNewViewer1 As New frmRptViewer
        'arInput = Array(CDate(txtToDate.Text))
        frmNewViewer1.rptFileName = App.Path & "\Reports\rptBalanceSheet.rpt"
        frmNewViewer1.WindowState = vbMaximized
        frmNewViewer1.WindowState = vbMaximized
        frmNewViewer1.InputParameters = arrInput
        Call frmNewViewer1.ShowReport
        frmNewViewer1.Show
    End Sub

    Public Sub cmdCashBook_Click()
        Call fmeCashBook_Click
    End Sub

    Private Sub cmdCBShow_Click()
        arInput = Array(val(lblAccountHead.Tag), CDate(txtCBFrom.Text), CDate(txtCBTo.Text))
        If chkCashBookSummary.value = 1 Then
            frmNewRpt.rptFileName = App.Path & "\Reports\rptCashBookSummary.rpt"
        Else
            frmNewRpt.rptFileName = App.Path & "\Reports\rptCashBook.rpt"
        End If
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Private Sub cmdClose_Click()
        Unload Me
    End Sub

    Private Sub cmdIEShow_Click()
        arInput = Array(CDate(txtIEFrom.Text), CDate(txtIETo.Text), val(txtFund.Tag))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptIESchedules.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
        frmNewViewer.rptFileName = App.Path & "\Reports\rptIncomeAndExpenditure.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Public Sub cmdIncomeExpenditure_Click()
        Call fmeIncomeExpenditure_Click
    End Sub

    Private Sub cmdJBShow_Click()
        arInput = Array(CDate(txtJBFrom.Text), CDate(txtJBTo.Text), gbFundID)
        frmNewRpt.rptFileName = App.Path & "\Reports\rptJournalBook.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Public Sub cmdJounalBook_Click()
        Call fmeJournalBook_Click
    End Sub

    Private Sub cmdLBBrowse_Click()
        Dim mSql As String
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads"
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
        txtLBAccountHeadCode.SetFocus
    End Sub

    Private Sub cmdLBShow_Click()
        arInput = Array(val(txtLBAccountHeadCode.Tag), CDate(txtLBFrom.Text), CDate(txtLBTo.Text))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptGeneralLedger.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Public Sub cmdLedgerBook_Click()
        Call fmeLedgerBook_Click
    End Sub

    Public Sub cmdReceiptPayment_Click()
        Call fmeReceiptPayment_Click
    End Sub

    Private Sub cmdRPShow_Click()
        Dim mYear As Integer
        Dim mMonth  As Integer
        
        If chkRPDescrepancy.value = 1 Then
            If cmbMonth.ListIndex > 0 Then
                mMonth = cmbMonth.ItemData(cmbMonth.ListIndex)
                
            Else
                MsgBox "Please Select Month", vbApplicationModal
                Exit Sub
            End If
            
            If cmbYear.ListIndex > 0 Then
                mYear = cmbYear.ItemData(cmbYear.ListIndex)
            Else
                 MsgBox "Please Select Year", vbApplicationModal
                Exit Sub
            End If
            arInput = Array(mMonth, mYear)
            frmNewViewer.rptFileName = App.Path & "\Reports\rptRPDescrepancy.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show
            
            
        Else
            arInput = Array(CDate(txtRpFrom.Text), CDate(txtRpTo.Text), val(txtFund.Tag))
            frmNewRpt.rptFileName = App.Path & "\Reports\rptRPSchedules.rpt"
            frmNewRpt.WindowState = vbMaximized
            frmNewRpt.InputParameters = arInput
            Call frmNewRpt.ShowReport
            frmNewRpt.Show
            
            frmNewViewer.rptFileName = App.Path & "\Reports\rptRP.rpt"
            frmNewViewer.WindowState = vbMaximized
            frmNewViewer.InputParameters = arInput
            Call frmNewViewer.ShowReport
            frmNewViewer.Show
            
        End If
        
    End Sub

    Private Sub cmdSearchSubLedger_Click()
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtSubLedgerCode.Text = gbSearchCode
        txtSubLedgerCode.Tag = gbSearchID
        txtSubLedger.Text = gbSearchStr
        gbSearchID = -1
        gbSearchStr = ""
    End Sub

    Private Sub cmdShowSubLedger_Click()
        If val(txtSubLedgerCode.Tag) = 0 Then
            MsgBox "Please Select the SubLedger", vbInformation
            cmdSearchSubLedger.Visible = True
            Exit Sub
        End If
        If txtSubLedgerFromDate = "" Then
            MsgBox "Please Give the From Date", vbInformation
            txtSubLedgerFromDate.SetFocus
            Exit Sub
        End If
        If txtSubLedgerToDate = "" Then
            MsgBox "Please Give the To Date", vbInformation
            txtSubLedgerToDate.SetFocus
            Exit Sub
        End If
        arInput = Array(CDate(txtSubLedgerFromDate.Text), CDate(txtSubLedgerToDate.Text), val(txtSubLedgerCode.Tag))
        frmNewRpt.rptFileName = App.Path & "\Reports\rptSubLedger.rpt"
        frmNewRpt.WindowState = vbMaximized
        frmNewRpt.InputParameters = arInput
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Private Sub cmdSubLedger_Click()
        Call fmeSubLedger_Click
    End Sub

    Private Sub cmdTBShow_Click()
        arInput = Array(CDate(txtTBFrom.Text), CDate(txtTBTo.Text), val(txtFund.Tag))
                frmNewRpt.rptFileName = App.Path & "\Reports\rptLedgerTrialBalance.rpt"
                frmNewRpt.WindowState = vbMaximized
                frmNewRpt.InputParameters = arInput
                Call frmNewRpt.ShowReport
                frmNewRpt.Show
    End Sub

    Public Sub cmdTrialBalance_Click()
        Call fmeTrialBalance_Click
    End Sub





Private Sub cmdYear_Click()

End Sub

    Private Sub dtpBBFrom_CloseUp()
        If CDate(dtpBBFrom.value) Then
        If CDate(dtpBBTo.value) Then
                If CDate(dtpBBFrom.value) > CDate(dtpBBTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpBBFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtBBFrom.Text = CheckDateInMMM(dtpBBFrom.value)
        End If
        txtBBFrom.Text = DdMmmYy(dtpBBFrom.value)
    End Sub
    Private Sub dtpBBTo_CloseUp()
        If CDate(dtpBBTo.value) Then
            If CDate(dtpBBFrom.value) Then
                If CDate(dtpBBFrom.value) > CDate(dtpBBTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpBBTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtBBTo.Text = CheckDateInMMM(dtpBBTo.value)
        End If
        txtBBTo.Text = DdMmmYy(dtpBBTo.value)
    End Sub



    Private Sub dtpBSTo_CloseUp()
        txtBSTo.Text = DdMmmYy(dtpBSTo.value)
    End Sub

    Private Sub dtpCBFrom_CloseUp()
        If CDate(dtpCBFrom.value) Then
        If CDate(dtpCBTo.value) Then
                If CDate(dtpCBFrom.value) > CDate(dtpCBTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpCBFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtCBFrom.Text = CheckDateInMMM(dtpCBFrom.value)
        End If
        txtCBFrom.Text = DdMmmYy(dtpCBFrom.value)
    End Sub

    Private Sub dtpCBTo_CloseUp()
        If CDate(dtpCBTo.value) Then
            If CDate(dtpCBFrom.value) Then
                If CDate(dtpCBFrom.value) > CDate(dtpCBTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpCBTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtCBTo.Text = CheckDateInMMM(dtpCBTo.value)
        End If
        txtCBTo.Text = DdMmmYy(dtpCBTo.value)
    End Sub


    Private Sub dtpIEFrom_CloseUp()
        If CDate(dtpIEFrom.value) Then
        If CDate(dtpIETo.value) Then
                If CDate(dtpIEFrom.value) > CDate(dtpIETo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpIEFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtIEFrom.Text = CheckDateInMMM(dtpIEFrom.value)
        End If
        txtIEFrom.Text = DdMmmYy(dtpIEFrom.value)
    End Sub
    Private Sub dtpIETo_CloseUp()
        If CDate(dtpIETo.value) Then
            If CDate(dtpIEFrom.value) Then
                If CDate(dtpIEFrom.value) > CDate(dtpIETo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpIETo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtIETo.Text = CheckDateInMMM(dtpIETo.value)
        End If
        txtIETo.Text = DdMmmYy(dtpIETo.value)
    End Sub
    Private Sub dtpJBFrom_CloseUp()
        If CDate(dtpJBFrom.value) Then
        If CDate(dtpJBTo.value) Then
                If CDate(dtpJBFrom.value) > CDate(dtpJBTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpJBFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtJBFrom.Text = CheckDateInMMM(dtpJBFrom.value)
        End If
        txtJBFrom.Text = DdMmmYy(dtpJBFrom.value)
    End Sub

    Private Sub dtpJBTo_CloseUp()
        If CDate(dtpJBTo.value) Then
            If CDate(dtpJBFrom.value) Then
                If CDate(dtpJBFrom.value) > CDate(dtpJBTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpJBTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtJBTo.Text = CheckDateInMMM(dtpJBTo.value)
        End If
        txtJBTo.Text = DdMmmYy(dtpJBTo.value)
    End Sub


    Private Sub dtpLBFrom_CloseUp()
        If CDate(dtpLBFrom.value) Then
        If CDate(dtpLBTo.value) Then
                If CDate(dtpLBFrom.value) > CDate(dtpLBTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpLBFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtLBFrom.Text = CheckDateInMMM(dtpLBFrom.value)
        End If
        txtLBFrom.Text = DdMmmYy(dtpLBFrom.value)
    End Sub
    Private Sub dtpLBTo_CloseUp()
        If CDate(dtpLBTo.value) Then
            If CDate(dtpLBFrom.value) Then
                If CDate(dtpLBFrom.value) > CDate(dtpLBTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpLBTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtLBTo.Text = CheckDateInMMM(dtpLBTo.value)
        End If
        txtLBTo.Text = DdMmmYy(dtpLBTo.value)
    End Sub

    Private Sub dtpRPFrom_CloseUp()
        If CDate(dtpRPFrom.value) Then
        If CDate(dtpRPTo.value) Then
                If CDate(dtpRPFrom.value) > CDate(dtpRPTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpRPFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtRpFrom.Text = CheckDateInMMM(dtpRPFrom.value)
        End If
        txtRpFrom.Text = DdMmmYy(dtpRPFrom.value)
    End Sub
    Private Sub dtpRPTo_CloseUp()
        If CDate(dtpRPTo.value) Then
            If CDate(dtpRPFrom.value) Then
                If CDate(dtpRPFrom.value) > CDate(dtpRPTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpRPTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtRpTo.Text = CheckDateInMMM(dtpRPTo.value)
        End If
        txtRpTo.Text = DdMmmYy(dtpRPTo.value)
    End Sub

    Private Sub dtpSubFrom_CloseUp()
        If CDate(dtpSubFrom.value) Then
        If CDate(dtpSubTo.value) Then
                If CDate(dtpSubFrom.value) > CDate(dtpSubTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpSubFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtSubLedgerFromDate.Text = CheckDateInMMM(dtpSubFrom.value)
        End If
        txtSubLedgerFromDate.Text = DdMmmYy(dtpSubFrom.value)
    End Sub

    Private Sub dtpSubTo_CloseUp()
        If CDate(dtpSubTo.value) Then
            If CDate(dtpSubFrom.value) Then
                If CDate(dtpSubFrom.value) > CDate(dtpSubTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpSubTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtSubLedgerToDate.Text = CheckDateInMMM(dtpSubTo.value)
        End If
        txtSubLedgerToDate.Text = DdMmmYy(dtpSubTo.value)
    End Sub



    Private Sub dtpTBFrom_CloseUp()
        If CDate(dtpTBFrom.value) Then
        If CDate(dtpTBTo.value) Then
                If CDate(dtpTBFrom.value) > CDate(dtpTBTo.value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpTBFrom.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtTBFrom.Text = CheckDateInMMM(dtpTBFrom.value)
        End If
        txtTBFrom.Text = DdMmmYy(dtpTBFrom.value)
    End Sub

    Private Sub dtpTBTo_CloseUp()
        If CDate(dtpTBTo.value) Then
            If CDate(dtpTBFrom.value) Then
                If CDate(dtpTBFrom.value) > CDate(dtpTBTo.value) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    dtpTBTo.value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtTBTo.Text = CheckDateInMMM(dtpTBTo.value)
        End If
        txtTBTo.Text = DdMmmYy(dtpTBTo.value)
    End Sub

    Private Sub fmeBankBook_Click()
        Call GetOrderFrames
        fmeBankBook.ZOrder (0)
    End Sub

    Private Sub fmeCashBook_Click()
        Call GetOrderFrames
        fmeCashBook.ZOrder (0)
    End Sub

    Private Sub fmeLedgerBook_Click()
        Call GetOrderFrames
        fmeLedgerBook.ZOrder (0)
    End Sub
    
    Private Sub fmeJournalBook_Click()
        Call GetOrderFrames
        fmeJournalBook.ZOrder (0)
    End Sub
    
    Private Sub fmeTrialBalance_Click()
        Call GetOrderFrames
        fmeTrialBalance.ZOrder (0)
    End Sub
    
    Private Sub fmeBalanceSheet_Click()
        Call GetOrderFrames
        fmeBalanceSheet.ZOrder (0)
    End Sub
    
    Private Sub fmeIncomeExpenditure_Click()
        Call GetOrderFrames
        fmeIncomeExpenditure.ZOrder (0)
    End Sub
    
    Private Sub fmeReceiptPayment_Click()
        Call GetOrderFrames
        fmeReceiptPayment.ZOrder (0)
    End Sub
    
    Private Sub fmeSubLedger_Click()
        Call GetOrderFrames
        fmeSubLedger.ZOrder (0)
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        Me.ZOrder (0)
    End Sub

    Private Sub Form_Load()
        Call FormInitilise
        Call fillMonthCombo
        Call Fillyear
        If gbLocalBodyID = 167 Then
            fmeRPDiscrepancy.Visible = True
        End If
    End Sub
    Private Sub Fillyear()
        Call PopulateList(cmbYear, "SELECT intFinancialYear,intFinancialYearID FROM faFinancialYear", , , True, True)
    End Sub
    Private Sub FormInitilise()
        Dim objAcc As New clsAccounts
        Dim objFund As New clsFund
        Dim mCounter As Integer
        lblLocalBody.Caption = GetLocalBodyName()
        For mCounter = 1 To 70
            lblLocalBody.Caption = lblLocalBody.Caption + " "
        Next mCounter
        lblLocalBody.Caption = Left(lblLocalBody.Caption, 70)
        lblDate.Caption = Format(gbTransactionDate, "dd/MMM/YYYY")
        objFund.SetFund (gbFundID)
        objAcc.SetAccounts (gbAcHeadCodeCash) ' ("450100100")
        lblAccountHead.Tag = objAcc.AccountHeadID
        lblAccountHead.Caption = objAcc.AccountCode + "    " + objAcc.AccountHead
        If objFund.FundID > -1 Then
            txtFund.Text = objFund.FundName
            txtFund.Tag = objFund.FundID
        Else
            txtFund.Text = ""
            txtFund.Tag = ""
        End If
        txtCBFrom.Text = DdMmmYy(gbStartingDate)
        txtBBFrom.Text = DdMmmYy(gbStartingDate)
        txtLBFrom.Text = DdMmmYy(gbStartingDate)
        txtJBFrom.Text = DdMmmYy(gbStartingDate)
        txtTBFrom.Text = DdMmmYy(gbStartingDate)
        txtIEFrom.Text = DdMmmYy(gbStartingDate)
        txtRpFrom.Text = DdMmmYy(gbStartingDate)
        
        dtpCBFrom.value = DdMmmYy(gbStartingDate)
        dtpBBFrom.value = DdMmmYy(gbStartingDate)
        dtpLBFrom.value = DdMmmYy(gbStartingDate)
        dtpTBFrom.value = DdMmmYy(gbStartingDate)
        dtpJBFrom.value = DdMmmYy(gbStartingDate)
        dtpIEFrom.value = DdMmmYy(gbStartingDate)
        dtpRPFrom.value = DdMmmYy(gbStartingDate)
        
        txtCBTo.Text = DdMmmYy(gbTransactionDate)
        txtBBTo.Text = DdMmmYy(gbTransactionDate)
        txtLBTo.Text = DdMmmYy(gbTransactionDate)
        txtTBTo.Text = DdMmmYy(gbTransactionDate)
        txtJBTo.Text = DdMmmYy(gbTransactionDate)
        txtBSTo.Text = DdMmmYy(gbTransactionDate)
        txtIETo.Text = DdMmmYy(gbTransactionDate)
        txtRpTo.Text = DdMmmYy(gbTransactionDate)
        
        dtpCBTo.value = DdMmmYy(gbTransactionDate)
        dtpBBTo.value = DdMmmYy(gbTransactionDate)
        dtpLBTo.value = DdMmmYy(gbTransactionDate)
        dtpJBTo.value = DdMmmYy(gbTransactionDate)
        dtpTBTo.value = DdMmmYy(gbTransactionDate)
        dtpBSTo.value = DdMmmYy(gbTransactionDate)
        dtpIETo.value = DdMmmYy(gbTransactionDate)
        dtpRPTo.value = DdMmmYy(gbTransactionDate)
        dtpSubFrom.value = DdMmmYy(gbTransactionDate)
        dtpSubTo.value = DdMmmYy(gbTransactionDate)
        
        cmdShowSubLedger.Enabled = False  'Added on 19.10.2011 By Poornima
        
        Call GetOrderFrames
        
    End Sub
    
    Private Function GetLocalBodyName() As String
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim mVOut As Variant
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "Connection to Saankhya Not Present", vbCritical
            Exit Function
        End If
        objdb.ExecuteSP "Select IsNull(chvTitle,'') [vchTitle] From faLBSettings", , mVOut, , mCnn, adCmdText
        GetLocalBodyName = CStr(mVOut(0, 0))
    End Function


    Private Sub tmrLocalBody_Timer()
        lblLocalBody.Caption = Right(lblLocalBody.Caption, Len(lblLocalBody) - 1) + Left(lblLocalBody.Caption, 1)
    End Sub

    Private Sub txtBBAccountHeadCode_GotFocus()
        If Len(gbSearchStr) Then
            Dim objAccHead As New clsAccounts
            objAccHead.SetAccountCode (Token(gbSearchStr, " "))
            If objAccHead.AccountHeadID > 0 Then
                txtBBAccountHeadCode.Text = objAccHead.AccountCode
                txtBBAccountHeadCode.Tag = objAccHead.AccountHeadID
                txtBBAccountHead = objAccHead.AccountHead
                txtBBAccountHead.ToolTipText = txtBBAccountHead.Text
            End If
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub txtBBAccountHeadCode_LostFocus()
        If Trim(txtBBAccountHeadCode.Text) <> "" Then
            Dim objAcc As New clsAccounts
            objAcc.SetAccounts (Trim(txtBBAccountHeadCode.Text))
            txtBBAccountHeadCode.Tag = objAcc.AccountHeadID
            txtBBAccountHead.Text = objAcc.AccountHead
            txtBBAccountHead.ToolTipText = txtBBAccountHead.Text
        End If
    End Sub

    Private Sub txtCBFrom_LostFocus()
        txtCBFrom.Text = CheckDateInMMM(txtCBFrom.Text)
        If CDate(txtCBFrom.Text) Then
            If CDate(txtCBTo.Text) Then
                If CDate(txtCBFrom.Text) > CDate(txtCBTo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtCBFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtCBFrom.Text = CheckDateInMMM(txtCBFrom.Text)
        End If
    End Sub

    Private Sub txtBBFrom_LostFocus()
        txtBBFrom.Text = CheckDateInMMM(txtBBFrom.Text)
    If CDate(txtBBFrom.Text) Then
                If CDate(txtBBTo.Text) Then
                    If CDate(txtBBFrom.Text) > CDate(txtBBTo.Text) Then
                        MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                        txtBBFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                        Exit Sub
                    End If
                Else
                    MsgBox "Please enter To date", vbInformation
                    Exit Sub
                End If
            Else
                txtBBFrom.Text = CheckDateInMMM(txtBBFrom.Text)
            End If
        End Sub
    
    Private Sub txtLBFrom_LostFocus()
        txtLBFrom.Text = CheckDateInMMM(txtLBFrom.Text)
        If CDate(txtLBFrom.Text) Then
            If CDate(txtLBTo.Text) Then
                If CDate(txtLBFrom.Text) > CDate(txtLBTo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtLBFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtLBFrom.Text = CheckDateInMMM(txtLBFrom.Text)
        End If
    End Sub
    
    Private Sub txtSubLedgerFromDate_GotFocus()
        txtSubLedgerFromDate.SelStart = 0
        txtSubLedgerFromDate.SelLength = Len(txtSubLedgerFromDate)
    End Sub

    Private Sub txtSubLedgerFromDate_LostFocus()
        txtSubLedgerFromDate.Text = CheckDateInMMM(txtSubLedgerFromDate.Text)
        If CDate(txtSubLedgerFromDate.Text) Then
            If CDate(txtSubLedgerToDate.Text) Then
                If CDate(txtSubLedgerFromDate.Text) > CDate(txtSubLedgerToDate.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtSubLedgerFromDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtSubLedgerFromDate.Text = CheckDateInMMM(txtSubLedgerFromDate.Text)
        End If
    End Sub
    
    Private Sub txtSubLedgerToDate_GotFocus()
        txtSubLedgerToDate.SelStart = 0
        txtSubLedgerToDate.SelLength = Len(txtSubLedgerToDate)
    End Sub

    Private Sub txtSubLedgerToDate_LostFocus()
        txtSubLedgerToDate.Text = CheckDateInMMM(txtSubLedgerToDate.Text)
        If CDate(txtSubLedgerToDate.Text) Then
            If CDate(txtSubLedgerFromDate.Text) Then
                If CDate(txtSubLedgerFromDate.Text) > CDate(txtSubLedgerToDate.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtSubLedgerFromDate.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtSubLedgerToDate.Text = CheckDateInMMM(txtSubLedgerFromDate.Text)
        End If
    End Sub

    Private Sub txtTBFrom_LostFocus()
        txtTBFrom.Text = CheckDateInMMM(txtTBFrom.Text)
        If CDate(txtTBFrom.Text) Then
            If CDate(txtTBTo.Text) Then
                If CDate(txtTBFrom.Text) > CDate(txtTBTo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtTBFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtTBFrom.Text = CheckDateInMMM(txtTBTo.Text)
        End If
    End Sub
    
    Private Sub txtJBFrom_LostFocus()
    txtJBFrom.Text = CheckDateInMMM(txtJBFrom.Text)
    If CDate(txtJBFrom.Text) Then
            If CDate(txtJBTo.Text) Then
                If CDate(txtJBFrom.Text) > CDate(txtJBTo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtJBFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtJBFrom.Text = CheckDateInMMM(txtJBFrom.Text)
        End If
    End Sub
    
    Private Sub txtIEFrom_LostFocus()
        txtIEFrom.Text = CheckDateInMMM(txtIEFrom.Text)
        If CDate(txtIEFrom.Text) Then
            If CDate(txtIETo.Text) Then
                If CDate(txtIEFrom.Text) > CDate(txtIETo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtIEFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtIEFrom.Text = CheckDateInMMM(txtIEFrom.Text)
        End If
    End Sub
    
    Private Sub txtRPFrom_LostFocus()
        txtRpFrom.Text = CheckDateInMMM(txtRpFrom.Text)
        If CDate(txtRpFrom.Text) Then
            If CDate(txtRpTo.Text) Then
                If CDate(txtRpFrom.Text) > CDate(txtRpTo.Text) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    txtRpFrom.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            txtRpFrom.Text = CheckDateInMMM(txtRpFrom.Text)
        End If
    End Sub
    
    Private Sub txtCBTo_LostFocus()
        txtCBTo.Text = CheckDateInMMM(txtCBFrom.Text)
        If CDate(txtCBTo.Text) Then
            If CDate(txtCBFrom.Text) Then
                If CDate(txtCBFrom.Text) > CDate(txtCBTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtCBTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtCBTo.Text = CheckDateInMMM(txtCBFrom.Text)
        End If
    End Sub
    
    Private Sub txtBBTo_LostFocus()
        txtBBTo.Text = CheckDateInMMM(txtBBTo.Text)
        If CDate(txtBBTo.Text) Then
            If CDate(txtBBFrom.Text) Then
                If CDate(txtBBFrom.Text) > CDate(txtBBTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtBBTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtBBTo.Text = CheckDateInMMM(txtBBFrom.Text)
        End If
    End Sub
    
    Private Sub txtLBTo_LostFocus()
        txtLBTo.Text = CheckDateInMMM(txtLBTo.Text)
        If CDate(txtLBTo.Text) Then
            If CDate(txtLBFrom.Text) Then
                If CDate(txtLBFrom.Text) > CDate(txtLBTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtLBTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtLBTo.Text = CheckDateInMMM(txtLBTo.Text)
        End If
    End Sub
    
    Private Sub txtJBTo_LostFocus()
        txtJBTo.Text = CheckDateInMMM(txtJBTo.Text)
        If CDate(txtJBTo.Text) Then
            If CDate(txtJBFrom.Text) Then
                If CDate(txtJBFrom.Text) > CDate(txtJBTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtJBTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtJBTo.Text = CheckDateInMMM(txtJBFrom.Text)
        End If
    End Sub
    
    Private Sub txtTBTo_LostFocus()
        txtTBTo.Text = CheckDateInMMM(txtTBTo.Text)
        If CDate(txtTBTo.Text) Then
            If CDate(txtTBFrom.Text) Then
                If CDate(txtTBFrom.Text) > CDate(txtTBTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtTBTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtTBTo.Text = CheckDateInMMM(txtTBTo.Text)
        End If
    End Sub
    
    Private Sub txtBSTo_LostFocus()
        txtBSTo.Text = CheckDateInMMM(txtBSTo.Text)
    End Sub
    
    Private Sub txtIETo_LostFocus()
        txtIETo.Text = CheckDateInMMM(txtIETo.Text)
        If CDate(txtIETo.Text) Then
            If CDate(txtIEFrom.Text) Then
                If CDate(txtIEFrom.Text) > CDate(txtIETo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtIETo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtIETo.Text = CheckDateInMMM(txtIETo.Text)
        End If
    End Sub
    
    Private Sub txtRPTo_LostFocus()
        txtRpTo.Text = CheckDateInMMM(txtRpTo.Text)
        If CDate(txtRpTo.Text) Then
            If CDate(txtRpFrom.Text) Then
                If CDate(txtRpFrom.Text) > CDate(txtRpTo.Text) Then
                    MsgBox "Please Enter a value Greater than or equal to To Date", vbInformation
                    txtRpTo.Text = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please Enter From date", vbInformation
                Exit Sub
            End If
        Else
            txtRpTo.Text = CheckDateInMMM(txtRpFrom.Text)
        End If
    End Sub
    
    Private Sub txtLBAccountHeadCode_GotFocus()
        If Len(gbSearchStr) Then
            Dim objAccHead As New clsAccounts
            objAccHead.SetAccountCode (Token(gbSearchStr, " "))
            If objAccHead.AccountHeadID > 0 Then
                txtLBAccountHeadCode.Text = objAccHead.AccountCode
                txtLBAccountHeadCode.Tag = objAccHead.AccountHeadID
                txtLBAccountHead = objAccHead.AccountHead
                txtLBAccountHead.ToolTipText = txtLBAccountHead.Text
            End If
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub txtLBAccountHeadCode_LostFocus()
        If Trim(txtLBAccountHeadCode.Text) <> "" Then
            Dim objAcc As New clsAccounts
            objAcc.SetAccounts (Trim(txtLBAccountHeadCode.Text))
            txtLBAccountHeadCode.Tag = objAcc.AccountHeadID
            txtLBAccountHead.Text = objAcc.AccountHead
            txtLBAccountHead.ToolTipText = txtLBAccountHead.Text
        End If
    End Sub
    
    Private Sub GetOrderFrames()
        fmeCashBook.ZOrder (0)
        fmeBankBook.ZOrder (0)
        fmeLedgerBook.ZOrder (0)
        fmeJournalBook.ZOrder (0)
        fmeTrialBalance.ZOrder (0)
        fmeBalanceSheet.ZOrder (0)
        fmeIncomeExpenditure.ZOrder (0)
        fmeReceiptPayment.ZOrder (0)
        chkCashBookSummary.value = 0
    End Sub


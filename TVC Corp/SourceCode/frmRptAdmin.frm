VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report to Administrator"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeZonalCollection 
      Caption         =   "ZonalCollection"
      Height          =   1335
      Left            =   90
      TabIndex        =   107
      Top             =   9360
      Width           =   5505
      Begin VB.ComboBox cmbzone 
         Height          =   390
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   240
         Width           =   3420
      End
      Begin VB.CommandButton cmdZonalCollection 
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4620
         TabIndex        =   108
         Top             =   720
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpZonedateFrom 
         Height          =   345
         Left            =   570
         TabIndex        =   110
         Top             =   720
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   44097
      End
      Begin MSComCtl2.DTPicker dtpZonedateTo 
         Height          =   345
         Left            =   2745
         TabIndex        =   111
         Top             =   735
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   44097
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zonal Office-"
         Height          =   270
         Left            =   540
         TabIndex        =   114
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   2370
         TabIndex        =   112
         Top             =   750
         Width           =   300
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   120
         TabIndex        =   109
         Top             =   780
         Width           =   375
      End
   End
   Begin VB.Frame fmeCardPayment 
      Caption         =   "CardPayment Vouchers"
      Height          =   1350
      Left            =   5640
      TabIndex        =   99
      Top             =   8940
      Width           =   5445
      Begin VB.CommandButton cmdCard 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   100
         Top             =   795
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpCardFrom 
         Height          =   345
         Left            =   1215
         TabIndex        =   101
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   43555
      End
      Begin MSComCtl2.DTPicker dtpCardTo 
         Height          =   345
         Left            =   3615
         TabIndex        =   102
         Top             =   315
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   43555
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   615
         TabIndex        =   104
         Top             =   345
         Width           =   525
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3180
         TabIndex        =   103
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame fmePreviousDatesReceiptCancellation 
      Caption         =   "Previous Date's Receipt Cancellation"
      Height          =   1230
      Left            =   120
      TabIndex        =   92
      Top             =   8070
      Width           =   5445
      Begin VB.CommandButton cmdfmePreviousDatesReceiptCancellation 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   93
         Top             =   795
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpPredateFrom 
         Height          =   345
         Left            =   1950
         TabIndex        =   94
         Top             =   300
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpPredateTo 
         Height          =   345
         Left            =   3915
         TabIndex        =   95
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3540
         TabIndex        =   97
         Top             =   330
         Width           =   300
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1305
         TabIndex        =   96
         Top             =   345
         Width           =   525
      End
   End
   Begin VB.Frame fmeReceiptCount 
      Caption         =   "ReceiptCount"
      Height          =   1620
      Left            =   5700
      TabIndex        =   80
      Top             =   7410
      Width           =   5445
      Begin VB.CommandButton cmdReceiptCount 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   83
         Top             =   1110
         Width           =   795
      End
      Begin VB.ComboBox cmbRecceiptCountUser 
         Height          =   390
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1080
         Width           =   2130
      End
      Begin VB.ComboBox cmbReceiptCounter 
         Height          =   390
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   1080
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker dtpCountFrom 
         Height          =   345
         Left            =   2040
         TabIndex        =   84
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208320
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpCountTo 
         Height          =   345
         Left            =   2070
         TabIndex        =   85
         Top             =   585
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpTimeFrom 
         Height          =   345
         Left            =   3555
         TabIndex        =   89
         Top             =   225
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208322
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpTimeTo 
         Height          =   345
         Left            =   3555
         TabIndex        =   90
         Top             =   585
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208322
         CurrentDate     =   39808
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1485
         TabIndex        =   88
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   1740
         TabIndex        =   87
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter"
         Height          =   270
         Left            =   90
         TabIndex        =   86
         Top             =   855
         Width           =   690
      End
   End
   Begin VB.Frame fraChitta 
      Caption         =   "Chitta"
      Height          =   1335
      Left            =   5610
      TabIndex        =   69
      Top             =   2550
      Width           =   5445
      Begin VB.ComboBox cmbChittaCounter 
         Height          =   390
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   585
         Width           =   2040
      End
      Begin VB.CommandButton cmdChitta 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4545
         TabIndex        =   71
         Top             =   945
         Width           =   795
      End
      Begin VB.ComboBox cmbUser 
         Height          =   390
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   585
         Width           =   2130
      End
      Begin MSComCtl2.DTPicker dtpChittaFromDate 
         Height          =   330
         Left            =   825
         TabIndex        =   72
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin MSComCtl2.DTPicker dtpChittaToDate 
         Height          =   330
         Left            =   3060
         TabIndex        =   73
         Top             =   195
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date-"
         Height          =   270
         Left            =   360
         TabIndex        =   75
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter -"
         Height          =   270
         Left            =   30
         TabIndex        =   74
         Top             =   630
         Width           =   780
      End
   End
   Begin VB.Frame fmeheadwiseReport 
      Caption         =   "Head wise report"
      Height          =   1575
      Left            =   120
      TabIndex        =   60
      Top             =   6630
      Width           =   5475
      Begin VB.CheckBox chkRent 
         Caption         =   "Rent"
         Height          =   270
         Left            =   150
         TabIndex        =   106
         Top             =   570
         Width           =   825
      End
      Begin VB.CheckBox chkGst 
         Caption         =   "Gst"
         Height          =   285
         Left            =   150
         TabIndex        =   105
         Top             =   330
         Width           =   1065
      End
      Begin VB.TextBox txtHeadwiseAccounthead 
         Height          =   390
         Left            =   1410
         TabIndex        =   63
         Top             =   750
         Width           =   3555
      End
      Begin VB.CommandButton cmdHeadwiseView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4470
         TabIndex        =   62
         Top             =   1170
         Width           =   795
      End
      Begin VB.CommandButton cmdSearchAccHead 
         Caption         =   ".."
         Height          =   315
         Left            =   4980
         TabIndex        =   61
         Top             =   780
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpHeadwisefrom 
         Height          =   330
         Left            =   1890
         TabIndex        =   64
         Top             =   345
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin MSComCtl2.DTPicker dtpHeadwiseTo 
         Height          =   330
         Left            =   3855
         TabIndex        =   65
         Top             =   345
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head-"
         Height          =   270
         Left            =   120
         TabIndex        =   68
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3525
         TabIndex        =   67
         Top             =   345
         Width           =   300
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1290
         TabIndex        =   66
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame fraDayBook 
      Caption         =   "Day Book"
      Height          =   1335
      Left            =   5640
      TabIndex        =   52
      Top             =   5970
      Width           =   5445
      Begin VB.CommandButton cmdCounterDay 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3690
         TabIndex        =   98
         Top             =   765
         Width           =   345
      End
      Begin VB.CommandButton cmdDayBookview 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         TabIndex        =   54
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox txtDayBookCounter 
         Height          =   390
         Left            =   1920
         TabIndex        =   53
         Top             =   720
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpDayBookDate 
         Height          =   330
         Left            =   1890
         TabIndex        =   55
         Top             =   315
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   43555
      End
      Begin MSComCtl2.DTPicker dtpDayBookToDate 
         Height          =   330
         Left            =   3990
         TabIndex        =   57
         Top             =   330
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   43555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1260
         TabIndex        =   59
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3660
         TabIndex        =   58
         Top             =   330
         Width           =   300
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter -"
         Height          =   270
         Left            =   960
         TabIndex        =   56
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.Frame fraLedger 
      Caption         =   "Transaction Type wise Ledger"
      Height          =   2115
      Left            =   5640
      TabIndex        =   40
      Top             =   3870
      Width           =   5415
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   ".."
         Height          =   315
         Left            =   5010
         TabIndex        =   50
         Top             =   750
         Width           =   315
      End
      Begin VB.ComboBox cmbLedTransactionTypes 
         Height          =   390
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1185
         Width           =   3420
      End
      Begin VB.CommandButton cmdLedgerView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         TabIndex        =   42
         Top             =   1665
         Width           =   795
      End
      Begin VB.TextBox txtAccountHead 
         Height          =   390
         Left            =   1380
         TabIndex        =   41
         Top             =   765
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtpLedFrom 
         Height          =   330
         Left            =   1410
         TabIndex        =   44
         Top             =   345
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin MSComCtl2.DTPicker dtpLedTo 
         Height          =   330
         Left            =   3645
         TabIndex        =   45
         Top             =   345
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   810
         TabIndex        =   49
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3315
         TabIndex        =   48
         Top             =   345
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head-"
         Height          =   270
         Left            =   120
         TabIndex        =   47
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Types-"
         Height          =   270
         Left            =   105
         TabIndex        =   46
         Top             =   1230
         Width           =   1680
      End
   End
   Begin VB.Frame fmeDailyCounterConsolidation 
      Caption         =   "Daily Counter Consolidation"
      Height          =   1230
      Left            =   5610
      TabIndex        =   36
      Top             =   75
      Width           =   5445
      Begin VB.CommandButton cmdViewDailyCounterConsolidation 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         TabIndex        =   37
         Top             =   780
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpdate 
         Height          =   330
         Left            =   2250
         TabIndex        =   38
         Top             =   435
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date-"
         Height          =   270
         Left            =   1650
         TabIndex        =   39
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame fmeTotalHeadWiseCollection 
      Caption         =   "Total Headwise Collection"
      Height          =   1230
      Left            =   5595
      TabIndex        =   30
      Top             =   1320
      Width           =   5445
      Begin VB.CommandButton cmdViewTotalHeadWiseCollection 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         TabIndex        =   31
         Top             =   780
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpTotHeadFrom 
         Height          =   330
         Left            =   1890
         TabIndex        =   32
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin MSComCtl2.DTPicker dtpTotHeadTo 
         Height          =   330
         Left            =   3855
         TabIndex        =   33
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1290
         TabIndex        =   35
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3525
         TabIndex        =   34
         Top             =   315
         Width           =   300
      End
   End
   Begin VB.Frame fmeTotalCounterCollection 
      Caption         =   "Total Counter Collection"
      Height          =   1230
      Left            =   90
      TabIndex        =   24
      Top             =   4110
      Width           =   5445
      Begin VB.CommandButton cmdTotalCollectionView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4530
         TabIndex        =   25
         Top             =   795
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtptotColFrom 
         Height          =   345
         Left            =   1905
         TabIndex        =   26
         Top             =   300
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpTotColTo 
         Height          =   345
         Left            =   3915
         TabIndex        =   27
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1305
         TabIndex        =   29
         Top             =   345
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3540
         TabIndex        =   28
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame fmeCancelled 
      Caption         =   "Cancelled Receipts"
      Height          =   1260
      Left            =   120
      TabIndex        =   17
      Top             =   5340
      Width           =   5445
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   285
         Left            =   90
         TabIndex        =   79
         Top             =   240
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.ComboBox cmbCounter 
         Height          =   390
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   675
         Width           =   2220
      End
      Begin VB.ComboBox cmbCancelledUser 
         Height          =   390
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   675
         Width           =   2130
      End
      Begin VB.CommandButton cmdCancelledView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   23
         Top             =   705
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpCanFrom 
         Height          =   345
         Left            =   2085
         TabIndex        =   20
         Top             =   210
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpCanTo 
         Height          =   345
         Left            =   3915
         TabIndex        =   21
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter"
         Height          =   270
         Left            =   90
         TabIndex        =   77
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3585
         TabIndex        =   19
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1485
         TabIndex        =   18
         Top             =   255
         Width           =   525
      End
   End
   Begin VB.Frame fmeDept 
      Caption         =   "Department wise Report"
      Height          =   1665
      Left            =   135
      TabIndex        =   10
      Top             =   2250
      Width           =   5415
      Begin VB.CommandButton cmdDeptView 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   22
         Top             =   1230
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpDeptTo 
         Height          =   345
         Left            =   3900
         TabIndex        =   16
         Top             =   330
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin MSComCtl2.DTPicker dtpDeptFrom 
         Height          =   345
         Left            =   1890
         TabIndex        =   15
         Top             =   315
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39808
      End
      Begin VB.ComboBox cmbDepartMent 
         Height          =   390
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   765
         Width           =   3420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1290
         TabIndex        =   14
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3525
         TabIndex        =   13
         Top             =   345
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department-"
         Height          =   270
         Left            =   660
         TabIndex        =   12
         Top             =   840
         Width           =   1110
      End
   End
   Begin VB.Frame fmeWardWiseTransactionTypes 
      Caption         =   "Wardwise Transaction Types"
      Height          =   2100
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5415
      Begin VB.CheckBox chkDetailed 
         Caption         =   "Detailed"
         Height          =   315
         Left            =   1860
         TabIndex        =   51
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.TextBox txtWard 
         Height          =   345
         Left            =   1890
         TabIndex        =   9
         Top             =   765
         Width           =   3405
      End
      Begin VB.CommandButton cmdViewWardwise 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4485
         TabIndex        =   8
         Top             =   1665
         Width           =   795
      End
      Begin VB.ComboBox cmbTransactionTypes 
         Height          =   390
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1185
         Width           =   3420
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   330
         Left            =   1890
         TabIndex        =   5
         Top             =   345
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   330
         Left            =   3855
         TabIndex        =   6
         Top             =   345
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   39806
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Types-"
         Height          =   270
         Left            =   105
         TabIndex        =   4
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward-"
         Height          =   270
         Left            =   1215
         TabIndex        =   3
         Top             =   810
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To-"
         Height          =   270
         Left            =   3525
         TabIndex        =   2
         Top             =   345
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From-"
         Height          =   270
         Left            =   1290
         TabIndex        =   1
         Top             =   360
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmRptAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mHeadWiseReport As Integer
    Private Sub setTransactionDate()
        dtpCanFrom.Value = gbTransactionDate
        dtpCanTo.Value = gbTransactionDate
        dtpChittaFromDate.Value = gbTransactionDate
        dtpdate.Value = gbTransactionDate
        dtpDeptFrom.Value = gbTransactionDate
        dtpDeptTo.Value = gbTransactionDate
        dtpFrom.Value = gbTransactionDate
        dtpTo.Value = gbTransactionDate
        dtptotColFrom.Value = gbTransactionDate
        dtpTotColTo.Value = gbTransactionDate
        dtpTotHeadFrom.Value = gbTransactionDate
        dtpTotHeadTo.Value = gbTransactionDate
        dtpCountFrom.Value = gbTransactionDate
        dtpCountTo.Value = gbTransactionDate
    End Sub

   


Private Sub chkGst_Click()
        If chkRent.Value = 1 Then chkGst.Value = 0
    End Sub

    Private Sub chkRent_Click()
        If chkGst.Value = 1 Then chkRent.Value = 0
    End Sub

    Private Sub cmbChittaCounter_Click()
        FillUsers
    End Sub

    Private Sub cmbCounter_Click()
        Call FillCancelUser
    End Sub




    Private Sub cmbReceiptCounter_Click()
        Dim mSql As String
        Dim mCounter As String
        If cmbReceiptCounter.ListIndex < 1 Then
            mCounter = "%"
        Else
            mCounter = cmbReceiptCounter.ItemData(cmbReceiptCounter.ListIndex)
        End If
        mSql = "Select Distinct FaUser.vchUserName,FaUser.numUserID From FaUser "
        mSql = mSql + " Inner Join faVouchers On faVouchers.intUserID = faUser.numUserID"
        mSql = mSql + " Where dtDate BetWeen '" & CheckDateInMMM(dtpCountFrom.Value) & "' And '" & CheckDateInMMM(dtpCountTo.Value) & "'"
        mSql = mSql + " And Convert(varchar(3),intCounterID) Like '" & CStr(mCounter) & "'"
        PopulateList cmbRecceiptCountUser, mSql, , True, , True, enuSourceString.Saankhya
    End Sub

    Private Sub cmdCancelledView_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCounterID As String
        Dim mUserID As String
        ''frmMenu.Transactions.Enabled = False
        If cmbCounter.ListIndex < 1 Then
            mCounterID = "%"
        Else
            mCounterID = CStr(cmbCounter.ItemData(cmbCounter.ListIndex))
        End If
        If cmbCancelledUser.ListIndex < 1 Then
            mUserID = "%"
        Else
            mUserID = cmbCancelledUser.ItemData(cmbCancelledUser.ListIndex)
        End If
        If chkAll.Value = 1 Then
            arInput = Array(mCounterID, dtpCanFrom.Value, dtpCanTo.Value, mUserID)
            frmNewViewer.rptFileName = App.Path & "\Reports\rptCancelledReceipts.rpt"
        Else
            arInput = Array(dtpCanFrom.Value, dtpCanTo.Value)
            frmNewViewer.rptFileName = App.Path & "\Reports\rptCancelledCounterReceipts.rpt"
        End If
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdCard_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCounterID As String
        Dim mUserID As String
        Dim mDtFrom As String
        Dim mDtTo As String
        ''frmMenu.Transactions.Enabled = False
        
        mDtFrom = dtpCardFrom.Value & " " & TimeValue(dtpTimeFrom)
        mDtTo = dtpCardTo.Value & " " & TimeValue(dtpTimeTo)

        arInput = Array(CDate(mDtFrom), CDate(mDtTo))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCardPayment.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show

    End Sub

    Private Sub cmdChitta_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mUser As String
        Dim mCounter   As String
        ''frmMenu.Transactions.Enabled = False
        'arInput = Array(gbTransactionDate, gbTransactionDate, CStr(gbCounterID), "%", "%", "%")
        If cmbChittaCounter.ListIndex < 1 Then
            mCounter = "%"
        Else
            mCounter = cmbChittaCounter.ItemData(cmbChittaCounter.ListIndex)
        End If
        
        If cmbUser.ListIndex < 1 Then
            mUser = "%"
        Else
            mUser = cmbUser.ItemData(cmbUser.ListIndex)
        End If
        
        arInput = Array(dtpChittaFromDate.Value, dtpChittaToDate.Value, CStr(mCounter), "%", "%", "%", CStr(mUser))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptChitta.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub
    Private Sub cmdCounterDay_Click()
        Dim mSql As String
        mSql = "SELECT     intCounterID,vchDescription + '(' +CAST(intCounterNo AS varchar(10)) + ')' AS counter FROM   faCounters "
'        frmMaster.Show
        frmSearchMasters.SQLQry = mSql
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = 1
        frmSearchMasters.Show vbModal
        txtDayBookCounter.Text = gbSearchStr
        txtDayBookCounter.Tag = gbSearchID
    End Sub

    Private Sub cmdDayBookview_Click()
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
       ' If txtDayBookCounter.Text = "" Then txtDayBookCounter.Text = "%"
        arInput = Array(dtpDayBookDate.Value, dtpDayBookToDate.Value, IIf(CStr(Trim(txtDayBookCounter.Tag)) = "", "%", CStr((txtDayBookCounter.Tag))), "%", "%", "%", "%")
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCounterDayBook.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdDeptView_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        ''frmMenu.Transactions.Enabled = False
        If cmbDepartMent.ListIndex < 1 Then
            mDepartment = "%"
        Else
            mDepartment = CStr(cmbDepartMent.ItemData(cmbDepartMent.ListIndex))
        End If
        arInput = Array(dtpDeptFrom.Value, dtpDeptTo.Value, "%", "%", "%", mDepartment, "%")
        frmNewViewer.rptFileName = App.Path & "\Reports\rptDepartmentWise.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub


    Private Sub cmdfmePreviousDatesReceiptCancellation_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        arInput = Array(dtpPredateFrom.Value, dtpPredateTo.Value)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCancelledPreDateReceipts.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdHeadwiseView_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        arInput = Array(val(txtHeadwiseAccounthead.Tag), dtpHeadwisefrom.Value, dtpHeadwiseTo.Value)
        If chkGst.Value = False And chkRent.Value = False Then
            If mHeadWiseReport = 1 Then
                frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadWiseReport.rpt"
            Else
                frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadWiseReportforVAT.rpt"
            End If
        
        ElseIf chkRent.Value = False Then
            frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadWiseReportforGST.rpt"
        ElseIf chkGst.Value = False Then
            arInput = Array(dtpHeadwisefrom.Value, dtpHeadwiseTo.Value)
            frmNewViewer.rptFileName = App.Path & "\Reports\rptRentCollectedDetails.rpt"
        End If
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdLedgerView_Click()
        Dim mLedTransactionTypes As Variant
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        If Trim(txtAccountHead) = "" Then
            MsgBox "Select Account Head", vbInformation
            Exit Sub
        End If
        If cmbLedTransactionTypes.ListIndex < 1 Then
            mLedTransactionTypes = "0"
        Else
            mLedTransactionTypes = cmbLedTransactionTypes.ItemData(cmbLedTransactionTypes.ListIndex)
        End If
        arInput = Array(val(txtAccountHead.Tag), dtpLedFrom.Value, dtpLedTo.Value, val(mLedTransactionTypes))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadWiseReport.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdReceiptCount_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCounterID As String
        Dim mUserID As String
        Dim mDtFrom As String
        Dim mDtTo As String
        ''frmMenu.Transactions.Enabled = False
        If cmbReceiptCounter.ListIndex < 1 Then
            mCounterID = "%"
        Else
            mCounterID = CStr(cmbReceiptCounter.ItemData(cmbReceiptCounter.ListIndex))
        End If
        If cmbRecceiptCountUser.ListIndex < 1 Then
            mUserID = "%"
        Else
            mUserID = cmbRecceiptCountUser.ItemData(cmbRecceiptCountUser.ListIndex)
        End If
        mDtFrom = dtpCountFrom.Value & " " & TimeValue(dtpTimeFrom)
        mDtTo = dtpCountTo.Value & " " & TimeValue(dtpTimeTo)
'        arInput = Array(mCounterID, mUserID, dtpCountFrom.Value, dtpCountTo.Value)
        arInput = Array(mCounterID, mUserID, CDate(mDtFrom), CDate(mDtTo))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptReceiptCount.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdSearchAccHead_Click()
        If mHeadWiseReport <> 1 Then
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where vchAccountHeadCode in ('350300400','350300700','350300800')"
        End If
        frmSearchAccountHeads.Show 1
        Dim objAccHead As New clsAccounts
        objAccHead.SetAccountCode (Token(gbSearchStr, " "))
        If objAccHead.AccountHeadID > 0 Then
            txtHeadwiseAccounthead.Text = objAccHead.AccountHead
            txtHeadwiseAccounthead.Tag = objAccHead.AccountHeadID
        End If
        gbSearchStr = ""
        gbSearchID = -1
    End Sub

Private Sub cmdSearchAccountHead_Click()
    frmSearchAccountHeads.Show 1
        
        Dim objAccHead As New clsAccounts
        objAccHead.SetAccountCode (Token(gbSearchStr, " "))
        If objAccHead.AccountHeadID > 0 Then
            txtAccountHead.Text = objAccHead.AccountHead
            txtAccountHead.Tag = objAccHead.AccountHeadID
        End If
        gbSearchStr = ""
        gbSearchID = -1
End Sub

    Private Sub cmdTotalCollectionView_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        ''frmMenu.Transactions.Enabled = False
        If cmbDepartMent.ListIndex < 1 Then
            mDepartment = "%"
        Else
            mDepartment = CStr(cmbDepartMent.ItemData(cmbDepartMent.ListIndex))
        End If
        arInput = Array(dtptotColFrom.Value, dtpTotColTo.Value)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptTotalCounterCollection.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdViewDailyCounterConsolidation_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        ''frmMenu.Transactions.Enabled = False
'        If cmbDepartMent.ListIndex < 1 Then
'            mDepartment = "%"
'        Else
'            mDepartment = CStr(cmbDepartMent.ItemData(cmbDepartMent.ListIndex))
'        End If
        arInput = Array(dtpdate.Value, "%", "%", "%", "%")
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCounterReports.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdViewTotalHeadWiseCollection_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        ''frmMenu.Transactions.Enabled = False
'        If cmbDepartMent.ListIndex < 1 Then
'            mDepartment = "%"
'        Else
'            mDepartment = CStr(cmbDepartMent.ItemData(cmbDepartMent.ListIndex))
'        End If
        arInput = Array(dtpTotHeadFrom.Value, dtpTotHeadTo.Value)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadwiseCollection.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdViewWardwise_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mTransactionType As Variant
        ''frmMenu.Transactions.Enabled = False
        If cmbTransactionTypes.ListIndex < 1 Then
            mTransactionType = "%"
        Else
            mTransactionType = CStr(cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex))
        End If
        If chkDetailed.Value = 0 Then
            arInput = Array(dtpFrom.Value, dtpTo.Value, "%", IIf(Trim(txtWard) = "", "%", CStr(txtWard)), mTransactionType, "%", "%")
            frmNewViewer.rptFileName = App.Path & "\Reports\rptTransactionTypeWise.rpt"
        Else
            
            arInput = Array(dtpFrom.Value, dtpTo.Value, "%", IIf(Trim(txtWard) = "", "%", CStr(txtWard)), mTransactionType, "%", "%")
            frmNewViewer.rptFileName = App.Path & "\Reports\rptDepartmentwiseWardWise.rpt"
        End If
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub cmdZonalCollection_Click()

        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mCounterID As String
        Dim mUserID As String
        Dim mDtFrom As String
        Dim mDtTo As String
        ''frmMenu.Transactions.Enabled = False
        Dim mZone As Integer
        
        mDtFrom = dtpZonedateFrom.Value
        mDtTo = dtpZonedateTo.Value
        If cmbzone.ListIndex < 1 Then
            mZone = 0
            MsgBox "Please Select Zonal office", vbApplicationModal
            Exit Sub
        Else
            mZone = cmbzone.ItemData(cmbzone.ListIndex)
        End If
        arInput = Array(CDate(mDtFrom), CDate(mDtTo), mZone)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptZonalReceipts.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub dtpCanFrom_Click()
        Call FillCancelUser
    End Sub


    Private Sub dtpCanTo_Click()
        Call FillCancelUser
    End Sub

    Private Sub dtpChittaFromDate_Click()
        Call FillUsers
    End Sub

    Private Sub dtpChittaFromDate_LostFocus()
         dtpChittaFromDate.Value = CheckDateInMMM(dtpChittaFromDate.Value)
        If CDate(dtpChittaFromDate.Value) Then
            If CDate(dtpChittaToDate.Value) Then
                If CDate(dtpChittaFromDate.Value) > CDate(dtpChittaToDate.Value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpChittaFromDate.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            dtpChittaFromDate.Value = CheckDateInMMM(dtpChittaFromDate.Value)
        End If
    End Sub

    Private Sub dtpChittaToDate_Click()
        FillUsers
    End Sub

    Private Sub dtpChittaToDate_LostFocus()
        If CDate(dtpChittaFromDate.Value) > CDate(dtpChittaToDate.Value) Then
            MsgBox "Please Enter a valid Date", vbInformation
            dtpChittaFromDate.Value = gbTransactionDate
            dtpChittaFromDate.SetFocus
            Exit Sub
        End If
    End Sub

    Private Sub dtpDayBookDate_LostFocus()
        dtpDayBookDate.Value = CheckDateInMMM(dtpDayBookDate.Value)
        If CDate(dtpDayBookDate.Value) Then
            If CDate(dtpDayBookToDate.Value) Then
                If CDate(dtpDayBookDate.Value) > CDate(dtpDayBookToDate.Value) Then
                    MsgBox "Please Enter a value Less than or equal to To Date", vbInformation
                    dtpDayBookDate.Value = Format(gbTransactionDate, "dd-mmm-yyyy")
                    Exit Sub
                End If
            Else
                MsgBox "Please enter To date", vbInformation
                Exit Sub
            End If
        Else
            dtpDayBookDate.Value = CheckDateInMMM(dtpDayBookDate.Value)
        End If
    End Sub

    Private Sub dtpDayBookToDate_LostFocus()
        If CDate(dtpDayBookDate.Value) > CDate(dtpDayBookToDate.Value) Then
            MsgBox "Please Enter a valid Date", vbInformation
            dtpDayBookDate.Value = gbTransactionDate
            dtpDayBookDate.SetFocus
            Exit Sub
        End If
    End Sub


    Private Sub dtpDeptTo_LostFocus()
        If CDate(dtpDeptFrom.Value) > CDate(dtpDeptTo.Value) Then
            MsgBox "Please Enter a valid Date", vbInformation
            dtpDeptFrom.Value = gbTransactionDate
            dtpDeptFrom.SetFocus
            Exit Sub
        End If
    End Sub
    
    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
        
        fmeCancelled.Left = 0
        fmeCancelled.Top = 0
        
        fmeDailyCounterConsolidation.Left = 0
        fmeDailyCounterConsolidation.Top = 0
        
        fmeDept.Left = 0
        fmeDept.Top = 0
        
        fmeTotalCounterCollection.Left = 0
        fmeTotalCounterCollection.Top = 0
        
        fmeTotalHeadWiseCollection.Left = 0
        fmeTotalHeadWiseCollection.Top = 0
        
        fmeWardWiseTransactionTypes.Top = 0
        fmeWardWiseTransactionTypes.Left = 0
        
        fmePreviousDatesReceiptCancellation.Top = 0
        fmePreviousDatesReceiptCancellation.Left = 0
        
        fmeDailyCounterConsolidation.Top = 0
        fmeDailyCounterConsolidation.Left = 0
        
        fraChitta.Top = 0
        fraChitta.Left = 0
        
        fraLedger.Top = 0
        fraLedger.Left = 0
        
        fraDayBook.Top = 0
        fraDayBook.Left = 0
        
        fmeZonalCollection.Top = 0
        fmeZonalCollection.Left = 0
    End Sub
    
    Private Sub Form_Load()
        LoadWardwise
        LoadDepartMentwise
        Call setTransactionDate
        LoadZonalOffice
        '------------------------------------------------------'
        '           Setting Current Date                       '
        '------------------------------------------------------'
        dtpDayBookDate = Date
        dtpLedFrom = Date
        dtpLedTo = Date
        dtpChittaFromDate = Date
        dtpChittaToDate = Date
        dtpTotHeadFrom = Date
        dtpTotHeadTo = Date
        dtpdate = Date
        dtpCanFrom = Date
        dtpCanTo = Date
        dtptotColFrom = Date
        dtpTotColTo = Date
        dtpDeptFrom = Date
        dtpDeptTo = Date
        dtpFrom = Date
        dtpTo = Date
        dtpDayBookToDate = Date
        dtpHeadwisefrom = Date
        dtpHeadwiseTo = Date
        
        Call FillCounters
        '------------------------------------------------------'
        '------------------------------------------------------'
        
    End Sub
    Public Sub LoadWardwise()
        ''''''' For Ward wise''''''''
        dtpFrom.Value = gbTransactionDate
        dtpTo.Value = gbTransactionDate
        PopulateList cmbTransactionTypes, "Select vchtransactionType,intTransactionTypeID From faTransactionType Order By vchtransactionType", "Property Tax", True, , True
        PopulateList cmbLedTransactionTypes, "Select vchtransactionType,intTransactionTypeID From faTransactionType Order By vchtransactionType", , True, , True
        ''''''' For Ward wise''''''''
    End Sub
    
    Public Sub LoadDepartMentwise()
        ''''''''Load Department wise'''''''
        PopulateList cmbDepartMent, "Select vchSectionName,intSectionID From faSection Order By vchSectionName", "Revenue Department (Property Tax & Profession Tax)", True, , True
        ''''''''Load Department wise'''''''
    End Sub
    Public Sub LoadZonalOffice()
        ''''''''Load Department wise'''''''
        PopulateList cmbzone, "Select chvZoneNameEnglish, numZoneID From GM_Zone WHERE Right(numZoneID,2)<>1 AND intLBID =" & gbLocalBodyID & " Order By chvZoneNameEnglish", gbLocation, True, True, True, DBMaster
        ''''''''Load Department wise'''''''
    End Sub
    Public Sub frameVisible()
        fmeCancelled.Visible = False
        fmeDailyCounterConsolidation.Visible = False
        fmeDept.Visible = False
        fmeTotalCounterCollection.Visible = False
        fmeTotalHeadWiseCollection.Visible = False
        fmeWardWiseTransactionTypes.Visible = False
        fmeheadwiseReport.Visible = False
        fraChitta.Visible = False
        fraLedger.Visible = False
        fraDayBook.Visible = False
        fmeReceiptCount.Visible = False
        fmePreviousDatesReceiptCancellation.Visible = False
        fmeZonalCollection.Visible = False
    End Sub
    Public Property Let HeadWiseReport(mData As Integer)
        mHeadWiseReport = mData
    End Property
    Private Sub FillUsers()
        Dim mSql As String
        Dim mCounter As String
        If Trim(cmbChittaCounter.ListIndex) < 1 Then
            mCounter = "%"
        Else
            mCounter = cmbChittaCounter.ItemData(cmbChittaCounter.ListIndex)
        End If
        mSql = "Select Distinct FaUser.vchUserName,FaUser.numUserID From FaUser "
        mSql = mSql + " Inner Join faVouchers On faVouchers.intUserID = faUser.numUserID"
        mSql = mSql + " Where dtDate BetWeen '" & CheckDateInMMM(dtpChittaFromDate.Value) & "' And '" & CheckDateInMMM(dtpChittaToDate.Value) & "'"
        mSql = mSql + " And Convert(varchar(3),intCounterID) Like '" & CStr(mCounter) & "'"
        PopulateList cmbUser, mSql, , True, , True, enuSourceString.Saankhya
        
    End Sub
    Private Sub FillCounters()
        Dim mSql As String
        mSql = "SELECT     vchDescription + '(' +CAST(intCounterNo AS varchar(10)) + ')' AS counter, intCounterID FROM   faCounters"
        PopulateList cmbCounter, mSql, , True, , True, enuSourceString.Saankhya
        PopulateList cmbReceiptCounter, mSql, , True, , True, enuSourceString.Saankhya
        PopulateList cmbChittaCounter, mSql, , True, , True, enuSourceString.Saankhya
    End Sub
    Private Sub FillCancelUser()
        Dim mSql As String
        Dim mCounter As String
        If cmbCounter.ListIndex < 1 Then
            mCounter = "%"
        Else
            mCounter = cmbCounter.ItemData(cmbCounter.ListIndex)
        End If
        mSql = "Select Distinct FaUser.vchUserName,FaUser.numUserID From FaUser "
        mSql = mSql + " Inner Join faVouchers On faVouchers.intUserID = faUser.numUserID"
        mSql = mSql + " Where dtDate BetWeen '" & CheckDateInMMM(dtpCanFrom.Value) & "' And '" & CheckDateInMMM(dtpCanTo.Value) & "'"
        mSql = mSql + " And Convert(varchar(3),intCounterID) Like '" & CStr(mCounter) & "'"
        PopulateList cmbCancelledUser, mSql, , True, , True, enuSourceString.Saankhya
    End Sub
   

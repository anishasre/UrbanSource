VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProjectRegisterDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Register Details"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOldVoucher 
      Caption         =   "Payment Done in Saankhya"
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
      Height          =   195
      Left            =   11700
      TabIndex        =   94
      Top             =   4905
      Width           =   2355
   End
   Begin VB.Frame fraReqOnlineDate 
      Height          =   2040
      Left            =   2115
      TabIndex        =   79
      Top             =   5040
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton cmdPreviousDateVerify 
         Caption         =   "VERIFY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7740
         TabIndex        =   91
         Top             =   810
         Width           =   1680
      End
      Begin VB.CommandButton cmdPreviousBank 
         Caption         =   "..."
         Height          =   330
         Left            =   7155
         TabIndex        =   90
         Top             =   1440
         Width           =   285
      End
      Begin VB.TextBox txtPreviousInstDate 
         Height          =   330
         Left            =   5085
         TabIndex        =   89
         Top             =   900
         Width           =   2040
      End
      Begin VB.TextBox txtPreviousInstNo 
         Height          =   330
         Left            =   5085
         TabIndex        =   87
         Top             =   360
         Width           =   2040
      End
      Begin VB.TextBox txtPreviousBank 
         Height          =   330
         Left            =   1395
         TabIndex        =   85
         Top             =   1440
         Width           =   5730
      End
      Begin VB.TextBox txtPreviousVoucherDate 
         Height          =   330
         Left            =   1395
         TabIndex        =   83
         Top             =   900
         Width           =   2040
      End
      Begin VB.TextBox txtPreviousVoucher 
         Height          =   330
         Left            =   1395
         TabIndex        =   81
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label lblClose 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   240
         Left            =   9315
         TabIndex        =   93
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Instrument Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   88
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Instrument No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3915
         TabIndex        =   86
         Top             =   405
         Width           =   1140
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank"
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
         Left            =   855
         TabIndex        =   84
         Top             =   1530
         Width           =   510
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   82
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   80
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame frmPayVDetails 
      Caption         =   "Payment Voucher"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7200
      TabIndex        =   56
      Top             =   4995
      Width           =   6990
      Begin VB.Frame frmPayVoucherDetails 
         Enabled         =   0   'False
         Height          =   1140
         Left            =   45
         TabIndex        =   70
         Top             =   900
         Width           =   6900
         Begin VB.TextBox txtPVDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   78
            Top             =   315
            Width           =   1590
         End
         Begin VB.TextBox txtPVTrType 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1620
            TabIndex        =   73
            Top             =   315
            Width           =   2850
         End
         Begin VB.TextBox txtPVBankHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1620
            TabIndex        =   72
            Top             =   675
            Width           =   2850
         End
         Begin VB.TextBox txtPVAmt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   71
            Top             =   675
            Width           =   1590
         End
         Begin VB.Label Label22 
            Caption         =   " Date"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4590
            TabIndex        =   77
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Transaction Type"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   76
            Top             =   270
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank/ Treasury"
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
            Left            =   225
            TabIndex        =   75
            Top             =   675
            Width           =   1230
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
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
            Left            =   4410
            TabIndex        =   74
            Top             =   720
            Width           =   690
         End
      End
      Begin VB.TextBox txtPaymentVoucher 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1665
         TabIndex        =   58
         Top             =   360
         Width           =   2850
      End
      Begin VB.CommandButton cmdSearchPV 
         Caption         =   "..."
         Height          =   330
         Left            =   4545
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label14 
         Caption         =   "Payment Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   60
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblmsgPV 
         Caption         =   "*Wrong Payment Voucher"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4905
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdVerifyPaymentVoucher 
      Caption         =   "VERIFY PAYMENT VOUCHER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4950
      TabIndex        =   55
      Top             =   7245
      Width           =   2400
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   -1080
      Top             =   7695
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fraPVDetails 
      Caption         =   "Payment Order"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   45
      TabIndex        =   34
      Top             =   4995
      Width           =   7125
      Begin VB.Frame frmPODetails 
         Enabled         =   0   'False
         Height          =   1140
         Left            =   90
         TabIndex        =   61
         Top             =   900
         Width           =   6990
         Begin VB.TextBox txtPSource 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   69
            Top             =   270
            Width           =   1725
         End
         Begin VB.TextBox txtPAmt 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   68
            Top             =   675
            Width           =   1725
         End
         Begin VB.TextBox txtPExpHead 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            TabIndex        =   67
            Top             =   675
            Width           =   2985
         End
         Begin VB.TextBox txtPTrType 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            TabIndex        =   66
            Top             =   270
            Width           =   2985
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Fund"
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
            Left            =   4680
            TabIndex        =   65
            Top             =   270
            Width           =   420
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
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
            Left            =   4455
            TabIndex        =   64
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Expenditure Head"
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
            Left            =   45
            TabIndex        =   63
            Top             =   720
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Transaction Type"
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
            Left            =   90
            TabIndex        =   62
            Top             =   270
            Width           =   1365
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdSearchPO 
         Caption         =   "..."
         Height          =   330
         Left            =   4680
         TabIndex        =   37
         Top             =   405
         Width           =   285
      End
      Begin VB.TextBox txtPaymentOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   35
         Top             =   405
         Width           =   2985
      End
      Begin VB.Label lblmsgPO 
         Caption         =   "*Wrong Payment Order"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   5040
         TabIndex        =   42
         Top             =   405
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label Label13 
         Caption         =   "Payment Order"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   36
         Top             =   405
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14235
      TabIndex        =   33
      Top             =   0
      Width           =   14235
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   870
      Left            =   0
      TabIndex        =   24
      Top             =   495
      Width           =   14235
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11835
         TabIndex        =   53
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtRequisitionNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox txtAllotmentNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8100
         TabIndex        =   26
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox txtAllotmentDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5040
         TabIndex        =   25
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label Label16 
         Caption         =   "Requisition Amount "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         TabIndex        =   54
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Allotment Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         TabIndex        =   32
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Requisition No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   450
         TabIndex        =   29
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Allotment No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6885
         TabIndex        =   28
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "VERIFY REQUISITION"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12285
      TabIndex        =   31
      Top             =   4410
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7470
      TabIndex        =   30
      Top             =   7245
      Width           =   915
   End
   Begin VB.Frame fraReqDetails 
      Caption         =   "Requisition Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   45
      TabIndex        =   0
      Top             =   1395
      Width           =   14145
      Begin VB.CheckBox chkRevisedPrjUpdate 
         Caption         =   "Revised Project"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12195
         TabIndex        =   95
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton cmdUndoRequisition 
         Caption         =   "UNDO VERIFICATION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12105
         TabIndex        =   92
         Top             =   495
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtProjectNameEng 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2745
         TabIndex        =   52
         Top             =   720
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.TextBox txtProjCost 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9630
         TabIndex        =   50
         Top             =   540
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdFunctionary 
         Caption         =   "..."
         Height          =   330
         Left            =   13725
         TabIndex        =   46
         Top             =   1620
         Width           =   285
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9135
         TabIndex        =   45
         Top             =   1620
         Width           =   4560
      End
      Begin VB.CommandButton cmdCategory 
         Caption         =   "..."
         Height          =   330
         Left            =   6030
         TabIndex        =   44
         Top             =   1620
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         TabIndex        =   43
         Top             =   1605
         Width           =   4380
      End
      Begin VB.CommandButton cmdIMPO 
         Caption         =   "..."
         Height          =   330
         Left            =   6030
         TabIndex        =   40
         Top             =   2970
         Width           =   285
      End
      Begin VB.TextBox txtIMPO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   39
         Top             =   2955
         Width           =   4380
      End
      Begin VB.TextBox txtMicroHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtSourceOfFund 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   16
         Top             =   1155
         Width           =   4380
      End
      Begin VB.TextBox txtSubSector 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         TabIndex        =   15
         Top             =   2055
         Width           =   4380
      End
      Begin VB.TextBox txtMicroSector 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         TabIndex        =   14
         Top             =   2505
         Width           =   4380
      End
      Begin VB.CommandButton cmdSourceOfFund 
         Caption         =   "..."
         Height          =   330
         Left            =   6030
         TabIndex        =   13
         Top             =   1170
         Width           =   285
      End
      Begin VB.CommandButton cmdSubSector 
         Caption         =   "..."
         Height          =   330
         Left            =   6030
         TabIndex        =   12
         Top             =   2070
         Width           =   285
      End
      Begin VB.CommandButton cmdMicroSector 
         Caption         =   "..."
         Height          =   330
         Left            =   6030
         TabIndex        =   11
         Top             =   2520
         Width           =   285
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9135
         TabIndex        =   10
         Top             =   1170
         Width           =   4560
      End
      Begin VB.TextBox txtGrossExpenditureHead 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10710
         TabIndex        =   9
         Top             =   2070
         Width           =   3030
      End
      Begin VB.TextBox txtTreasury 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9135
         TabIndex        =   8
         Top             =   2550
         Width           =   4605
      End
      Begin VB.CommandButton cmdTreasury 
         Caption         =   "..."
         Height          =   330
         Left            =   13725
         TabIndex        =   7
         Top             =   2565
         Width           =   285
      End
      Begin VB.CommandButton cmdExpHead 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   330
         Left            =   13725
         TabIndex        =   6
         Top             =   2070
         Width           =   285
      End
      Begin VB.CommandButton cmdFunction 
         Caption         =   "..."
         Height          =   330
         Left            =   13725
         TabIndex        =   5
         Top             =   1170
         Width           =   285
      End
      Begin VB.TextBox txtProjectNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearchProject 
         Caption         =   "..."
         Height          =   330
         Left            =   7335
         TabIndex        =   3
         Top             =   720
         Width           =   285
      End
      Begin VB.TextBox txtProjectName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "ML-TTRevathi"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2745
         TabIndex        =   2
         Top             =   720
         Width           =   4560
      End
      Begin VB.TextBox txtExpAccHeadCode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9135
         TabIndex        =   1
         Top             =   2070
         Width           =   1590
      End
      Begin VB.Label Label17 
         Caption         =   "New Proj.Cost"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8370
         TabIndex        =   51
         Top             =   540
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblmsg 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   330
         Left            =   45
         TabIndex        =   49
         Top             =   315
         Width           =   8070
      End
      Begin VB.Label Label8 
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8055
         TabIndex        =   48
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   765
         TabIndex        =   47
         Top             =   1620
         Width           =   780
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "IMPO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   450
         TabIndex        =   41
         Top             =   2985
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "Source Of Fund"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   23
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8280
         TabIndex        =   22
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label Label9 
         Caption         =   "SubSector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   675
         TabIndex        =   21
         Top             =   2070
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "MicroSector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   495
         TabIndex        =   20
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label11 
         Caption         =   "Gross Expenditure Head"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7065
         TabIndex        =   19
         Top             =   2115
         Width           =   2040
      End
      Begin VB.Label Label12 
         Caption         =   "Treasury/Bank"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7830
         TabIndex        =   18
         Top             =   2610
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Project No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   540
         TabIndex        =   17
         Top             =   765
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmProjectRegisterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mWHERE  As String
Dim mProjectID As Variant
Dim mSourceOfFundID As Variant
Dim mReqID As Integer
Public dtOnlinedate As String

Private Sub chkOldVoucher_Click()
    If chkOldVoucher.value = 1 Then
        fraReqOnlineDate.Visible = False
        fraPVDetails.Visible = True
        frmPayVDetails.Visible = True
        fraPVDetails.Enabled = True
        frmPayVDetails.Enabled = True
        cmdVerifyPaymentVoucher.Visible = True
    Else
        fraReqOnlineDate.Visible = True
        fraPVDetails.Visible = False
        frmPayVDetails.Visible = False
        cmdVerifyPaymentVoucher.Visible = False
    End If
End Sub

Private Sub chkRevisedPrjUpdate_Click()
     Dim mCnn    As New ADODB.Connection
     Dim mCnnSulekha As New ADODB.Connection
     Dim objdb   As New clsDB
     Dim mSql As String
     Dim msqlSulekha As String
     Dim mCount As Integer
     Dim RecSulekha As New ADODB.Recordset
     Dim Rec As New ADODB.Recordset
     Dim mArrIn As Variant
     
     If objdb.SetConnection(mCnn) Then
         mSql = "Select * from suProjectDEtails where decProjectID = " & val(txtProjectNo.Tag)
         Rec.Open mSql, mCnn
         If Not (Rec.EOF And Rec.BOF) Then
             mCount = 1
         Else
             mCount = 0
         End If
         Rec.Close
         If mCount = 1 Then
            mSql = "Update suProjectDEtails set intPlanID=1  From suProjectDEtails"
            mSql = mSql + " Inner Join faAllotments On suProjectDEtails.decProjectID=faAllotments.numProjectID"
            mSql = mSql + " Where faAllotments.numProjectID=" & val(txtProjectNo.Tag) & " And  intID= " & txtRequisitionNo.Tag
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
         Else
                 If (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                    msqlSulekha = "Select * from ProjectDetails "
                    msqlSulekha = msqlSulekha + " left join SubjectCheckList On SubjectCheckList.decProjectID=ProjectDetails.decProjectID Where ProjectDetails.decProjectID= " & val(txtProjectNo.Tag)
                    RecSulekha.Open msqlSulekha, mCnnSulekha
                    If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                        mArrIn = Array(Trim(val(txtProjectNo.Tag)), _
                                              gbLBID, _
                                              gbFinancialYearID - 1, _
                                              IIf(IsNull(RecSulekha!intProjectSlNo), "", RecSulekha!intProjectSlNo), _
                                              IIf(IsNull(RecSulekha!chvProjectSlNo), "", RecSulekha!chvProjectSlNo), _
                                              IIf(IsNull(RecSulekha!chvProjectName), "", RecSulekha!chvProjectName), _
                                              IIf(IsNull(RecSulekha!chvProjectNameEng), "", RecSulekha!chvProjectNameEng), _
                                              IIf(IsNull(RecSulekha!intProjCatID), "", RecSulekha!intProjCatID), _
                                              IIf(IsNull(RecSulekha!nchApprovalNo), "", RecSulekha!nchApprovalNo), _
                                              IIf(IsNull(RecSulekha!dtApprovalDate), "", RecSulekha!dtApprovalDate), _
                                              IIf(IsNull(RecSulekha!intSecID), "", RecSulekha!intSecID), _
                                              IIf(IsNull(RecSulekha!intImplOfficerID), "", RecSulekha!intImplOfficerID), _
                                              IIf(IsNull(RecSulekha!intSubSecID), "", RecSulekha!intSubSecID), _
                                              9, _
                                              IIf(IsNull(RecSulekha!chvFullName), "", RecSulekha!chvFullName), _
                                              IIf(IsNull(RecSulekha!chvDesignation), "", RecSulekha!chvDesignation), _
                                              1 _
                                            )
                        objdb.ExecuteSP "spUpdateProjectDetails", mArrIn, , , mCnn, adCmdStoredProc
                    End If
                End If
                mCnnSulekha.Close
            End If
                
    
    '''     mSQL = "Update faAllotments set vchBillNo=1  "
    '''     mSQL = mSQL + " Where faAllotments.numProjectID=" & txtProjectNo.Tag & " And intID= " & txtRequisitionNo.Tag
    '''     objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
         
         cmdVerify.Enabled = False
         fraPVDetails.Enabled = False
         frmPayVDetails.Enabled = False
         cmdVerifyPaymentVoucher.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    frmProjectDetails.FillGrid
    Unload Me
End Sub
Private Sub cmdExpHead_Click()
    Dim mToken   As String
    
    If Len(cmdExpHead.Tag) > 0 Then
        frmSearchAccountHeads.SQLString = cmdExpHead.Tag
    Else
        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where  tinHiddenFlag = 0 and vchAccountHeadCode Not Like '1%' Order By faAccountHeads.vchAccountHeadCode"
    End If
    frmSearchAccountHeads.Show vbModal
    mToken = Token(gbSearchStr, " ")
      If gbSearchID <> -1 Then
          txtExpAccHeadCode.Text = mToken
          txtGrossExpenditureHead.Text = Trim(gbSearchStr)
          txtGrossExpenditureHead.Tag = gbSearchID
          gbSearchID = -1
          gbSearchStr = ""
      End If
End Sub
Private Function ProjectExpHeadValidation() As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mAccHeadId As Integer
       
    If objdb.SetConnection(mCnn) Then
'        mSQL = " Select * from faAllotments "
'        mSQL = mSQL + " Where intID= " & txtRequisitionNo.Tag & " "
        mSql = " Select * from faPayOrder "
        mSql = mSql + " Where intAllotmentID= " & txtRequisitionNo.Tag & " "
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
             mAccHeadId = IIf(IsNull(Rec!intCashOrBankHeadID), 0, Rec!intCashOrBankHeadID)
'            mAccHeadId = Rec!intAccountHeadID
'            txtPaymentOrder.Text = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
'            txtPaymentOrder.Tag = IIf(IsNull(Rec!intPayOrderID), 0, Rec!intPayOrderID)
'            txtPaymentVoucher.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
'            txtPaymentVoucher.Tag = IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)
        End If
        If mAccHeadId <> 0 Then
            If mAccHeadId = txtGrossExpenditureHead.Tag Then
                ProjectExpHeadValidation = True
            Else
                ProjectExpHeadValidation = False
            End If
        Else
            ProjectExpHeadValidation = True
        End If
        Rec.Close
    End If
    'mCnn.Close
End Function
Private Sub cmdFunction_Click()
    Dim mToken   As String
    frmSearchFunction.Show vbModal
    mToken = Token(gbSearchStr, " ")
    txtFunction.Text = Trim(gbSearchStr)
    txtFunction.Tag = gbSearchID
    gbSearchStr = ""
    gbSearchID = -1
End Sub
Private Sub cmdFunctionary_Click()
    Dim mToken As String
    frmSearchFunctionary.Show vbModal
    mToken = Token(gbSearchStr, " ")
    txtFunctionary.Text = Trim(gbSearchStr)
    txtFunctionary.Tag = gbSearchID
    gbSearchStr = ""
    gbSearchID = -1
End Sub
Private Sub cmdIMPO_Click()
    gbSearchID = -1                                         ''  Setting the Search ID to -1
    frmSearchSubsidiaryAccountHeads.SubLedgerType = 1       ''  1. Implementing Officer
    frmSearchSubsidiaryAccountHeads.Show vbModal
    txtImpo.SetFocus
End Sub
Private Sub cmdMicroSector_Click()
    frmSearchMasters.Connection = enuSourceString.Sulekha
    If txtMicroHead.Tag = "" Then
        frmSearchMasters.SQLQry = "Select intMicroSecID,chvEngMicroSector from M_MicroSector Where intMicroSecID IN ( " & txtMicroHead.Text & ")"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtMicroSector.Text = gbSearchStr
            txtMicroSector.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    Else
        txtMicroSector.Text = ""
        txtMicroSector.Tag = ""
    End If
End Sub

Private Sub cmdPreviousBank_Click()
    Dim mSql As String
     mSql = ""
     mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
     mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID   "
     frmSearchAccountHeads.SQLString = mSql
     frmSearchAccountHeads.cmdSearch.Enabled = False
     frmSearchAccountHeads.Show vbModal
      If gbSearchID > 0 Then
        Dim objAc   As New clsAccounts
        objAc.SetAccounts (gbSearchID)
        gbSearchID = -1
        gbSearchStr = ""
        If objAc.AccountHeadID > 0 Then
            txtPreviousBank.Text = objAc.AccountHead
            txtPreviousBank.Tag = objAc.AccountHeadID
        End If
     End If
End Sub

Private Sub cmdPreviousDateVerify_Click()
   Call FillRequisitionBeforedtOnlineDate
End Sub

Private Sub cmdSearchPO_Click()
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mAccID As Integer
    Dim mFunctionId As Integer
    Dim mFunctionaryID As Integer
    Dim mSourceOfFundID As Integer
    Dim mImpID As Integer
    
    lblmsgPO.Visible = False
    lblmsgPV.Visible = False
    frmSearchPaymentOrder.PendingTask = 99
    frmSearchPaymentOrder.Show vbModal
    If gbSearchID > 0 Then
        txtPaymentOrder.Tag = gbSearchID
        txtPaymentOrder.Text = gbSearchStr
        If objdb.SetConnection(mCnn) Then
            mSql = " Select * from faPayOrder Where intPayOrderID=" & gbSearchID & " "
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtPaymentVoucher.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                txtPaymentVoucher.Tag = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                mFunctionId = Rec!intFunctionID
                mFunctionaryID = Rec!intFunctionaryID
                mSourceOfFundID = Rec!intSourceOfFundID
                mAccID = Rec!intCashOrBankHeadID
                mImpID = Rec!intImplementingOfficerID
            
            End If
            Rec.Close
        End If
        'mCnn.Close
'        If mFunctionID <> txtFunction.Tag Or mFunctionaryID <> txtFunctionary.Tag Or mSourceOfFundID <> txtSourceOfFund.Tag Or _
'        mAccID <> txtGrossExpenditureHead.Tag Or mImpID <> txtIMPO.Tag Then
        If mAccID <> txtGrossExpenditureHead.Tag Then
            MsgBox "The Payment Order Linked is wrong...Please Link correct Payment Order", vbInformation
            txtPaymentOrder.Tag = ""
            txtPaymentOrder.Text = ""
            txtPaymentVoucher.Text = ""
            txtPaymentVoucher.Tag = ""
            Exit Sub
        End If
        fillPODetails (txtPaymentOrder.Tag)
        fillPVDetails (txtPaymentVoucher.Tag)
        gbSearchID = -1
        gbSearchStr = ""
    End If
End Sub
Private Function checkPayOrderLink(mPayorderID As Variant) As Boolean
    Dim mCnnChild  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim RecChild   As New ADODB.Recordset
    Dim mSql  As String
    Dim mAccID As Integer
    Dim mFunctionId As Integer
    Dim mFunctionaryID As Integer
    Dim mSourceOfFundID As Integer
    Dim mImpID As Integer
    
     If objdb.SetConnection(mCnnChild) Then
            mSql = " Select * from faPayOrder Where intAllotmentID= " & txtRequisitionNo.Tag & " "
            RecChild.Open mSql, mCnnChild
            If Not (RecChild.EOF And RecChild.BOF) Then
                mFunctionId = RecChild!intFunctionID
                mFunctionaryID = RecChild!intFunctionaryID
                mSourceOfFundID = RecChild!intSourceOfFundID
                mAccID = RecChild!intCashOrBankHeadID
                mImpID = RecChild!intImplementingOfficerID
            
            End If
            RecChild.Close
      End If
        mCnnChild.Close
        If mFunctionId <> 0 Then
'            If mFunctionID <> txtFunction.Tag Or mFunctionaryID <> txtFunctionary.Tag Or mSourceOfFundID <> txtSourceOfFund.Tag Or _
'            mAccID <> txtGrossExpenditureHead.Tag Or mImpID <> txtIMPO.Tag Then
            If mAccID <> txtGrossExpenditureHead.Tag Then
                MsgBox "The Payment Order Linked is wrong(Since the Function/Functionary/SourceOfFund selected may be not matching).Cancel the PaymentOrder", vbInformation
                checkPayOrderLink = False
            Else
                checkPayOrderLink = True
            End If
        Else
            checkPayOrderLink = True
        End If
End Function
Private Sub cmdSearchProject_Click()
    txtSubSector.Enabled = True
    txtMicroSector.Enabled = True
    cmdMicroSector.Enabled = True
    frmSearchProjects.PreviousYearMode = 1
    frmSearchProjects.Show vbModal
    txtProjectNo.SetFocus
    lblmsg.Caption = ""
    cmdVerify.Enabled = True
End Sub

Private Sub cmdSearchPV_Click()
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mSql    As String
    Dim Rec As New ADODB.Recordset
    
    frmSearchVouchers.chkContra.Visible = False
    frmSearchVouchers.chkReceipt.Visible = False
    frmSearchVouchers.chkJournal.Visible = False
    frmSearchVouchers.chkPayment.value = 1
    frmSearchVouchers.Show vbModal
    If gbSearchID <> -1 Then
        txtPaymentVoucher.Text = gbSearchCode
        txtPaymentVoucher.Tag = gbSearchID
        gbSearchCode = ""
        gbSearchID = -1
    End If
    If objdb.SetConnection(mCnn) Then
        If txtPaymentVoucher.Tag <> "" Then
                mSql = ""
                mSql = " SELECT     intVoucherNo, intVoucherID,intKeyID2 From faVouchers"
                mSql = mSql + " WHERE intVoucherID  = " & txtPaymentVoucher.Tag
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    Dim mPayOrderNo As Variant
                    If Rec!intVoucherNo = txtPaymentVoucher.Text Then
                        mPayOrderNo = Rec!intKeyID2
                        If mPayOrderNo > 0 Then
                            If mPayOrderNo <> txtPaymentOrder.Text Then
                                MsgBox "The Voucher selected is not matching", vbInformation, "Saankhya"
                                txtPaymentVoucher.Text = ""
                                txtPaymentVoucher.Tag = ""
                                Exit Sub
                            End If
                        End If
                    End If
                 End If
            Rec.Close
        End If
    End If
End Sub

Private Sub cmdSourceOfFund_Click()
    frmSearchMasters.Connection = enuSourceString.Saankhya
    frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
    frmSearchMasters.QrySP = Qyery
    frmSearchMasters.Show vbModal
    If gbSearchID <> -1 Then
        txtSourceOfFund.Text = gbSearchStr
        txtSourceOfFund.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        ProjectSourceOfFundValidation (txtSourceOfFund.Tag)
    End If
End Sub
Private Function ProjectSourceOfFundValidation(mSourceID As Integer)
Dim objAccounts As New clsAccounts
    If mSourceID = 1 Then
        txtTreasury.Tag = gbAcHeadIDTreasuryAccount2
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 29 Then
        txtTreasury.Tag = gbAcHeadIDTreasuryAccount6
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount6)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 30 Then
        txtTreasury.Tag = gbAcHeadIDTreasuryAccount7
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount7)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 16 Then
        txtTreasury.Tag = gbAcHeadIDTreasuryAccount3
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount3)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 17 Then
    txtTreasury.Tag = gbAcHeadIDTreasuryAccount3
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount3)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 25 Then
    txtTreasury.Tag = gbAcHeadIDTreasuryAccount4
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount4)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 26 Then
    txtTreasury.Tag = gbAcHeadIDTreasuryAccount5
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount5)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 27 Then
    txtTreasury.Tag = gbAcHeadIDTreasuryAccount2
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
        txtTreasury.Text = objAccounts.AccountHead
    ElseIf mSourceID = 28 Then
    txtTreasury.Tag = gbAcHeadIDTreasuryAccount2
        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
        txtTreasury.Text = objAccounts.AccountHead
    Else
        cmdTreasury.Enabled = True
        txtTreasury.Tag = ""
        txtTreasury.Text = ""
        'Call cmdTreasury_Click
    End If
End Function
Private Sub cmdSubSector_Click()
    frmSearchMasters.Connection = enuSourceString.Saankhya
    frmSearchMasters.SQLQry = "Select intSubSecID,vchSubSectorEng from faSubSector "
    frmSearchMasters.QrySP = Qyery
    frmSearchMasters.Show vbModal
    If gbSearchID <> -1 Then
        txtSubSector.Text = gbSearchStr
        txtSubSector.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End If
End Sub
Private Function SaveValidation() As Boolean
      
     If Trim(txtFunction.Text) = "" Then
        MsgBox "Select Function", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     If Trim(txtFunctionary.Text) = "" Then
        MsgBox "Select Functionary", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     If Trim(txtProjectNo.Text) = "" Then
        MsgBox "Select Project", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     If Trim(txtSourceOfFund.Text) = "" Then
        MsgBox "Select Source Of Fund", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     If Trim(txtSubSector.Text) = "" Then
        MsgBox "Select SubSector", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     If Trim(txtMicroSector.Text) = "" Then
        MsgBox "Select Micro sector", vbInformation, "Saankhya"
        SaveValidation = False
        Exit Function
     End If
     SaveValidation = True
End Function

Private Sub cmdUndoRequisition_Click()
    cmdUndoRequisition.Enabled = False
    cmdSearchProject.Enabled = True
    cmdTreasury.Enabled = True
    cmdVerify.Enabled = True
End Sub

Private Sub cmdVerify_Click()
     Dim mCnn    As New ADODB.Connection
     Dim mCnnSulekha   As New ADODB.Connection
     Dim objdb   As New clsDB
     Dim mArrIn  As Variant
     Dim mArrInChild  As Variant
     Dim mArrOut  As Variant
     Dim mSql As String
     Dim mMsg  As String
     
     Call GetDtOnlineDate
     If SaveValidation = False Then
        Exit Sub
     Else
        If objdb.SetConnection(mCnn) Then
            If CDate(txtAllotmentDate.Text) < CDate(dtOnlinedate) Then
                fraPVDetails.Visible = False
                frmPayVDetails.Visible = False
                cmdVerifyPaymentVoucher.Visible = False
                fraReqOnlineDate.Visible = True
                mSql = "Update faAllotments set  numProjectID =" & val(txtProjectNo.Tag) & ", vchProjectNo ='" & txtProjectNo.Text & "', intSourceID =" & txtSourceOfFund.Tag & ","
                mSql = mSql + " intFundCategoryID =" & txtCategory.Tag & ", intFunctionaryID =" & txtFunctionary.Tag & ", intFunctionID =" & txtFunction.Tag & ", intAccountHeadID =" & txtGrossExpenditureHead.Tag & ", vchAccountHeadCode ='" & txtExpAccHeadCode.Text & "' "
                mSql = mSql + " ,intSubSecID=" & txtSubSector.Tag & ",intMircoSectorID=" & txtMicroSector.Tag & " ,"
                mSql = mSql + " tnyProjectStatus=1 "
                mSql = mSql + " Where intID=" & txtRequisitionNo.Tag & "  "
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                MsgBox "Updated Successfully ", vbInformation
                cmdVerify.Enabled = False
                chkOldVoucher.Enabled = True
            Else
                 If ProjectExpHeadValidation = False Then
                    mMsg = "AccountHead Selected is Not Matching" + vbCrLf
                    mMsg = mMsg + "a) If the Selected AccountHead is Wrong in Requisition,Cancel the Requisition" + vbCrLf
                    mMsg = mMsg + "b) Else change the AccountHead In Project."
                    MsgBox mMsg, vbInformation, "Saankhya"
                    'MsgBox "The Payment Order Linked is incorrect(Since the selected head is not matching).Cancel the PaymentOrder and link the correct one", vbInformation, "Saankhya"
                    Exit Sub
                    'cmdVerify.Enabled = False
                    lblmsgPO.Visible = True
                    lblmsgPV.Visible = True
                    Exit Sub
                 End If
                 If checkPayOrderLink(cmdSearchProject.Tag) = False Then
                    lblmsgPO.Visible = True
                    lblmsgPV.Visible = True
                 End If
                 'If objDB.SetConnection(mCnn) Then
                     mSql = "Update faAllotments set  numProjectID =" & val(txtProjectNo.Tag) & ", vchProjectNo ='" & txtProjectNo.Text & "', intSourceID =" & txtSourceOfFund.Tag & ","
                     mSql = mSql + " intFundCategoryID =" & txtCategory.Tag & ", intFunctionaryID =" & txtFunctionary.Tag & ", intFunctionID =" & txtFunction.Tag & ", intAccountHeadID =" & txtGrossExpenditureHead.Tag & ", vchAccountHeadCode ='" & txtExpAccHeadCode.Text & "' "
                     mSql = mSql + " ,intSubSecID=" & txtSubSector.Tag & ",intMircoSectorID=" & txtMicroSector.Tag & " ,"
                     mSql = mSql + " tnyProjectStatus=1 "
                     mSql = mSql + " Where intID=" & txtRequisitionNo.Tag & "  "
                     objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                     MsgBox "Updated Successfully ", vbInformation
                     cmdVerify.Enabled = False
                     cmdUndoRequisition.Enabled = True
                     fraPVDetails.Enabled = True
                     cmdVerifyPaymentVoucher.Enabled = True
                     cmdSearchPO.Enabled = True
                     cmdSearchPV.Enabled = True
                'End If
                'mCnn.Close
            End If
        End If
     End If
       If (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
            mSql = "Update RequisitionDetails set decProjectID= " & val(txtProjectNo.Tag) & " ,tnyTransfer=0 where intReqID = " & val(txtRequisitionNo.Tag) & "  "
            objdb.ExecuteSP mSql, , , , mCnnSulekha, adCmdText
          
       Else
            MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
            Exit Sub
        End If
        
End Sub
Private Function ProjectPayVoucherValidation() As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mTrAccHeadId As Integer

    
    If objdb.SetConnection(mCnn) Then
        mSql = " Select * from faVouchers  Where intVoucherID=" & val(txtPaymentVoucher.Tag)
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mTrAccHeadId = Rec!intKeyID1
        End If
        Rec.Close
    End If
    If val(txtSourceOfFund.Tag) = 1 Or val(txtSourceOfFund.Tag) = 29 Or val(txtSourceOfFund.Tag) = 30 Or val(txtSourceOfFund.Tag) = 16 _
         Or val(txtSourceOfFund.Tag) = 17 Or val(txtSourceOfFund.Tag) = 25 Or val(txtSourceOfFund.Tag) = 26 Or val(txtSourceOfFund.Tag) = 28 _
          Or val(txtSourceOfFund.Tag) = 29 Then
        If mTrAccHeadId <> val(txtTreasury.Tag) Then
            ProjectPayVoucherValidation = False
        Else
            ProjectPayVoucherValidation = True
        End If
    Else
        ProjectPayVoucherValidation = True
    End If
    'mCnn.Close
End Function
Private Sub cmdTreasury_Click()
    Dim mSql As String
    
     mSql = ""
     mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
     mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID WHERE  tinHiddenFlag = 0 AND (faAccountHeads.vchAccountHeadCode Like '450%' Or faAccountHeads.vchAccountHeadCode Like '45065%' Or faAccountHeads.vchAccountHeadCode Like '45025%' ) "
     frmSearchAccountHeads.SQLString = mSql
     frmSearchAccountHeads.cmdSearch.Enabled = False
     frmSearchAccountHeads.Show vbModal
      If gbSearchID > 0 Then
        Dim objAc   As New clsAccounts
        objAc.SetAccounts (gbSearchID)
        gbSearchID = -1
        gbSearchStr = ""
        If objAc.AccountHeadID > 0 Then
            txtTreasury.Text = objAc.AccountHead
            txtTreasury.Tag = objAc.AccountHeadID
        End If
     End If
End Sub

Private Sub cmdVerifyPaymentVoucher_Click()
    Dim mCnn    As New ADODB.Connection
    Dim mCnnSulekha    As New ADODB.Connection
    Dim objdb   As New clsDB
    Dim mArrIn  As Variant
    Dim mArrInChild  As Variant
    Dim mArrOut  As Variant
    Dim mSql As String
    Dim arrInput As Variant
    Dim Rec   As New ADODB.Recordset
    Dim mCount As Integer

    
     If Trim(txtPaymentOrder.Text) = "" Then
        MsgBox "Select Payment PO", vbInformation, "Saankhya"
        Exit Sub
     End If
     If Trim(txtPaymentVoucher.Text) = "" Then
        MsgBox "Select Payment Voucher", vbInformation, "Saankhya"
        Exit Sub
     End If
     If ProjectPayVoucherValidation = False Then
        MsgBox "The Treasury mapped is different", vbInformation, "Saankhya"
        
        If val(txtSourceOfFund.Tag) <> 1 Or val(txtSourceOfFund.Tag) <> 29 Or val(txtSourceOfFund.Tag) <> 30 Or val(txtSourceOfFund.Tag) <> 16 _
         Or val(txtSourceOfFund.Tag) <> 17 Or val(txtSourceOfFund.Tag) <> 25 Or val(txtSourceOfFund.Tag) <> 26 Or val(txtSourceOfFund.Tag) <> 28 _
          Or val(txtSourceOfFund.Tag) <> 29 Then
            cmdUndoRequisition.Visible = True
            cmdUndoRequisition.Enabled = True
            fraReqDetails.Enabled = True
            Exit Sub
        Else
            cmdVerifyPaymentVoucher.Enabled = False
            Exit Sub
        End If
     End If
     If objdb.SetConnection(mCnn) Then
        mSql = "Update faAllotments set  tnyProjectStatus=2 Where intID=" & txtRequisitionNo.Tag & "  "
        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        
        mSql = "Update faPayOrder set  intAllotmentID=" & txtRequisitionNo.Tag & ""
        mSql = mSql + " ,intImplementingOfficerID=" & txtImpo.Tag & ",intSourceOfFundID =" & txtSourceOfFund.Tag & ",intFunctionaryID =" & txtFunctionary.Tag & ", intFunctionID =" & txtFunction.Tag & ", intCashOrBankHeadID =" & txtGrossExpenditureHead.Tag & ""
        mSql = mSql + " Where intPayOrderID=" & txtPaymentOrder.Tag & ""
        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        
        mSql = "Update faVouchers  set  intKeyID2='" & txtPaymentOrder.Text & "' Where intVoucherID=" & txtPaymentVoucher.Tag & "  "
        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        
        
        If (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
'            mSQL = "Update RequisitionDetails set decProjectID= " & txtProjectNo.Tag & " ,tnyTransfer=0 where intReqID = " & val(txtRequisitionNo.Tag) & "  "
'            objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
            
            
            mSql = "Select * from ExpenseDetails where intVoucherID = " & txtPaymentVoucher.Tag & "  "
            Rec.Open mSql, mCnnSulekha
            If Not (Rec.EOF And Rec.BOF) Then
                mCount = 1
            Else
                mCount = 0
            End If
            Rec.Close
            If mCount = 1 Then
                mSql = "Update ExpenseDetails set decProjectID= " & val(txtProjectNo.Tag) & " where intVoucherID = " & txtPaymentVoucher.Tag & "  "
                objdb.ExecuteSP mSql, , , , mCnnSulekha, adCmdText
            Else
                arrInput = Array(gbLBID, _
                                    gbFinancialYearID - 1, _
                                    val(txtProjectNo.Tag), _
                                    -1, val(txtSourceOfFund.Tag), _
                                    val(txtAmount), _
                                    txtPaymentVoucher.Tag)
    
                    objdb.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
            End If
            
        Else
            MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
            Exit Sub
        End If
        
        
        MsgBox "Updated Successfully ", vbInformation
        cmdVerifyPaymentVoucher.Enabled = False
        chkOldVoucher.Enabled = False
        cmdUndoRequisition.Enabled = False
    End If
End Sub
Private Sub Form_Activate()
    Call GetRequisitionDetails
End Sub
Private Sub FormInitialize()
    
    If CheckProjectValidation = True Then
        
        'cmdSearchProject.Enabled = False
        gbSearchStr = mProjectID
        gbSearchID = mSourceOfFundID
    Else
        lblmsg.Caption = "Please Map the appropruate Project"
        cmdSearchProject.Enabled = True
        cmdSourceOfFund.Enabled = True
        cmdImpo.Enabled = True
        cmdFunction.Enabled = True
        cmdFunctionary.Enabled = True
        chkOldVoucher.Enabled = False
        'cmdExpHead.Enabled = True
        'cmdTreasury.Enabled = True
    End If
End Sub
Private Function CheckProjectValidation() As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mNewProjectId As Variant
    Dim mNewProjectNo As Variant
    Dim mNewSubSectorId As Integer
    Dim mReqProjectId As Variant
    Dim mReqProjectNo As Variant
    Dim mReqSubSectorId As Integer
    
    If objdb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
        mSql = " SELECT ProjectDetails.decProjectID ProjectID,intProjectSlNo,chvProjectSlNo,intFundSrcID,intSubSecID FROM ProjectDetails"
        mSql = mSql + " INNER JOIN FundDetails ON FundDetails.decProjectID=ProjectDetails.decProjectID"
        mSql = mSql + " Where ProjectDetails.chvProjectSlNo = '" & txtProjectNo.Text & "'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
           
            mNewProjectId = IIf(IsNull(Rec!ProjectID), 0, Rec!ProjectID)
            mNewProjectNo = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
            mNewSubSectorId = IIf(IsNull(Rec!intSubSecID), "", Rec!intSubSecID)
        Else
            mNewProjectId = 0
            mNewProjectNo = ""
            mNewSubSectorId = 0
        End If
        If mNewProjectId <> val(txtProjectNo.Tag) Or mNewProjectNo <> txtProjectNo.Text Then
            CheckProjectValidation = False
        Else
            CheckProjectValidation = True
        End If
        Rec.Close
    End If
    mCnn.Close
End Function
Private Sub Form_Load()
    XPC.InitSubClassing
    lblmsg.Caption = ""
End Sub
Private Sub lblClose_Click()
    fraReqOnlineDate.Visible = False
    fraPVDetails.Visible = True
    frmPayVDetails.Visible = True
    fraPVDetails.Enabled = True
    frmPayVDetails.Enabled = True
    cmdVerifyPaymentVoucher.Visible = True
    chkOldVoucher.Visible = True
End Sub

Private Sub txtFunction_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub
Private Sub txtFunctionary_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub

Private Sub txtGrossExpenditureHead_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub

Private Sub txtIMPO_GotFocus()
    If gbSearchID > 0 Then
        Dim objSubLedger As New clsSubLedger
        objSubLedger.SetSubLedgerDetails (gbSearchID)
        If objSubLedger.SubsidiaryAccountHeadID Then
            txtImpo.Tag = IIf(IsNull(objSubLedger.SubsidiaryAccountHeadID), 0, objSubLedger.SubsidiaryAccountHeadID)
            txtImpo.Text = IIf(IsNull(objSubLedger.NameOfSubLedger), "", objSubLedger.NameOfSubLedger)
        Else
            txtImpo.Tag = ""
            txtImpo.Text = ""
        End If
    End If
        gbSearchID = -1
End Sub

Private Sub txtIMPO_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub
Private Sub txtPreviousBank_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtPreviousInstDate_LostFocus()
        If txtPreviousInstDate.Text <> "" Then
            If CheckDateInMMM(txtPreviousInstDate.Text) >= gbTransactionDate Then
                txtPreviousInstDate.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
            Else
                txtPreviousInstDate.Text = Format(CheckDateInMMM(txtPreviousInstDate.Text), "dd/mmm/yyyy")
            End If
        End If
End Sub

Private Sub txtPreviousInstNo_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
               KeyAscii = 0
  End If
End Sub

Private Sub txtPreviousVoucher_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
               KeyAscii = 0
    End If
End Sub
Private Sub txtPreviousVoucherDate_LostFocus()
    If txtPreviousVoucherDate.Text <> "" Then
            If CheckDateInMMM(txtPreviousVoucherDate.Text) >= gbTransactionDate Then
                txtPreviousVoucherDate.Text = Format(gbTransactionDate, "dd/mmm/yyyy")
            Else
                txtPreviousVoucherDate.Text = Format(CheckDateInMMM(txtPreviousVoucherDate.Text), "dd/mmm/yyyy")
            End If
       End If
End Sub

Private Sub txtProjectNameEng_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub
Private Sub txtProjectName_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub

Private Sub txtProjectNo_GotFocus()
    Dim objProj As New clsProject
    Dim objProFund As New clsProjectFund
    'Dim mProjectID As Variant
    'Dim mSourceOfFundID As Variant
    Dim mSubsectorID As Integer
    Dim mintCategoryID As Integer
    Dim mCol As Collection
    Dim mRow As Integer
    Dim mCnn    As New ADODB.Connection
    Dim mCnPlan As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim RecChild     As New ADODB.Recordset
    Dim obj     As New clsDB
    Dim mSql    As String
    Dim mWHERE  As String
    Dim mCapitalExpFlag As Boolean
    Dim mMicroSectorCount As Integer
    Dim mMicroHeads As Integer
    Dim mMicroSectorID As Integer

    mProjectID = gbSearchStr
    mSourceOfFundID = gbSearchID
'    If ProjectPOValiadtion(mProjectID) = False Then
'        MsgBox "The PAyment Order linked for the Selected Project is not correct....", vbInformation, "Saankhya"
'    End If
    If val(gbSearchStr) > 0 Then
        objProj.SetProject mProjectID, gbFinancialYearID - 1
        If objProj.ProjectID > 0 Then
            txtProjectName.Visible = False
            txtProjectNameEng.Visible = True
            txtProjectNameEng.Text = objProj.ProjectNameEnglish
            txtProjectNo.Text = objProj.ProjectSerialNo
            txtProjectNo.Tag = objProj.ProjectID
            txtCategory.Tag = objProj.ProjCatID
            mintCategoryID = objProj.ProjCatID
            txtSourceOfFund.Tag = mSourceOfFundID
            txtSourceOfFund.Text = objProj.FindSourceOfFund(mSourceOfFundID)
            txtSourceOfFund.Enabled = False
            ProjectSourceOfFundValidation (txtSourceOfFund.Tag)
            mSubsectorID = objProj.SubSectorID
            
            Set mCol = objProj.GetFundDetails(CInt(gbFinancialYearID - 1), objProj.ProjectID)
            For mRow = 1 To mCol.count
                Set objProFund = mCol.Item(mRow)
                If objProFund.SourceOfFundID = mSourceOfFundID Then
                    txtProjCost.Text = objProFund.SourceWiseAmount
                    Exit For
                End If
            Next mRow
             
          
        End If
    
        mSql = " SELECT faSubSectorHeads.intCategoryID,vchTransactionCategory, intSubSectorID, vchSubSectorCode, vchSubSector, "
        mSql = mSql + " faSubSectorHeads.intAccountHeadID, faAccountHeads.vchAccountHeadCode, vchAccountHead, "
        mSql = mSql + " faSubSectorHeads.intFunctionID, faFunctions.vchFunctionCode, vchFunction, "
        mSql = mSql + " faFunctionaryFunctions.intFunctionaryID, vchFunctionary, "
        mSql = mSql + " faSubSectorHeads.intTransactionTypeID , vchTransactionType "
        mSql = mSql + " FROM faSubSectorHeads "
        mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faSubSectorHeads.intAccountHeadID "
        mSql = mSql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faSubSectorHeads.intFunctionID "
        mSql = mSql + " LEFT JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
        mSql = mSql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID "
        mSql = mSql + " INNER JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faSubSectorHeads.intCategoryID"
        mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faSubSectorHeads.intTransactionTypeID "
        mSql = mSql + " Where faSubSectorHeads.intSubSectorID = " & mSubsectorID & " And faSubSectorHeads.intCategoryID = " & val(txtCategory.Tag)
        
        If obj.SetConnection(mCnn) Then
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
            If Not (Rec.EOF And Rec.BOF) Then
                txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                txtCategory.Enabled = False
                
                txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            
                txtGrossExpenditureHead.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                txtExpAccHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                txtGrossExpenditureHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            Else
                txtCategory.Text = ""
                txtFunction.Tag = ""
                txtFunction.Text = ""
                cmdFunction.Enabled = True
            
                txtGrossExpenditureHead.Tag = ""
                txtExpAccHeadCode.Text = ""
                txtGrossExpenditureHead.Text = ""
                cmdExpHead.Enabled = True
                
                txtFunctionary.Tag = ""
                txtFunctionary.Text = ""
                
                cmdFunction.Enabled = True
                cmdFunctionary.Enabled = True
            
            End If
            Rec.Close
        End If
        'mCnn.Close
        'Finding SubSector from faSubSector
        mSql = "Select intSubSecID,vchSubSector,vchSubSectorEng from faSubSector Where intSubSecID=" & mSubsectorID & " "
        If obj.SetConnection(mCnn) Then
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
            If Not (Rec.EOF And Rec.BOF) Then
                txtSubSector.Text = Rec!vchSubSectorEng
                txtSubSector.Tag = Rec!intSubSecID
                txtSubSector.Enabled = False
            End If
            Rec.Close
        End If
        'mCnn.Close
        ' EXTRACTING MicrosectorIDs From Plan-Project
        If obj.CreateNewConnection(mCnPlan, enuSourceString.Sulekha) Then
            mSql = " SELECT MicroSector.intMicroSecID  FROM MicroSector WHERE decProjectID = " & val(txtProjectNo.Tag)  '118600160078
            Rec.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
            mMicroSectorCount = 0
            If Not (Rec.BOF And Rec.EOF) Then
                mMicroSectorID = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                While Not Rec.EOF
                    mMicroSectorCount = mMicroSectorCount + 1
                    If Len(mWHERE) > 0 Then
                        mWHERE = mWHERE & ", " & Rec!intMicroSecID
                    Else
                        mWHERE = Trim(str(Rec!intMicroSecID))
                    End If
                    Rec.MoveNext
                Wend
            End If
            Rec.Close
            If mMicroSectorCount = 1 Then
                mSql = "SELECT intMicroSecID,chvEngMicroSector FROM M_MicroSector WHERE intMicroSecID= " & mMicroSectorID & ""
                RecChild.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
                If Not (RecChild.BOF And RecChild.EOF) Then
                    txtMicroSector.Tag = RecChild!intMicroSecID
                    txtMicroSector.Text = RecChild!chvEngMicroSector
                    cmdMicroSector.Enabled = False
                    txtMicroSector.Enabled = False
                End If
            Else
                cmdMicroSector.Enabled = True
                txtMicroHead.Text = mWHERE
                Call cmdMicroSector_Click
            End If
            'RecChild.Close
        End If
        mCnPlan.Close
    
        'Finding Account Heads from MicroSectors - CAPITAL EXPENDITURE
        If mMicroSectorCount > 0 Then
            If obj.SetConnection(mCnn) Then
            mSql = "SELECT Distinct intAccountHeadID FROM faMicroSectorHeads WHERE intMircoSectorID IN ( " & mWHERE & ")" '323,324,325,326,327,350)
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
                If Not (Rec.EOF And Rec.BOF) Then
                    mWHERE = ""
                    mMicroHeads = 0
                    mCapitalExpFlag = True
                    While Not Rec.EOF
                        mMicroHeads = mMicroHeads + 1
                        If Len(mWHERE) > 0 Then
                            mWHERE = mWHERE & ", " & Rec!intAccountHeadID
                        Else
                            mWHERE = Trim(str(Rec!intAccountHeadID))
                        End If
                        Rec.MoveNext
                    Wend
                Else
                    mCapitalExpFlag = False
                End If
                Rec.Close
            End If
            'mCnn.Close
        End If
    
        'mSQL = "SELECT * FROM faAccountHeads WHERE intAccountHeadID IN  (" & mWHERE & ")"
        If mMicroHeads > 0 Then
            mSql = "SELECT (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where  intAccountHeadID IN  (" & mWHERE & ") Order By faAccountHeads.vchAccountHeadCode"
            cmdExpHead.Tag = mSql
        Else
            cmdExpHead.Tag = ""
        End If
    
        If mCapitalExpFlag Then
            Select Case mMicroHeads
            Case Is = 0
                cmdExpHead.Enabled = True
            Case Is = 1
                Dim objAcc As New clsAccounts
                objAcc.SetAccountID val(mWHERE)
                If objAcc.AccountHeadID > 0 Then
                    txtGrossExpenditureHead.Tag = objAcc.AccountHeadID
                    txtExpAccHeadCode.Text = objAcc.AccountCode
                    txtGrossExpenditureHead.Text = objAcc.AccountHead
                    cmdExpHead.Enabled = True
                End If
            Case Else
                txtGrossExpenditureHead.Tag = ""
                txtExpAccHeadCode.Text = ""
                txtGrossExpenditureHead.Text = ""
                cmdExpHead.Enabled = True
                
            End Select
        End If
        If ProjectPOValiadtion(mProjectID) = False Then
            MsgBox "The PAyment Order linked for the Selected Project is not correct....", vbInformation, "Saankhya"
        End If
    End If
End Sub
Private Function ProjectPOValiadtion(mProjectID As Variant) As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mReqID As Integer
    Dim mFlag As Boolean
    Dim mFunctionId As Integer
    Dim mFunctionaryID As Integer
    Dim mSourceOfFundID As Integer
    Dim mAccHeadId As Integer
    Dim mImpID As Integer
    Dim mReqFunctionID As Integer
    Dim mReqFunctionaryID As Integer
    Dim mReqSourceOfFundID As Integer
    Dim mReqAccHeadID As Integer
    Dim mReqImpID As Integer
    Dim mPayorderID As Integer
    Dim mAllotmentID As Integer
    Dim mNewAmt As Double
    Dim mBalAvail As Double
    
    If mProjectID <> "" Then
        If objdb.SetConnection(mCnn) Then
            mSql = " Select * from faPayOrder "
            mSql = mSql + " Inner Join  faVouchers On faVouchers.intVoucherID=faPayOrder.intVoucherID"
            'mSQL = mSQL + " Inner Join faAllotments on faPayOrder.numProjectNo=faAllotments.numProjectID "
            'mSQL = mSQL + " Where numProjectNo = " & mProjectId & " And intSourceOfFundID=" & val(txtSourceOfFund.Tag) & " "
            mSql = mSql + " Inner Join faAllotments on faAllotments.intID = faPayOrder.intAllotmentID"
            mSql = mSql + " Where faPayOrder.intSourceOfFundID = " & val(txtSourceOfFund.Tag) & " And numProjectID = " & mProjectID & ""
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFunctionId = IIf(IsNull(Rec!intFunctionID), 0, Rec!intFunctionID)
                mFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), 0, Rec!intFunctionaryID)
                mSourceOfFundID = IIf(IsNull(Rec!intSourceOfFundID), 0, Rec!intSourceOfFundID)
                mAccHeadId = IIf(IsNull(Rec!intCashOrBankHeadID), 0, Rec!intCashOrBankHeadID)
                mImpID = IIf(IsNull(Rec!intImplementingOfficerID), 0, Rec!intImplementingOfficerID)
                mPayorderID = IIf(IsNull(Rec!intPayOrderID), 0, Rec!intPayOrderID)
                mAllotmentID = IIf(IsNull(Rec!intAllotmentID), 0, Rec!intAllotmentID)
                mNewAmt = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
           Else
                ProjectPOValiadtion = True
           End If
           Rec.Close
           If mAllotmentID <> 0 Then
                mSql = " Select * from faAllotments "
                mSql = mSql + " Where intID = " & mAllotmentID & ""
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mReqID = Rec!intID
                    mReqFunctionID = IIf(IsNull(Rec!intFunctionID), 0, Rec!intFunctionID)
                    mReqFunctionaryID = IIf(IsNull(Rec!intFunctionaryID), 0, Rec!intFunctionaryID)
                    mReqSourceOfFundID = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
                    mReqAccHeadID = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
                    mReqImpID = IIf(IsNull(Rec!intImplementingOfficersID), 0, Rec!intImplementingOfficersID)
                End If
                Rec.Close
                If mReqFunctionID <> mFunctionId Or mReqFunctionaryID <> mFunctionaryID Or mReqSourceOfFundID <> mSourceOfFundID Or _
                mReqAccHeadID <> mAccHeadId Or mReqImpID <> mImpID Then
                    ProjectPOValiadtion = False
                Else
                    ProjectPOValiadtion = True
                End If
            End If
            If val(mNewAmt) = 0 Then
                 mSql = " Select * from faAllotments "
                 mSql = mSql + " Where numProjectID = " & mProjectID & " And tnyStatus=1 And tnyStage=2 "
                 mSql = mSql + " And intSourceID=" & val(txtSourceOfFund.Tag)
                 Rec.Open mSql, mCnn
                 If Not (Rec.EOF And Rec.BOF) Then
                    While Not Rec.EOF
                        mNewAmt = mNewAmt + IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
                        Rec.MoveNext
                    Wend
                 End If
                 Rec.Close
            End If
            'Amount Validation
            If val(txtAmount.Tag) <> val(txtProjectNo.Tag) Then
                mBalAvail = val(txtProjCost.Text) - (val(txtAmount.Text) + val(mNewAmt))
            Else
                mBalAvail = val(txtProjCost.Text) - val(mNewAmt)
            End If
            'If val(txtAmount.Text) + val(mNewAmt) > val(txtProjCost.Text) Then
            If mBalAvail < 0 Then
                MsgBox "Amount Exceeded", vbInformation, "Saankhya"
                cmdVerify.Enabled = False
                Exit Function
            End If
        End If
        'mCnn.Close
    End If
End Function
Private Function ProjectLinkPOValidation(mProjectID As Variant) As Boolean
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mReqID As Integer
    Dim mFlag As Boolean
    Dim mFunctionId As Integer
    Dim mFunctionaryID As Integer
    Dim mSourceOfFundID As Integer
    Dim mAccHeadId As Integer
    Dim mImpID As Integer
    Dim mReqFunctionID As Integer
    Dim mReqFunctionaryID As Integer
    Dim mReqSourceOfFundID As Integer
    Dim mReqAccHeadID As Integer
    Dim mReqImpID As Integer
    Dim mPayorderID As Integer

    
    If mProjectID <> "" Then
        If objdb.SetConnection(mCnn) Then
            mSql = " Select * from faAllotments "
            mSql = mSql + " Where numProjectID =" & mProjectID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mReqID = Rec!intID
                mReqFunctionID = Rec!intFunctionID
                mReqFunctionaryID = Rec!intFunctionaryID
                mReqSourceOfFundID = Rec!intSourceID
                mReqAccHeadID = Rec!intAccountHeadID
                mReqImpID = Rec!intImplementingOfficersID
            End If
            Rec.Close
            mSql = " Select * from faPayOrder "
            mSql = mSql + " Where intAllotmentID = " & mReqID & ""
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mFunctionId = Rec!intFunctionID
                mFunctionaryID = Rec!intFunctionaryID
                mSourceOfFundID = Rec!intSourceOfFundID
                mAccHeadId = Rec!intCashOrBankHeadID
                mImpID = Rec!intImplementingOfficerID
                mPayorderID = Rec!intPayOrderID
            End If
            Rec.Close
        End If
        mCnn.Close
        If mReqFunctionID <> mFunctionId Or mReqFunctionaryID <> mFunctionaryID Or mReqSourceOfFundID <> mSourceOfFundID Or _
        mReqAccHeadID <> mAccHeadId Or mReqImpID <> mImpID Then
            ProjectLinkPOValidation = False
        Else
            ProjectLinkPOValidation = True
        End If
    End If
End Function

Private Sub GetRequisitionDetails()
    Dim mCnn  As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec   As New ADODB.Recordset
    Dim mSql  As String
    Dim mProjectStatus As Integer
    Dim objProj As New clsProject
    If objdb.SetConnection(mCnn) Then
        mSql = " Select *,faBanks.intAccountHeadID TreasuryAccHeadID,faTransactionCategory.intCategoryID CategoryID,faFunctions.intFunctionID FunctionID,faAccountHeads.intAccountHeadID AccountHeadID,faAccountHeads.vchAccountHeadCode AccountHeadCode from faAllotments "
        mSql = mSql + " INNER Join suSourceOfFund On faAllotments.intSourceID = suSourceOfFund.intSourceFundID"
        mSql = mSql + " Left Join faTransactionCategory On faAllotments.intFundCategoryID = faTransactionCategory.intCategoryID"
        mSql = mSql + " LEFT Join faPayOrder On faPayOrder.intAllotmentID=faAllotments.intID"
        mSql = mSql + " LEFT JOIN faVouchers on faVouchers.intVoucherID=faPayOrder.intVoucherID"
        mSql = mSql + " LEFT JOIN faBanks on faBanks.intAccountHeadID=faVouchers.intKeyID1"
        mSql = mSql + " INNER JOIN faFunctions on faFunctions.intFunctionID=faAllotments.intFunctionID"
        mSql = mSql + " INNER JOIN faFunctionaries On faFunctionaries.intFunctionaryID=faAllotments.intFunctionaryID"
        mSql = mSql + " INNER JOIN faAccountHeads On faAccountHeads.intAccountHeadID=faAllotments.intAccountHeadID "
        mSql = mSql + " Left JOIN faSubSector On faSubSector.intSubSecID=faAllotments.intSubSecID"
        mSql = mSql + " Left JOIN faMicroSectorHeads On faMicroSectorHeads.intMircoSectorID=faAllotments.intMircoSectorID"
        mSql = mSql + " Left JOIN suProjectDetails On suProjectDetails.decProjectID=faAllotments.numProjectID"
        mSql = mSql + " Where faAllotments.intID=" & mReqID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtRequisitionNo.Text = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
            txtRequisitionNo.Tag = IIf(IsNull(Rec!intID), 0, Rec!intID)
            txtAllotmentNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
            txtAllotmentDate.Text = IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate)
            txtProjectNo.Tag = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
            txtAmount.Tag = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
            cmdSearchProject.Tag = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
            txtSourceOfFund.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            txtSourceOfFund.Tag = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
            txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
            txtCategory.Tag = IIf(IsNull(Rec!CategoryID), 0, Rec!CategoryID)
            If txtCategory.Tag = 0 Then
                txtCategory.Enabled = False
            End If
            txtImpo.Tag = IIf(IsNull(Rec!intImplementingOfficersID), 0, Rec!intImplementingOfficersID)
            txtImpo.Text = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
            txtFunction.Tag = IIf(IsNull(Rec!FunctionID), 0, Rec!FunctionID)
            txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), 0, Rec!intFunctionaryID)
            txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            txtGrossExpenditureHead.Tag = IIf(IsNull(Rec!AccountHeadID), 0, Rec!AccountHeadID)
            txtExpAccHeadCode.Text = IIf(IsNull(Rec!AccountHeadCode), "", Rec!AccountHeadCode)
            txtGrossExpenditureHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            ProjectSourceOfFundValidation (txtSourceOfFund.Tag)
'            txtTreasury.Tag = IIf(IsNull(Rec!TreasuryAccHeadID), 0, Rec!TreasuryAccHeadID)
'            txtTreasury.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
            mProjectStatus = IIf(IsNull(Rec!tnyProjectStatus), 0, Rec!tnyProjectStatus)
            txtAmount.Text = IIf(IsNull(Rec!fltAuthorizedAmt), 0, Rec!fltAuthorizedAmt)
            txtSubSector.Enabled = False
            cmdSubSector.Enabled = False
            txtMicroSector.Enabled = False
            cmdMicroSector.Enabled = False
            cmdSearchProject.Enabled = False
            cmdSourceOfFund.Enabled = False
            cmdImpo.Enabled = False
            cmdFunction.Enabled = False
            cmdFunctionary.Enabled = False
            cmdExpHead.Enabled = True
            cmdTreasury.Enabled = False
            txtProjectName.Text = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
            txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
            txtPaymentOrder.Text = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
            txtPaymentOrder.Tag = IIf(IsNull(Rec!intPayOrderID), 0, Rec!intPayOrderID)
            txtPaymentVoucher.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            txtPaymentVoucher.Tag = IIf(IsNull(Rec!intVoucherID), 0, Rec!intVoucherID)
            fillPODetails (txtPaymentOrder.Tag)
            fillPVDetails (txtPaymentVoucher.Tag)
            If mProjectStatus = 0 Then
                objProj.SetProject val(txtProjectNo.Tag), gbFinancialYearID - 1
                If objProj.ProjectID > 0 Then
                    txtProjectName.Text = objProj.ProjectName
                    txtProjectNo.Text = objProj.ProjectSerialNo
                    If objProj.Status = 9 Then
                        objProj.SetProject val(txtProjectNo.Tag), gbFinancialYearID - 1
                        If objProj.ProjectID > 0 Then
                            txtProjectName.Visible = False
                            txtProjectNameEng.Visible = True
                            txtProjectNameEng.Text = objProj.ProjectNameEnglish
                            txtProjectNo.Text = objProj.ProjectSerialNo
                            Call GetSubSectorDetails(val(objProj.SubSectorID))
                            Call GetMicroSectorDetails
                        End If
                    Else
                        cmdSearchProject.Enabled = True
                    End If
                Else
                    objProj.SetProject val(txtProjectNo.Tag), gbFinancialYearID - 1
                    If objProj.ProjectID > 0 Then
                        txtProjectName.Visible = False
                        txtProjectNameEng.Visible = True
                        txtProjectNameEng.Text = objProj.ProjectNameEnglish
                        txtProjectNo.Text = objProj.ProjectSerialNo
                        Call GetSubSectorDetails(val(objProj.SubSectorID))
                        Call GetMicroSectorDetails
                    End If
                End If
                cmdVerifyPaymentVoucher.Enabled = False
                fraPVDetails.Enabled = False
                cmdSearchPO.Enabled = False
                cmdSearchPV.Enabled = False
                'chkOldVoucher.Enabled = False
                Call FormInitialize
            ElseIf mProjectStatus = 1 Then
                objProj.SetProject val(txtProjectNo.Tag), gbFinancialYearID - 1
                If objProj.ProjectID > 0 Then
                    txtProjectName.Visible = False
                    txtProjectNameEng.Visible = True
                    txtProjectNameEng.Text = objProj.ProjectNameEnglish
                    txtProjectNo.Text = objProj.ProjectSerialNo
                End If
                txtSubSector.Tag = IIf(IsNull(Rec!intSubSecID), 0, Rec!intSubSecID)
                txtSubSector.Text = IIf(IsNull(Rec!vchSubSectorEng), "", Rec!vchSubSectorEng)
                Call GetMicroSectorDetails
                'fraReqDetails.Enabled = False
                cmdVerifyPaymentVoucher.Enabled = True
                cmdSearchPO.Enabled = True
                cmdSearchPV.Enabled = True
                cmdVerify.Enabled = False
                If checkPayOrderLink(cmdSearchProject.Tag) = False Then
                   lblmsgPO.Visible = True
                   lblmsgPV.Visible = True
                End If
                'If txtPaymentVoucher.Tag <> 0 Then
                    Call GetDtOnlineDate
                    If CDate(txtAllotmentDate.Text) < CDate(dtOnlinedate) Then
                        fraPVDetails.Visible = False
                        frmPayVDetails.Visible = False
                        cmdVerifyPaymentVoucher.Visible = False
                        fraReqOnlineDate.Visible = True
                        cmdVerify.Enabled = False
                        chkOldVoucher.Enabled = True
                        Call GetRequisitionBeforedtOnlineDate
                        cmdPreviousDateVerify.Enabled = True
                    End If
                'End If
                cmdUndoRequisition.Visible = True
                fraReqDetails.Enabled = True
                chkRevisedPrjUpdate.Enabled = False
            ElseIf mProjectStatus = 2 Then
                objProj.SetProject val(txtProjectNo.Tag), gbFinancialYearID - 1
                If objProj.ProjectID > 0 Then
                    txtProjectName.Visible = False
                    txtProjectNameEng.Visible = True
                    txtProjectNameEng.Text = objProj.ProjectNameEnglish
                    txtProjectNo.Text = objProj.ProjectSerialNo
                End If
                Call GetSubSectorDetails(val(objProj.SubSectorID))
                Call GetMicroSectorDetails
                fraReqDetails.Enabled = False
                cmdVerifyPaymentVoucher.Enabled = True
                cmdSearchPO.Enabled = True
                cmdSearchPV.Enabled = True
                cmdVerify.Enabled = False
                cmdVerifyPaymentVoucher.Enabled = False
                fraPVDetails.Enabled = False
                frmPayVDetails.Enabled = False
                If txtPaymentVoucher.Tag = 0 Then
                    Call GetDtOnlineDate
                    If CDate(txtAllotmentDate.Text) < CDate(dtOnlinedate) Then
                        fraPVDetails.Visible = False
                        frmPayVDetails.Visible = False
                        cmdVerifyPaymentVoucher.Visible = False
                        fraReqOnlineDate.Visible = True
                        cmdVerify.Enabled = False
                        Call GetRequisitionBeforedtOnlineDate
                    End If
                End If
            End If
        Else
            txtAmount.Tag = ""
        End If
        'Rec.Close
    End If
    'mCnn.Close
End Sub
Private Sub GetSubSectorDetails(mSubsectorID As Integer)
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objdb     As New clsDB
    Dim mSql    As String
    
    
    If objdb.SetConnection(mCnn) Then
        mSql = " SELECT intSubSecID,vchSubSectorEng  FROM faSubSector WHERE intSubSecID = " & val(mSubsectorID)
        Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
            txtSubSector.Tag = IIf(IsNull(Rec!intSubSecID), 0, Rec!intSubSecID)
            txtSubSector.Text = IIf(IsNull(Rec!vchSubSectorEng), "", Rec!vchSubSectorEng)
        End If
        Rec.Close
    End If
    'mCnn.Close
End Sub
Private Sub GetMicroSectorDetails()
   
    Dim mCnPlan As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim RecChild     As New ADODB.Recordset
    Dim obj     As New clsDB
    Dim mSql    As String
    Dim mMicroSectorID As Integer
    
    If obj.CreateNewConnection(mCnPlan, enuSourceString.Sulekha) Then
        mSql = " SELECT MicroSector.intMicroSecID  FROM MicroSector WHERE decProjectID = " & val(txtProjectNo.Tag)  '118600160078
        Rec.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
            mMicroSectorID = Rec!intMicroSecID
        End If
        Rec.Close
        mSql = "SELECT intMicroSecID,chvEngMicroSector FROM M_MicroSector WHERE intMicroSecID= " & mMicroSectorID & ""
        RecChild.Open mSql, mCnPlan, adOpenStatic, adLockReadOnly
            If Not (RecChild.BOF And RecChild.EOF) Then
                txtMicroSector.Tag = RecChild!intMicroSecID
                txtMicroSector.Text = RecChild!chvEngMicroSector
                cmdMicroSector.Enabled = False
                txtMicroSector.Enabled = False
            End If
        RecChild.Close
    End If
    mCnPlan.Close
End Sub
Private Sub fillPVDetails(mVoucherID As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objdb     As New clsDB
    Dim mSql    As String
    
    If mVoucherID <> 0 Then
        frmPayVoucherDetails.Enabled = True
        If objdb.SetConnection(mCnn) Then
            mSql = " SELECT *  FROM faVouchers "
            mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faVouchers.intKeyID1"
            mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID"
            mSql = mSql + " Where intVoucherID = " & val(mVoucherID)
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
            If Not (Rec.BOF And Rec.EOF) Then
                txtPVTrType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                txtPVBankHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtPVAmt.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                txtPVDate.Text = DdMmmYy(IIf(IsNull(Rec!dtDate), 0, Rec!dtDate))
            End If
            Rec.Close
        End If
        'mCnn.Close
    End If
End Sub
Private Sub fillPODetails(mPayorderID As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objdb     As New clsDB
    Dim mSql    As String
    
    If mPayorderID <> 0 Then
        frmPODetails.Enabled = True
        If objdb.SetConnection(mCnn) Then
            mSql = " SELECT *  FROM faPayOrder "
            mSql = mSql + " INNER JOIN faPayOrderChild ON faPayOrderChild.intPayOrderID=faPayOrder.intPayOrderID"
            mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faPayOrder.intCashOrBankHeadID"
            mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID=faPayOrder.intTransactionTypeID"
            mSql = mSql + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faPayOrder.intSourceOfFundID"
            mSql = mSql + " Where faPayOrderChild.intSlNo=3    And faPayOrder.intPayOrderID = " & val(mPayorderID)
            Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
            If Not (Rec.BOF And Rec.EOF) Then
                txtPTrType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                txtPExpHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                txtPAmt.Text = IIf(IsNull(Rec!numAmount), 0, Rec!numAmount)
                txtPSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
            End If
            Rec.Close
        End If
        'mCnn.Close
    End If
End Sub
Private Sub GetRequisitionBeforedtOnlineDate()
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objdb     As New clsDB
    Dim mSql    As String

    If objdb.SetConnection(mCnn) Then
        mSql = " SELECT *  FROM suExpenditures "
        mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=suExpenditures.intBankID"
        mSql = mSql + " Where intAllotmentID = " & val(txtRequisitionNo.Tag)
        Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
            txtPreviousVoucher.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            txtPreviousVoucherDate.Text = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
            txtPreviousInstNo.Text = IIf(IsNull(Rec!vchInstrumentNo), 0, Rec!vchInstrumentNo)
            txtPreviousInstDate.Text = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
            txtPreviousBank.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            txtPreviousBank.Tag = IIf(IsNull(Rec!intBankID), "", Rec!intBankID)
        End If
        Rec.Close
    End If
    'mCnn.Close
    cmdPreviousDateVerify.Enabled = False
End Sub
Private Sub FillRequisitionBeforedtOnlineDate()
   Dim mCnn    As New ADODB.Connection
    Dim mCnnSulekha    As New ADODB.Connection
    Dim objdb   As New clsDB
    Dim mArrIn  As Variant
    Dim mArrInChild  As Variant
    Dim mArrOut  As Variant
    Dim mSql As String
    Dim arrInput As Variant
    Dim Rec   As New ADODB.Recordset
    Dim mCount As Integer
    Dim mVoucherID, mNewVoucherID As Long
    Dim mID As Integer
    Dim mRecCount As Integer
    Dim mVrZerCheck As Integer
    Dim mNegCheck As Integer
    
    If Trim(txtPreviousVoucher.Text) = "" Then
        MsgBox "Enter the Voucher Number", vbInformation, "Saankhya"
        Exit Sub
    End If
'    If Trim(txtPreviousVoucherDate.Text) = "" Then
'        MsgBox "Enter the Voucher Date", vbInformation, "Saankhya"
'        Exit Sub
'    End If
    
    If Trim(txtPreviousVoucherDate.Text) <> "" Then
        If CDate(txtPreviousVoucherDate.Text) > CDate(dtOnlinedate) Then
            MsgBox "VoucherDate should be less than Online Date ", vbInformation, "Saankhya"
            Exit Sub
        End If
    Else
        MsgBox "Enter the Voucher Date", vbInformation, "Saankhya"
        Exit Sub
    End If
    
    
    If Trim(txtPreviousInstNo.Text) = "" Then
        MsgBox "Enter the Instrument Number", vbInformation, "Saankhya"
        Exit Sub
    End If
    If Trim(txtPreviousInstDate.Text) = "" Then
        MsgBox "Enter the Instrument Date", vbInformation, "Saankhya"
        Exit Sub
    End If
    If Trim(txtPreviousBank.Text) = "" Then
        MsgBox "Select  the  Bank", vbInformation, "Saankhya"
        Exit Sub
    End If
    mRecCount = 0
    objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
    'mSql = "Select * from suExpenditures Where intVoucherID= " & IIf(txtPaymentVoucher.Tag = "", 0, txtPaymentVoucher.Tag) & "  And intAllotmentID= " & txtRequisitionNo.Tag
    mSql = "Select * from suExpenditures Where intAllotmentID= " & txtRequisitionNo.Tag
    
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        While Not Rec.EOF
            mRecCount = mRecCount + 1
        If (Rec!intVoucherID = 0) Then
           mVrZerCheck = 1
           mVoucherID = Rec!intVoucherID
        Else
            mVrZerCheck = 0
        End If
'        Else
        If (Rec!intVoucherID < 0) Then
            mNegCheck = 1
            mVoucherID = Rec!intVoucherID
'        Else
'            mID = 1
        End If
'    Else
'        mID = -1
        Rec.MoveNext
        Wend
    End If
    If mRecCount > 1 Then
        mID = 1
    ElseIf mRecCount = 1 Then
        If mVrZerCheck = 1 Then
            If mVoucherID = 0 Then
                mID = -1
            Else
                mID = 1
            End If
        End If
        If mNegCheck = 1 Then
            mID = 1
        End If
    Else
        mID = -1
    End If
    
    mArrIn = Array(mID, _
    gbFinancialYearID - 1, _
    gbLBID, _
    val(txtProjectNo.Tag), _
    mVoucherID, _
    val(txtPreviousVoucher.Text), _
    txtPreviousInstDate.Text, _
    val(txtPreviousInstNo.Text), _
    txtPreviousInstDate.Text, _
    val(txtPreviousBank.Tag), _
    val(txtAmount.Text), _
    val(txtSourceOfFund.Tag), _
    val(txtRequisitionNo.Tag) _
    )
    
    objdb.ExecuteSP "spUpdateExpenditureDetails", mArrIn, mArrOut, , mCnn, adCmdStoredProc
    mNewVoucherID = mArrOut(0, 0) 'IIf(txtPaymentVoucher.Tag = "", 0, txtPaymentVoucher.Tag)
    Rec.Close
    mSql = "Update faAllotments set  numProjectID =" & val(txtProjectNo.Tag) & ", vchProjectNo ='" & txtProjectNo.Text & "', intSourceID =" & txtSourceOfFund.Tag & ","
    mSql = mSql + " intFundCategoryID =" & txtCategory.Tag & ", intFunctionaryID =" & txtFunctionary.Tag & ", intFunctionID =" & txtFunction.Tag & ", intAccountHeadID =" & txtGrossExpenditureHead.Tag & ", vchAccountHeadCode ='" & txtExpAccHeadCode.Text & "' "
    mSql = mSql + " ,intSubSecID=" & txtSubSector.Tag & ",intMircoSectorID=" & txtMicroSector.Tag & " ,"
    mSql = mSql + " tnyProjectStatus=2 "
    mSql = mSql + " Where intID=" & txtRequisitionNo.Tag & "  "
    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
    
    If (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
        mSql = "Update RequisitionDetails set decProjectID= " & val(txtProjectNo.Tag) & " ,tnyTransfer=0 where intReqID = " & val(txtRequisitionNo.Tag) & "  "
        objdb.ExecuteSP mSql, , , , mCnnSulekha, adCmdText
        
        
        mSql = "Select * from ExpenseDetails where intVoucherID = " & mNewVoucherID & "  And decProjectID= " & val(txtProjectNo.Tag) & " And tnyCancelation is null "
        Rec.Open mSql, mCnnSulekha
        If Not (Rec.EOF And Rec.BOF) Then
            mCount = 1
        Else
            mCount = 0
        End If
        Rec.Close
        If mCount = 0 Then
'            mSQL = "Update ExpenseDetails set decProjectID= " & txtProjectNo.Tag & " where intVoucherID = " & txtPaymentVoucher.Tag & "  "
'            objDB.ExecuteSP mSQL, , , , mCnnSulekha, adCmdText
'        Else
            arrInput = Array(gbLBID, _
                                gbFinancialYearID - 1, _
                                val(txtProjectNo.Tag), _
                                -1, val(txtSourceOfFund.Tag), _
                                val(txtAmount), _
                                mNewVoucherID)

                objdb.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
        End If
        
    Else
        MsgBox "Connection to Sulekha Database doesnot exist", vbInformation, "Saankhya"
        Exit Sub
    End If
    cmdPreviousDateVerify.Enabled = False
End Sub
Private Sub GetDtOnlineDate()
    Dim mCnn As New ADODB.Connection
    Dim Rec     As New ADODB.Recordset
    Dim objdb     As New clsDB
    Dim mSql    As String
  
    If objdb.SetConnection(mCnn) Then
        mSql = " SELECT dtRPOpeningDate  FROM faConfig"
        Rec.Open mSql, mCnn, adOpenStatic, adLockReadOnly
        If Not (Rec.BOF And Rec.EOF) Then
            dtOnlinedate = IIf(IsNull(Rec!dtRPOpeningDate), 0, Rec!dtRPOpeningDate)
        End If
        Rec.Close
    End If
    'mCnn.Close
End Sub
Public Property Let ReqID(mData As Variant)
    mReqID = mData
End Property
Public Property Get ReqID() As Variant
    ReqID = mReqID
End Property

Private Sub txtSourceOfFund_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub
Private Sub txtTreasury_KeyPress(KeyAscii As Integer)
      If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub

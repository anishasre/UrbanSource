VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmIntegratedPayments 
   BackColor       =   &H00E9F8F8&
   Caption         =   "P a y m e n t s"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCreditHdIDR 
      Height          =   285
      Left            =   12480
      TabIndex        =   119
      Text            =   "txtCreditHdIDR"
      Top             =   3120
      Width           =   1875
   End
   Begin VB.TextBox txtWebExtractIDforP 
      Height          =   285
      Left            =   12480
      TabIndex        =   118
      Text            =   "numBillControCodeID"
      Top             =   3360
      Width           =   1875
   End
   Begin VB.TextBox txtBillControCodeID 
      Height          =   285
      Left            =   12480
      TabIndex        =   117
      Text            =   "numBillControCodeID"
      Top             =   3630
      Width           =   1875
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5250
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8970
      Width           =   1380
   End
   Begin VB.ListBox lstMasters 
      BackColor       =   &H00E7F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3180
      Left            =   14340
      TabIndex        =   63
      Top             =   3300
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.TextBox txtSubCashCode 
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
      Left            =   6180
      TabIndex        =   73
      Top             =   4110
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0EEEE&
      Height          =   1305
      Left            =   15
      TabIndex        =   68
      Top             =   975
      Width           =   11925
      Begin VB.TextBox txtBranch 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   9930
         TabIndex        =   18
         Top             =   840
         Width           =   1905
      End
      Begin VB.TextBox txtNameOfBank 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5730
         TabIndex        =   17
         Top             =   855
         Width           =   2685
      End
      Begin VB.TextBox txtAccountNo 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   855
         Width           =   2685
      End
      Begin VB.CommandButton cmdInstrument 
         Caption         =   "..."
         Height          =   315
         Left            =   4620
         TabIndex        =   13
         Top             =   510
         Width           =   270
      End
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5730
         TabIndex        =   14
         Top             =   510
         Width           =   2685
      End
      Begin VB.TextBox txtInstrument 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   510
         Width           =   2685
      End
      Begin VB.CommandButton cmdCrAccountHead 
         Caption         =   "..."
         Height          =   315
         Left            =   8430
         TabIndex        =   10
         Top             =   150
         Width           =   300
      End
      Begin VB.TextBox txtCrAccountHead 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   165
         Width           =   4905
      End
      Begin VB.TextBox txtCrHeadCode 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   165
         Width           =   1590
      End
      Begin VB.TextBox txtDated 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   9930
         TabIndex        =   15
         Top             =   495
         Width           =   1905
      End
      Begin VB.TextBox txtCrAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   9930
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   165
         Width           =   1905
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8850
         TabIndex        =   79
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Left            =   8850
         TabIndex        =   78
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Left            =   5265
         TabIndex        =   77
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Left            =   480
         TabIndex        =   76
         Top             =   885
         Width           =   1410
      End
      Begin VB.Label lblDrAccountHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Type"
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
         Left            =   450
         TabIndex        =   72
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cr (Acc.Head)"
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
         Left            =   675
         TabIndex        =   71
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Inst. No"
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
         Left            =   5025
         TabIndex        =   70
         Top             =   555
         Width           =   675
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Inst.Date"
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
         Left            =   8850
         TabIndex        =   69
         Top             =   555
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdPaymentOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   300
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   9000
      Width           =   270
   End
   Begin VB.TextBox txtPayOrder 
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
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   1410
      TabIndex        =   55
      Top             =   9015
      Width           =   1665
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4485
      Left            =   12000
      Picture         =   "frmIntegratedPayments.frx":0000
      ScaleHeight     =   4425
      ScaleWidth      =   2625
      TabIndex        =   65
      Top             =   3900
      Width           =   2685
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the Payment order No to make a Payment"
         Height          =   3525
         Left            =   150
         TabIndex        =   67
         Top             =   735
         Width           =   2235
      End
      Begin VB.Label lblHelpTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAYMENTS"
         Height          =   255
         Left            =   540
         TabIndex        =   66
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cance&L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8085
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8970
      Width           =   1380
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8970
      Width           =   1380
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   11865
      Top             =   8505
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0EEEE&
      Height          =   960
      Left            =   15
      TabIndex        =   62
      Top             =   30
      Width           =   11925
      Begin VB.CommandButton cmdSearchVoucher 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   180
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   7515
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   570
         Width           =   270
      End
      Begin VB.TextBox txtTransactionType 
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   6075
      End
      Begin VB.TextBox txtVoucherNo 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1410
         TabIndex        =   0
         Top             =   180
         Width           =   1665
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   165
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchFunction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   540
         Width           =   270
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   2760
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   555
         Width           =   2760
      End
      Begin VB.TextBox txtDate 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4245
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   1
         Top             =   165
         Width           =   1620
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   64
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "  Date"
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
         Left            =   3645
         TabIndex        =   58
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
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
         Left            =   105
         TabIndex        =   61
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
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
         Left            =   7785
         TabIndex        =   60
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
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
         Left            =   8055
         TabIndex        =   59
         Top             =   585
         Width           =   705
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1650
      Left            =   1500
      TabIndex        =   19
      Top             =   2310
      Width           =   9510
      _cx             =   16775
      _cy             =   2910
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   13559526
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14349042
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   3
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmIntegratedPayments.frx":030A
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
   Begin TabDlg.SSTab SSTab 
      Height          =   4500
      Left            =   30
      TabIndex        =   80
      Top             =   4410
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   7938
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Payment Order"
      TabPicture(0)   =   "frmIntegratedPayments.frx":03DB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "<Subledger>"
      TabPicture(1)   =   "frmIntegratedPayments.frx":03F7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "..."
      TabPicture(2)   =   "frmIntegratedPayments.frx":0413
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H00F4FCFC&
         BorderStyle     =   0  'None
         Height          =   4080
         Left            =   60
         TabIndex        =   82
         Top             =   60
         Width           =   11835
         Begin VB.TextBox txtScheme 
            Height          =   285
            Left            =   5685
            TabIndex        =   116
            Text            =   "Scheme ID"
            Top             =   2625
            Width           =   975
         End
         Begin VB.TextBox txtGo 
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
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   495
            Width           =   3420
         End
         Begin VB.CommandButton cmdGo 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   11250
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   480
            Width           =   285
         End
         Begin VB.TextBox txtNarration 
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
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   1575
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   3180
            Width           =   9630
         End
         Begin VB.CommandButton cmdAsset 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   300
            Left            =   12375
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   1665
            Width           =   345
         End
         Begin VB.CommandButton cmdSubsidiaryCash 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   11235
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   165
            Width           =   285
         End
         Begin VB.CommandButton cmdSourceOfFund 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   11235
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   780
            Width           =   300
         End
         Begin VB.CommandButton cmdSeat 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   11265
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   3705
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtSourceOfFund 
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
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   810
            Width           =   3420
         End
         Begin VB.TextBox txtSubsidiaryCash 
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
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   180
            Width           =   3420
         End
         Begin VB.TextBox txtSeat 
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
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   3720
            Width           =   1725
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F4FCFC&
            Caption         =   "Final Bill"
            Height          =   195
            Left            =   1605
            TabIndex        =   100
            Top             =   3720
            Width           =   930
         End
         Begin VB.Frame fraProject 
            BackColor       =   &H00F4FCFC&
            BorderStyle     =   0  'None
            Height          =   1950
            Left            =   5700
            TabIndex        =   93
            Top             =   1155
            Width           =   5910
            Begin VB.TextBox txtCategory 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1380
               Width           =   3405
            End
            Begin VB.TextBox txtSector 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   1680
               Width           =   3405
            End
            Begin VB.TextBox txtProjectNo 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1050
               Width           =   3405
            End
            Begin VB.TextBox txtAgreementNo 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   750
               Width           =   3405
            End
            Begin VB.TextBox txtAllotmentLetterNo 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   450
               Width           =   3405
            End
            Begin VB.CommandButton cmdAgreementNo 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   735
               Width           =   300
            End
            Begin VB.CommandButton cmdAllotmentLetterNo 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   300
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   420
               Width           =   300
            End
            Begin VB.CommandButton cmdProjectNo 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   1050
               Width           =   300
            End
            Begin VB.TextBox txtImplementingOfficer 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   2100
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   150
               Width           =   3420
            End
            Begin VB.CommandButton cmdImplementingOfficer 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5550
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   120
               Width           =   300
            End
            Begin VB.Label lblAgreementNo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Agreement No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   690
               TabIndex        =   99
               Top             =   780
               Width           =   1395
            End
            Begin VB.Label lblAllotmentLetterNo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Allotment Letter No"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   98
               Top             =   480
               Width           =   1905
            End
            Begin VB.Label lblCategory 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Category"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1185
               TabIndex        =   97
               Top             =   1410
               Width           =   885
            End
            Begin VB.Label lblSector 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Sector"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1440
               TabIndex        =   96
               Top             =   1740
               Width           =   630
            End
            Begin VB.Label lblProjectNo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Project Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   540
               TabIndex        =   95
               Top             =   1080
               Width           =   1530
            End
            Begin VB.Label lblImplementingOfficer 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Implementing Officer"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   -30
               TabIndex        =   94
               Top             =   165
               Width           =   2085
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00F4FCFC&
            BorderStyle     =   0  'None
            Height          =   2730
            Left            =   90
            TabIndex        =   83
            Top             =   360
            Width           =   5430
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   30
               TabIndex        =   34
               Top             =   2445
               Width           =   2220
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   4110
               MaxLength       =   6
               TabIndex        =   33
               Top             =   2130
               Width           =   915
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   50
               TabIndex        =   32
               Top             =   2130
               Width           =   2220
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   4395
               MaxLength       =   1
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   555
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   4065
               MaxLength       =   1
               TabIndex        =   24
               Top             =   555
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   3735
               MaxLength       =   1
               TabIndex        =   23
               Top             =   555
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   31
               Top             =   1815
               Width           =   3525
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   30
               Top             =   1500
               Width           =   3525
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   29
               Top             =   1185
               Width           =   3525
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   28
               Top             =   870
               Width           =   3525
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   22
               Top             =   555
               Width           =   2190
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
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   4725
               MaxLength       =   1
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   555
               Width           =   315
            End
            Begin VB.TextBox txtSubLedgerType 
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
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   20
               Top             =   60
               Width           =   3540
            End
            Begin VB.CommandButton cmdSubLederType 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   300
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   30
               Width           =   345
            End
            Begin VB.CommandButton cmdSearchName 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   300
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   540
               Width           =   345
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Phone No"
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
               Left            =   645
               TabIndex        =   92
               Top             =   2475
               Width           =   810
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Pin"
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
               Left            =   3810
               TabIndex        =   91
               Top             =   2160
               Width           =   255
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Local Place"
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
               Left            =   525
               TabIndex        =   90
               Top             =   1530
               Width           =   945
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Main Place"
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
               Left            =   555
               TabIndex        =   89
               Top             =   1845
               Width           =   900
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Post"
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
               Left            =   1065
               TabIndex        =   88
               Top             =   2160
               Width           =   360
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Payee"
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
               Left            =   150
               TabIndex        =   87
               Top             =   570
               Width           =   1305
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "House/Office"
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
               Left            =   360
               TabIndex        =   86
               Top             =   885
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Street"
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
               Left            =   900
               TabIndex        =   85
               Top             =   1215
               Width           =   525
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "SubLedger Type"
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
               Left            =   0
               TabIndex        =   84
               Top             =   120
               Width           =   1395
            End
         End
         Begin VB.TextBox txtPayeeType 
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
            Height          =   270
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   405
            Width           =   3420
         End
         Begin VB.TextBox txtPayee 
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
            Height          =   270
            Left            =   1665
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   945
            Width           =   3420
         End
         Begin VB.Label lblGo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "GO No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7140
            TabIndex        =   115
            Top             =   510
            Width           =   585
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Narration"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   585
            TabIndex        =   109
            Top             =   3210
            Width           =   930
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Seat"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9000
            TabIndex        =   108
            Top             =   3750
            Width           =   435
         End
         Begin VB.Label lblSubsidiaryCash 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Subsidiary Cash"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6180
            TabIndex        =   107
            Top             =   195
            Width           =   1560
         End
         Begin VB.Label lblSourceOfFund 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Source of Fund"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6270
            TabIndex        =   106
            Top             =   825
            Width           =   1470
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1050
            TabIndex        =   105
            Top             =   705
            Width           =   480
         End
         Begin VB.Label lblNameOfPayee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Payee"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   60
            TabIndex        =   104
            Top             =   420
            Width           =   1470
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4FCFC&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   3315
         Left            =   -74940
         TabIndex        =   81
         Top             =   60
         Width           =   11850
      End
   End
   Begin VB.Label lblLedgerBalance 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   112
      Top             =   4050
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8775
      TabIndex        =   110
      Top             =   3990
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payable To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   75
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SubsidiaryCashBook Hidden"
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
      Left            =   3690
      TabIndex        =   74
      Top             =   4170
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00592525&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay &Order No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   75
      TabIndex        =   57
      Top             =   9060
      Width           =   1290
   End
End
Attribute VB_Name = "frmIntegratedPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Dim mSelect As Boolean
    Dim mBudgetBalanceAmt As Variant
    Dim mvarGrossSalryID As Variant
    Dim intPayOrderID As Variant
    Dim intLoadMode As Integer
    Dim mViewPayOrderListFormIsLoaded As Boolean
    Dim mSaveFlag As Boolean
    Dim mHelpTips As String
                                          
    Dim intVoucherID As Variant
    Dim intVoucherNo As Variant
    Dim mWaterBillPVMode As Boolean
    Dim mPayOrderNo As Variant
                                          '
    Dim mPaymentOrderBasedMode As Boolean ' To Decide Whether this Payment
                                          ' Detail is Fetch from Payment Order
    Dim mOfficialUserID As Variant
    Dim mOfficialSeatID As Variant
    Dim mintIntID As Variant
    Dim mintTransferID As Variant
    Public mPreYearMode    As Integer  ' set 1 for Pay order with moduleId 96
    Dim mViewMode       As Integer
    Dim mNewACRMode As Integer
    Dim mUnAuthorized As Integer
    Public mWebExtract  As Boolean
    
    Private Function CheckOfficial() As Boolean
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
'        Dim mEmpID      As Integer
        Dim mUserID     As Variant
        
        On Error GoTo err
        CheckOfficial = False
        mOfficialUserID = ""
        mOfficialSeatID = ""
'        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
'            mSQL = "Select numEmpID From faSubsidiaryAccountHeads Where intSubsidiaryAccountHeadID = " & txtName.Tag
'            Rec.Open mSQL, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mEmpID = IIf(IsNull(Rec!numEmpID), 0, Rec!numEmpID)
'            End If
'            Rec.Close
'            mCnn.Close
'        End If
'        If mEmpID <> 0 Then
        If objdb.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
            mSql = "Select GM_User.numUserID[UserID],GL_UserSeats.numSeatID[SeatID] From GM_User"
            mSql = mSql + " Inner Join GL_UserSeats On GM_User.numUserID = GL_UserSeats.numUserID"
            mSql = mSql + " Where numEmpID = " & txtName.Tag
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mOfficialUserID = IIf(IsNull(Rec!UserID), "", Rec!UserID)
                mOfficialSeatID = IIf(IsNull(Rec!SeatID), "", Rec!SeatID)
                CheckOfficial = True
            End If
            Rec.Close
            mCnn.Close
        End If
'        End If
        Exit Function
err:
        MsgBox err.Description
    End Function
    
    Private Sub ShowBudgetBalance(mAcHeadID As Long)
        Dim objAc As New clsAccounts
        Dim objBudjet As New clsBudgetCentre
        Dim mFunctionaryID As Variant
        Dim mFunctionId As Variant
        Dim mBudgetAmt As Variant
        Dim mUtilizedAmt As Variant
        
        mBudgetBalanceAmt = 0
        If val(txtFunctionary.Tag) > 0 Then
            mFunctionaryID = txtFunctionary.Tag
        End If
        If val(txtFunction.Tag) > 0 Then
            mFunctionId = txtFunctionary.Tag
        End If
        mBudgetAmt = objBudjet.GetBudgetAmount(mFunctionId, mFunctionaryID, mAcHeadID)
        'lblBudget.Caption = "Budget Amount :" & Format(mBudgetAmt, "0.00")
        mUtilizedAmt = objAc.GetLedgerBalance(mAcHeadID, , mFunctionaryID, mFunctionId)
        'lblBudget.Caption = lblBudget.Caption & "/   Budget Utilized : " & Format(mUtilizedAmt, "0.00")
        mBudgetBalanceAmt = Format(mBudgetAmt, "0.00") - Format(mUtilizedAmt, "0.00")
    End Sub
    
    Private Sub InitializeSelectedHeads()
        vsGrid.Clear 1, 1
        cmdCrAccountHead.Enabled = True
    End Sub

    Private Sub FormInitialize()
        Dim ctrl As Control
        Dim mCnn As New ADODB.Connection
        Dim Rec As New Recordset
        Dim objAc As New clsAccounts
        
        mSaveFlag = False
        mPaymentOrderBasedMode = False
        '-----------------------------------------------------------'
        'Enagling FRAM for the Particulars Fields
        '-----------------------------------------------------------'
        Frame4.Enabled = True
        txtTransactionType.Enabled = True
        cmdSearchTransactionType.Enabled = True
        cmdSearchFunction.Enabled = True
        cmdSearchFunctionary.Enabled = True
        cmdPaymentOrder.Enabled = True
        
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
            ElseIf TypeOf ctrl Is OptionButton Then
                ctrl.Value = False
            ElseIf TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                ctrl.Tag = ""
            End If
        Next
        txtDate.Text = DdMmmYy(gbTransactionDate)
        txtDate.Tag = ""
        txtDated.Text = ""
        txtDated.Enabled = True
        'Note:- User Type wise Functionality should enable or Disabled

        vsGrid.Clear 1, 1
        txtDated.Text = ""
        cmdCrAccountHead.Enabled = True
        mSelect = False
        mBudgetBalanceAmt = 0
        
        
        If gbSectionID <> gbJSKSectionID Then
            txtInstrument.Tag = 5
            txtInstrument.Text = "Cheque"
            If gbDefaultBankID = Null Then
                objAc.SetAccountCode ("450210100")
                If objAc.AccountHeadID > -1 Then
                    txtCrHeadCode.Text = objAc.AccountCode
                    txtCrHeadCode.Tag = objAc.AccountHeadID
                End If
            Else
                objAc.SetAccountID Trim(val(gbDefaultBankID))                       'Added on 13.05.2010
                txtCrHeadCode.Text = objAc.AccountCode
                txtCrHeadCode.Tag = objAc.AccountHeadID
            End If
            Call txtCrHeadCode_LostFocus
            txtCrHeadCode.Tag = gbDefaultBankID
        Else
            Call LockingClearBanks(True)
            txtInstrument.Tag = 1
            txtInstrument.Text = "Cash"
            objAc.SetAccountCode (gbAcHeadCodeCash)
            If objAc.AccountHeadID > -1 Then
               txtCrHeadCode.Text = objAc.AccountCode
               txtCrHeadCode.Tag = objAc.AccountHeadID
               Call txtCrHeadCode_LostFocus
            End If
        End If
        txtSubsidiaryCash.Enabled = False
        cmdSubsidiaryCash.Enabled = False
        
        'cmdSave.Enabled = True
        txtSeat.Text = gbSeatName
        txtSeat.Tag = gbSeatID
        txtVoucherNo.Tag = -1
        lblTotal.Caption = ""
        LoadMode = 0
        cmdSave.Caption = "&Save"
        mPreYearMode = 0
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Enabled = False
        End If
    End Sub
    Private Function CalculateAmt() As Variant
        Dim mCount As Integer
        Dim mTot As Variant
        mTot = 0
        
        For mCount = 1 To vsGrid.Rows
            If Trim(vsGrid.TextMatrix(mCount, 3)) = "" Then Exit For
            mTot = val(mTot) + val(vsGrid.TextMatrix(mCount, 3))
        Next
        CalculateAmt = mTot
    End Function
    Private Sub FillGridCombo()
        Dim objdb As New clsDB
        Dim RecAccHead As New ADODB.Recordset
        Dim mItem As String

        RecAccHead.CursorLocation = adUseClient
        Set RecAccHead = GetRecordSet("spGetAccHead4Payments", adOpenStatic, adLockReadOnly)
        While Not RecAccHead.EOF
            mItem = mItem + "|" + RecAccHead!vchAccountHead
            RecAccHead.MoveNext
        Wend
        RecAccHead.Close
        vsGrid.ColComboList(2) = mItem
    End Sub

    Private Sub cmdAgreementNo_Click()              'ADDED BY MINU FOR AGREEMENTS ON 25-05-2011
        frmSearchAgreements.Show vbModal
        If gbSearchID <> -1 Then
            txtAgreementNo.Text = gbSearchStr
            txtAgreementNo.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdAllotmentLetterNo_Click()
        frmListOfAllotmentLetters.Visible = True
        frmListOfAllotmentLetters.ZOrder (0)
    End Sub

    Private Sub cmdCancel_Click()
        ''If mViewPayOrderListFormIsLoaded Then
        ''    Unload Me
        ''Else
        ''    Call FormInitialize
        ''End If
        Unload Me
    End Sub
    Private Sub cmdHeadSearch_Click()
        Dim mSql As String
        Dim objAcc As New clsAccounts
        
        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faCash
        frmSearchAccountHeads.SQLString = mSql
        frmSearchAccountHeads.Show vbModal
    End Sub
    Private Sub TrTypeUnUtilizedAmountInFunds()
        Dim mSql As String
        If val(txtSourceofFund.Tag) > 0 Then
            If val(txtSourceofFund.Tag) = 1 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount2
            ElseIf val(txtSourceofFund.Tag) = 25 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount4
            ElseIf val(txtSourceofFund.Tag) = 26 Or val(txtSourceofFund.Tag) = 41 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount5
            ElseIf val(txtSourceofFund.Tag) = 27 Then
                 mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount2
            ElseIf val(txtSourceofFund.Tag) = 28 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount2
            ElseIf val(txtSourceofFund.Tag) = 29 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount6
            ElseIf val(txtSourceofFund.Tag) = 30 Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE intAccountHeadID=" & gbAcHeadIDTreasuryAccount7
            End If
            txtInstrument.Tag = gbInstrumentCheque
            txtInstrument.Text = "Cheque"
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.chkListAll.Enabled = False
            frmSearchAccountHeads.VoucherMode = 200
            frmSearchAccountHeads.Show vbModal
            If gbSearchID <> -1 Then
                txtCrHeadCode.Text = Left(gbSearchStr, 9)
                txtCrHeadCode.Tag = gbSearchID
                txtCrAccountHead.Text = gbSearchStr
                txtCrHeadCode.SetFocus
                txtCrAccountHead.SetFocus
                If txtInstrument.Tag <> gbInstrumentCash Then
                    If gbLocalBodyID <> 167 Then
                        lblLedgerBalance.Visible = True
                        lblLedgerBalance.Caption = "Bank Balance: " + Format(LedgerBalance(val(txtCrHeadCode.Tag)), "0.00")
                    End If
                Else
                    lblLedgerBalance.Visible = False
                End If
                gbSearchID = -1
                gbSearchStr = ""
            End If
        End If
    End Sub
    Private Function mCheckNewACRMode(mPayOrderNo As Long)
        
        Dim Rec As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
    
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = " SELECT * FROM faPAyOrder WHERE vchPayOrderNo=" & mPayOrderNo
        
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mCheckNewACRMode = IIf(IsNull(Rec!intKeyID), 0, Rec!intKeyID)
            'Exit Function
        End If
        Rec.Close
        mCnn.Close
    End Function
    Private Sub cmdCrAccountHead_Click()
                
        '*******************VALIDATIONS FOR RECEIPTS FROM OTHER LSGI's*****************************

        Dim RecDate As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim objBank As New clsBank
        Dim objAc   As New clsAccounts
        objdb.SetConnection mCnn
        
        If val(txtTransactionType.Tag) = gbTransactionTypePayBills Then
        
        Else
            If mCheckNewACRMode(val(txtPayOrder.Text)) = 1 Then
                '''  Modified for Joint Venture 10, 11, 12, 13, 14
                If txtSourceofFund.Tag = 10 Or txtSourceofFund.Tag = 11 Or txtSourceofFund.Tag = 12 Or val(txtSourceofFund.Tag) = 13 _
                Or val(txtSourceofFund.Tag) = 14 Or val(txtSourceofFund.Tag) = 2 Then
                
                Else
                
                    MsgBox "Instrument Type cannot be edited for this Payment", vbInformation
                Exit Sub
                End If
            End If
        End If
        If mPreYearMode <> 1 Then
            If val(txtSourceofFund.Tag) = 1 Or val(txtSourceofFund.Tag) = 29 Or val(txtSourceofFund.Tag) = 30 _
            Or val(txtSourceofFund.Tag) = 10 Or val(txtSourceofFund.Tag) = 11 _
            Or val(txtSourceofFund.Tag) = 12 Or val(txtSourceofFund.Tag) = 13 _
            Or val(txtSourceofFund.Tag) = 14 Then
                If mUnAuthorized <> 3 Then
                    ' MODIFIED by AIBY 24th Aug, 2014 Allowing IAY Projects to Change their Banks.
                    ' Project Schemes will be identified as IAY General. Scheme Ids are taken from Saankhya.
                    If val(txtScheme) <> 41 And _
                        val(txtScheme) <> 52 And _
                        val(txtScheme) <> 54 Then
                        
                        RecDate.Open "Select *,getDate() as CurDate From faLBSettings", mCnn, adOpenDynamic, adLockOptimistic
                        If Not (RecDate.EOF And RecDate.BOF) Then
                            If CDate(RecDate!CurDate) < gbBankChangePermitDate Then
                               mSql = "You are going to Edit the Default treasury For this Source Of Fund"
                               MsgBox mSql, vbInformation
                            ElseIf gbLocalBodyID = 1286 And gbTransactionDate < CDate("16/Apr/2019") Then
                                mSql = "Editing of Default treasury For this Source Of Fund is limited to 15/apr/2019"
                                MsgBox mSql, vbInformation
                            Else
                                mSql = "The Default treasury Can not be Edited For this Source Of Fund"
                                MsgBox mSql, vbInformation
                                cmdCrAccountHead.Enabled = False
                                Exit Sub
                            End If
                        End If
                        RecDate.Close
                    End If
            
                End If
            ElseIf val(txtSourceofFund.Tag) = 4 Then  ''' added on 15 mar 17 for own fund project payment
                If val(txtTransactionType.Tag) = 1141 Or val(txtTransactionType.Tag) = 1151 Or val(txtTransactionType.Tag) = 1161 Or _
                        val(txtTransactionType.Tag) = 1171 Or val(txtTransactionType.Tag) = 1181 Or val(txtTransactionType.Tag) = 1191 Then
                            
                        If objBank.BankID > 0 Then
                            objAc.SetAccountCode gbAcHeadCodeTreasuryAccountTSB
                        Else
                            objAc.SetAccountID 0
                        End If
                    Else
                         objAc.SetAccountID 1504
                    End If
            
            End If
        End If
        
        '**********************************************************************
          
        If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
            Call TrTypeUnUtilizedAmountInFunds
        Else
            If val(txtInstrument.Tag) = 0 Then
                MsgBox "Please Select the Instrument Type", vbInformation
                cmdInstrument.SetFocus
                Exit Sub
            End If
            'Dim msql As String
            If val(txtInstrument.Tag) > 0 Then
                Select Case val(txtInstrument.Tag)
                    Case 1 '[Cash]
                     mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & faCash
    '                Case 7 '[Treasury Bills]
    '                 mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faBank
                    Case Else
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & faBank
                                    
                        mSql = " Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
                        mSql = mSql + " INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID"
                        mSql = mSql + " WHERE  tinHiddenFlag = 0 And faAccountHeads.intGroupID = " & faBank & " Order By vchAccountHeadCode"
                            
                End Select
                frmSearchAccountHeads.SQLString = mSql
                frmSearchAccountHeads.chkListAll.Enabled = False
                frmSearchAccountHeads.VoucherMode = 200
                
                frmSearchAccountHeads.Show vbModal
                'txtAccountCode.SetFocus
                If gbSearchID <> -1 Then
                    txtCrHeadCode.Text = Left(gbSearchStr, 9)
                    txtCrHeadCode.Tag = gbSearchID
                    txtCrAccountHead.Text = gbSearchStr
                    txtCrHeadCode.SetFocus
                    txtCrAccountHead.SetFocus
                    If txtInstrument.Tag <> gbInstrumentCash Then
                        If gbLocalBodyID <> 167 Then
                            lblLedgerBalance.Visible = True
                            lblLedgerBalance.Caption = "Bank Balance: " + Format(LedgerBalance(val(txtCrHeadCode.Tag)), "0.00")
                        End If
                    Else
                        lblLedgerBalance.Visible = False
                    End If
                    gbSearchID = -1
                    gbSearchStr = ""
                End If
            End If
        End If
    End Sub
    
    Private Function LedgerBalance(mAcID As Integer) As Variant
        Dim mSql    As String
        Dim Rec     As New ADODB.Recordset
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mArIn   As Variant
        Dim mArOut  As Variant
        Dim mLegBalance As Variant
        
        mArIn = Array(mAcID, gbTransactionDate)
        If gbLocalBodyID = 167 Then
            LedgerBalance = 0
        Else
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                objdb.ExecuteSP "spGetLedgerBalance", mArIn, mArOut, , mCnn, adCmdStoredProc
                If IsNumeric(mArOut(2, 0)) Then
                    mLegBalance = mArOut(2, 0)
                    LedgerBalance = mLegBalance
                Else
                    LedgerBalance = 0
                End If
    
            End If
        End If
    End Function
    
    Private Sub cmdImplementingOfficer_Click()
        frmSearchMasters.SQLQry = "Select intFunctionaryID, vchFunctionary +'[' + vchFunctionaryCode + ']' From faFunctionaries Where intFunctionaryID > 13"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        txtImplementingOfficer.SetFocus
    End Sub
    
    Private Sub cmdInstrument_Click()
        Dim mSql As String
        ''----------------Modified By Anisha On 31.10.2011
        mSql = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes Where intInstrumentTypeID not in (6,8,9)"
        ''-------------------------------------------------
        Call PopulateList(lstMasters, mSql, "", , , True)
        lstMasters.Left = 5040
        lstMasters.Top = 3300
        lstMasters.Height = 2000
        lstMasters.Width = txtInstrument.Width
        lstMasters.Visible = True
        lstMasters.SetFocus
    End Sub
    
    Private Sub cmdNew_Click()
        Call FormInitialize
    End Sub
    
    Private Sub cmdPaymentOrder_Click()
        frmSearchPaymentOrder.Mode = 50
        frmSearchPaymentOrder.Staus = 1
        frmSearchPaymentOrder.chkListToApprove.Value = 1
        frmSearchPaymentOrder.Show vbModal
        If gbSearchID > 0 Then
            txtPayOrder.Tag = gbSearchID
            txtPayOrder.Text = gbSearchStr
            gbSearchID = -1
            gbSearchStr = ""
            Call txtPayOrder_LostFocus
        End If
    End Sub

    Private Sub cmdProjectNo_Click()
        frmSulekhaIntegration.Show vbModal
        txtProjectNo.SetFocus
    End Sub
    
    Private Sub cmdSave_Click()
        'On Error GoTo Err:
        If LoadMode = 2 Then
            Dim objdb As New clsDB
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim mSql As String
            mSql = "Select tnyStatus From faPayOrder Where vchPayOrderNo = " & val(txtPayOrder) & " And tnyStatus = 1"
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            If Not (Rec.BOF And Rec.EOF) Then
                If Rec!tnyStatus = 0 Then
                    MsgBox "Please Approve the Payment Order before making the Payment", vbInformation
                    Exit Sub
                Else
                    If SaveValidation Then
                        cmdSave.Enabled = False
                        Call Saving
                        If mSaveFlag Then
                            'MsgBox "Payment Voucher Generated Successfully", vbInformation
                        End If
                    End If
                End If
            End If
        Else
            If SaveValidation Then
                cmdSave.Enabled = False
                Call Saving
                If mSaveFlag Then
                    MsgBox "Payment Voucher Generated Successfully", vbInformation
                End If
            End If
        End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdSearchFunction_Click()
        frmSearchFunction.Show vbModal
        If gbSearchID <> -1 Then
            txtFunction.Text = gbSearchStr
            txtFunction.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdSearchFunctionary_Click()
        frmSearchFunctionary.Show vbModal
        If gbSearchID <> -1 Then
            txtFunctionary.Text = gbSearchStr
            txtFunctionary.Tag = gbSearchID
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub
    
    Private Sub cmdSearchName_Click()
        Dim objSubLedger As New clsSubLedger
        frmSearchSubsidiaryAccountHeads.SubLedgerType = val(txtSubLedgerType.Tag)
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchStr <> "" Then
            txtName.Tag = gbSearchID
            objSubLedger.SetSubLedgerDetails (val(txtName.Tag))
            If Not IsNull(objSubLedger.NameOfSubLedger) Then
                txtName.Text = IIf(IsNull(objSubLedger.NameOfSubLedger), "", objSubLedger.NameOfSubLedger)
                txtHouse.Text = IIf(IsNull(objSubLedger.HouseOrOffice), "", objSubLedger.HouseOrOffice)
                txtStreet.Text = IIf(IsNull(objSubLedger.Street), "", objSubLedger.Street)
                txtLocalPlace.Text = IIf(IsNull(objSubLedger.LocalPlace), "", objSubLedger.LocalPlace)
                txtMainPlace.Text = IIf(IsNull(objSubLedger.MainPlace), "", objSubLedger.MainPlace)
                txtPost.Text = IIf(IsNull(objSubLedger.PostOffice), "", objSubLedger.PostOffice)
                txtPin.Text = IIf(IsNull(objSubLedger.PinCode), "", objSubLedger.PinCode)
                txtPhone.Text = IIf(IsNull(objSubLedger.Phone), "", objSubLedger.Phone)
            Else
                txtName.Text = gbSearchStr
            End If
        End If
        txtName.SetFocus
        gbSearchID = -1
        gbSearchStr = ""
        
        If val(txtSubLedgerType.Tag) = 10 Then
            txtPayee.Text = txtName.Text
        End If
    End Sub
    
    Private Sub cmdSearchSeat_Click()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        lstMasters.Tag = 4
        mSql = "Select chvSeatTitle , numSeatID, intID From GL_Seats Where intLocalBodyID = " & gbLocalBodyID
        If objdb.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
            Rec.Open mSql, mCnn
            While Not (Rec.EOF Or Rec.BOF)
                lstMasters.AddItem Rec!chvSeatTitle
                lstMasters.ItemData(lstMasters.NewIndex) = Rec!intID
                
                
                Rec.MoveNext
            Wend
        Else
            MsgBox "Connection To Global Master Doesnot Establish, Contact Your System Administrator", vbInformation
            Exit Sub
        End If
        lstMasters.Visible = True
    End Sub

    Private Sub cmdSearchTransactionType_Click()
        gbSearchStr = ""
        gbSearchID = -1
        gbSearchCode = ""
        
        frmSearchTransactionType.ModeOfTransaction = 2
        If mWebExtract = True Then
             frmSearchTransactionType.SQLQry = "Select  vchTransactionType, intTransactionTypeID From faTransactiontype  Where intGroupID=20 and intTransactionTypeId in (1141,1151,1161,1171,1181,1191)"
        End If
        frmSearchTransactionType.Show vbModal
        If gbSearchID < 1 Then
            Call InitializeSelectedHeads
        End If
        txtTransactionType.Text = gbSearchStr
        txtTransactionType.Tag = gbSearchID
        
        'Note:- Project and Non Project Payment Orders
        If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
            fraProject.Visible = True
            Check1.Visible = True
        Else
            fraProject.Visible = False
            Check1.Visible = False
        End If
        
        'Call ShowDetailsForSubCashBook
        
        Dim objTrns As New clsTransactionType
        objTrns.SetSourceOfFund (txtTransactionType.Tag)
        If Not IsEmpty(objTrns.SourceFundID) Then
            txtSourceofFund.Text = objTrns.SourceOfFund
            txtSourceofFund.Tag = objTrns.SourceFundID
        Else
            txtSourceofFund.Text = "Own Fund"
            txtSourceofFund.Tag = 4
        End If
        
        gbSearchStr = ""
        gbSearchID = -1
        gbSearchCode = ""
    End Sub
    
    Private Sub cmdSearchVoucher_Click()
        frmSearchVouchers.chkContra.Visible = False
        frmSearchVouchers.chkReceipt.Visible = False
        frmSearchVouchers.chkJournal.Visible = False
        frmSearchVouchers.chkPayment.Value = 1
        frmSearchVouchers.Show vbModal
        If gbSearchID <> -1 Then
            txtVoucherNo.Text = gbSearchCode
            txtVoucherNo.Tag = gbSearchID
            
            gbSearchCode = ""
            gbSearchID = -1
            
            Call txtVoucherNo_LostFocus
        End If
    End Sub
    
    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
    End Sub

    Private Sub cmdSourceOfFund_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        txtSourceofFund.SetFocus
    End Sub

    Private Sub cmdSubLederType_Click()
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            txtSubLedgerType.Text = gbSearchStr
            txtSubLedgerType.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub

    Private Sub cmdSubsidiaryCash_Click()
        '================================================================================================'
        '               Added By Cijith On 20/11/2009   For Integrating Subsidiary Cash Book             '
        '================================================================================================'
        On Error GoTo err:
            frmSearchSubsidiaryAccountHeads.SubLedgerType = 12
            frmSearchSubsidiaryAccountHeads.Show vbModal
            If gbSearchStr <> "" Then
                txtSubsidiaryCash.Text = gbSearchStr
                txtSubsidiaryCash.Tag = gbSearchID
            End If
            gbSearchID = -1
            gbSearchStr = ""
            
            '''If Val(txtSubsidiaryCash.Tag) <> 0 Then
            '''    txtPayeeType.Text = txtSubsidiaryCash.Text
            '''    Call ShowDetailsForSubCashBook
            '''End If
            txtSubCashCode.Text = vsGrid.TextMatrix(1, 1)
            txtSubCashCode.Tag = vsGrid.TextMatrix(1, 4)
            
            Dim objAc As New clsAccounts
            
            objAc.SetAccountID (1550)
            
            vsGrid.Cell(flexcpText, 1, 0) = 1
            vsGrid.Cell(flexcpText, 1, 1) = objAc.AccountCode
            vsGrid.Cell(flexcpText, 1, 2) = objAc.AccountHead
            'vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
            vsGrid.Cell(flexcpText, 1, 4) = objAc.AccountHeadID
        Exit Sub
err:
        MsgBox (Error$)
        '================================================================================================'
    End Sub

    Private Sub Form_Activate()
        Dim intcnt As Integer
        XPC.InitSubClassing
    End Sub
    
    Private Sub Form_Click()
        lstMasters.Visible = False
    End Sub

    Private Sub Form_Load()
        Call FormInitialize
        vsGrid.ColComboList(1) = "|..."
        
        '''Note:- Filling Combo Transaction Type and Instrument Type
        
        
        '''If intLoadMode = 1 Then           'Normal Save Mode    '
        '''    cmdSave.Enabled = True
        '''ElseIf intLoadMode = 2 Then       'Approving Stage     '
        '''    cmdSave.Enabled = False
        '''End If
        '''LoadMode = 1
        If PayOrderNo <> "" Then
            txtPayOrder.Text = PayOrderNo
            DisplayPayOrder (txtPayOrder.Text)
        End If
        SSTab.TabVisible(1) = False
        SSTab.TabVisible(2) = False
        
    End Sub
        
    Private Sub Form_Resize()
        Me.WindowState = 2
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
'        txtPayOrder.Text = ""
'        intPayOrderID = ""
'        If mViewPayOrderListFormIsLoaded Then
'            mViewPayOrderListFormIsLoaded = False
'            frmViewPaymentOrder.Visible = True
'        End If
    
    End Sub

    Private Sub lstMasters_DblClick()
        '-----------------------------------------------------------------'
        '               Added On 07/04/2009 By Cijith Sreedharan          '
        '-----------------------------------------------------------------'
        '''If lstMasters.Tag = 1 Then
        '''    txtFunctionary.Text = lstMasters.Text
        '''    txtFunctionary.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        '''ElseIf lstMasters.Tag = 2 Then
        '''    txtFunction.Text = lstMasters.Text
        '''    txtFunction.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        '''ElseIf lstMasters.Tag = 3 Then
        '''    txtSubLedgerType.Text = lstMasters.Text
        '''    txtSubLedgerType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        '''ElseIf lstMasters.Tag = 4 Then
        '''
        '''End If
        
        
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim mDate As Date

    
        
        
        If gbSectionID = gbJSKSectionID Then
            If lstMasters.ItemData(lstMasters.ListIndex) <> gbInstrumentCash Then
                lstMasters.Visible = False
                MsgBox "This Section Allows Cash Payments Only", vbInformation
                Exit Sub
            End If
        ElseIf gbSectionID <> gbJSKSectionID Then
        ''-----------------------------------------------------------------------------
        '''Validation skipped on 29.8.2011  to Allow Cash Payment in Accounts Section
        '' Decision Made on 28.8.2011 at Kila review Meeting (c bulb Team) By SPO,UBK ...
        ''-----------------------------------------------------------------------------
'           'If (gbLBPanchayat = 1 Or gbLBType = 4) And DateValue(gbOnlinedate) > gbTransactionDate And gbSeatGroupID = gbSeatGroupAccountsClerk Then
'           If DateValue(gbOnlinedate) > gbTransactionDate And gbSeatGroupID = gbSeatGroupAccountsClerk Then
'                    'lstMasters.Visible = False
'           Else
'                If lstMasters.ItemData(lstMasters.ListIndex) = gbInstrumentCash Then
'                    lstMasters.Visible = False
'                    MsgBox "This Section Does not Allows Cash Payments", vbInformation
'                    Exit Sub
'                End If
'
'          End If
        End If
        If lstMasters.ItemData(lstMasters.ListIndex) <> -1 Then
            txtInstrument.Text = lstMasters.Text
            txtInstrument.Tag = lstMasters.ItemData(lstMasters.ListIndex)
            lstMasters.Visible = False
            
'            txtCrHeadCode.Text = ""
'            txtCrHeadCode.Tag = ""
'            txtCrAccountHead.Text = ""
        End If
        Call txtInstrument_LostFocus
    End Sub

    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub lstMasters_LostFocus()
        lstMasters.Visible = False
    End Sub

    Private Sub txtCrAmount_GotFocus()
        '''If gbSearchStr <> "" Then
        '''    Dim mStr As String
        '''    txtCrHeadCode.Text = Token(gbSearchStr, " ")
        '''    txtCrAccountHead.Text = Trim(gbSearchStr)
        '''    gbSearchStr = ""
        '''End If
        txtCrAmount.Text = Format(CalculateAmt, "0.00")
    End Sub
    
    Private Sub txtCrAmount_KeyPress(KeyAscii As Integer)
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
        Else
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtCrAmount_LostFocus()
        txtCrAmount.Text = Format(txtCrAmount.Text, "0.00")
    End Sub
    
    Private Sub txtCrHeadCode_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtCrHeadCode.Text = Token(gbSearchStr, " ")
            txtCrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
        End If
    End Sub
    
    Private Sub txtCrAccountHead_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtCrHeadCode.Text = Token(gbSearchStr, " ")
            txtCrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
        End If
    End Sub
    
    Private Sub txtCrHeadCode_LostFocus()
        Dim mChequeNo As Variant
        Dim mBalanceAmt As Variant
        Dim objCr As New clsAccounts
        Dim objBk As New clsBank
        
        objCr.SetAccountCode Trim(txtCrHeadCode.Text)
        If objCr.AccountHeadID > 0 Then
            If objCr.AccountHeadID = 1504 Then
               
                txtCrHeadCode.Text = gbAcHeadCodeCash
                txtCrAccountHead.Text = "Cash"
                txtCrHeadCode.Tag = gbAcHeadIDCash
                txtCrAccountHead.Tag = gbAcHeadIDCash
                txtInstrument.Text = "Cash"
                txtNameOfBank.Text = ""
                txtInstrumentNo.Text = ""
            Else
                txtCrAccountHead.Text = objCr.AccountHead
                txtCrHeadCode.Tag = objCr.AccountHeadID
                txtCrHeadCode.Text = objCr.AccountCode
                objBk.SetBankInfoByAccID objCr.AccountHeadID
                If objBk.BankAccountHeadID > -1 Then
                    txtNameOfBank.Text = objBk.BankName
                    txtBranch.Text = objBk.Branch
                    txtAccountNo.Text = objBk.AccountNumber
                    'mChequeNo = objBk.GetNeWChequeNumber
                    'txtRef.Text = IIf(IsNull(mChequeNo), "", mChequeNo)
                Else
                    txtNameOfBank.Text = ""
                    txtBranch.Text = ""
                    txtAccountNo.Text = ""
                    txtInstrument.Text = ""
                    'txtRef.Text = ""
                End If
            End If
            '''mBalanceAmt = objCr.GetLedgerBalance(objCr.AccountHeadID)
            '''If Not IsNull(mBalanceAmt) Then
            '''    txtDr.Tag = mBalanceAmt
            '''End If
        Else
            txtCrAccountHead.Text = ""
            txtCrHeadCode.Text = ""
            txtCrHeadCode.Tag = ""
        End If
    End Sub



Private Sub txtDate_LostFocus()
    Dim mDate As Date
    
    If IsDate(txtDate.Text) Then
        mDate = CDate(txtDate)
    Else
        MsgBox "Please check the Transaction Date!", vbInformation
        Exit Sub
    End If
    
    If mPreYearMode = 0 Then
        If Not (mDate >= gbStartingDate And mDate <= gbEndingDate) Then
            MsgBox "Not a valid transaction date!", vbInformation
            Exit Sub
        End If
    Else
        If Not (mDate >= DateAdd("yyyy", -1, gbStartingDate) And mDate <= DateAdd("yyyy", -1, gbEndingDate)) Then
            MsgBox "Not a valid transaction date!", vbInformation
            Exit Sub
        End If
    End If
    txtDate.Text = DdMmmYy(mDate)
End Sub

    Private Sub txtDated_LostFocus()
        If val(txtInstrument.Tag) <> gbInstrumentCash Then
            txtDated.Text = CheckDateInMMM(txtDated.Text)
        End If
    End Sub
    Private Sub txtImplementingOfficer_GotFocus()
        If gbSearchID > 0 Then
            txtImplementingOfficer.Text = gbSearchStr
            txtImplementingOfficer.Tag = gbSearchID
            gbSearchID = -1
            gbSearchCode = ""
            gbSearchStr = ""
        End If
    End Sub

    Private Sub txtInit1_Change()
        If Len(Trim(txtInit1.Text)) > 0 Then
            txtInit2.SetFocus
        End If
    End Sub
    Private Sub txtInit2_Change()
        If Len(Trim(txtInit2.Text)) > 0 Then
            txtInit3.SetFocus
        End If
    End Sub
    Private Sub txtInit3_Change()
        If Len(Trim(txtInit3.Text)) > 0 Then
            txtInit4.SetFocus
        End If
    End Sub
    Private Sub txtInit4_Change()
        If Len(Trim(txtInit4.Text)) > 0 Then
            txtHouse.SetFocus
        End If
    End Sub

    Private Sub txtInstrument_LostFocus()
        Dim objBk As New clsBank
        Dim objAc As New clsAccounts
        '-----------------------------------
        ' Default Account Head             '
        '-----------------------------------
        
        If val(txtInstrument.Tag) = gbInstrumentCash Then
            Call LockingClearBanks(True)
            If txtCrHeadCode.Tag = "" Then
                txtCrHeadCode.Text = gbAcHeadCodeCash
                txtCrAccountHead.Text = "Cash"
                txtCrHeadCode.Tag = gbAcHeadIDCash
                txtCrAccountHead.Tag = gbAcHeadIDCash
                txtInstrument.Text = "Cash"
                txtNameOfBank.Text = ""
                txtInstrumentNo.Text = ""
                txtAccountNo.Text = ""
            ElseIf txtCrHeadCode.Tag <> gbAcHeadIDCash Then
                txtCrHeadCode.Text = gbAcHeadCodeCash
                txtCrAccountHead.Text = "Cash"
                txtCrHeadCode.Tag = gbAcHeadIDCash
                txtCrAccountHead.Tag = gbAcHeadIDCash
                txtNameOfBank.Text = ""
                txtInstrumentNo.Text = ""
            End If
            
        Else
            If txtCrHeadCode.Text = gbAcHeadCodeCash Then
                Call LockingClearBanks(False)
                '''Modified on 7 mar 2017 Special TSB Acc For Joint Venture Project
                 If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                    If val(txtSourceofFund.Tag) = 10 Or val(txtSourceofFund.Tag) = 11 Or val(txtSourceofFund.Tag) = 12 Or val(txtSourceofFund.Tag) = 13 Or val(txtSourceofFund.Tag) = 14 Then
                        objBk.SetBankInfoByAccID (IIf(val(txtCrHeadCode.Tag) > 0, txtCrHeadCode.Tag, gbAcHeadIDTreasuryAccountSpecialTSB))
                    Else
                        objBk.SetBankInfoByAccID (IIf(val(txtCrHeadCode.Tag) > 0, txtCrHeadCode.Tag, gbDefaultBankID))
                    End If
                End If
                txtCrAccountHead.Text = objBk.BankName
                txtCrHeadCode.Text = objBk.BankAccountHeadCode
                txtCrAccountHead.Tag = objBk.BankAccountHeadID
                txtCrHeadCode_LostFocus
            
            Else
                If Trim(txtCrHeadCode.Text) = "" Then
                    Call LockingClearBanks(False)
                End If
                 If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                '''Modified on 7 mar 2017 Special TSB Acc For Joint Venture Project
                If txtSourceofFund.Tag = 10 Or txtSourceofFund.Tag = 11 Or txtSourceofFund.Tag = 12 Or txtSourceofFund.Tag = 13 Or txtSourceofFund.Tag = 14 Then
                    objBk.SetBankInfoByAccID (IIf(val(txtCrHeadCode.Tag) > 0, txtCrHeadCode.Tag, gbAcHeadIDTreasuryAccountSpecialTSB))
                    txtCrAccountHead.Text = objBk.BankName
                    txtCrHeadCode.Text = objBk.BankAccountHeadCode
                    txtCrAccountHead.Tag = objBk.BankAccountHeadID
                    txtCrHeadCode_LostFocus
                End If
                '''.........................
                End If
            End If
            
        End If
        
    End Sub

    Private Sub LockingClearBanks(val As Boolean)
        txtNameOfBank.Locked = val
        txtInstrumentNo.Locked = val
        txtAccountNo.Locked = val
        txtDated.Locked = val
        txtBranch.Locked = val
        
'        If val Then
            txtNameOfBank.Text = ""
            txtInstrumentNo.Text = ""
            txtAccountNo.Text = ""
            txtDated.Text = ""
            txtDated.Enabled = True
            txtCrHeadCode.Tag = ""
            txtBranch.Text = ""
'        End If
    End Sub



    Private Sub txtInstrumentNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtInstrumentNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
        txtInstrumentNo.Locked = True
    Else
        txtInstrumentNo.Locked = False
    End If
    End Sub

Private Sub txtName_GotFocus()
    If gbSearchID > -1 Then
        Dim ObjSubLed As New clsSubLedger
        ObjSubLed.SetSubLedgerDetails (gbSearchID)
        gbSearchID = -1
        gbSearchCode = ""
        gbSearchStr = ""
        If ObjSubLed.SubsidiaryAccountHeadID > -1 Then
            txtSubLedgerType.Text = ObjSubLed.SubLedgerType
            txtSubLedgerType.Tag = ObjSubLed.SubLedgerTypeID
            txtName.Tag = ObjSubLed.SubsidiaryAccountHeadID
            txtName.Text = ObjSubLed.NameOfSubLedger
            
            txtPayee.Text = ObjSubLed.NameOfSubLedger
            txtPayeeType.Text = ObjSubLed.SubLedgerType
            
            txtHouse.Text = ObjSubLed.HouseOrOffice
            txtStreet.Text = ObjSubLed.Street
            txtLocalPlace.Text = ObjSubLed.LocalPlace
            txtMainPlace.Text = ObjSubLed.MainPlace
            txtPost.Text = ObjSubLed.PostOffice
            txtPin.Text = ObjSubLed.PinCode
            txtPhone.Text = ObjSubLed.Phone
        End If
    End If
End Sub

Private Sub txtPayee_DblClick()
    SSTab.Tab = 1
    frmSearchSubsidiaryAccountHeads.Show vbModal
End Sub

Private Sub txtPayee_GotFocus()
    lblTipText.Caption = "Double Click for select a payee from subledger.." & vbLf
    lblTipText.Caption = lblTipText.Caption + " Other wise you can type and give a Payee Name and Address in the Second Tab named <Subledger>"
End Sub



Private Sub txtPayeeType_DblClick()
    SSTab.Tab = 1
End Sub




Private Sub txtPayOrder_LostFocus()
    If val(txtPayOrder) > 0 Then
        mPaymentOrderBasedMode = False
        Call DisplayPayOrder(val(txtPayOrder))
        On Error Resume Next
        txtInstrumentNo.SetFocus
        On Error GoTo 0
        cmdSave.Enabled = True
        LoadMode = 2
    End If
End Sub
    Private Sub txtPhone_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtPin_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

Private Sub txtProjectNo_GotFocus()
    If gbProject.decProjectID > 0 Then
        Dim objProj As New clsProject
        objProj.SetProject gbProject.decProjectID
        If objProj.ProjectID > 0 Then
            txtProjectNo = objProj.ProjectSerialNo
            txtProjectNo.Tag = objProj.ProjectID
            
            'txtCategory.Text = objProj.Category
            'txtCategory.Tag = objProj.ProjCatID
            
            txtSector.Text = objProj.Sector
            txtSector.Tag = objProj.SectorTypeID
            
            txtSourceofFund.Text = objProj.FindSourceOfFund(gbProject.intSourceOfFundID)
            txtSourceofFund.Tag = gbProject.intSourceOfFundID
        End If
        With gbProject
            .decProjectID = Null
            .intLBID = Null
            .intYearID = Null
            .intProjectSlNo = Null
            .chvProjectSlNo = Null
            .chvProjectName = Null
            .chvProjectnameEnglish = Null
            .intProjCatID = Null
            .chvDPCOrderNo = Null
            .dtDPCOrderDate = Null
            .intSectorTypeID = Null
            .intPlanID = Null
            .intSourceOfFundID = Null
            .fltEstSourceAmt = Null
        End With
    End If
End Sub

Private Sub txtSourceOfFund_GotFocus()
    If gbSearchID > -1 And Trim(gbSearchStr) <> "" Then
        txtSourceofFund.Text = gbSearchStr
        txtSourceofFund.Tag = gbSearchID
        gbSearchID = -1
        gbSearchCode = ""
        gbSearchStr = ""
    Else
        txtSourceofFund.Text = ""
        txtSourceofFund.Tag = ""
    End If
End Sub

Private Sub txtVoucherNo_LostFocus()
    If txtVoucherNo.Text <> "" Then
       If mID(txtVoucherNo.Text, 1, 1) = "2" Then
            Call DisplayVoucherDetails(txtVoucherNo.Text)
       Else
            MsgBox "Please Enter Valid payment Voucher", vbInformation
       End If
    End If
End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) = False And vsGrid.Col = 3 Then
            vsGrid.TextMatrix(vsGrid.Row, 3) = ""
            MsgBox "Enter Numeric values"
        End If
        Dim mTot As Variant
        If vsGrid.Col = 3 Then
                If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) Then
                    vsGrid.TextMatrix(vsGrid.Row, 3) = Format(val(vsGrid.TextMatrix(vsGrid.Row, 3)), "0.00")
                    mTot = CalculateAmt
                    '''If Val(txtCrAmount.Text) > Val(mTOt) Then
                    '''    txtCrAmount.Text = Val(txtCrAmount.Text) - Val(mTOt)
                    '''Else
                    '''    MsgBox "Amount Out of Range"
                    '''End If
                    txtCrAmount.Text = Format(mTot, "0.00")
                    lblTotal.Caption = txtCrAmount.Text
                End If
        End If
        
    End Sub
    
    Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        vsGrid.TextMatrix(vsGrid.Row, 3) = Format(vsGrid.TextMatrix(vsGrid.Row, 3), "0.00")
        
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If mPaymentOrderBasedMode Then
            Cancel = True
            Exit Sub
        End If
        If vsGrid.Row <> 1 Then
            If vsGrid.TextMatrix(Row - 1, 1) = "" Then
                Cancel = True
            End If
        End If
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        ''''    Dim objAc As New clsAccounts
        ''''    Dim objDb As New clsDb
        ''''    Dim mCnn As New ADODB.Connection
        ''''    Dim Rec As New ADODB.Recordset
        ''''    Dim mSQL As String
        ''''
        ''''        If Val(txtTransactionType.Tag) = 1001 Or Val(txtTransactionType.Tag) = 1007 Then
        ''''            frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faTransactionTypeChild INNER JOIN faAccountHeads ON faAccountHeads.vchAccountHeadCode= faTransactionTypeChild.vchAccountHeadCode WHERE faTransactionTypeChild.intTransactionTypeID=1001"
        ''''            'mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
        ''''            'mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
        ''''            'mSql = mSql + " Where intTransactionTypeID = " & Val(cmbTransactionType.Tag) & " Order By faTransactionTypeChild.intOrder"
        ''''            'frmSearchAccountHeads.SQLString = mSql
        ''''
        ''''            frmSearchAccountHeads.Show vbModal
        ''''            vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        ''''            vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        ''''            objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
        ''''            vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
        ''''            objDb.SetConnection mCnn
        ''''            Rec.Open "Select intGroupID From faTransactionTypeChild Where vchaccountheadcode ='" & vsGrid.TextMatrix(vsGrid.Row, 1) & "'", mCnn
        ''''                If Not (Rec.EOF And Rec.BOF) Then
        ''''                     vsGrid.TextMatrix(vsGrid.Row, 5) = Rec!intGroupID
        ''''                End If
        ''''                mCnn.Close
        ''''            vsGrid.Col = 3
        ''''            gbSearchStr = ""
        ''''
        ''''        ElseIf Val(txtTransactionType.Tag) = 1002 Or Val(txtTransactionType.Tag) = 1003 Then
        ''''            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where faaccountHeads.vchAccountHeadCode Between '350200201' and '350200299'"
        ''''            frmSearchAccountHeads.Show vbModal
        ''''            vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        ''''            vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        ''''            objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
        ''''            vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
        ''''            objDb.SetConnection mCnn
        ''''            Rec.Open "Select intGroupID From fatransactiontypechild Where vchaccountheadcode ='" & vsGrid.TextMatrix(vsGrid.Row, 1) & "'", mCnn
        ''''                If Not (Rec.EOF And Rec.BOF) Then
        ''''                     vsGrid.TextMatrix(vsGrid.Row, 5) = Rec!intGroupID
        ''''                End If
        ''''                mCnn.Close
        ''''            vsGrid.Col = 3
        ''''            gbSearchStr = ""
        ''''
        ''''        Else
        ''''            If Trim(txtCrHeadCode) <> "" Then
        ''''                'Note:-
        ''''                'Debit Head is selected
        ''''                Dim mGroupID As Integer
        ''''                mGroupID = 0
        ''''                objDb.SetConnection mCnn
        ''''                mSQL = "Select * From faTransactionTypeChild Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " AND vchAccountHeadCode = '" & txtCrHeadCode.Text & "'"
        ''''                Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic
        ''''                If Not (Rec.BOF And Rec.EOF) Then
        ''''                    mGroupID = Rec!intGroupID
        ''''                    If mGroupID = 0 Then GoTo ListAllDeductions:
        ''''                    mSQL = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        ''''                    mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        ''''                    mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        ''''                    mSQL = mSQL + " And (faTransactionTypeChild.intGroupID = 0 Or faTransactionTypeChild.intGroupID = " & mGroupID & " )"
        ''''                    mSQL = mSQL + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        ''''                    frmSearchAccountHeads.SQLString = mSQL
        ''''                Else
        ''''ListAllDeductions:
        ''''                    mSQL = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        ''''                    mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        ''''                    mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        ''''                    mSQL = mSQL + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        ''''                    frmSearchAccountHeads.SQLString = mSQL
        ''''                End If
        ''''                Rec.Close
        ''''
        ''''            Else
        ''''                mSQL = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        ''''                mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        ''''                mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        ''''                mSQL = mSQL + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        ''''                frmSearchAccountHeads.SQLString = mSQL
        ''''            End If
        ''''            'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
        ''''
        ''''            frmSearchAccountHeads.Show vbModal
        ''''            vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        ''''            vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        ''''            objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
        ''''            vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
        ''''            vsGrid.Col = 3
        ''''            gbSearchStr = ""
        ''''        End If
        '-------------------------------------------------------------------------------------------------------'
        
        frmSearchAccountHeads.VoucherMode = 201
        
        If val(txtInstrument.Tag) <> 0 Then
            If txtInstrument.Tag = 1 Then
                'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where intGroupID  IN (2) And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where intGroupID is Null And tinType IN (1,4) And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
                frmSearchAccountHeads.Show vbModal
            Else
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (2,3,4) AND intGroupID is Null And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
                frmSearchAccountHeads.Show vbModal
            End If
        Else
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where  tinType IN (2,3,4) AND intGroupID is Null And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
        End If
        
        vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        vsGrid.TextMatrix(vsGrid.Row, 4) = gbSearchID
        vsGrid.Col = 3
        
        gbSearchID = -1
        gbSearchStr = ""
        
    End Sub
    Private Sub vsGrid_GotFocus()
'        If Trim(txtDrAmount.Text = "") Then
'            MsgBox "It is Mandatory to Enter"
'            txtDrAmount.SetFocus
'            Exit Sub
'        End If
    End Sub
    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        Dim mTot As Variant
        If KeyAscii = 13 And vsGrid.Col = 3 And Trim(vsGrid.TextMatrix(vsGrid.Row, 3)) <> "" And IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) And val(vsGrid.TextMatrix(vsGrid.Row, 3)) > 0 Then
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.Row = vsGrid.Row + 1
            vsGrid.Col = 1
            mTot = CalculateAmt
            
            If val(txtCrAmount.Text) > val(mTot) Then
                'txtCrAmount.Text = Val(txtDrAmount.Text) - Val(mTOt)
            End If
        End If
    End Sub

'''Private Function CheckValidation() As Boolean
'''    If Val(txtCrHeadCode.Tag) = 0 Then
'''        MsgBox "Please Select The Credit Account Head", vbInformation
'''        cmdCrAccountHead.SetFocus
'''        CheckValidation = False
'''        Exit Function
'''    End If
'''    If txtCrAmount.Text = "" Then
'''        MsgBox "Please Give the Credit Amount", vbInformation
'''        txtCrAmount.SetFocus
'''        CheckValidation = False
'''        Exit Function
'''    End If
'''    CheckValidation = True
'''End Function



Public Property Get PayOrderID() As Variant
    PayOrderID = intPayOrderID
End Property

Public Property Get LoadMode() As Integer
    LoadMode = intLoadMode
End Property

Public Property Let LoadMode(Data As Integer)
    intLoadMode = Data
End Property

Public Property Let ListLoaded(Data As Boolean)
    mViewPayOrderListFormIsLoaded = Data
End Property

Public Sub DisplayPayOrder(intPayOrderNo As Variant)
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objdb As New clsDB
    Dim mSql As String
    Dim objAc As New clsAccounts
    Dim mLoopCount As Long
    Dim objSubLedger As New clsSubLedger
    Dim objInst As New clsInstruments
    Dim objProj As New clsProject
    Dim mAffectedRecords As Integer
    Dim RecTask As New Recordset
    Dim mIncludeDeductionsFlag As Boolean

    
    Call FormInitialize
    
    mLoopCount = 0
    
    mSql = " Select isNull(faPayOrder.intModuleID,0) PreYearMode, faPayOrder.intFinancialYearID intPOFinancialYearID,faPayOrder.tnyStatus POStatus,  faPayOrder.*, faPayOrderChild.*, faPayOrderAddress.*, faFunctionaries.vchFunctionary, faFunctions.vchFunction,faPayOrderChild.tnyExcludeFromSourceFlag, " & vbNewLine
    mSql = mSql + " faTransactionType.vchTransactionType, faInstrumentTypes.vchInstrumentType, faPayOrder.vchDescription as PODesc, chvSeatTitle,faAgreements.*,faAllotments.vchAllotmentNo,  " & vbNewLine
    mSql = mSql + " faAllotments.intSchemeID,faAllotments.tnyTypeID mUnAuthTypeID, suGOForFunds.intRefID GoID,suGOForFunds.vchRefNo GoNo  From faPayOrder Inner Join " & vbNewLine
    mSql = mSql + " faPayOrderChild On faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID Left Join " & vbNewLine
    mSql = mSql + " faPayOrderAddress On faPayOrderAddress.intPayOrderID = faPayOrder.intPayOrderID Left Join" & vbNewLine
    mSql = mSql + " faFunctionaries On faFunctionaries.intFunctionaryID = faPayOrder.intFunctionaryID Left Join" & vbNewLine
    mSql = mSql + " faFunctions On faFunctions.intFunctionID = faPayOrder.intFunctionID Left Join" & vbNewLine
    mSql = mSql + " faTransactionType On faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID Left Join" & vbNewLine
    mSql = mSql + " faInstrumentTypes On faInstrumentTypes.intInstrumentTypeID = faPayOrder.intInstrumentTypeID Left Join " & vbNewLine
    mSql = mSql + " faSeats On faPayOrder.numSeatID = faSeats.numSeatID " & vbNewLine
    mSql = mSql + " Left Join faAllotments On faPayOrder.intAllotmentID = faAllotments.intID" & vbNewLine
    mSql = mSql + " Left Join faAgreements On faAgreements.intAgreementID=faPayOrder.intAgreementID" & vbNewLine
    mSql = mSql + " Left Join suGOForFunds On suGOForFunds.intPayOrderID=faPayOrder.intPayOrderID" & vbNewLine
    mSql = mSql + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID=faPayOrder.intSourceOfFundID" & vbNewLine
    mSql = mSql + " Where faPayOrder.vchPayOrderNo = '" & intPayOrderNo & "'"
    mSql = mSql + " Order by tnyCategoryFlag, intSlNo"
    
    objdb.SetConnection mCnn
    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
    If Not (Rec.BOF And Rec.EOF) Then
    
        If Rec!POstatus = 0 Then
            MsgBox "This Payment Order is not Approved", vbInformation
            FormInitialize
            Exit Sub
        End If
        
        If Not IsNull(Rec!intVoucherID) Then
            MsgBox "Payment is already recorded for this Pay Order No", vbInformation
            FormInitialize
            Exit Sub
        End If
    
        mPaymentOrderBasedMode = True '
        txtTransactionType.Enabled = False
        cmdSearchTransactionType.Enabled = False
        
        If Rec!intPOFinancialYearID <> gbFinancialYearID Then
            If IsDate(Rec!dtDueDate) Then
                txtDate.Text = DdMmmYy(Rec!dtDueDate)
                txtDate.Tag = DdMmmYy(Rec!dtDueDate)
                txtDate.Enabled = True
                'mSql = "SELECT * FROM faPendingTaskRequest WHERE vchInstrumentNo = '" & intPayOrderNo & "'"
                mSql = "SELECT * FROM faPendingTaskRequest WHERE numDemandID  = " & Rec!intPayOrderID
                RecTask.Open mSql, mCnn, adOpenStatic, adLockReadOnly
                If Not (RecTask.BOF And RecTask.EOF) Then
                    If RecTask!intTransactionTypeID = 1001 And Month(RecTask!dtTransactionDate) = 3 Then
                        mPreYearMode = 1 ' CHANGED ON FEB,2014 BUILD 2.2.14
                    Else
                        mPreYearMode = 1
                    End If
                Else
                    mPreYearMode = 0
                    txtDate.Text = gbTransactionDate
                End If
                RecTask.Close
            End If
        Else
            txtDate.Tag = ""
            txtDate.Enabled = False
            mPreYearMode = 0
        End If
        
        ' BLOCK PREVIOUS YEARS PAYMENT ORDERS
'        If Rec!PreYearMode = 96 Then
'            mPreYearMode = 1
'        Else
'            mPreYearMode = 0
'        End If
        
        If mPreYearMode = 0 Then ''Module id 96 for Pre year PayOrder
            If Rec!dtPayOrderDate < gbStartingDate Or Rec!dtPayOrderDate > gbEndingDate Then
                If Rec!intTransactionTypeID = gbTransactionTypePayBills Then
                    If Rec!dtPayOrderDate < DateAdd("m", -1, gbStartingDate) Or Rec!dtPayOrderDate > gbEndingDate Then
                        MsgBox "This Payment belongs to the previous year And month less March, plz verify", vbInformation
                        cmdSave.Enabled = False
                        Exit Sub
                    End If
                Else
                    MsgBox "This Payment belongs to the previous year, plz verify", vbInformation
                    cmdSave.Enabled = False
                    Exit Sub
                End If
    '            MsgBox "This seems to be previous years Payment Order, plz verify", vbInformation
    '            Exit Sub
            End If
        End If
        
        If Rec!tnyStatus = 0 Then
            MsgBox "Please Approve the Payment Order before making Payment", vbInformation
            'MsgBox "Payment Voucher is already generated for this Payment Order", vbInformation
            txtPayOrder.SetFocus
            Exit Sub
        End If
        'If Not IsNull(Rec!intVoucherID) Then
        If Rec!tnyStatus = 2 Then
            MsgBox "Payment Voucher is already generated for this Payment Order", vbInformation
            txtPayOrder.SetFocus
            Exit Sub
        End If
        
        If IsNull(Rec!tnyCancelled) = False Then
            If Rec!tnyCancelled = 1 Then
                MsgBox "The payorder is Cancelled", vbInformation
                Exit Sub
            End If
        End If
        
        txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        txtVoucherNo.Tag = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
        
    
        txtPayOrder.Text = Rec!vchPayOrderNo
        txtPayOrder.Tag = Rec!intPayOrderID
        
        If mPreYearMode = 0 Then
            txtDated.Text = txtDate.Text  'DdMmmYy(gbTransactionDate)
        Else
            
            txtDated.Text = Format(Rec!dtDueDate, "dd/mmm/yyyy")
            txtDated.Enabled = False
            
        End If
        
        If Not IsNull(Rec!vchFunctionary) Then
            txtFunctionary.Text = Rec!vchFunctionary
            txtFunctionary.Tag = Rec!intFunctionaryID
        Else
            txtFunctionary.Text = ""
            txtFunctionary.Tag = ""
        End If
        
        If Not IsNull(Rec!vchFunction) Then
            txtFunction.Text = Rec!vchFunction
            txtFunction.Tag = Rec!intFunctionID
        Else
            txtFunction.Text = ""
            txtFunction.Tag = ""
        End If
        
        If Not IsNull(Rec!vchTransactionType) Then
            txtTransactionType.Text = Rec!vchTransactionType
            txtTransactionType.Tag = Rec!intTransactionTypeID
            If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Or val(txtTransactionType.Tag) = gbTransactionTypeProjectExpGO Then
                txtGo.Text = IIf(IsNull(Rec!GoNo), "", Rec!GoNo)
                txtGo.Tag = IIf(IsNull(Rec!Goid), "", Rec!Goid)
            End If
        Else
            txtTransactionType.Tag = -1
            txtTransactionType.Text = ""
        End If
        ''  Modified For SaankhyaWeb Updation

        If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
            If Not IsNull(Rec!intSourceOfFundID) Then
                objProj.FindSourceOfFund (Rec!intSourceOfFundID)
                txtSourceofFund.Text = objProj.SourceOfFund
                txtSourceofFund.Tag = objProj.SourceOfFundID
            End If
            
        Else
            If Not IsNull(Rec!intSourceOfFundID) Then
                objProj.FindSourceOfFund (Rec!intSourceOfFundID)
                txtSourceofFund.Text = objProj.SourceOfFund
                txtSourceofFund.Tag = objProj.SourceOfFundID
            Else
                txtSourceofFund.Text = "OwnFund"
                txtSourceofFund.Tag = 4
                cmdSourceOfFund.Enabled = True
                'txtSourceOfFund.Enabled = True
            End If
        End If
        
        'BLOCK[1]
        'NOTE: IDENTIFY THE SOURCES WHICH FOR WHICH DEDUCTIONS
        '      ARE ALSO CREDITED FROM THE BANK WHILE PAYMENT
        Select Case Rec!intSourceOfFundID
            Case Is = 1, 29, 30, 16, 17, 25, 26, 10, 11, 12, 13, 14, 41
                mIncludeDeductionsFlag = True
            Case Else
                mIncludeDeductionsFlag = False
        End Select
        'END OF BLOCK[1]
        
        objAc.SetAccountID IIf(IsNull(Rec!intCashOrBankHeadID), 0, Rec!intCashOrBankHeadID)
        While Not Rec.EOF
            If Rec!intSlNo = 1 And Rec!tnyCategoryFlag = 1 Then
                objAc.SetAccountID Rec!intAccountHeadID
                If objAc.AccountHeadID > 0 Then
                    
                Else
                    MsgBox "Error: Head Not Found", vbInformation
                End If
            End If

            '------------------------------------------------------------------'
            'Note:- Deduction Block
            '------------------------------------------------------------------'
            ' If Source Of Fund is Development Fund Or Maintenance Road Or
            ' Maintenance Non-Road , all deductions will also be paid directly
            ' from the Treasury. So there Deductions will also be fetched for
            ' Payments - Net Payable Amount will be Gross Amount
            '------------------------------------------------------------------'
            
            'If Rec!intSourceOfFundID = 1 Or Rec!intSourceOfFundID = 16 Or Rec!intSourceOfFundID = 17 _
            '  Or Rec!intSourceOfFundID = 25 Or Rec!intSourceOfFundID = 26 Then
            If mIncludeDeductionsFlag = True Then
                If Rec!tnyCategoryFlag = 2 Then
                    objAc.SetAccountID Rec!intAccountHeadID
                    If objAc.AccountHeadID > 0 And Rec!tnyExcludeFromSourceFlag = 0 Then
                        mLoopCount = mLoopCount + 1
                        vsGrid.Cell(flexcpText, mLoopCount, 0) = mLoopCount
                        vsGrid.Cell(flexcpText, mLoopCount, 1) = objAc.AccountCode
                        vsGrid.Cell(flexcpText, mLoopCount, 2) = objAc.AccountHead
                        vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
                        vsGrid.Cell(flexcpText, mLoopCount, 4) = objAc.AccountHeadID
                    End If
                End If
            End If
            
            'Note:- End of Deduction Block
            
            'Note:- Net Payable
'             If Rec!tnyCategoryFlag = 3 Then
'                objAc.SetAccountID Rec!intAccountHeadID
'                If objAc.AccountHeadID > 0 Then
'                    mLoopCount = mLoopCount + 1
'                    vsGrid.Cell(flexcpText, mLoopCount, 0) = mLoopCount
'                    vsGrid.Cell(flexcpText, mLoopCount, 1) = objAc.AccountCode
'                    vsGrid.Cell(flexcpText, mLoopCount, 2) = objAc.AccountHead
'                    vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
'                    vsGrid.Cell(flexcpText, mLoopCount, 4) = objAc.AccountHeadID
'                End If
'            End If
            
            
            If Rec!tnyCategoryFlag = 3 Then
                If Not (IsNull(Rec!intSubsidiaryCashBookID)) Then
                    If (Rec!intSubsidiaryCashBookID > 0) Then
                        objAc.SetAccountID gbAcHeadIDMiscAdvance
                    Else
                        objAc.SetAccountID Rec!intAccountHeadID
                    End If
                Else
                    objAc.SetAccountID Rec!intAccountHeadID
                End If

                If objAc.AccountHeadID > 0 Then
                    mLoopCount = mLoopCount + 1
                    vsGrid.Cell(flexcpText, mLoopCount, 0) = mLoopCount
                    vsGrid.Cell(flexcpText, mLoopCount, 1) = objAc.AccountCode
                    vsGrid.Cell(flexcpText, mLoopCount, 2) = objAc.AccountHead
                    vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
                    vsGrid.Cell(flexcpText, mLoopCount, 4) = objAc.AccountHeadID
                End If
            End If

            '================================================='
            If Rec!tnyCategoryFlag = 4 Then
                objAc.SetAccountID Rec!intAccountHeadID
                If objAc.AccountHeadID > 0 Then
                    txtSubCashCode.Text = objAc.AccountCode
                    txtSubCashCode.Tag = objAc.AccountHeadID
                End If
            End If
            '================================================='
            
            Rec.MoveNext
        Wend
        
        
        'BLOCK[2] CREDIT BANK ACCOUNT
        '------------------------------------------------------------------'
        'Note:- Auto Selection of  Bank Account Heads
        '------------------------------------------------------------------'
        Rec.MoveFirst
        
      
        mNewACRMode = IIf(IsNull(Rec!intKeyID), 0, Rec!intKeyID)
        
        If mNewACRMode = 1 Then
            '------------------------------------------------------------------'
            'Note:- No Source Of Fund Bank Filtering is for NEW TREASURY MODE
            '------------------------------------------------------------------'
            Dim objBank As New clsBank
            'objBank.SetBankInfo CInt(gbDefaultBankID) 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID")
            'If objBank.BankID > 0 Then
                objAc.SetAccountID 1504 'objBank.BankAccountHeadID ''Commented on 7 mar 2017 Special TSB Acc For Joint Venture Project
                '''Modified on 7 mar 2017 Special TSB Acc For Joint Venture Project
                ''-----------------------------------------------------------------
                Select Case Rec!intSourceOfFundID
                    Case Is = 10, 11, 12, 13, 14 ' RECEIPTS FROM OTHER LSIG's
                        objAc.SetAccountCode gbAcHeadCodeTreasuryAccountSpecialTSB
                    Case Else
                        objAc.SetAccountID 1504
                End Select
'                ''-----------------------------------------------------------------
            'Else
                'objAc.SetAccountID 0
            'End If
         Else
            Select Case Rec!intSourceOfFundID
                Case Is = 1  ' Development Fund (Plan Fund)
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount2
                Case Is = 29
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount6
                Case Is = 30
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount7
                Case Is = 16 ' Maintenance Fund (Road)
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount3
                Case Is = 17 ' Maintenance Fund (Non-Road)
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount3
                Case Is = 25 ' CFC
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount4
                Case Is = 26, 41 ' KLGSDP
                    objAc.SetAccountCode gbAcHeadCodeTreasuryAccount5
                
                Case Is = 10, 11, 12, 13, 14 ' RECEIPTS FROM OTHER LSIG's
                
                    '''Modified on 7 mar 2017 Special TSB Acc For Joint Venture Project
''
''                    If Rec!tnyCategoryID = 1 Then
''                        objAc.SetAccountCode gbAcHeadCodeTreasuryAccount2
''                    ElseIf Rec!tnyCategoryID = 2 Then
''                         objAc.SetAccountCode gbAcHeadCodeTreasuryAccount6
''                    ElseIf Rec!tnyCategoryID = 3 Then
''                         objAc.SetAccountCode gbAcHeadCodeTreasuryAccount7
''                    Else
''                         objAc.SetAccountID gbDefaultBankID
''                    End If

                      objAc.SetAccountCode gbAcHeadCodeTreasuryAccountSpecialTSB
                   
                Case Is = 4
                    If val(txtTransactionType.Tag) = 1141 Or val(txtTransactionType.Tag) = 1151 Or val(txtTransactionType.Tag) = 1161 Or _
                        val(txtTransactionType.Tag) = 1171 Or val(txtTransactionType.Tag) = 1181 Or val(txtTransactionType.Tag) = 1191 Then
                           objAc.SetAccountCode gbAcHeadCodeTreasuryAccountTSB
                    Else
                         objAc.SetAccountID 1504
                    End If
                Case Else    ' Default Bank
                    'Dim objBank As New clsBank
                    objBank.SetBankInfo CInt(gbDefaultBankID) 'ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID")
                    If objBank.BankID > 0 Then
                        objAc.SetAccountID objBank.BankAccountHeadID
                    Else
                        objAc.SetAccountID 0
                    End If
            End Select
         End If
        
        '========================================================'
        If txtSubCashCode.Text <> "" Then
            objAc.SetAccountCode gbAcHeadCodeCash
        End If
        '========================================================'
        
        If objAc.AccountHeadID > 0 Then
            txtCrHeadCode.Text = objAc.AccountCode
            txtCrHeadCode.Tag = objAc.AccountHeadID
            txtCrAccountHead.Text = objAc.AccountHead
            Call txtCrHeadCode_LostFocus
        End If
        'END OF BLOCK [2]
        '
        
        
        'Note:- Net Amount Re-Calculate From Grid and Diplayed
        txtCrAmount.Text = Format(CalculateAmt, "0.00")
        lblTotal.Caption = txtCrAmount.Text
        
        'If mNewACRMode <> 1 Then
            'Note:- Selecting InstrumentType based on Fund (Bank )
            Select Case Rec!intSourceOfFundID
                Case 1, 16, 17, 29, 30, 25, 26, 27, 28, 10  '', 11, 12, 13, 14, 41 'ON 23-12-2012
                    If mNewACRMode = 1 Then
                        objInst.SetInstrumentType 1
                    Else
                        objInst.SetInstrumentType 5
                        txtInstrumentNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    End If
                Case Is = 10, 11, 12, 13, 14 ' RECEIPTS FROM OTHER LSIG's
                        objInst.SetInstrumentType 5  ''Modified from 1 to 5 on 14/Mar/2017
                        cmdInstrument.Enabled = False
                Case Else
                    If gbSectionID = gbJSKSectionID Then
                        objInst.SetInstrumentType 1
                    Else
                        objInst.SetInstrumentType 5
                    End If
                     
            End Select
            '========================================================'
            If txtSubCashCode.Text <> "" Then
                objInst.SetInstrumentType 1
            End If
            '========================================================'
            If mNewACRMode <> 1 Then
                If objInst.InstrumentTypeID > 0 Then
                    txtInstrument.Text = objInst.InstrumentType
                    txtInstrument.Tag = objInst.InstrumentTypeID
                    
                End If
            Else
'                If txtSourceOfFund.Tag = 10 Or txtSourceOfFund.Tag = 11 Or txtSourceOfFund.Tag = 12 Or txtSourceOfFund.Tag = 13 _
'                    Or txtSourceOfFund.Tag = 14 Or txtSourceOfFund.Tag = 2 Then
'                    objInst.SetInstrumentType 5
'                    txtInstrument.Tag = 5
'                Else
'                    objInst.SetInstrumentType 1
'                    txtInstrument.Tag = 1
'                End If
            End If
            
            If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                
            Else
                objInst.SetInstrumentType 5
                txtInstrument.Tag = 5
            End If
            Call txtInstrument_LostFocus
        'Note:- End of selecting Instrument Type
'        Else
'            objInst.SetInstrumentType 1
'        End If
        
        If Not IsNull(Rec!vchName) Then
            txtName.Text = Rec!vchName
            txtPayee.Text = Rec!vchName
        Else
            txtName.Text = ""
            txtPayee.Text = ""
        End If
        On Error Resume Next
        If Not IsNull(Rec!vchInit1) Then
            txtInit1.Text = Rec!vchInit1
        Else
            txtInit1.Text = ""
        End If
        If Not IsNull(Rec!vchInit2) Then
            txtInit2.Text = Rec!vchInit2
        Else
            txtInit2.Text = ""
        End If
        If Not IsNull(Rec!vchInit3) Then
            txtInit3.Text = Rec!vchInit3
        Else
            txtInit3.Text = ""
        End If
        If Not IsNull(Rec!vchInit4) Then
            txtInit4.Text = Rec!vchInit4
        Else
            txtInit4.Text = ""
        End If
        On Error GoTo 0
        If Not IsNull(Rec!vchHouseName) Then
            txtHouse.Text = Rec!vchHouseName
        Else
            txtHouse.Text = ""
        End If
        If Not IsNull(Rec!vchStreet) Then
            txtStreet.Text = Rec!vchStreet
        Else
            txtStreet.Text = ""
        End If
        If Not IsNull(Rec!vchLocalPlace) Then
            txtLocalPlace.Text = Rec!vchLocalPlace
        Else
            txtLocalPlace.Text = ""
        End If
        If Not IsNull(Rec!vchMainPlace) Then
            txtMainPlace.Text = Rec!vchMainPlace
        Else
            txtMainPlace.Text = ""
        End If
        
        If Not IsNull(Rec!vchPost) Then
            txtPost.Text = Rec!vchPost
        Else
            txtPost.Text = ""
        End If
        If Not IsNull(Rec!vchPinCode) Then
            txtPin.Text = Rec!vchPinCode
        Else
            txtPin.Text = ""
        End If
        If Not IsNull(Rec!vchPhone) Then
            txtPhone.Text = Rec!vchPhone
        Else
            txtPhone.Text = ""
        End If
        
        objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intSubsidiaryAccountHeadID), 0, Rec!intSubsidiaryAccountHeadID)
        If objSubLedger.SubLedgerTypeID > 0 Then
            txtSubLedgerType.Text = objSubLedger.SubLedgerType
            txtSubLedgerType.Tag = objSubLedger.SubLedgerTypeID
            
            txtPayeeType.Text = objSubLedger.SubLedgerType
            txtPayeeType.Tag = objSubLedger.SubLedgerTypeID
            'To get the empID from Sthapana for SubCashBook
            txtName.Tag = objSubLedger.EmpID
        End If
        
        txtNarration.Text = IIf(IsNull(Rec!PODesc), "", Rec!PODesc)
        txtNarration.Tag = IIf(IsNull(Rec!intModuleID), "", Rec!intModuleID)
        txtSeat.Tag = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
        txtSeat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
        
        If Not IsNull(Rec!intSubsidiaryCashBookID) Then
            objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intSubsidiaryCashBookID), 0, Rec!intSubsidiaryCashBookID)
            txtSubsidiaryCash.Text = IIf(IsNull(objSubLedger.Title), "", objSubLedger.Title)
            txtSubsidiaryCash.Tag = Rec!intSubsidiaryCashBookID
        End If
        
        If Not IsNull(Rec!intImplementingOfficerID) Then
            txtImplementingOfficer.Text = IIf(Rec!intImplementingOfficerID = 0, "", Rec!vchFunctionary)
            txtImplementingOfficer.Tag = Rec!intImplementingOfficerID
        Else
            txtImplementingOfficer.Text = ""
            txtImplementingOfficer.Tag = ""
        End If
        '-------------------------------------------------  'ADDED BY MINU ON 25-05-2011
        If Not IsNull(Rec!intAgreementID) Then
            txtAgreementNo.Text = Rec!vchAgreementNo
            txtAgreementNo.Tag = Rec!intAgreementID
        Else
            txtAgreementNo.Text = ""
            txtAgreementNo.Tag = ""
        End If
        '-------------------------------------------------
        If Not IsNull(Rec!numProjectNo) Then
            If Rec!numProjectNo > 0 Then
                If mPreYearMode Then
                    objProj.SetProject Rec!numProjectNo, gbFinancialYearID - 1
                Else
                    objProj.SetProject Rec!numProjectNo
                End If
                If objProj.ProjectID > 0 Then
                    txtProjectNo.Text = objProj.ProjectSerialNo
                    txtProjectNo.Tag = objProj.ProjectID
                    txtCategory.Text = objProj.Category
                    txtCategory.Tag = objProj.ProjCatID
                    txtSector.Text = objProj.SubSector
                    txtSector.Tag = objProj.SubSectorID
                    objProj.FindSourceOfFund Rec!intSourceOfFundID
                    txtSourceofFund.Text = objProj.SourceOfFund
                    txtSourceofFund.Tag = objProj.SourceOfFundID
                    txtScheme.Text = objProj.SchemeID
                    txtScheme.Text = IIf(IsNull(Rec!intSchemeID), 0, Rec!intSchemeID)
                End If
            End If
        Else
            txtProjectNo.Text = ""
            txtProjectNo.Tag = ""
        End If
        '-----------------------------------------------------------'
        'Locking the Particular Details Frame
        '-----------------------------------------------------------'
'        Frame4.Enabled = False
        If val(txtTransactionType.Tag) = 1001 Then 'Pay & Allowance
            Frame4.Enabled = True
            txtNarration.Enabled = False
        Else
            Frame4.Enabled = False
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '                   If Allotment Letter is Linked                '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not (IsNull(Rec!intAllotmentID)) Then
            If Rec!intAllotmentID > 0 Then
                txtAllotmentLetterNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                txtAllotmentLetterNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                mUnAuthorized = IIf(IsNull(Rec!mUnAuthTypeID), 0, Rec!mUnAuthTypeID)

                
                Dim objAllotment As New clsAllotmentLetter
                objAllotment.SetAllotment (txtAllotmentLetterNo.Tag)
                txtSourceofFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
                txtSourceofFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), "", objAllotment.SourceOfFundID)
                
                txtCategory.Text = IIf(IsNull(objAllotment.Category), "", objAllotment.Category)
                txtCategory.Tag = IIf(IsNull(objAllotment.CategoryID), "", objAllotment.CategoryID)
                
                txtImplementingOfficer.Text = IIf(IsNull(objAllotment.ImplementingOfficer), "", objAllotment.ImplementingOfficer)
                txtImplementingOfficer.Tag = IIf(IsNull(objAllotment.ImplementingOfficersID), "", objAllotment.ImplementingOfficersID)
            End If
        End If
        Rec.Close
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            If txtVoucherNo.Text <> "" Then
                   mSql = "SELECT faSubsidiaryCashBook.intID, faSubsidiaryCashBook.intTransferID, faSubsidiaryCashBook.intSubsidiaryAccountHeadID,faSubsidiaryCashBook.intTypeID, faSubsidiaryCashBook.intVoucherID,faVouchers.intVoucherNo From faSubsidiaryCashBook"
                   mSql = mSql + " INNER JOIN faVouchers ON faSubsidiaryCashBook.intVoucherID = faVouchers.intVoucherID"
                   mSql = mSql + " Where intTypeID = 50 And faVouchers.intVoucherNo = " & txtVoucherNo.Text
                   Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
                    If Not (Rec.EOF Or Rec.BOF) Then
                        mintIntID = IIf(Rec!intID = "", -1, Rec!intID)
                        mintTransferID = IIf(Rec!intTransferID = "", -1, Rec!intTransferID)
                    End If
            Else
                mintIntID = -1
                mintTransferID = -1
            End If
    End If
'    Rec.Close
    PayOrderNo = ""
    Set mCnn = Nothing
    
End Sub

Private Function SaveValidation() As Boolean
        Dim objBank     As New clsBank
        Dim mStr As String
  '  On Error GoTo err:
        SaveValidation = False
        Dim mTot As Variant
        'If gbSeatGroupID <> gbSeatGroupAccountsOfficer Then
         If mWebExtract = False Then
         
            If val(txtPayOrder.Tag) < 1 Then
                If val(txtVoucherNo.Tag) < 1 Then
                    MsgBox "Please Make Payment through Payment Order ", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
            End If
        End If
        'End If
        If Not IsDate(txtDate) Then
            MsgBox "Please Check the Transaction Date", vbInformation
            txtDated.SetFocus
            SaveValidation = False
            Exit Function
        End If
    
        If val(txtFunctionary.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Functionary", vbInformation
            cmdSearchFunctionary.SetFocus
            SaveValidation = False
            Exit Function
        End If
        
        If val(txtFunction.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Function", vbInformation
            cmdSearchFunction.SetFocus
            SaveValidation = False
            Exit Function
        End If
        
        If val(txtTransactionType.Tag) < 1 Then
            MsgBox "Please Select Proper Transaction Type for this Transaction", vbInformation
            txtTransactionType.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If txtInstrument.Tag <> 1 And txtInstrument.Tag <> 7 Then
            If IsDate(txtDated.Text) = False Then
                MsgBox "Please Give Due Date", vbInformation
                txtDated.SetFocus
                SaveValidation = False
                Exit Function
            End If
            
            If txtInstrument.Tag <> 10 Then ''' this line added On 31.10.2011 By Anisha to Skip instrument No Validation for Directly debited By Bank
                If txtInstrumentNo.Text = "" Then
                    MsgBox "Please Give the Instrument/Cheque Number", vbInformation
                    txtInstrumentNo.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            End If                          ''' this line added On 31.10.2011 By Anisha to Skip instrument No Validation for Directly debited By Bank
        End If
       
        If txtInstrument.Tag = 1 Then
             If val(txtCrHeadCode.Tag) <> gbAcHeadIDCash Then
                 MsgBox "Please Select Cash Head for Cash Instrument", vbInformation
                 txtCrAccountHead.SetFocus
                 SaveValidation = False
                 Exit Function
             End If
        End If
        If val(txtCrHeadCode.Tag) < 1 Then
            MsgBox "Please Select The Credit Account Head", vbInformation
            cmdCrAccountHead.SetFocus
            SaveValidation = False
            Exit Function
        End If
        
        If txtCrAmount.Text = "" Then
            MsgBox "Please Give the Credit Amount", vbInformation
            txtCrAmount.SetFocus
            SaveValidation = False
            Exit Function
        End If
        
        If vsGrid.TextMatrix(1, 1) = "" Then
            MsgBox "Please Select the Debit Account Head", vbInformation
            vsGrid.SetFocus
            SaveValidation = False
            Exit Function
        End If
       
        mTot = CalculateAmt
        If val(txtCrAmount.Text) <> val(mTot) Then
            MsgBox "Debit and Credit amount Should be Equal, Please Correct the Amount", vbInformation
            txtCrAmount.SetFocus
            SaveValidation = False
            Exit Function
        End If
       If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
            If val(txtSourceofFund.Tag) < 1 Then
            txtSourceofFund.Enabled = True
                MsgBox "Please Select the Source Of Fund", vbInformation
                txtSourceofFund.SetFocus
                SaveValidation = False
                Exit Function
            End If
        End If
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            MsgBox "Please Enter the Name of Payee..", vbInformation
            SaveValidation = False
            Exit Function
        End If
        If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
        
        If mUnAuthorized <> 3 Then
            If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
                Dim mCnnSulekha As New ADODB.Connection
                Dim objdb As New clsDB
                If Not (objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                    MsgBox "Connection To Plan [Sulekha] Module not found", vbCritical
                    SaveValidation = False
                    Exit Function
                End If
                
                If val(txtProjectNo.Tag) < 1 Then
                    MsgBox "Please select a Project", vbInformation
                    txtProjectNo.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                
                If val(txtCategory.Tag) < 1 Then
                    MsgBox "Please select a Category", vbInformation
                    txtCategory.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                
                If val(txtSector.Tag) < 1 Then
                    MsgBox "Please select a Sector", vbInformation
                    txtSector.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                If val(txtAgreementNo.Tag) < 1 Then
                    MsgBox "Please select an Agreement", vbInformation
                    'txtAgreementNo.SetFocus
                    'SaveValidation = False
                    'Exit Function
                End If
            End If
        End If
        End If
        If txtSubsidiaryCash.Text <> "" Then
            If val(txtSubLedgerType.Tag) = 10 And txtName.Text = "" Then
                MsgBox "Please Select the Official for disbursing the Subsidiary Cash"
                txtName.SetFocus
                SaveValidation = False
                Exit Function
            End If
        End If
        
        If val(txtCrHeadCode.Tag) <> 0 Then
            If val(txtCrHeadCode.Tag) <> 1504 Then
                If CDate(txtDate.Text) <= GetLastReconDate(val(txtCrHeadCode.Tag)) Then
                    mStr = ""
                    mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                    mStr = mStr + " No new Transaction is allowed to Enter during the period."
                    MsgBox mStr, vbInformation
                    txtCrHeadCode.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            End If
        End If
        If mWebExtract = True Then
            If VerifyWebExtract = True Then
                SaveValidation = True
                Exit Function
            Else
                MsgBox "E bill Receipt deatails not Sync Properly "
                SaveValidation = False
                Exit Function
            End If
        End If
        
        SaveValidation = True
        
    Exit Function
err:
    MsgBox (Error$)
End Function

Private Function VerifyWebExtract() As Boolean
    Dim mCnn            As New ADODB.Connection
    Dim Rec             As New ADODB.Recordset
    Dim objdb           As New clsDB
    Dim mSql As String
    VerifyWebExtract = False
    mSql = "Select * From faWebExtracts "
    mSql = mSql + "Inner Join faWebExtractChild On faWebExtracts.intWebExtractID=faWebExtractChild.intWebExtractID"
    mSql = mSql + " Where tnyVoucherTypeID=1 And intAccountHeadID<>1504 And numBillControlID=" & txtBillControCodeID.Text
    objdb.SetConnection mCnn
    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
    If Not (Rec.BOF And Rec.EOF) Then
        txtBillControCodeID.Tag = IIf(IsNull(Rec!intWebExtractID), 0, Rec!intWebExtractID)
        txtCreditHdIDR.Tag = IIf(IsNull(Rec!intAccountHeadID), 0, Rec!intAccountHeadID)
        VerifyWebExtract = True
        Exit Function
    Else
        VerifyWebExtract = False
    End If
    
End Function
Private Sub Saving()
    Dim mV              As uVoucher
    Dim mVC             As uVChild
    Dim mVA             As uVoucherAddress
    Dim mVS             As uVoucherSub
    Dim mT              As uTr
    Dim mTC             As uTrChild
    Dim arrInput        As Variant
    Dim arrOutPut       As Variant
    Dim mintVoucherID   As Variant
    Dim mintVoucherNo   As Variant
    Dim mintTransactionID As Variant
    Dim objdb           As New clsDB
    Dim mCnn            As New ADODB.Connection
    Dim Rec             As New ADODB.Recordset
    Dim mLoop           As Integer
    Dim mSql            As String
    Dim mCnnSulekha     As New ADODB.Connection
    
    Dim mYearID         As Integer
    Dim mDate           As Date
    Dim mStr            As String
    
    
    mSaveFlag = False
    If txtName.Tag <> "" Then
        If txtSubLedgerType.Text = "Officials" Then
            If CheckOfficial = False Then
                MsgBox "This official can't operate Saankhya Application, Please add this Official as a User in Saankhya through Admin Module", vbInformation
                Exit Sub
            End If
        End If
    End If
    '''Added Anisha on 2/Jul/2014
    If val(txtCrHeadCode.Tag) <> 0 Then
        If val(txtCrHeadCode.Tag) <> gbAcHeadIDCash Then
            If mPreYearMode = 0 Then
                If gbTransactionDate <= GetLastReconDate(val(txtCrHeadCode.Tag)) Then
                    mStr = ""
                    mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                    mStr = mStr + " No new Transaction is allowed to Enter during the period."
                    MsgBox mStr, vbInformation
                    txtCrHeadCode.SetFocus
                    Exit Sub
                 End If
            Else
                If CDate(Format(txtDate.Text, "dd/mmm/yyyy")) <= GetLastReconDate(val(txtCrHeadCode.Tag)) Then
                    mStr = ""
                    mStr = mStr + " Selected Bank or Treasury is reconciled for the month." & vbCrLf
                    mStr = mStr + " No new Transaction is allowed to Enter during the period."
                    MsgBox mStr, vbInformation
                    txtCrHeadCode.SetFocus
                    Exit Sub
                 End If
            End If
        End If
    End If
    
    If cmdSave.Caption <> "&Edit" Then
    '    On Error Resume Next
        '=============================================='
        ' Getting Active Connection                    '
        '=============================================='
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '----------------------------------------------'
        
        With mV
            .intVoucherID_1 = txtVoucherNo.Tag
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = val(txtTransactionType.Tag)
            .tnyVoucherTypeID_5 = 20
            .intVoucherNo_6 = IIf(txtVoucherNo.Text = "", Null, txtVoucherNo.Text)
            .intBookNo_7 = Null
            If mWebExtract = True Then
                .dtDate_8 = Format(txtDate.Text, "dd/mmm/yyyy")
                mDate = .dtDate_8
            ElseIf mPreYearMode = 0 Then
                .dtDate_8 = gbTransactionDate
                mDate = gbTransactionDate
            Else
                .dtDate_8 = Format(txtDate.Text, "dd/mmm/yyyy")
                mDate = .dtDate_8
            End If
            .fltAmount_9 = val(txtCrAmount)
            .intInstrumentTypeID_10 = val(txtInstrument.Tag)
            .vchInstrumentNo_11 = Trim(txtInstrumentNo)
            .dtInstrumentDate_12 = IIf(IsDate(txtDated), txtDated.Text, gbTransactionDate)
            .vchDescription_13 = Trim(txtNarration)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .intKeyID2_23 = txtPayOrder.Text
            If mWebExtract = True Then
               .numSubLedgerID_21 = txtWebExtractIDforP.Tag
               .intKeyID2_23 = txtBillControCodeID.Text
            Else
                .numSubLedgerID_21 = Null
            End If
            .intKeyID1_22 = val(txtCrHeadCode.Tag)
            
            If mWebExtract = True Then
                .intExternalApplicationID_24 = 118
            Else
                .intExternalApplicationID_24 = Null
            End If
            If mNewACRMode = 1 Then         'TO IDENTIFY PAYMENTS FROM NEW ACR MODE
                .intExternalModuleID_25 = 1
            Else
                .intExternalModuleID_25 = txtNarration.Tag
            End If
            If mPreYearMode = 0 Then
                .intFinancialYearID_26 = gbFinancialYearID
                mYearID = gbFinancialYearID
            Else
                .intFinancialYearID_26 = gbFinancialYearID - 1
                mYearID = gbFinancialYearID - 1
            End If
            .tnyShiftID_27 = gbShiftID
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = 0
            .vchBank_33 = txtNameOfBank.Text
            .vchBankPlace_34 = txtBranch.Text
            .intFundID_35 = 1
            .numSeatID = val(txtSeat.Tag)
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            
                              ' Changed By Aiby On 23-Apr-2011 to Introduce Approval for Payment VOuchers
                              ' tnyStatus_32=6
            .tnyStatus_32 = 0 ' Before Approval For Payment Voucher
                              ' This will skip the Voucher Number Generation
                              
            .numLocationID = gbLocationID
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intTransactionID_3, _
            .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
            .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
            .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
            .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
            .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
            .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, _
            .tnyShiftID_27, .tnyPrintFlag_28, _
            .tnyCancelFlag_29, .vchBank_33, _
            .vchBankPlace_34, .intFundID_35, _
            .numSeatID, .intSessionID, _
            .vchRefNo, .fltRoundOff, _
            .fltAdvAmtAdj, .numInwardNo, _
            .tnyStatus_32, .numLocationID)
            
            '=============================================='
            ' T r a n s a c t i o n   B e g i n            '
            '=============================================='
            mCnn.BeginTrans
            On Error GoTo ErrRollBank:
            '=============================================='
            objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn, adCmdStoredProc
            
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
            mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mintVoucherNo = Rec!intVoucherNo 'IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            End If
            Rec.Close
            VoucherID = mintVoucherID
            VoucherNo = mintVoucherNo
            'mintVoucherNo = arrOutPut(1, 0)
        Else
            GoTo ErrRollBank:
        End If
        ''SaankhyaWeb
        If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
        
            If val(txtAllotmentLetterNo.Tag) > 0 Then
                arrInput = Array(mYearID, _
                            gbLocalBodyID, _
                            txtProjectNo.Tag, _
                            mintVoucherID, _
                            mintVoucherNo, _
                            mDate, _
                            txtInstrumentNo.Text, _
                            txtDated.Text, _
                            txtCrHeadCode.Tag, _
                            val(txtCrAmount.Text), _
                            val(txtSourceofFund.Tag), _
                            val(txtAllotmentLetterNo.Tag), _
                            Null _
                            )
                objdb.ExecuteSP "spSaveSUExpenditures", arrInput, , , mCnn, adCmdStoredProc
                mSql = "Update faAllotments set  tnyProjectStatus=2 Where intID=" & val(txtAllotmentLetterNo.Tag) & "  "
                mCnn.Execute mSql
            End If
        Else
        End If
        mSql = "Delete From faVoucherAddress Where intVoucherID = " & mintVoucherID
        mCnn.Execute mSql
        
        With mVA
            .intVoucherID = mintVoucherID
            .intLocalBodyID = gbLocalBodyID
            .vchName = Trim(txtName)
            .vchInit1 = Trim(txtInit1)
            .vchInit2 = Trim(txtInit2)
            .vchInit3 = Trim(txtInit3)
            .vchInit4 = Trim(txtInit4)
            .vchHouseName = Trim(txtHouse)
            .vchStreetName = Trim(txtStreet)
            .vchLocalPlace = Trim(txtLocalPlace)
            .vchMainPlace = Trim(txtMainPlace)
            .vchPostOffice = Trim(txtPost)
            .vchDistrict = Null
            .vchPinNumber = Trim(txtPin)
            .vchPhone = Trim(txtPhone)
            .intWardNo = Null
            .intDoorNo = Null
            .vchDoorNo2 = Null
        
            arrInput = Array(.intVoucherID, _
                .intLocalBodyID, _
                .vchName, _
                .vchInit1, _
                .vchInit2, _
                .vchInit3, _
                .vchInit4, _
                .vchHouseName, _
                .vchStreetName, _
                .vchLocalPlace, _
                .vchMainPlace, _
                .vchPostOffice, _
                .vchDistrict, _
                .vchPinNumber, _
                .vchPhone, _
                .intWardNo, _
                .intDoorNo, _
                .vchDoorNo2)
                
            objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn, adCmdStoredProc
        End With
        
          
        
            With mVS
                .intVoucherID = mintVoucherID
                .intLocalBodyID = gbLocalBodyID
                
                If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                    .decProjectID = IIf(val(txtProjectNo.Tag) > 0, val(txtProjectNo.Tag), Null)
                     .intCreditorTypeID = IIf(val(txtSubLedgerType.Tag) > 0, val(txtSubLedgerType.Tag), Null)
                    .intCreditorsID = IIf(val(txtName.Tag) > 0, val(txtName.Tag), Null)
                Else
                    .decProjectID = Null
                    .intSourceOfFundID = IIf(val(txtSourceofFund.Tag) > 0, val(txtSourceofFund.Tag), Null)
                    .intCategoryID = IIf(val(txtCategory.Tag) > 0, val(txtCategory.Tag), Null)
                    .intSectorID = IIf(val(txtSector.Tag) > 0, val(txtSector.Tag), Null)
                    .intAllotmentID = IIf(val(txtAllotmentLetterNo.Tag) > 0, val(txtAllotmentLetterNo.Tag), Null)
                    .intAgreementID = IIf(val(txtAgreementNo.Tag) > 0, val(txtAgreementNo.Tag), Null)
                    .intCashBookID = IIf(val(txtSubsidiaryCash.Tag) > 0, val(txtSubsidiaryCash.Tag), Null)
                    .intImplementingOfficerID = IIf(val(txtImplementingOfficer.Tag) > 0, val(txtImplementingOfficer.Tag), Null)
                    .intCreditorTypeID = Null
                    .intCreditorsID = Null
                    .intTypeID = Null
                    '.intLocalBodyID = gbLocalBodyID
                End If
           
            
                mSql = "Delete From faVoucherSub Where intVoucherID = " & .intVoucherID
                mCnn.Execute mSql
                
                arrInput = Array(.intVoucherID, _
                .intLocalBodyID, _
                .decProjectID, _
                .intSourceOfFundID, _
                .intCategoryID, _
                .intSectorID, _
                .intAllotmentID, _
                .intAgreementID, _
                .intCashBookID, _
                .intImplementingOfficerID, _
                .intCreditorTypeID, _
                .intCreditorsID, _
                .intTypeID)
                objdb.ExecuteSP "spSaveVoucherSub", arrInput, , , mCnn, adCmdStoredProc
            
            End With
'        Else
'
'        End If
        If val(txtSubsidiaryCash.Tag) > 0 Then
            arrInput = Array(IIf(IsNull(mintIntID), -1, mintIntID), _
                            IIf(IsNull(mintTransferID), -1, mintTransferID), _
                            val(txtSubsidiaryCash.Tag), _
                            50, _
                            Format(CDate(txtDate.Text), "dd/mmm/yyyy"), _
                            mOfficialUserID, _
                            mOfficialSeatID, _
                            val(txtSubCashCode.Tag), _
                            val(txtFunctionary.Tag), _
                            val(txtFunction.Tag), _
                            val(txtCrAmount.Text), _
                            Null, _
                            Null, _
                            Null, _
                            txtNarration.Text, 0, 1, mintVoucherID)
            objdb.ExecuteSP "spSaveSubsidiaryCashBook", arrInput, , , mCnn, adCmdStoredProc
        End If
        
      
        
        mSql = "Select intTransactionID From faTransactions Where intVoucherID = " & mintVoucherID
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            mintTransactionID = Rec!intTransactionID
        Else
            mintTransactionID = -1
        End If
        If Rec.State = 1 Then Rec.Close
        With mT
            .intTransactionID = mintTransactionID
            .intLocalBodyID = gbLocalBodyID
'            If mWebExtract = True Then
'
'                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
'                .intFinancialYearID = gbFinancialYearID
'            ElseIf mPreYearMode = 0 Then
'               ' .intFinancialYearID = gbFinancialYearID
'                .dtTransactionDate = gbTransactionDate
'            Else
'              '  .intFinancialYearID = gbFinancialYearID - 1
'                '.dtTransactionDate = Format(txtDated.Text, "dd/mmm/yyy")
'                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
'            End If

            If mWebExtract = True Then
                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
               ' mDate = .dtDate_8
            ElseIf mPreYearMode = 0 Then
                .dtTransactionDate = gbTransactionDate
               ' mDate = gbTransactionDate
            Else
                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
                'mDate = .dtDate_8
            End If
            
            If mPreYearMode = 0 Then
                .intFinancialYearID = gbFinancialYearID
                '.intFinancialYearID_26 = gbFinancialYearID
                mYearID = gbFinancialYearID
            Else
                .intFinancialYearID = gbFinancialYearID - 1
                mYearID = gbFinancialYearID - 1
            End If
            If mWebExtract = True Then
                
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = val(txtFunction.Tag)
            .intFunctionaryID = val(txtFunctionary.Tag)
            .intFieldID = Null
            .intFundID = 1
            .intBudgetCentreID = Null
            .vchNarration = Trim(txtNarration)
            .intTransactionTypeID = val(txtTransactionType.Tag)
            .intProcessID = Null
            .vchGroup = "P"
            .intGroupID = 20
            .intKeyID = Null
            .numSubLedgerID = IIf(val(txtName.Tag) > 0, val(txtName.Tag), Null)
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
            objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        Else
            
        End If
        
        mSql = "Delete from faTransactionChild Where intTransactionID = " & mintTransactionID
        mCnn.Execute mSql
        
        mSql = "Delete From faVoucherChild Where intVoucherID = " & mintVoucherID
        mCnn.Execute mSql
        
        With mTC
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = val(txtCrHeadCode.Tag)
            .fltAmount = Format(val(txtCrAmount), "0.00")
            .tinDebitOrCreditFlag = 0
            .intByAccountHeadID = Null
            .vchNarration = Trim(txtNarration)
            .intFundID = 1
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        
        For mLoop = 1 To vsGrid.Rows - 1
            If vsGrid.TextMatrix(mLoop, 4) <> "" Then
                With mVC
                    .intVoucherID_1 = mintVoucherID
                    .intLocalBodyID_2 = gbLocalBodyID
                    .intSlNo_3 = mLoop
        
                    .intAccountHeadID_4 = val(vsGrid.TextMatrix(mLoop, 4))
                    .tnyDebitOrCredit_5 = 1
                    .intYearID_6 = Null
                    .tnyPeriodID_7 = Null
                    .tnyArrearFlag_8 = Null
                    .numDemandID_9 = Null
                    .fltAmount_10 = Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
        
                    arrInput = Array(.intVoucherID_1, _
                    .intLocalBodyID_2, _
                    .intSlNo_3, _
                    .intAccountHeadID_4, _
                    .tnyDebitOrCredit_5, _
                    .intYearID_6, _
                    .tnyPeriodID_7, _
                    .tnyArrearFlag_8, _
                    .numDemandID_9, _
                    .fltAmount_10)
                    objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn, adCmdStoredProc
                End With
        
                With mTC
                    .intTransactionID = mintTransactionID
                    .intSerialNo = mLoop + 1
                    .intAccountHeadID = val(vsGrid.TextMatrix(mLoop, 4))
                    .fltAmount = Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
                    .tinDebitOrCreditFlag = 1
                    .intByAccountHeadID = val(txtCrHeadCode.Tag)
                    .vchNarration = Null
                    .intFundID = 1
        
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
                End With
            Else
                Exit For
            End If
        Next
        
        
        'mSql = "Update faPayOrder Set tnyStatus = 2, intVoucherID = " & mintVoucherID & " , intVoucherNo = " & mintVoucherNo & " Where vchPayOrderNo = '" & txtPayOrder.Text & "'" ::: Changed by Aiby for PV approval and Avoiding Voucher No Generation
        'mSql = "Update faPayOrder Set tnyStatus = 2, intVoucherID = " & mintVoucherID & " Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
        mSql = "Update faPayOrder Set intVoucherID = " & mintVoucherID & " , intVoucherNo = " & mintVoucherNo & " Where vchPayOrderNo = '" & txtPayOrder.Text & "'" 'This line of code is replace with  2.2.6 version. Modified By Anisha On 30/5/2011
        mCnn.Execute mSql
        
        If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
            If val(txtGo.Tag) > 0 Then
                mSql = "Update suGOForFunds Set intVoucherID = " & mintVoucherID & " Where intRefID= " & txtGo.Tag
                mCnn.Execute mSql
            End If
        End If
        '' Expense Details to Sulekha
        '' On 22-12-2012 By Anisha
        ''
        
        
        '=============================================='
        ' T r a n s a c t i o n   C o m m i t t i n g
        '=============================================='
        mCnn.CommitTrans
        '=============================================='
            ''  Modified For SaankhyaWeb Updation (blocking Updation in sulekha)
           
'            If mWebExtract = True Then
'                Call GenerateEbillReceipt(mintVoucherNo)
'            End If
            If mWebExtract = True Then
                Call GenerateEbillReceipt
                If mPreYearMode = 1 Then
                    objdb.ExecuteSP "Update faWebExtracts set intExtractTypeID=1,numKeyID=" & mintVoucherID & " ,tnyPendingTask=1,dtPendingDate=getdate() Where intWebExtractID=" & txtName.Tag, , , , mCnn, adCmdText
                Else
                
                    objdb.ExecuteSP "Update faWebExtracts set intExtractTypeID=1,numKeyID=" & mintVoucherID & " Where intWebExtractID=" & txtName.Tag, , , , mCnn, adCmdText
                End If
                MsgBox "Saved Sucessfully"
                
                Unload Me
            
            'Exit Sub
            
            End If
            If mPreYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                ''''''''''''FOR NEW ACR MODE''''''''''''''''''''''
                
                If mNewACRMode = 1 Then
                
                    Call GenerateReceipt(mintVoucherNo)
                    
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''
            
      
                On Error GoTo SkipSulekha:
                'If cmdSave.Caption <> "&Edit" Then
                 Dim mSourceFundID As Integer
                 mSourceFundID = val(txtSourceofFund.Tag)
                   If mSourceFundID = 41 And gbFinancialYearID - 1 = 2016 Then ''Map Expense details of KLGSDP State share to Central Share For the year 2016-17 only
                        mSourceFundID = 26
                   End If
                   If val(txtProjectNo.Tag) > 0 Then
                       If objdb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha) Then
                           arrInput = Array(gbLBID, _
                                           mYearID, _
                                           val(txtProjectNo.Tag), _
                                           -1, mSourceFundID, _
                                           val(txtCrAmount), _
                                           mintVoucherID, _
                                           mDate)
                           objdb.ExecuteSP "ExpenseDetails_I", arrInput, , , mCnnSulekha, adCmdStoredProc
                       End If
                   End If
                   txtVoucherNo.Text = mintVoucherNo
            End If
        End If
        
        mSaveFlag = True
       
        If val(txtNarration.Tag) = 75 Then 'If ModuleID = 75 then i.e Module = WaterBill
            Call UpdatePVStatus
        End If
        If mWebExtract = True Then
            mWebExtract = False
            Unload Me
        Else
            If gbLocalBodyID = 167 Then
                txtVoucherNo.Text = mintVoucherNo
                MsgBox "Payment Voucher " & mintVoucherNo & " Saved Successfully", vbApplicationModal
                'txtVoucherNo.Text = mintVoucherNo
            Else
                frmViewVoucher.FormName = "PaymentVoucher"
                If cmdSave.Caption <> "&Edit" Then
                    frmViewVoucher.ArrayIn = Array(CStr(mintVoucherID))
                Else
                    frmViewVoucher.ArrayIn = Array(CStr(txtVoucherNo.Tag))
                End If
                frmViewVoucher.Show vbModal
                
                WaterBillPVMode = True
            End If
        End If
    Exit Sub
ErrRollBank:
    mCnn.RollbackTrans
    Set mCnn = Nothing
    Exit Sub
SkipSulekha:
    MsgBox "Didn't able to update in Sulekha"

End Sub

Private Sub UpdatePVStatus()
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim objdb As New clsDB
        Dim i As Integer
        
        On Error GoTo err
        If objdb.CreateNewConnection(mCnn, enuSourceString.iSaankhyaMasters) Then
            mSql = "Update snWrBillDetails"
            mSql = mSql + " Set tnyStatus = 5,"
            mSql = mSql + " intVoucherID = " & VoucherID & ","
            mSql = mSql + " intVoucherNo ='" & VoucherNo & "'"
            mSql = mSql + " Where vchPayOrderNo = " & txtPayOrder.Text
            mCnn.Execute mSql
            frmIntegratedPayments.WaterBillPVMode = False
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
Public Sub DisplayVoucherDetails(mVoucherNo As String)
        Dim mCnn            As New ADODB.Connection
        Dim objdb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mRec            As New ADODB.Recordset
        Dim mSql            As String
        Dim mRowCount       As Double
        Dim mArrearFlag     As Variant
        Dim RecAccHeads     As New ADODB.Recordset
        Dim mSqlAccHeads    As String
        Dim mSeatID         As Variant
        Dim mStatus         As Variant
        
        Call FormInitialize
        
        If Not IsNumeric(mVoucherNo) Then
            MsgBox "Invalid Voucher Number!", vbInformation
            txtVoucherNo.SetFocus
            Exit Sub
        End If
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mSql = " Select faVouchers.*, faFunctionaries.vchFunctionary, faFunctions.vchFunction, faFunds.vchFund, faInstrumentTypes.vchInstrumentType,"
        mSql = mSql + " faVoucherChild.*, faVoucherAddress.*, faTransactionType.vchTransactionType, " & vbNewLine
        mSql = mSql + " faTransactions.intTransactionID, faTransactions.intFunctionaryID, faTransactions.intFunctionID, " & vbNewLine
        mSql = mSql + " faAccountHeads.intAccountHeadID, faAccountHeads.vchAccountHeadCode, faAccountHeads.vchAccountHead, faVoucherSub.* ,isNull(faVouchers.intExternalModuleID,0) ModuleID"
        mSql = mSql + " From faVouchers Inner Join" & vbNewLine
        mSql = mSql + " faTransactions On faTransactions.intVoucherID = faVouchers.intVoucherID Left Join" & vbNewLine
        mSql = mSql + " faTransactionType On faTransactionType.intTransactionTypeID = faVouchers.intTransactionTypeID Left Join" & vbNewLine
        mSql = mSql + " faFunctionaries On faFunctionaries.intFunctionaryID = faTransactions.intFunctionaryID Left Join" & vbNewLine
        mSql = mSql + " faFunctions On faFunctions.intFunctionID = faTransactions.intFunctionID Left Join" & vbNewLine
        mSql = mSql + " faFunds On faFunds.intFundID = faVouchers.intFundID Left Join" & vbNewLine
        mSql = mSql + " faInstrumentTypes On faInstrumentTypes.intInstrumentTypeID = faVouchers.intInstrumentTypeID Left Join" & vbNewLine
        mSql = mSql + " faVoucherChild On faVoucherChild.intVoucherID = faVouchers.intVoucherID Left Join" & vbNewLine
        mSql = mSql + " faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID Left Join " & vbNewLine
        mSql = mSql + " faAccountHeads On faAccountHeads.intAccountHeadID = faVouchers.intKeyID1 " & vbNewLine
        mSql = mSql + " Left Join faVoucherSub On faVouchers.intVoucherID = faVoucherSub.intVoucherID " & vbNewLine
        mSql = mSql + " Where faVouchers.intVoucherNo = " & mVoucherNo
        
        Rec.Open mSql, mCnn
        If (Rec.EOF And Rec.BOF) Then
            Exit Sub
        End If
        
'        If Rec!dtDate < gbStartingDate Or Rec!dtDate > gbEndingDate Then
'            MsgBox "This Payment belongs to the previous year, plz verify", vbInformation
'            cmdSave.Enabled = False
'        End If
'        Added on 1/06/2012 By Anisha
        If mViewMode <> 1 Then
            If Rec!dtDate < gbStartingDate Or Rec!dtDate > gbEndingDate Then
                If Rec!intTransactionTypeID = gbTransactionTypePayBills Then
                    If Rec!dtDate < DateAdd("m", -1, gbStartingDate) Or Rec!dtDate > gbEndingDate Then
                        MsgBox "This Payment belongs to the previous year And month less March, plz verify", vbInformation
                        cmdSave.Enabled = False
                    End If
                Else
                    MsgBox "This Payment belongs to the previous year, plz verify", vbInformation
                    cmdSave.Enabled = False
                End If
            End If
        
            If Rec!tnyReversed = 1 Then
                MsgBox "You cannot Edit a Cancelled Payment!", vbInformation
                Exit Sub
            End If
        End If
        If Rec!tnyVoucherTypeID <> 20 Then
            MsgBox "Can not Edit Invalid Voucher!", vbInformation
            Exit Sub
        End If
        If mStatus = 1 Then
            MsgBox "Can not Edit this Voucher!", vbInformation
            Exit Sub
        End If
        If Rec!ModuleID = 55 Then
            MsgBox "You Are not Allowed to Edit Reversed Vouchers!", vbInformation
            cmdSave.Enabled = False
            Exit Sub
        End If
        txtVoucherNo.Text = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
        txtPayOrder.Text = IIf(IsNull(Rec!intKeyID2), "", Rec!intKeyID2)
        
        If Not IsNull(Rec!intTransactionTypeID) Then
            If Not IsNull(Rec!vchTransactionType) Then
                txtTransactionType.Text = Rec!vchTransactionType
                txtTransactionType.Tag = Rec!intTransactionTypeID
            End If
        End If
        
            If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
                fraProject.Visible = True
                Check1.Visible = True
            Else
                fraProject.Visible = False
                Check1.Visible = False
            End If
            
            'Call ShowDetailsForSubCashBook
            
            Dim objTrns As New clsTransactionType
            objTrns.SetSourceOfFund (txtTransactionType.Tag)
            If Not IsEmpty(objTrns.SourceFundID) Then
                txtSourceofFund.Text = objTrns.SourceOfFund
                txtSourceofFund.Tag = objTrns.SourceFundID
            Else
                txtSourceofFund.Text = "Own Fund"
                txtSourceofFund.Tag = 4
            End If
        
        txtVoucherNo.Tag = Rec.Fields(0) 'intVoucherID
        txtDate.Text = DdMmmYy(Rec!dtDate)
        txtDate.Tag = IIf(IsNull(Rec!intTransactionID), "", Rec!intTransactionID)
        
        'txtFund.Text = IIf(IsNull(Rec!vchFund), "", Rec!vchFund)
        'txtFund.Tag = IIf(IsNull(Rec!intFundID), "", Rec!intFundID)
        txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
        txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
        txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
        txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
        
        txtInstrument.Text = IIf(IsNull(Rec!vchInstrumentType), "", Rec!vchInstrumentType)
        txtInstrument.Tag = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
        
        txtInstrumentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
        
        txtCrHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
        txtCrAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
        txtCrHeadCode.Tag = IIf(IsNull(Rec!intKeyID1), "", Rec!intKeyID1)
        
        
        Call txtCrHeadCode_LostFocus
        'txtRef.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
        txtDated.Text = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
        'dtpDueDate.Value = IIf(IsNull(Rec!dtInstrumentDate), Date, Rec!dtInstrumentDate)
        txtNarration.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
        
        
        If Not IsNull(Rec!vchName) Then
            txtName.Text = Rec!vchName
            txtPayee.Text = Rec!vchName
        Else
            txtName.Text = ""
            txtPayee.Text = ""
        End If
        On Error Resume Next
        If Not IsNull(Rec!vchInit1) Then
            txtInit1.Text = Rec!vchInit1
        Else
            txtInit1.Text = ""
        End If
        If Not IsNull(Rec!vchInit2) Then
            txtInit2.Text = Rec!vchInit2
        Else
            txtInit2.Text = ""
        End If
        If Not IsNull(Rec!vchInit3) Then
            txtInit3.Text = Rec!vchInit3
        Else
            txtInit3.Text = ""
        End If
        If Not IsNull(Rec!vchInit4) Then
            txtInit4.Text = Rec!vchInit4
        Else
            txtInit4.Text = ""
        End If
        On Error GoTo 0
        If Not IsNull(Rec!vchHouseName) Then
            txtHouse.Text = Rec!vchHouseName
        Else
            txtHouse.Text = ""
        End If
        If Not IsNull(Rec!vchStreetName) Then
            txtStreet.Text = Rec!vchStreetName
        Else
            txtStreet.Text = ""
        End If
        If Not IsNull(Rec!vchLocalPlace) Then
            txtLocalPlace.Text = Rec!vchLocalPlace
        Else
            txtLocalPlace.Text = ""
        End If
        If Not IsNull(Rec!vchMainPlace) Then
            txtMainPlace.Text = Rec!vchMainPlace
        Else
            txtMainPlace.Text = ""
        End If
        
        If Not IsNull(Rec!vchPostOffice) Then
            txtPost.Text = Rec!vchPostOffice
        Else
            txtPost.Text = ""
        End If
        If Not IsNull(Rec!vchPinNumber) Then
            txtPin.Text = Rec!vchPinNumber
        Else
            txtPin.Text = ""
        End If
        If Not IsNull(Rec!vchPhone) Then
            txtPhone.Text = Rec!vchPhone
        Else
            txtPhone.Text = ""
        End If
        
        
        
        Dim objSubLedger As New clsSubLedger
        Dim objProj As New clsProject
        
        objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intCreditorTypeID), 0, Rec!intCreditorTypeID)
        If objSubLedger.SubLedgerTypeID > 0 Then
            txtSubLedgerType.Text = objSubLedger.SubLedgerType
            txtSubLedgerType.Tag = objSubLedger.SubLedgerTypeID
            
            txtPayeeType.Text = objSubLedger.SubLedgerType
            txtPayeeType.Tag = objSubLedger.SubLedgerTypeID
            'To get the empID from Sthapana for SubCashBook
            txtName.Tag = objSubLedger.EmpID
        End If
             
        If Not IsNull(Rec!intCashBookID) Then
            objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intCashBookID), 0, Rec!intCashBookID)
            txtSubsidiaryCash.Text = IIf(IsNull(objSubLedger.Title), "", objSubLedger.Title)
            txtSubsidiaryCash.Tag = Rec!intCashBookID
        End If
        
        '--------------- added on 1.6.12 By Anisha-----'
        If txtVoucherNo.Text <> "" Then
            mSql = "SELECT faSubsidiaryCashBook.intID, faSubsidiaryCashBook.intTransferID, faSubsidiaryCashBook.intSubsidiaryAccountHeadID,faSubsidiaryCashBook.intTypeID, faSubsidiaryCashBook.intVoucherID,faVouchers.intVoucherNo From faSubsidiaryCashBook"
            mSql = mSql + " INNER JOIN faVouchers ON faSubsidiaryCashBook.intVoucherID = faVouchers.intVoucherID"
            mSql = mSql + " Where intTypeID = 50 And faVouchers.intVoucherNo = " & txtVoucherNo.Text
            mRec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adLockReadOnly
            If Not (mRec.EOF Or mRec.BOF) Then
                mintIntID = IIf(mRec!intID = "", -1, mRec!intID)
                mintTransferID = IIf(mRec!intTransferID = "", -1, mRec!intTransferID)
            End If
            mRec.Close
        Else
            mintIntID = -1
            mintTransferID = -1
        End If
        '---------------------------------------------'
        
        If Not IsNull(Rec!intImplementingOfficerID) Then
            txtImplementingOfficer.Text = IIf(Rec!intImplementingOfficerID = 0, "", Rec!vchFunctionary)
            txtImplementingOfficer.Tag = Rec!intImplementingOfficerID
        Else
            txtImplementingOfficer.Text = ""
            txtImplementingOfficer.Tag = ""
        End If
        
        If Not IsNull(Rec!decProjectID) Then
            objProj.SetProject Rec!decProjectID
            If objProj.ProjectID > 0 Then
                txtProjectNo.Text = objProj.ProjectSerialNo
                txtProjectNo.Tag = objProj.ProjectID
                txtCategory.Text = objProj.Category
            
                txtCategory.Tag = objProj.ProjCatID
                txtSector.Text = objProj.SubSector
                txtSector.Tag = objProj.SubSectorID
                objProj.FindSourceOfFund Rec!intSourceOfFundID
                txtSourceofFund.Text = objProj.SourceOfFund
                txtSourceofFund.Tag = objProj.SourceOfFundID
            End If
        Else
            fraProject.Visible = False
            txtProjectNo.Text = ""
            txtProjectNo.Tag = ""
        End If
        
        '-----------------------------------------------'
        '                 Source of Fund                '
        '-----------------------------------------------'
        If Not IsNull(Rec!intSourceOfFundID) Then
            objProj.FindSourceOfFund Rec!intSourceOfFundID
            txtSourceofFund.Text = objProj.SourceOfFund
            txtSourceofFund.Tag = objProj.SourceOfFundID
        End If
        '-----------------------------------------------'

       
        mSqlAccHeads = "Select * From faTransactionChild"
        mSqlAccHeads = mSqlAccHeads + " Inner Join faAccountHeads On faTransactionChild.intAccountHeadID=faAccountHeads.intAccountHeadID"
        mSqlAccHeads = mSqlAccHeads + " Where intTransactionID = " & txtDate.Tag
        mSqlAccHeads = mSqlAccHeads + " And intSerialNo <> 1"
        RecAccHeads.Open mSqlAccHeads, mCnn
        
        mRowCount = 1
        While Not Rec.EOF
            While Not RecAccHeads.EOF
                vsGrid.TextMatrix(mRowCount, 1) = IIf(IsNull(RecAccHeads!vchAccountHeadCode), "", RecAccHeads!vchAccountHeadCode)
                vsGrid.TextMatrix(mRowCount, 2) = IIf(IsNull(RecAccHeads!vchAccountHead), "", RecAccHeads!vchAccountHead)
                'vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!vchNarration), "", RecAccHeads!vchNarration)
                vsGrid.TextMatrix(mRowCount, 3) = IIf(IsNull(RecAccHeads!fltAmount), "", RecAccHeads!fltAmount)
                vsGrid.TextMatrix(mRowCount, 4) = IIf(IsNull(RecAccHeads!intAccountHeadID), "", RecAccHeads!intAccountHeadID)
                vsGrid.Rows = vsGrid.Rows + 1
                mRowCount = mRowCount + 1
                RecAccHeads.MoveNext
            Wend
            Rec.MoveNext
        Wend
        RecAccHeads.Close
        'Call Calculate
        
        txtCrAmount.Text = Format(CalculateAmt, "0.00")
        lblTotal.Caption = txtCrAmount.Text
        'Frame3.Enabled = False
        vsGrid.Editable = flexEDNone
        'Frame5.Enabled = False
        Frame4.Enabled = False
        cmdPaymentOrder.Enabled = False
        cmdSearchFunction.Enabled = False
        cmdSearchFunctionary.Enabled = False
        cmdSearchTransactionType.Enabled = False
        txtDate.Enabled = False
        cmdSave.Caption = "&Edit"
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Enabled = False
        End If
        
    Rec.Close
        If mViewMode = 1 Then
         cmdSave.Enabled = False
         cmdNew.Enabled = False
        End If
End Sub

 Private Function GetLastReconDate(intBankID As Integer) As Variant
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mMonthID As Integer
        Dim mFinYear As Integer
               
        mSql = "Select * From faBanks Where intAccountHeadID=" & intBankID
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSql)
            If Not (Rec.BOF And Rec.EOF) Then
                GetLastReconDate = IIf(IsNull(Rec!dtReconEndDate), Null, Rec!dtReconEndDate)
            End If
        Rec.Close
        
 End Function
Public Property Get VoucherID() As Variant
    VoucherID = intVoucherID
End Property

Public Property Let VoucherID(mData As Variant)
    intVoucherID = mData
End Property

Public Property Get VoucherNo() As Variant
    VoucherNo = intVoucherNo
End Property

Public Property Let VoucherNo(mData As Variant)
    intVoucherNo = mData
End Property

Public Property Let WaterBillPVMode(Data As Boolean)
    mWaterBillPVMode = Data
End Property

Public Property Get WaterBillPVMode() As Boolean
    WaterBillPVMode = mWaterBillPVMode
End Property

Public Property Let PayOrderNo(Data As Variant)
    mPayOrderNo = Data
End Property

Public Property Get PayOrderNo() As Variant
    PayOrderNo = mPayOrderNo
End Property
Public Property Let ViewMode(Data As Variant)
    mViewMode = Data
End Property
Public Property Get ViewMode() As Variant
    ViewMode = mViewMode
End Property

'''Private Function SaveToSubsidiaryCashBook() As Boolean
'''    On Error GoTo Err:
'''        Dim mCnn As New ADODB.Connection
'''        Dim aryIn As Variant
'''        Dim objDb As New clsDB
'''
'''        If objDb.SetConnection(mCnn) Then
'''            aryIn = Array(-1, _
'''                        12 _
'''                        , 50, _
'''                        Format(CDate(Date), "dd/mmm/yyyy"), _
'''                        gbUserID, _
'''                        gbSeatID, _
'''                        Val(txtCrHeadCode.Tag), _
'''                        Val(txtFunctionary.Tag), _
'''                        Val(txtFunction.Tag), _
'''                        Format(txtCrAmount.Text, 0#), _
'''                        Null, _
'''                        Null, _
'''                        Null, _
'''                        txtNarration.Text, Null, 0, 1, txtVoucherNo.Text)
'''            objDb.ExecuteSP "spSaveSubsidiaryCashBook", aryIn, , , mCnn, adCmdStoredProc
'''        Else
'''            MsgBox "Connection to Finance does not exits, Please contact your System Administrator", vbInformation
'''        End If
'''    Exit Function
'''Err:
'''    MsgBox (Error$)
'''End Function
    Private Function CreditHead(intSourceID As Integer) As Integer
        ''to Automate Receipt Credit Head according to Source of Fund
        If gbLBPanchayat = 1 Then
            Select Case intSourceID
                        Case 1 'DEVELOPMENT FUND - General
                           CreditHead = 1065
                        Case 10 'RECEIPTS FROM Other LSIG's-GRAMA PANCHAYAT
                           CreditHead = 1103
                        Case 11 'RECEIPTS FROM Other LSIG's-BLOCK PANCHAYAT
                           CreditHead = 1104
                        Case 12 'RECEIPTS FROM Other LSIG's-DISTRICT PANCHAYAT
                           CreditHead = 1105
                        Case 13 'RECEIPTS FROM Other LSIG's-MUNICIPALITIES
                           CreditHead = 1101
                        Case 14 'RECEIPTS FROM Other LSIG's-MUNICIPAL CORPORATIONS
                           CreditHead = 1102
                        Case 16 'MAINTENANCE FUND- ROAD ASSESTS
                            CreditHead = 1667
                        Case 17 'MAINTENANCE FUND- NON ROAD ASSESTS
                            CreditHead = 1668
                        Case 21 'BEST PANCHAYAT
                            CreditHead = 1688
                        Case 26, 41 'KLGSDP GRANT
                            CreditHead = 1612
                        Case 27 'SPECIAL GRANT
                            CreditHead = 1617
                        Case 28 'ROAD RENOVATION
                            CreditHead = 1618
                        Case 29 'DEVELOPMENT FUND - SCP
                            CreditHead = 1066
                        Case 30 'DEVELOPMENT FUND - TSP
                            CreditHead = 1067
                        Case 25 'CFC GRANT
                            CreditHead = 1068
                        Case Else
                            CreditHead = 0
                End Select
        Else
            Select Case intSourceID
                        Case 1 'DEVELOPMENT FUND - General
                           CreditHead = 928
                        Case 16 'MAINTENANCE FUND- ROAD ASSESTS
                            CreditHead = 2347
                        Case 17 'MAINTENANCE FUND- NON ROAD ASSESTS
                            CreditHead = 2348
                        Case 26, 41 'KLGSDP GRANT
                            CreditHead = 1762
                        Case 27 'SPECIAL GRANT
                            CreditHead = 1763
                        Case 28 'ROAD RENOVATION
                            CreditHead = 1764
                        Case 29 'DEVELOPMENT FUND - SCP
                            CreditHead = 929
                        Case 30 'DEVELOPMENT FUND - TSP
                            CreditHead = 930
                        Case 25 'CFC GRANT
                            CreditHead = 931
                        Case Else
                            CreditHead = 0
                End Select
        End If
    End Function

    Private Sub GenerateReceipt(mintPVoucherNo As Variant)  'FUNCTION TO RENERATE RECEIPT FOR ALLOTMENT PAYMENT
        Dim mV              As uVoucher
        Dim mVC             As uVChild
        Dim mVA             As uVoucherAddress
        Dim mVS             As uVoucherSub
        Dim mT              As uTr
        Dim mTC             As uTrChild
        Dim arrInput        As Variant
        Dim arrOutPut       As Variant
        Dim mintVoucherID   As Variant
        Dim mintVoucherNo   As Variant
        Dim mintTransactionID As Variant
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim mLoop           As Integer
        Dim mSql            As String
        Dim mYearID         As Integer
        Dim mDate           As Date
        Dim mCreditHeadID   As Integer
        
        mCreditHeadID = CreditHead(val(txtSourceofFund.Tag))
        If mCreditHeadID = 0 Then
            mCreditHeadID = val(vsGrid.TextMatrix(mLoop, 4))
        End If
        
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        With mV
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = val(txtTransactionType.Tag)
            .tnyVoucherTypeID_5 = 10
            .intVoucherNo_6 = IIf(txtVoucherNo.Text = "", Null, txtVoucherNo.Text)
            .intBookNo_7 = Null
            If mPreYearMode = 0 Then
                .dtDate_8 = gbTransactionDate
                mDate = gbTransactionDate
            Else
                .dtDate_8 = Format(txtDate.Text, "dd/mmm/yyyy")
                mDate = .dtDate_8
            End If
            .fltAmount_9 = val(txtCrAmount)
            .intInstrumentTypeID_10 = 1 'val(txtInstrument.Tag)
            .vchInstrumentNo_11 = Trim(txtInstrumentNo)
            .dtInstrumentDate_12 = IIf(IsDate(txtDated), txtDated.Text, gbTransactionDate)
            .vchDescription_13 = Trim(txtNarration)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = val(txtCrHeadCode.Tag)
            .intKeyID2_23 = txtPayOrder.Text
            .intExternalApplicationID_24 = Null
            .intExternalModuleID_25 = 1
            If mPreYearMode = 0 Then
                .intFinancialYearID_26 = gbFinancialYearID
                mYearID = gbFinancialYearID
            Else
                .intFinancialYearID_26 = gbFinancialYearID - 1
                mYearID = gbFinancialYearID - 1
            End If
            .tnyShiftID_27 = gbShiftID
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = 0
            .vchBank_33 = txtNameOfBank.Text
            .vchBankPlace_34 = txtBranch.Text
            .intFundID_35 = 1
            .numSeatID = val(txtSeat.Tag)
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = gbLocationID
            
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intTransactionID_3, _
            .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
            .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
            .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
            .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
            .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
            .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, _
            .tnyShiftID_27, .tnyPrintFlag_28, _
            .tnyCancelFlag_29, .vchBank_33, _
            .vchBankPlace_34, .intFundID_35, _
            .numSeatID, .intSessionID, _
            .vchRefNo, .fltRoundOff, _
            .fltAdvAmtAdj, .numInwardNo, _
            .tnyStatus_32, .numLocationID)
        
        '=============================================='
        ' T r a n s a c t i o n   B e g i n            '
        '=============================================='
        mCnn.BeginTrans
        '=============================================='
            objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
            mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mintVoucherNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            End If
            Rec.Close
        End If
        
        With mVA
            .intVoucherID = mintVoucherID
            .intLocalBodyID = gbLocalBodyID
            .vchName = Trim(txtName)
            .vchInit1 = Trim(txtInit1)
            .vchInit2 = Trim(txtInit2)
            .vchInit3 = Trim(txtInit3)
            .vchInit4 = Trim(txtInit4)
            .vchHouseName = Trim(txtHouse)
            .vchStreetName = Trim(txtStreet)
            .vchLocalPlace = Trim(txtLocalPlace)
            .vchMainPlace = Trim(txtMainPlace)
            .vchPostOffice = Trim(txtPost)
            .vchDistrict = Null
            .vchPinNumber = Trim(txtPin)
            .vchPhone = Trim(txtPhone)
            .intWardNo = Null
            .intDoorNo = Null
            .vchDoorNo2 = Null
        
            arrInput = Array(.intVoucherID, _
                .intLocalBodyID, _
                .vchName, _
                .vchInit1, _
                .vchInit2, _
                .vchInit3, _
                .vchInit4, _
                .vchHouseName, _
                .vchStreetName, _
                .vchLocalPlace, _
                .vchMainPlace, _
                .vchPostOffice, _
                .vchDistrict, _
                .vchPinNumber, _
                .vchPhone, _
                .intWardNo, _
                .intDoorNo, _
                .vchDoorNo2)
                
            objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        With mVS
            .intVoucherID = mintVoucherID
            .intLocalBodyID = gbLocalBodyID
            .decProjectID = IIf(val(txtProjectNo.Tag) > 0, val(txtProjectNo.Tag), Null)
            .intSourceOfFundID = IIf(val(txtSourceofFund.Tag) > 0, val(txtSourceofFund.Tag), Null)
            .intCategoryID = IIf(val(txtCategory.Tag) > 0, val(txtCategory.Tag), Null)
            .intSectorID = IIf(val(txtSector.Tag) > 0, val(txtSector.Tag), Null)
            .intAllotmentID = IIf(val(txtAllotmentLetterNo.Tag) > 0, val(txtAllotmentLetterNo.Tag), Null)
            .intAgreementID = IIf(val(txtAgreementNo.Tag) > 0, val(txtAgreementNo.Tag), Null)
            .intCashBookID = IIf(val(txtSubsidiaryCash.Tag) > 0, val(txtSubsidiaryCash.Tag), Null)
            .intImplementingOfficerID = IIf(val(txtImplementingOfficer.Tag) > 0, val(txtImplementingOfficer.Tag), Null)
            .intCreditorTypeID = IIf(val(txtSubLedgerType.Tag) > 0, val(txtSubLedgerType.Tag), Null)
            .intCreditorsID = IIf(val(txtName.Tag) > 0, val(txtName.Tag), Null)
            .intTypeID = Null
            
            arrInput = Array(.intVoucherID, _
            .intLocalBodyID, _
            .decProjectID, _
            .intSourceOfFundID, _
            .intCategoryID, _
            .intSectorID, _
            .intAllotmentID, _
            .intAgreementID, _
            .intCashBookID, _
            .intImplementingOfficerID, _
            .intCreditorTypeID, _
            .intCreditorsID, _
            .intTypeID)
            objdb.ExecuteSP "spSaveVoucherSub", arrInput, , , mCnn, adCmdStoredProc
        End With
        
    
        With mT
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            If mPreYearMode = 0 Then
                .intFinancialYearID = gbFinancialYearID
                .dtTransactionDate = gbTransactionDate
            Else
                .intFinancialYearID = gbFinancialYearID - 1
                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = val(txtFunction.Tag)
            .intFunctionaryID = val(txtFunctionary.Tag)
            .intFieldID = Null
            .intFundID = 1
            .intBudgetCentreID = Null
            .vchNarration = Trim(txtNarration)
            .intTransactionTypeID = val(txtTransactionType.Tag)
            .intProcessID = Null
            .vchGroup = "R"
            .intGroupID = 10
            .intKeyID = Null
            .numSubLedgerID = IIf(val(txtName.Tag) > 0, val(txtName.Tag), Null)
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
            objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTC
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = val(txtCrHeadCode.Tag)
            .fltAmount = Format(val(txtCrAmount), "0.00")
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = Trim(txtNarration)
            .intFundID = 1
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        
        'For mLoop = 1 To vsGrid.Rows - 1
           ' If vsGrid.TextMatrix(mLoop, 4) <> "" Then
                With mVC
                    .intVoucherID_1 = mintVoucherID
                    .intLocalBodyID_2 = gbLocalBodyID
                    .intSlNo_3 = 1
        
                    .intAccountHeadID_4 = mCreditHeadID  'val(vsGrid.TextMatrix(mLoop, 4))
                    .tnyDebitOrCredit_5 = 0
                    .intYearID_6 = Null
                    .tnyPeriodID_7 = Null
                    .tnyArrearFlag_8 = Null
                    .numDemandID_9 = Null
                    .fltAmount_10 = Format(val(txtCrAmount), "0.00") 'Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
        
                    arrInput = Array(.intVoucherID_1, _
                    .intLocalBodyID_2, _
                    .intSlNo_3, _
                    .intAccountHeadID_4, _
                    .tnyDebitOrCredit_5, _
                    .intYearID_6, _
                    .tnyPeriodID_7, _
                    .tnyArrearFlag_8, _
                    .numDemandID_9, _
                    .fltAmount_10)
                    objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn, adCmdStoredProc
                End With
        
                With mTC
                    .intTransactionID = mintTransactionID
                    .intSerialNo = 2
                    .intAccountHeadID = mCreditHeadID 'val(vsGrid.TextMatrix(mLoop, 4))
                    .fltAmount = Format(val(txtCrAmount), "0.00") 'Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
                    .tinDebitOrCreditFlag = 0
                    .intByAccountHeadID = val(txtCrHeadCode.Tag)
                    .vchNarration = Null
                    .intFundID = 1
        
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
                End With
           ' Else
           '     Exit For
           ' End If
        'Next
        
        
        '=============================================='
        ' T r a n s a c t i o n   C o m m i t t i n g
        '=============================================='
        mCnn.CommitTrans
        '=============================================='
        
    End Sub


    Private Sub GenerateEbillReceipt()  'FUNCTION TO RENERATE RECEIPT FOR E Bill Payment
        Dim mV              As uVoucher
        Dim mVC             As uVChild
        Dim mVA             As uVoucherAddress
        Dim mVS             As uVoucherSub
        Dim mT              As uTr
        Dim mTC             As uTrChild
        Dim arrInput        As Variant
        Dim arrOutPut       As Variant
        Dim mintVoucherID   As Variant
        Dim mintVoucherNo   As Variant
        Dim mintTransactionID As Variant
        Dim objdb           As New clsDB
        Dim mCnn            As New ADODB.Connection
        Dim Rec             As New ADODB.Recordset
        Dim mLoop           As Integer
        Dim mSql            As String
        Dim mYearID         As Integer
        Dim mDate           As Date
        Dim mCreditHeadID   As Integer
        

        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mCreditHeadID = val(txtCreditHdIDR.Tag)
        With mV
            
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = val(txtTransactionType.Tag)
            .tnyVoucherTypeID_5 = 10
            .intVoucherNo_6 = Null 'IIf(txtVoucherNo.Text = "", Null, txtVoucherNo.Text)
            .intBookNo_7 = Null
'            If mPreYearMode = 0 Then
'                .dtDate_8 = gbTransactionDate
'                mDate = gbTransactionDate
'            Else
                .dtDate_8 = Format(txtDate.Text, "dd/mmm/yyyy")
                mDate = .dtDate_8
'            End If
            .fltAmount_9 = val(txtCrAmount)
            .intInstrumentTypeID_10 = 1 'val(txtInstrument.Tag)
            .vchInstrumentNo_11 = Trim(txtInstrumentNo)
            .dtInstrumentDate_12 = "" 'IIf(IsDate(txtDated), txtDated.Text, gbTransactionDate)
            .vchDescription_13 = Trim(txtNarration)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = txtBillControCodeID.Tag
            .intKeyID1_22 = val(txtCrHeadCode.Tag)
            .intKeyID2_23 = txtBillControCodeID.Text  'txtPayOrder.Text
            .intExternalApplicationID_24 = 118
            .intExternalModuleID_25 = 1
            If mPreYearMode = 0 Then
                .intFinancialYearID_26 = gbFinancialYearID
                mYearID = gbFinancialYearID
            Else
                .intFinancialYearID_26 = gbFinancialYearID - 1
                mYearID = gbFinancialYearID - 1
            End If
            .tnyShiftID_27 = gbShiftID
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = 0
            .vchBank_33 = "" 'txtNameOfBank.Text
            .vchBankPlace_34 = "" 'txtBranch.Text
            .intFundID_35 = 1
            .numSeatID = val(txtSeat.Tag)
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = gbLocationID
            
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intTransactionID_3, _
            .intTransactionTypeID_4, .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, _
            .dtDate_8, .fltAmount_9, .intInstrumentTypeID_10, _
            .vchInstrumentNo_11, .dtInstrumentDate_12, .vchDescription_13, .numZoneID_14, _
            .numWardID_15, .intDoorNoP1_16, .vchDoorNoP2_17, .vchDoorNoP3_18, _
            .intUserID_19, .intCounterID_20, .numSubLedgerID_21, .intKeyID1_22, _
            .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, _
            .tnyShiftID_27, .tnyPrintFlag_28, _
            .tnyCancelFlag_29, .vchBank_33, _
            .vchBankPlace_34, .intFundID_35, _
            .numSeatID, .intSessionID, _
            .vchRefNo, .fltRoundOff, _
            .fltAdvAmtAdj, .numInwardNo, _
            .tnyStatus_32, .numLocationID)
        
        '=============================================='
        ' T r a n s a c t i o n   B e g i n            '
        '=============================================='
        mCnn.BeginTrans
        '=============================================='
            objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
            mSql = "Select intVoucherNo From faVouchers Where intVoucherID = " & mintVoucherID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mintVoucherNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
            End If
            Rec.Close
        End If
        
        With mVA
            .intVoucherID = mintVoucherID
            .intLocalBodyID = gbLocalBodyID
            .vchName = Trim(txtName)
            .vchInit1 = Trim(txtInit1)
            .vchInit2 = Trim(txtInit2)
            .vchInit3 = Trim(txtInit3)
            .vchInit4 = Trim(txtInit4)
            .vchHouseName = Trim(txtHouse)
            .vchStreetName = Trim(txtStreet)
            .vchLocalPlace = Trim(txtLocalPlace)
            .vchMainPlace = Trim(txtMainPlace)
            .vchPostOffice = Trim(txtPost)
            .vchDistrict = Null
            .vchPinNumber = Trim(txtPin)
            .vchPhone = Trim(txtPhone)
            .intWardNo = Null
            .intDoorNo = Null
            .vchDoorNo2 = Null
        
            arrInput = Array(.intVoucherID, _
                .intLocalBodyID, _
                .vchName, _
                .vchInit1, _
                .vchInit2, _
                .vchInit3, _
                .vchInit4, _
                .vchHouseName, _
                .vchStreetName, _
                .vchLocalPlace, _
                .vchMainPlace, _
                .vchPostOffice, _
                .vchDistrict, _
                .vchPinNumber, _
                .vchPhone, _
                .intWardNo, _
                .intDoorNo, _
                .vchDoorNo2)
                
            objdb.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn, adCmdStoredProc
        End With
    
        With mT
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
'            If mPreYearMode = 0 Then
                .intFinancialYearID = gbFinancialYearID
                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
'            Else
'                .intFinancialYearID = gbFinancialYearID - 1
'                .dtTransactionDate = Format(txtDate.Text, "dd/mmm/yyyy")
'            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = val(txtFunction.Tag)
            .intFunctionaryID = val(txtFunctionary.Tag)
            .intFieldID = Null
            .intFundID = 1
            .intBudgetCentreID = Null
            .vchNarration = Trim(txtNarration)
            .intTransactionTypeID = val(txtTransactionType.Tag)
            .intProcessID = Null
            .vchGroup = "R"
            .intGroupID = 10
            .intKeyID = Null
            .numSubLedgerID = IIf(val(txtName.Tag) > 0, val(txtName.Tag), Null)
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
            objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTC
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = val(txtCrHeadCode.Tag)
            .fltAmount = Format(val(txtCrAmount), "0.00")
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = Trim(txtNarration)
            .intFundID = 1
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        
        'For mLoop = 1 To vsGrid.Rows - 1
           ' If vsGrid.TextMatrix(mLoop, 4) <> "" Then
                With mVC
                    .intVoucherID_1 = mintVoucherID
                    .intLocalBodyID_2 = gbLocalBodyID
                    .intSlNo_3 = 1
        
                    .intAccountHeadID_4 = mCreditHeadID  'val(vsGrid.TextMatrix(mLoop, 4))
                    .tnyDebitOrCredit_5 = 0
                    .intYearID_6 = Null
                    .tnyPeriodID_7 = Null
                    .tnyArrearFlag_8 = Null
                    .numDemandID_9 = Null
                    .fltAmount_10 = Format(val(txtCrAmount), "0.00") 'Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
        
                    arrInput = Array(.intVoucherID_1, _
                    .intLocalBodyID_2, _
                    .intSlNo_3, _
                    .intAccountHeadID_4, _
                    .tnyDebitOrCredit_5, _
                    .intYearID_6, _
                    .tnyPeriodID_7, _
                    .tnyArrearFlag_8, _
                    .numDemandID_9, _
                    .fltAmount_10)
                    objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn, adCmdStoredProc
                End With
        
                With mTC
                    .intTransactionID = mintTransactionID
                    .intSerialNo = 2
                    .intAccountHeadID = mCreditHeadID 'val(vsGrid.TextMatrix(mLoop, 4))
                    .fltAmount = Format(val(txtCrAmount), "0.00") 'Format(val(vsGrid.TextMatrix(mLoop, 3)), "0.00")
                    .tinDebitOrCreditFlag = 0
                    .intByAccountHeadID = val(txtCrHeadCode.Tag)
                    .vchNarration = Null
                    .intFundID = 1
        
                    arrInput = Array(.intTransactionID, _
                    .intSerialNo, _
                    .intAccountHeadID, _
                    .fltAmount, _
                    .tinDebitOrCreditFlag, _
                    .intByAccountHeadID, _
                    .vchNarration, _
                    .intFundID)
                    objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn, adCmdStoredProc
                End With
        
        If mPreYearMode = 1 Then
            objdb.ExecuteSP "Update faWebExtracts set intExtractTypeID=1,numKeyID=" & mintVoucherID & " ,tnyPendingTask=1,dtPendingDate=getdate() Where intWebExtractID=" & txtBillControCodeID.Tag, , , , mCnn, adCmdText
        Else
            objdb.ExecuteSP "Update faWebExtracts Set intExtractTypeID=1,numKeyID=" & mintVoucherID & " Where intWebExtractID=" & txtBillControCodeID.Tag, , , , mCnn, adCmdText
        End If
        '=============================================='
        ' T r a n s a c t i o n   C o m m i t t i n g
        '=============================================='
        mCnn.CommitTrans
        '=============================================='
        
    End Sub


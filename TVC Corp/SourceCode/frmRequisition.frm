VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmRequisition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requisition for Fund by Implementing Officer"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Height          =   420
      Left            =   1800
      TabIndex        =   74
      Top             =   7890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAuthorizationNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7695
      TabIndex        =   4
      Top             =   990
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   12480
      TabIndex        =   62
      Top             =   0
      Width           =   12480
      Begin VB.Label lblLabelstat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition is disabled 2017-18 onwords for e-bill submission"
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
         Height          =   195
         Left            =   1290
         TabIndex        =   91
         Top             =   180
         Width           =   5160
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   72
         Top             =   105
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   10800
         Picture         =   "frmRequisition.frx":0000
         Top             =   -15
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is to Record Requisition submitted by Implementing Officers including Secretary for release of fund."
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
         Height          =   195
         Left            =   1275
         TabIndex        =   71
         Top             =   480
         Width           =   8865
      End
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "Approve"
      Height          =   420
      Left            =   570
      TabIndex        =   73
      Top             =   7890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   12420
      TabIndex        =   67
      Top             =   7500
      Width           =   12480
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   70
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   69
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   68
         Top             =   0
         Width           =   1215
      End
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   -3195
         Top             =   375
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
   End
   Begin VB.Frame fraAllotments 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1FDFD&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6810
      Left            =   -45
      TabIndex        =   63
      Top             =   720
      Width           =   12495
      Begin VB.Frame frmNatureOfClaim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NATURE OF CLAIM"
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
         Height          =   975
         Left            =   10200
         TabIndex        =   89
         Top             =   5760
         Width           =   2175
         Begin VB.TextBox txtNatureofClaim 
            Height          =   615
            Left            =   120
            TabIndex        =   90
            ToolTipText     =   "For Printing In TR-59(C)"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtAmtIssued 
         Height          =   285
         Left            =   10650
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtTreasuryBalance 
         Height          =   285
         Left            =   10650
         TabIndex        =   87
         Top             =   450
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkNewMode 
         Caption         =   "New Mode(2015)"
         Height          =   210
         Left            =   10650
         TabIndex        =   86
         Top             =   225
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton cmdMicroSector 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9780
         TabIndex        =   85
         Top             =   4095
         Width           =   300
      End
      Begin VB.CommandButton cmdSubSector 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9765
         TabIndex        =   84
         Top             =   3780
         Width           =   300
      End
      Begin VB.TextBox txtSubSector 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   82
         Top             =   3780
         Width           =   6555
      End
      Begin VB.TextBox txtMicroSector 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   80
         Top             =   4095
         Width           =   6555
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9765
         TabIndex        =   79
         Top             =   3105
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtGo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   3105
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox chkUnAuthDrw 
         BackColor       =   &H00F1FDFD&
         Caption         =   "UnAuthorized Drawal"
         Height          =   195
         Left            =   5130
         TabIndex        =   76
         Top             =   2475
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CheckBox chkBFund 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F1FDFD&
         Caption         =   "B-Fund"
         Height          =   195
         Left            =   2505
         TabIndex        =   15
         Top             =   1485
         Width           =   900
      End
      Begin VB.TextBox txtAuthorizedAmt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   2070
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.TextBox txtScheme 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1725
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.CommandButton cmdSearchScheme 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9765
         TabIndex        =   18
         Top             =   1725
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox MaskDetailAccHead 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3210
         TabIndex        =   59
         Top             =   6450
         Width           =   3060
      End
      Begin VB.TextBox MaskAccHead 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3210
         TabIndex        =   55
         Top             =   6120
         Width           =   3060
      End
      Begin VB.TextBox txtStateHead 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6285
         TabIndex        =   56
         Top             =   6120
         Width           =   3465
      End
      Begin VB.TextBox txtStateSubHead 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6285
         TabIndex        =   60
         Top             =   6450
         Width           =   3465
      End
      Begin VB.CommandButton cmdSearchStateSubHead 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   61
         Top             =   6480
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchStateHead 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   57
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtProjName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2730
         Width           =   4530
      End
      Begin VB.TextBox txtProjCost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3435
         Width           =   1425
      End
      Begin VB.TextBox txtDDOCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtTreasuryCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3210
         MaxLength       =   15
         TabIndex        =   51
         Top             =   5820
         Width           =   1425
      End
      Begin VB.TextBox txttreasury 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4650
         TabIndex        =   52
         Top             =   5820
         Width           =   5100
      End
      Begin VB.CommandButton cmdSearchTreasury 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   53
         Top             =   5760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtDPCDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         TabIndex        =   39
         Top             =   4440
         Width           =   1425
      End
      Begin VB.TextBox txtDPCNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3210
         TabIndex        =   37
         Top             =   4440
         Width           =   3000
      End
      Begin VB.CheckBox chkNonPlan 
         BackColor       =   &H00F1FDFD&
         Caption         =   "Non -Plan"
         Height          =   195
         Left            =   3990
         TabIndex        =   23
         Top             =   2460
         Width           =   1110
      End
      Begin VB.CheckBox chkPlan 
         BackColor       =   &H00F1FDFD&
         Caption         =   "Plan"
         Height          =   195
         Left            =   3195
         TabIndex        =   22
         Top             =   2460
         Width           =   720
      End
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1140
         Width           =   6555
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   4200
      End
      Begin VB.TextBox txtProjectNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2730
         Width           =   1995
      End
      Begin VB.CommandButton cmdSearchProject 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9765
         TabIndex        =   29
         Top             =   2730
         Width           =   300
      End
      Begin VB.ComboBox cmbCategory 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3420
         Width           =   4245
      End
      Begin VB.TextBox txtAmountRequested 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         MaxLength       =   9
         TabIndex        =   20
         Top             =   2040
         Width           =   2010
      End
      Begin VB.TextBox txtRequisitiontDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   195
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtInstalmentNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2415
         Width           =   1425
      End
      Begin VB.TextBox txtRequisition 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3195
         TabIndex        =   1
         Top             =   195
         Width           =   1620
      End
      Begin VB.CommandButton cmdSearchAllotmentLetter 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4845
         TabIndex        =   2
         Top             =   195
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtIMPOName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   7
         Top             =   540
         Width           =   6555
      End
      Begin VB.CommandButton cmdSearchIMPO 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9780
         TabIndex        =   8
         Top             =   540
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9765
         TabIndex        =   42
         Top             =   4770
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtAccountHeadCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   5370
         Width           =   1290
      End
      Begin VB.CommandButton cmdSearchHead 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9765
         TabIndex        =   49
         Top             =   5385
         Width           =   300
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4455
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   5370
         Width           =   5295
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5070
         Width           =   6555
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4770
         Width           =   6555
      End
      Begin VB.CommandButton cmdSearchFunction 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9765
         TabIndex        =   45
         Top             =   5085
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpAllotmentDate 
         Height          =   315
         Left            =   9765
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60948481
         CurrentDate     =   40087
      End
      Begin VB.ComboBox cmbSource 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3075
         Width           =   4245
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Sector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2115
         TabIndex        =   83
         Top             =   3780
         Width           =   930
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Micro Sector"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1980
         TabIndex        =   81
         Top             =   4095
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8010
         TabIndex        =   78
         Top             =   3105
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblScheme 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scheme/Programe"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1530
         TabIndex        =   16
         Top             =   1725
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label lblAuthorizationNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization No"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6225
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label lblAuthorizedAmt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Authorized"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5985
         TabIndex        =   21
         Top             =   2055
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proj.Cost"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7440
         TabIndex        =   34
         Top             =   3435
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDO Code"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7440
         TabIndex        =   11
         Top             =   885
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detailed Head of Account (Apx. IV)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   58
         Top             =   6480
         Width           =   3060
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head of Account(State Budget)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   465
         TabIndex        =   54
         Top             =   6150
         Width           =   2715
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treasury"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2370
         TabIndex        =   50
         Top             =   5805
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7785
         TabIndex        =   38
         Top             =   4455
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sanction No "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2070
         TabIndex        =   36
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2085
         TabIndex        =   13
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label lblProjectNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2475
         TabIndex        =   26
         Top             =   2730
         Width           =   630
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2505
         TabIndex        =   30
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label lblAmountInFigures 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Requested"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1455
         TabIndex        =   19
         Top             =   2055
         Width           =   1650
      End
      Begin VB.Label lblAllotmentDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7740
         TabIndex        =   66
         Top             =   345
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblAllotmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1875
         TabIndex        =   0
         Top             =   195
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dfdsfdsf"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -585
         TabIndex        =   65
         Top             =   570
         Width           =   60
      End
      Begin VB.Label lblInstalmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instalment No"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7065
         TabIndex        =   24
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2325
         TabIndex        =   32
         Top             =   3435
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Implementing Officer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   465
         TabIndex        =   6
         Top             =   555
         Width           =   2670
      End
      Begin VB.Label lblAccountHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1905
         TabIndex        =   46
         Top             =   5385
         Width           =   1200
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2355
         TabIndex        =   43
         Top             =   5100
         Width           =   750
      End
      Begin VB.Label lblFunctionary 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2085
         TabIndex        =   40
         Top             =   4800
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2085
         TabIndex        =   9
         Top             =   870
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim ReqID   As Variant
    Dim mCategoryID As Integer
    Dim mCategory As Variant
    Dim mFundSlNo As Variant
    Dim mFundErSulekha As Integer
    Dim mPreviousYearMode As Integer
    Dim mPreviousYearRequestID As Variant
    Dim mTokenID As Variant
    Dim mReqInboxID As Long
    Dim mWHERE  As String
    Dim fltClosingTreasuryBalance As Variant

    Dim mLoadMode       As Integer '10-For UNAUTHORIZED DRAWAL '20-REQUISITION INBOX
    Dim mNewProcessFlag As Boolean ' 2015-16 Drawal from Consolidated Fund
    
    
    Public Sub GetRequisitionInboxDetails() '******************REQUISITION INBOX******************
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim msql        As String
        Dim Rec         As New ADODB.Recordset
        Dim objProj     As New clsProject
        Dim objSubLedger As New clsSubLedger
        
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            msql = " SELECT * FROM faRequisitionInbox "
            msql = msql + " Where intTockenNo = " & mTokenID
            Rec.Open msql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                cmbSource.Tag = IIf(IsNull(Rec!intSourceID), 0, Rec!intSourceID)
                cmbCategory.Tag = IIf(IsNull(Rec!intFundCategoryID), 0, Rec!intFundCategoryID)
                txtProjectNo.Tag = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
                txtAmountRequested.Text = IIf(IsNull(Rec!fltRequestedAmt), 0, Rec!fltRequestedAmt)
                
                'txtTreasuryCode.Text = IIf(IsNull(Rec!fltRequestedAmt), 0, Rec!fltRequestedAmt)
                'txttreasury

                objSubLedger.SetSubLedgerDetails (Rec!intImplementingOfficersID)
                txtIMPOName.Tag = IIf(IsNull(objSubLedger.SubsidiaryAccountHeadID), 0, objSubLedger.SubsidiaryAccountHeadID)
                txtIMPOName.Text = IIf(IsNull(objSubLedger.NameOfSubLedger), "", objSubLedger.NameOfSubLedger)
                txtDesig.Text = IIf(IsNull(objSubLedger.Designation), "", objSubLedger.Designation)
                txtDept.Text = objSubLedger.Department

                objProj.SetProject txtProjectNo.Tag, Rec!intFinancialYearID
                Call SetProjectDetails(objProj, val(cmbSource.Tag))
                Call SetAccountHeads(IIf(IsNull(Rec!intMircoSectorID), 0, Rec!intMircoSectorID))
                
            End If
        End If
    End Sub
    
    'Set mPreviousYearMode = 0
    Private Sub GetPreviousYearRequestDetails()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim msql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mRec         As New ADODB.Recordset
        Dim mSourceID   As Variant
        Dim mProjectID  As Variant
        Dim mCategoryID As Integer
        Dim mSubsectorID As Variant
        Dim objProj     As New clsProject
        Dim objProFund  As New clsProjectFund
        Dim mCol        As Collection
        Dim mRow        As Integer
        Dim mTaskID     As Integer
        Dim mYearID     As Integer

        'On Error GoTo Err
        
        If mPreviousYearRequestID > 0 Then
            If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                msql = "Select * from faPendingTaskRequest Where intRequestID= " & mPreviousYearRequestID
                Rec.Open msql, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!intKeyID > 0 Then
                        RequisitionID = Rec!intKeyID
                        FetchRequisitionDetails
                        Exit Sub
                        
                    End If
                    'KLGSDP Fund Modified on 17/4/17
                     mSourceID = IIf(IsNull(Rec!intSourceOfFundID), 0, Rec!intSourceOfFundID)
                     cmbSource.Tag = mSourceID
                   
                    If mSourceID = 26 And gbFinancialYearID - 1 = 2016 Then
                        msql = "SELECT vchSourceFundName,intSourceFundID from suSourceOfFund Where intSourceFundID in (26,41)"
                        PopulateList cmbSource, msql, , False, True, True, enuSourceString.Saankhya
                        cmbSource.Tag = mSourceID
                        cmbSource.Enabled = True
                    End If
                    
                    mCategoryID = IIf(IsNull(Rec!intCategoryID), 0, Rec!intCategoryID)
                    cmbCategory.Tag = mCategoryID
                    mProjectID = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
                    txtAmountRequested.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    txtRequisitiontDate.Text = DdMmmYy(Rec!dtTransactionDate)
                    txtAmountRequested.Enabled = True
                 '   txtAmountRequested.SetFocus
                    mTaskID = IIf(IsNull(Rec!intTaskID), 0, Rec!intTaskID)
                    txtRequisition.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    'Call txtAmountRequested_LostFocus
                End If
                Rec.Close
            End If
            
            txtAmountRequested.Enabled = False
            'txtProjectNo.Enabled = False
               
            If mTaskID = 3 Then
               chkPlan.value = 1
               objProj.SetProject mProjectID, gbFinancialYearID - 1
               If objProj.ProjectID > 0 Then
                
                   mYearID = objProj.YearID
                   If objProj.YearID > 2012 Then
                   If mSourceID = 26 And objProj.YearID = 2016 Then
                       Call SetProjectDetails(objProj, val(mSourceID))
                    Else
              
                       Call SetProjectDetails(objProj, val(mSourceID))
                    End If
                        GoTo VALIDATION:
                   End If
                   txtProjName.Text = objProj.ProjectNameEnglish
                   txtProjectNo.Text = objProj.ProjectSerialNo
                   txtProjectNo.Tag = objProj.ProjectID
                   cmbCategory.Tag = objProj.ProjCatID
                   cmbSource.Text = objProj.FindSourceOfFund(mSourceID)
                    If mSourceID = 26 And objProj.YearID = 2016 Then
                        cmbSource.Enabled = True
                    Else
                        cmbSource.Enabled = False
                    End If
                   mSubsectorID = objProj.SubSectorID
                   
                   Set mCol = objProj.GetFundDetails(CInt(gbFinancialYearID - 1), objProj.ProjectID)
                   For mRow = 1 To mCol.count
                       Set objProFund = mCol.Item(mRow)
                       If objProFund.SourceOfFundID = mSourceID Then
                           txtProjCost.Text = objProFund.SourceWiseAmount
                           Exit For
                       End If
                   Next mRow
               End If
               txtProjectNo.Enabled = False
               
               
               
               Dim mCnPlan As New ADODB.Connection
               'Dim mWHERE  As String
    
               Dim mCapitalExpFlag As Boolean
               Dim mMicroSectorCount As Integer
               Dim mMicroHeads As Integer
               
               msql = " SELECT faSubSectorHeads.intCategoryID,vchTransactionCategory, intSubSectorID, vchSubSectorCode, vchSubSector, "
               msql = msql + " faSubSectorHeads.intAccountHeadID, faAccountHeads.vchAccountHeadCode, vchAccountHead, "
               msql = msql + " faSubSectorHeads.intFunctionID, faFunctions.vchFunctionCode, vchFunction, "
               msql = msql + " faFunctionaryFunctions.intFunctionaryID, vchFunctionary, "
               msql = msql + " faSubSectorHeads.intTransactionTypeID , vchTransactionType "
               msql = msql + " FROM faSubSectorHeads "
               msql = msql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faSubSectorHeads.intAccountHeadID "
               msql = msql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faSubSectorHeads.intFunctionID "
               msql = msql + " LEFT JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
               msql = msql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID "
               msql = msql + " INNER JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faSubSectorHeads.intCategoryID"
               msql = msql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faSubSectorHeads.intTransactionTypeID "
               msql = msql + " Where faSubSectorHeads.intSubSectorID = " & mSubsectorID & " And faSubSectorHeads.intCategoryID = " & val(cmbCategory.Tag)
               
               If objDb.SetConnection(mCnn) Then
                   Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly
                   If Not (Rec.EOF And Rec.BOF) Then
                       cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                       cmbCategory.Enabled = False
                       txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                       txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                       cmdSearchFunction.Enabled = False
                       txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                       txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                       txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                       cmdSearchHead.Enabled = True 'False
                       txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                       txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                       cmdSearchFunctionary.Enabled = False
                       
                   Else
                       cmbCategory.ListIndex = -1
                       txtFunction.Tag = ""
                       txtFunction.Text = ""
                       cmdSearchFunction.Enabled = True
                       txtAccountHeadCode.Tag = ""
                       txtAccountHeadCode.Text = ""
                       txtAccountHead.Text = ""
                       cmdSearchHead.Enabled = True
                       txtFunctionary.Tag = ""
                       txtFunctionary.Text = ""
                       cmdSearchFunction.Enabled = True
                       cmdSearchFunctionary.Enabled = True
                       txtRequisition.Text = ""
                   End If
                   Rec.Close
               End If

               'AMOUNT VALIDATION
               Dim mCnnAmt    As New ADODB.Connection
               Dim RecAmt     As New ADODB.Recordset
               Dim objAmt     As New clsDB
               Dim mSQLAmt    As String
               Dim mAvailBalnz As Variant
       
               mAvailBalnz = 0
               objAmt.CreateNewConnection mCnnAmt, enuSourceString.Saankhya
               mSQLAmt = "Select fltRequestedAmt from faAllotments where tnyStatus <> 2 And numProjectID = " & val(txtProjectNo.Tag) & "  And intFinancialYearID = " & gbFinancialYearID - 1 & " And intSourceID=" & val(mSourceID) & " "
               RecAmt.Open mSQLAmt, mCnnAmt
               If Not (RecAmt.EOF And RecAmt.BOF) Then
                   While Not (RecAmt.EOF)
                       mAvailBalnz = mAvailBalnz + IIf(IsNull(RecAmt!fltRequestedAmt), 0, RecAmt!fltRequestedAmt)
                       RecAmt.MoveNext
                   Wend
               End If
               RecAmt.Close
       
               txtProjCost.Tag = Abs(objProFund.SourceWiseAmount - mAvailBalnz)
               If val(txtProjCost.Tag) >= 0 Then
                   If val(txtAmountRequested.Text) > val(txtProjCost.Tag) Then
                       msql = " Balance Available for " & cmbSource.Text & " in this Project" & vbCrLf  'Amount allocated
                       msql = msql + " is Rs. " & Format(val(txtProjCost.Tag), "0.00")
                       MsgBox msql
                       txtAmountRequested.Enabled = True
                       'txtAmountRequested.SetFocus
                       Exit Sub
                   End If
               End If
VALIDATION:
            ElseIf mTaskID = 13 Then
                chkBFund.value = 1
                cmbSource.Text = objProj.FindSourceOfFund(mSourceID)
                cmbSource.Enabled = False
                Call GetSchemeDetails(mCategoryID)
            End If
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
    
    '*********************************************************************************************'
    '                                   Form to make the Requisition                              '
    '*********************************************************************************************'
    Private Sub SetFunctionary()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim msql        As String
        Dim Rec         As New ADODB.Recordset
        
        '*********************************************************************************************'
        '                               Procedure to set the Functionary                              '
        '*********************************************************************************************'
        On Error GoTo Err
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            msql = "Select vchFunctionary,suImplementingOfficer.intFunctionaryID From suImplementingOfficer"
            msql = msql + " Inner Join faFunctionaries On suImplementingOfficer.intFunctionaryID = faFunctionaries.intFunctionaryID"
            msql = msql + " Where vchImplementingOfficerCode = '" & Trim(txtDDOCode.Text) & "'"
            Rec.Open msql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                cmdSearchFunctionary.Enabled = False
            End If
            Rec.Close
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
    
    Private Function Validations() As Boolean
        If txtIMPOName.Tag = "" Then
            MsgBox "Please select the Implementing Officers Name", vbInformation
            Validations = False
            cmdSearchIMPO.SetFocus
            Exit Function
        End If
        If txtAmountRequested.Text = "" Then
            MsgBox "Please enter the requested amount", vbInformation
            Validations = False
            txtAmountRequested.SetFocus
            Exit Function
        End If
         If val(txtAmountRequested.Text) < 0 Then
            MsgBox "Please check the amount", vbInformation
            Validations = False
            txtAmountRequested.SetFocus
            Exit Function
        End If
        If chkPlan.value = vbChecked Then
            If txtProjectNo.Text = "" Then
                MsgBox "Please enter the Project Details", vbInformation
                Validations = False
                cmdSearchProject.SetFocus
                Exit Function
            End If
        End If
''''        If chkUnAuthDrw.value = vbChecked Then
''''            'If txtGo.Text = "" Then
''''            If val(txtProjectNo.Tag) <= 0 Then
''''                MsgBox "Please enter the GO Details", vbInformation
''''                Validations = False
''''                'cmdGo.SetFocus
''''                txtProjectNo.SetFocus
''''                Exit Function
''''            End If
''''        End If
        If chkNonPlan.value = vbChecked Then
            If cmbSource.ListIndex < 1 Then
                MsgBox "Please select the Source of Fund", vbInformation
                Validations = False
                cmbSource.SetFocus
                Exit Function
            End If
            '''1,3,4,16,17,25,26,27,28, 10, 11, 12, 13, 14,29,30
            Select Case cmbSource.ListIndex
                Case 1, 3, 4, 16, 17, 25, 26, 27, 28, 10, 11, 12, 13, 14, 29, 30, 41
                 If cmbCategory.ListIndex < -1 Then
                    MsgBox "Please select the Category", vbInformation
                    Validations = False
                    cmbCategory.SetFocus
                    Exit Function
                End If
            End Select
'''            If cmbSource.ListIndex = 1 Then
'''                If cmbCategory.ListIndex < 1 Then
'''                    MsgBox "Please select the Category", vbInformation
'''                    Validations = False
'''                    cmbCategory.SetFocus
'''                    Exit Function
'''                End If
'''            End If
        End If
        If cmbSource.ListIndex = 2 And chkBFund.value = vbChecked Then
            If txtScheme.Text = "" Then
                MsgBox "Please select the B-Fund Scheme", vbInformation
                Validations = False
                txtScheme.SetFocus
                Exit Function
            End If
        End If
        
        
        If txtFunctionary.Text = "" Then
            MsgBox "Please select the Functionary", vbInformation
            Validations = False
            cmdSearchFunctionary.SetFocus
            Exit Function
        End If
        If txtFunction.Text = "" Then
            MsgBox "Please select the Function", vbInformation
            Validations = False
            cmdSearchFunction.SetFocus
            Exit Function
        End If
        If txtAccountHeadCode.Text = "" Then
            MsgBox "Please select the Account Head", vbInformation
            Validations = False
            cmdSearchHead.SetFocus
            Exit Function
        End If
        If txtNatureofClaim.Text = "" Then  'And mNewProcessFlag = True MODIFIED ON 07.Sep.2015
            MsgBox "Enter Nature Of Claim", vbInformation
            Validations = False
            txtNatureofClaim.SetFocus
            Exit Function
        End If
        
        Validations = True
    End Function
    
    Private Sub FetchRequisitionDetails()
        Dim mRequisitionID  As Variant
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mCnn            As New ADODB.Connection
        Dim mHeadAccCode    As String
        Dim mHeadAcc        As String
        
        '*********************************************************************************************'
        '                                   Form to fetch the Requisition Details                     '
        '*********************************************************************************************'
        'On Error GoTo Err
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        cmdNew.Visible = False
        If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
            cmdSave.Visible = False
            cmdApprove.Visible = True
            cmdApprove.Left = 5560
            cmdApprove.Top = 7575 '7605
            lblAuthorizationNo.Visible = True
            txtAuthorizationNo.Visible = True
            lblAuthorizedAmt.Visible = True
            txtAuthorizedAmt.Visible = True
        Else
            cmdSave.Visible = True
        End If
         
        If CheckPreviousYearRequisitions(RequisitionID) = 1 Then
            mRequisitionID = Array(RequisitionID, gbFinancialYearID - 1)
        Else
            mRequisitionID = Array(RequisitionID, gbFinancialYearID)
        End If
  
        Set Rec = objDb.ExecuteSP("spRptViewAllotmentLetter", mRequisitionID, , , mCnn, adCmdStoredProc)
        If Not (Rec.EOF And Rec.BOF) Then
            txtRequisition.Text = IIf(IsNull(Rec!vchRequisitionNo), "", Rec!vchRequisitionNo)
            txtRequisition.Tag = IIf(IsNull(Rec!intID), "", Rec!intID)
            txtRequisitiontDate.Text = IIf(IsNull(Rec!dtRequisitionDate), "", Rec!dtRequisitionDate)
            
            txtIMPOName.Text = IIf(IsNull(Rec!vchNameofIMPO), "", Rec!vchNameofIMPO)
            txtIMPOName.Tag = IIf(IsNull(Rec!intImplementingOfficersID), "", Rec!intImplementingOfficersID)
            txtDesig.Text = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
            txtDept.Text = IIf(IsNull(Rec!vchDepartment), "", Rec!vchDepartment)
            txtAmountRequested.Text = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
            txtAuthorizedAmt.Text = IIf(IsNull(Rec!fltRequestedAmt), "", Rec!fltRequestedAmt)
            If Rec!intTreasuryID = 1 Then
                chkNewMode.value = 1
            Else
                chkNewMode.value = 0
            End If
            If Not IsNull(Rec!intSchemeID) Then
                If Rec!intSourceID = 3 Then
                    chkBFund.value = vbChecked
                Else
                    chkBFund.value = vbUnchecked
                End If
                txtScheme.Visible = True
                txtScheme.Tag = Rec!intSchemeID
                txtScheme.Text = Rec!vchDescription
            End If
            
            '-----------------------------------------------------------------------------'
            ' BASED ON AS PER G.O. REQUISITIONS AND PLAN EXPENDITURE                      '
            '                                       MODIFIED ON : 17-MAR-2013             '
            '-----------------------------------------------------------------------------'
            If Rec!tnyPlanOrNonPlan = 2 Then
                chkUnAuthDrw.Visible = True
                chkUnAuthDrw.value = 1
                txtProjectNo.Text = ""
                txtProjectNo.Tag = ""
                txtProjName.Text = ""
                mLoadMode = 10
                
                cmbSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                cmbSource.Tag = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
                If Rec!vchTransactionCategory <> "" Then
                    cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                End If
                cmbCategory.Tag = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
                
                txtProjCost.Text = IIf(IsNull(Rec!fltProjectCost), "", Rec!fltProjectCost)
                txtDPCNo.Text = IIf(IsNull(Rec!vchDPCApprovalNo), "", Rec!vchDPCApprovalNo)
                txtDPCDate.Text = IIf(IsNull(Rec!dtDPCDate), "", Rec!dtDPCDate)
                
''''''                Dim RecGO As New ADODB.Recordset
''''''                Dim mSql As String
''''''                mSql = "SELECT intRefID, vchRefNo, dtGODate, intSourceOfFundID, fltAmount, intExpenditureHeadID, vchDescription, intPayOrderID, intVoucherID, fltAmountUpto, tnyStatus "
''''''                mSql = " From suGOForFunds WHERE intRefID = " & Rec!numProjectID
''''''                RecGO.Open mSql, mCnn, adOpenStatic, adLockReadOnly
''''''                If Not (RecGO.EOF And RecGO.BOF) Then
''''''
''''''                End If
''''''                RecGO.Close
'''''
               
            Else
                If Not (IsNull(Rec!vchProjectNo)) Then
                    chkPlan.value = vbChecked
                    txtProjectNo.Text = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
                    txtProjectNo.Tag = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                    txtProjName.Text = IIf(IsNull(Rec!chvProjectnameEnglish), "", Rec!chvProjectnameEnglish)
                    cmbSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    cmbSource.Tag = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
                    If Rec!vchTransactionCategory <> "" Then
                        cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    End If
                    cmbCategory.Tag = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
                    txtProjCost.Text = IIf(IsNull(Rec!fltProjectCost), "", Rec!fltProjectCost)
                    txtDPCNo.Text = IIf(IsNull(Rec!vchDPCApprovalNo), "", Rec!vchDPCApprovalNo)
                    txtDPCDate.Text = IIf(IsNull(Rec!dtDPCDate), "", Rec!dtDPCDate)
                Else

                    chkNonPlan.value = vbChecked
                    cmbSource.Text = IIf(IsNull(Rec!vchSourceFundName), "", Rec!vchSourceFundName)
                    cmbSource.Tag = IIf(IsNull(Rec!intSourceID), "", Rec!intSourceID)
                    
                    
                    If cmbSource.Tag = 1 Then 'Development Fund
                        cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                        cmbCategory.Tag = IIf(IsNull(Rec!intFundCategoryID), "", Rec!intFundCategoryID)
                    End If
                End If
            End If
            
            txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
            txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
            txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            
            '----------------------------------------------------------------------------------------'
            'Note:- Modified based on the Project Modificaiton in 2013-14 Microsector wise Mapping
            '----------------------------------------------------------------------------------------'
             If Not IsNull(Rec!intMircoSectorID) And Rec!intMircoSectorID <> 0 Then
                Call SetAccountHeads(Rec!intMircoSectorID)
             End If
            '----------------------------------------------------------------------------------------'
            
            
            txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
            cmdSearchFunctionary.Enabled = False
            
            txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            
            
            txttreasury.Text = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
            txttreasury.Tag = IIf(IsNull(Rec!intTreasuryID), "", Rec!intTreasuryID)
            txtTreasuryCode.Text = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
            MaskAccHead.Text = IIf(IsNull(Rec!vchGHeadofAccount), "", Rec!vchGHeadofAccount)
            txtStateHead.Text = IIf(IsNull(Rec!vchGBudgetHead), "", Rec!vchGBudgetHead)
            'MaskDetailAccHead.Text = IIf(IsNull(Rec!vchGDemandNo), "", Rec!vchGDemandNo)
            'txtStateSubHead.Text = IIf(IsNull(Rec!vchStateSubAcHead), "", Rec!vchStateSubAcHead)
            mHeadAcc = IIf(IsNull(Rec!vchGDemandNo), "", Rec!vchGDemandNo)
            mHeadAccCode = Token(mHeadAcc, "/")
            MaskDetailAccHead.Text = mHeadAccCode
            txtStateSubHead.Text = mHeadAcc
            txtNatureofClaim.Text = IIf(IsNull(Rec!vchNatureOfClaim), "", Rec!vchNatureOfClaim)
        End If
        Rec.Close
        'RequisitionID = ""
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
        
    Private Sub formInitialise()
        Dim ctrl As Control
            For Each ctrl In Me.Controls
                If TypeOf ctrl Is TextBox Then
                    ctrl.Text = ""
                    ctrl.Tag = ""
                ElseIf TypeOf ctrl Is OptionButton Then
                    ctrl.value = False
                ElseIf TypeOf ctrl Is CheckBox Then
                    ctrl.value = False
                ElseIf TypeOf ctrl Is ComboBox Then
                    If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                    ctrl.Tag = ""
                End If
            Next
            MaskAccHead.Text = ""
            MaskDetailAccHead.Text = ""
           ' RequisitionID = ""
           
           cmdMicroSector.Enabled = True
           cmdSubSector.Enabled = True
           
           cmbSource.Enabled = True
           cmbCategory.Enabled = True
           cmdSearchFunctionary.Enabled = True
           cmdSearchFunction.Enabled = True
           cmdSearchHead.Enabled = True
           If mLoadMode = 10 Then
                chkUnAuthDrw.value = vbChecked
           Else
                chkUnAuthDrw.Visible = False
           End If
    End Sub

    Private Sub chkBFund_Click()
        On Error GoTo Err
        If chkBFund.value = vbChecked Then
            lblScheme.Visible = True
            txtScheme.Visible = True
            cmdSearchScheme.Visible = True
            cmbSource.Text = "State Sponsored Scheme Fund"
            cmdSearchStateSubHead.Enabled = True
            cmbCategory.ListIndex = 1
            cmbCategory.Enabled = False

        Else
            lblScheme.Visible = False
            txtScheme.Visible = False
            cmdSearchScheme.Visible = False
            cmdSearchStateSubHead.Enabled = False
        End If
        Exit Sub
Err:
        MsgBox Error$
    End Sub

Private Sub chkUnAuthDrw_Click()
    If chkUnAuthDrw.value Then
        chkPlan.Enabled = False
        chkNonPlan.Enabled = False
        txtProjectNo.Enabled = False
        txtProjName.Enabled = False
        cmdSearchProject.Enabled = False
        cmbSource.Enabled = True
        cmbCategory.Enabled = True
        txtSubSector.Enabled = False
        cmdSubSector.Enabled = True
        'txtMicroSector.Enabled = False
        cmdMicroSector.Enabled = True
        cmdSearchHead.Enabled = True
        cmdSearchFunction.Enabled = True
        cmdSearchFunctionary.Visible = False
    Else
        cmdSearchProject.Enabled = False
        lblProjectNo.Caption = "Project"
    End If
End Sub

    Private Sub chkNonPlan_Click()
        Dim msql    As String
        
        If chkNonPlan.value = 1 Then
            chkPlan.value = 0
            chkUnAuthDrw.value = 0
            txtProjectNo.Text = ""
            txtProjName.Text = ""
            txtProjectNo.Tag = ""
            cmdSearchProject.Enabled = False
            'cmbSource.Enabled = False
            cmbSource.ListIndex = -1
            cmbCategory.Enabled = False
            cmbCategory.ListIndex = -1
            txtDPCNo.Enabled = False
            txtProjCost.Enabled = False
            txtDPCDate.Enabled = False
            cmbSource.Enabled = True
            cmbSource.Clear
            If gbLBPanchayat = 1 Then
                msql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In (1,3,4,16,17,25,26,27,28, 10, 11, 12, 13, 14,19,29,30,41)"
            Else
                msql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In (1,3,4,16,17,19,25,26,27,28,29,30,41)"
            End If
            PopulateList cmbSource, msql, , True, True, True, enuSourceString.Saankhya
        End If
    End Sub

    Private Sub chkPlan_Click()
        If chkPlan.value = 1 Then
            
            chkNonPlan.value = 0
            chkUnAuthDrw.value = 0
            
            cmdSearchProject.Enabled = True
            cmbSource.Enabled = True
            cmbCategory.Enabled = True
            
            txtDPCNo.Enabled = True
            txtProjCost.Enabled = True
            txtDPCDate.Enabled = True
            cmbSource.Clear
            
            txtProjectNo.Text = ""
            txtProjectNo.Tag = ""
            txtProjName.Text = ""
            
            Call FillSource
            Call FillCategory
        Else
            chkNonPlan.value = vbChecked
            cmdSearchProject.Enabled = False
            cmdSearchProject.Tag = -1
        End If
    End Sub

    Private Sub cmbCategory_Click()
        If mLoadMode = 10 Then
            If cmbCategory.ListIndex > -1 Then
                cmbCategory.Tag = cmbCategory.ItemData(cmbCategory.ListIndex)
            End If
        End If
    End Sub

    Private Sub cmbSource_Click()
        If cmbSource.ListIndex > -1 Then
'            If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
'            'If cmbSource.ListIndex = 1 Then
'                'cmbCategory.Enabled = True
'            Else
'                cmbCategory.ListIndex = 0
'                cmbCategory.Enabled = False
'            End If
                
                Select Case cmbSource.ItemData(cmbSource.ListIndex)
                    Case 1:
                        cmbCategory.Enabled = True
                    Case 29:
                        cmbCategory.Enabled = False
                        cmbCategory.ListIndex = 2
                    Case 30:
                        cmbCategory.Enabled = False
                        cmbCategory.ListIndex = 3
                    Case 4, 16, 17, 25, 26, 27, 28, 41:
                        cmbSource.Tag = cmbSource.ItemData(cmbSource.ListIndex)  'Modified on 22/12/16 KLGSDP
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                    Case 3:
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                    Case 10, 11, 12, 13, 14:
                        cmbCategory.Enabled = True
                    Case Else:
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                End Select
        End If
    End Sub
    Private Sub cmdApprove_Click()
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim mArrIN          As Variant
        Dim mArrOut         As Variant
        Dim mYearID         As Variant
        Dim dtAuthorizeDate As Date
        Dim dtAllotmentDate As Date
        '*********************************************************************************************'
        '                           Procedure to Approve the Requisition                              '
        '*********************************************************************************************'
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If Trim(txtAuthorizedAmt.Text) = "" Then
            MsgBox "Please Enter the Authorized Amount", vbInformation
            Exit Sub
        End If
        If mPreviousYearMode = 1 Or CheckPreviousYearRequisitions(RequisitionID) = 1 Then
            mYearID = gbFinancialYearID - 1
            dtAuthorizeDate = CDate(txtRequisitiontDate.Text)
            dtAllotmentDate = CDate(txtRequisitiontDate.Text)
        Else
            mYearID = gbFinancialYearID
            dtAuthorizeDate = gbTransactionDate
            dtAllotmentDate = gbTransactionDate
        End If
        
        mCnn.BeginTrans
        On Error GoTo Err:
        mArrIN = Array(txtRequisition.Tag, _
                        2, _
                        Null, _
                        dtAuthorizeDate, _
                        val(txtAuthorizedAmt.Text), _
                        gbUserID, _
                        Null, _
                        dtAllotmentDate, _
                        mYearID, _
                        IIf((mLoadMode = 10), 3, Null) _
                     )
        objDb.ExecuteSP "spSaveAuthorizeRequisition", mArrIN, mArrOut, , mCnn, adCmdStoredProc
        mCnn.CommitTrans
        cmdApprove.Enabled = False
        If IsArray(mArrOut) Then
            txtAuthorizationNo.Text = mArrOut(0, 0)
        End If
        frmViewAllotmentLetter.Mode = 2
        frmViewAllotmentLetter.ArrayIn = Array(CStr(val(txtRequisition.Tag)), CStr(mYearID))
        Unload Me
        frmViewAllotmentLetter.Show vbModal
        Exit Sub
Err:
        mCnn.RollbackTrans
        MsgBox Error$
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
        'Me.Hide
    End Sub

    Private Sub cmdGo_Click()
        If chkUnAuthDrw.value = vbChecked Then
            frmGoDetails.Show vbModal
            If gbSearchID <> -1 Then
                txtGo.Text = gbSearchStr
                txtGo.Tag = gbSearchID
                gbSearchStr = ""
                gbSearchID = -1
            End If
        End If
    End Sub

Private Sub cmdMicroSector_Click()
    Dim msql As String
    If Len(mWHERE) > 0 Then
        msql = "Select intMicroSecID, vchEngMicroSector From suMicroSectors WHERE intMicroSecID IN ( " & mWHERE & " ) "
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = msql
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        txtMicroSector.SetFocus
        mWHERE = ""
    ElseIf mLoadMode = 10 Then
        msql = "Select intMicroSecID, vchEngMicroSector From suMicroSectors "
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = msql
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        txtMicroSector.SetFocus
        If gbSearchID <> -1 Then
            txtMicroSector.Text = gbSearchStr
            txtMicroSector.Tag = gbSearchID
        End If
    End If
End Sub

'''    Private Sub cmdReject_Click()
'''        frmReject.Mode = 9
'''        frmReject.RequestTypeID = txtRequisition.Text
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''        cmdApprove.Enabled = False
'''    End Sub

    Private Sub cmdSearchFunction_Click()
        
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim obj    As New clsDB
        
        Dim msql As String
        Dim mToken1 As String
        
        
        frmSearchFunction.Show vbModal
        mToken1 = Token(gbSearchStr, " ")     'To place the Function Code seperately
        txtFunction.Text = Trim(gbSearchStr)
        txtFunction.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
        
        '****************UNAUTHORIZED DRAWAL****************************
        If mLoadMode = 10 Then
            obj.SetConnection mCnn
            
            msql = "SELECT faFunctions.intFunctionID,vchFunction,faFunctionaries.intFunctionaryID FunctionaryID,vchFunctionary  FROM faFunctionaryFunctions"
            msql = msql + " INNER JOIN faFunctions ON faFunctions.intFunctionID=faFunctionaryFunctions.intFunctionID"
            msql = msql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID=faFunctionaryFunctions.intFunctionaryID"
            msql = msql + " WHERE faFunctionaryFunctions.intFunctionID=" & val(txtFunction.Tag) & " "
            
            Rec.Open msql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtFunctionary.Text = Rec!vchFunctionary
                txtFunctionary.Tag = Rec!FunctionaryID
            End If
            Rec.Close
        End If
        '***************************************************************
        
    End Sub
    Private Sub cmdSearchFunctionary_Click()
        Dim msql As String
        Dim mToken1 As String
        frmSearchFunctionary.Show vbModal
        mToken1 = Token(gbSearchStr, " ")      'To Place the  Functionary Code sepereately
        txtFunctionary.Text = Trim(gbSearchStr)
        txtFunctionary.Tag = gbSearchID
        
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    
    Private Sub cmdSearchHead_Click()
        Dim msql As String
        'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where  tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
        '--MODIFIED BY MINU ON 31/Oct/2011 to remove listing of AccountHeadCode  starting with 1-----
        
        If Len(cmdSearchHead.Tag) > 0 Then
            frmSearchAccountHeads.SQLString = cmdSearchHead.Tag
        Else
            frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where  tinHiddenFlag = 0 and vchAccountHeadCode Not Like '1%' Order By faAccountHeads.vchAccountHeadCode"
        End If
        frmSearchAccountHeads.chkListAll.Enabled = False
        
        '----------------------------------------------------------------------------------
        frmSearchAccountHeads.Show vbModal

        If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    txtAccountHeadCode.Text = objAccHead.AccountCode
                    txtAccountHead.Text = objAccHead.AccountHead
                    txtAccountHeadCode.Tag = objAccHead.AccountHeadID
              End If
                gbSearchStr = ""
                gbSearchID = -1
            End If
    End Sub
    Private Sub FillCategory()
        Dim msql As String
        msql = "SELECT vchTransactionCategory,intCategoryID FROM faTransactionCategory"
        PopulateList cmbCategory, msql, True, True, True, True
        'PopulateList cmbCategory, mSQL, , False, True, True, enuSourceString.Saankhya
  
    End Sub
    Private Sub FillSource()
        Dim msql As String
        msql = "SELECT vchSourceFundName,intSourceFundID from suSourceOfFund"
        PopulateList cmbSource, msql, , False, True, True, enuSourceString.Saankhya
    End Sub
    Private Sub SaveRequisition()
        Dim mCnn        As New ADODB.Connection
        Dim objDb       As New clsDB
        Dim Rec         As New ADODB.Recordset
        Dim RecAmt      As New ADODB.Recordset
        Dim arrInput    As Variant
        Dim arrOutPut   As Variant
        Dim Reqn        As uRequisition
        Dim msql        As String
        Dim mRAmount    As Variant
        Dim mPAmount    As Variant
        Dim mAryIn      As Variant
        Dim mOpening    As Integer
        
        
        '****************GM_Treasury**************************
        
        Dim mCnnTR    As New ADODB.Connection
        Dim RecTr     As New ADODB.Recordset
        Dim objTr     As New clsDB
        Dim mSQLTR    As String
        Dim mArrIN    As Variant
        Dim mID       As Integer
        Dim mYearID   As Integer
        
        
        If mPreviousYearMode Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        
        objTr.CreateNewConnection mCnnTR, enuSourceString.DBMaster
        
        If txtTreasuryCode.Text <> "" And txttreasury.Text <> "" Then
            mID = IIf(txtTreasuryCode.Tag = "", -1, val(txtTreasuryCode.Tag))
            mSQLTR = "select * from GM_Treasury where chvTreasuryCode= '" & txtTreasuryCode.Text & "' "
            RecTr.Open mSQLTR, mCnnTR
            If (RecTr.EOF And RecTr.BOF) Then
                mArrIN = Array(mID, _
                        txtTreasuryCode.Text, _
                        txttreasury.Text)
                objTr.ExecuteSP "spSaveTreasury", mArrIN, , , mCnnTR, adCmdStoredProc
            End If
            RecTr.Close
        End If
        If cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
           If chkPlan.value = Unchecked And chkNonPlan = vbChecked Then
            If txtScheme.Tag = "" Then
                 MsgBox "Please select the Scheme", vbInformation, "Saankhya"
                 Exit Sub
            End If
           End If
        End If
        
        If mNewProcessFlag Then
                
        '*********************************************************************************************'
        '                                   Procedure to Save the Requisition   - NEW PROCESS                  '
        '*********************************************************************************************'
                Dim mCatID As Integer
                Dim mReqDate As Variant
                
                If cmbCategory.ListIndex > -1 Then
                    mCatID = cmbCategory.ItemData(cmbCategory.ListIndex)
                End If
                If IsDate(txtRequisitiontDate.Text) Then
                    mReqDate = txtRequisitiontDate.Text
                Else
                    mReqDate = gbTransactionDate
                End If
                If mPreviousYearMode = 1 Then
                    mArrIN = Array(cmbSource.ItemData(cmbSource.ListIndex), mCatID, mReqDate, gbFinancialYearID - 1)
                Else
                    mArrIN = Array(cmbSource.ItemData(cmbSource.ListIndex), mCatID, mReqDate)
                End If
                If objDb.SetConnection(mCnn) Then
                    Set Rec = objDb.ExecuteSP("spCheckACRBalance", mArrIN, , True, mCnn, adCmdStoredProc)
                    If val(cmbSource.Tag) <> 2 Then
                        If Not (Rec.EOF And Rec.BOF) Then
                            'MsgBox Rec!fltBalance
                            If IIf(IsNull(Rec!fltBalance), 0, Rec!fltBalance) < val(txtAmountRequested) Then
                                MsgBox "No Balance Available in ACR [ NEW-ACR ]", vbInformation
                                cmdSave.Enabled = False
                                Exit Sub
                            End If
                        Else
                            MsgBox "No Balance Available in ACR [ NEW-ACR ]", vbCritical
                            cmdSave.Enabled = False
                            Exit Sub
                        End If
                    End If
                End If
                cmdSave.Enabled = True
       Else
                '*********************************************************************************************'
                '                                   Procedure to Save the Requisition   - OLD PROCESS                      '
                '*********************************************************************************************'
                On Error GoTo Err
                objDb.SetConnection mCnn
                mRAmount = 0
                mPAmount = 0
'''''                If cmbSource.ItemData(cmbSource.ListIndex) <> 3 Then
'''''                    mSql = "Select faVouchers.intVoucherNo intVoucherNo,faAllotmentLetters.intSourceOfFundID,faAllotmentLetters.intCategoryID,faAllotmentLetters.tnyOpening,"
'''''                    mSql = mSql + " faVouchers.tnyCancelFlag tnyCancelFlag,faAllotmentLetters.tnyStatus as Status,  faAllotmentLetters.fltAmount As Amount From faAllotmentLetters"
'''''                    mSql = mSql + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID = faAllotmentLetters.intSourceOfFundID"
'''''                    mSql = mSql + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID"
'''''                    mSql = mSql + " Left Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID"
'''''                    mSql = mSql + "         And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
'''''                    mSql = mSql + "         And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
'''''                    mSql = mSql + " Left Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
'''''                    mSql = mSql + " Where faAllotmentLetters.intSourceOfFundID <> 3"
'''''                    mSql = mSql + " And faAllotmentLetters.tnyStatus <> 8 And faAllotmentLetters.intFinancialYearID=" & mYearID & " "
'''''                    mSql = mSql + "  AND ISNULL(tnyGroupID,0) NOT IN (40,90)"
'''''                    mSql = mSql + " Union All"
'''''                    mSql = mSql + " Select 10000000001 intVoucherNo,intSourceOfFundID,intCategoryID,tnyOpening,null tnyCancelFlag,tnyStatus As Status,"
'''''                    mSql = mSql + " fltAmount As Amount from faExtractAllotments where intFinancialYearID=" & mYearID & " "
'''''
'''''                    Rec.Open mSql, mCnn
'''''                    While Not Rec.EOF
'''''                        mOpening = IIf(IsNull(Rec!tnyOpening), 0, Rec!tnyOpening)
'''''                        If mYearID = gbFinancialYearID And Rec!Status = 1 Then
'''''                            If mOpening = 0 And Rec!Status = 1 Then
'''''                                    If IIf(IsNull(Rec!tnyCancelFlag), 0, Rec!tnyCancelFlag) <> 1 Then
'''''                                        If Rec!intSourceOfFundID = cmbSource.ItemData(cmbSource.ListIndex) Then
'''''                                            'If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
'''''                                            If (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
'''''                                                Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) Then
'''''                                                    If Rec!intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex) Then
'''''                                                        mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                                    End If
'''''                                            Else
'''''                                                mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                            End If
'''''                                        End If
'''''                                    End If
'''''                             ElseIf mOpening = 1 And Rec!Status = 1 Then     '**********ADDEDD FOR OPENING******************
'''''                                If Rec!intSourceOfFundID = cmbSource.ItemData(cmbSource.ListIndex) Then
'''''                                        'If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
'''''                                         If (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
'''''                                            Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) Then
'''''
'''''                                            If Rec!intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex) Then
'''''                                                mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                            End If
'''''                                        Else
'''''                                            mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                        End If
'''''                                 End If
'''''                            End If
'''''
'''''                        Else
'''''                            If Not IsNull(Rec!intVoucherNo) Then
'''''                                    If IIf(IsNull(Rec!tnyCancelFlag), 0, Rec!tnyCancelFlag) <> 1 Then
'''''                                        If Rec!intSourceOfFundID = cmbSource.ItemData(cmbSource.ListIndex) Then
'''''                                            'If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
'''''                                            If (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
'''''                                                Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) Then
'''''                                                    If Rec!intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex) Then
'''''                                                        mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                                    End If
'''''                                            Else
'''''                                                mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                            End If
'''''                                        End If
'''''                                    End If
'''''                             ElseIf mOpening = 1 And Rec!Status = 1 Then     '**********ADDEDD FOR OPENING******************
'''''                                If Rec!intSourceOfFundID = cmbSource.ItemData(cmbSource.ListIndex) Then
'''''                                        'If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
'''''                                         If (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
'''''                                            Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) Then
'''''
'''''                                            If Rec!intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex) Then
'''''                                                mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                            End If
'''''                                        Else
'''''                                            mRAmount = mRAmount + IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                                        End If
'''''                                 End If
'''''                            End If
'''''                        End If
'''''                        Rec.MoveNext
'''''                    Wend
'''''                    Rec.Close
'''''                Else  'SourceOfFund=3 State Sponsored Scheme Fund
'''''                    If chkPlan.value = vbChecked And chkNonPlan = Unchecked Then
'''''                        mSql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
'''''                        mSql = mSql + " And tnyStatus not in (8,0) And intFinancialYearID = " & mYearID & " "
'''''                        mSql = mSql + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
'''''                        Rec.Open mSql, mCnn
'''''                        If Not (Rec.EOF And Rec.BOF) Then
'''''                            mRAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                        End If
'''''                    Else
'''''                        mSql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
'''''                        mSql = mSql + " And tnyStatus not in (8,0) And intSchemeID = " & txtScheme.Tag & " And intFinancialYearID = " & mYearID & " "
'''''                        mSql = mSql + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
'''''                        Rec.Open mSql, mCnn
'''''                        If Not (Rec.EOF And Rec.BOF) Then
'''''                            mRAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
'''''                        End If
'''''                    End If
'''''                    Rec.Close
'''''                End If
                
                If cmbSource.ItemData(cmbSource.ListIndex) <> 3 Then
                    'Total Amount received
                    If cmbSource.ItemData(cmbSource.ListIndex) = 1 Or _
                        cmbSource.ItemData(cmbSource.ListIndex) = 19 Or cmbSource.ItemData(cmbSource.ListIndex) = 21 Or _
                        cmbSource.ItemData(cmbSource.ListIndex) = 27 Or cmbSource.ItemData(cmbSource.ListIndex) = 28 Or _
                        cmbSource.ItemData(cmbSource.ListIndex) = 10 _
                        Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
                        Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14 Then
                            msql = " SELECT SUM(fltAmtReceived) fltAmtReceived FROM ("
                            msql = msql + " Select fltAmount fltAmtReceived, intSourceOfFundId, intCategoryID, intFinancialYearID From faExtractAllotments "
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 2 AND intSourceOfFundID IN (1,21,27, 28, 10, 11, 12, 13, 14,19)"
                            msql = msql + " And intCategoryID = 1 and intFinancialYearID = " & mYearID & " "
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmtReceived,intSourceOfFundId,intCategoryID,intFinancialYearID FROM"
                            msql = msql + " ("
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID IN (21,27, 28, 10, 11, 12, 13, 14,19)"
                            msql = msql + "         And intCategoryID = 1 and intFinancialYearID =" & mYearID & " "
                            'mSql = mSql + "         And dtAllotmentDate <= @dtDate"
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID IN (1) AND ISNULL(tnyGroupID,0) =30"
                            msql = msql + " And intFinancialYearID = " & mYearID & " "
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID IN (1) "
                            msql = msql + " And intFinancialYearID = " & mYearID & " "
                            msql = msql + " ) A"
                            msql = msql + " )B"
                    ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 16 Or cmbSource.ItemData(cmbSource.ListIndex) = 17 Then
                            msql = " SELECT SUM(fltAmtReceived) fltAmtReceived FROM ("
                            msql = msql + " Select fltAmount fltAmtReceived, intSourceOfFundId, intCategoryID, intFinancialYearID From faExtractAllotments "
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 2 AND intSourceOfFundID IN (16,17)"
                            msql = msql + " And intCategoryID = 1 and intFinancialYearID = " & mYearID & " "
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmtReceived,intSourceOfFundId,intCategoryID,intFinancialYearID FROM"
                            msql = msql + " ("
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID IN (16,17)"
                            msql = msql + "         And intCategoryID = 1 and intFinancialYearID =" & mYearID & " "
                            'mSql = mSql + "         And dtAllotmentDate <= @dtDate"
                            msql = msql + " ) A"
                            msql = msql + " )B"
                    Else
                            msql = " SELECT SUM(fltAmtReceived) fltAmtReceived FROM ("
                            msql = msql + " Select fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID From faExtractAllotments"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 2 AND intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
                            msql = msql + "     and intFinancialYearID = " & mYearID & " "
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
                            msql = msql + "         and intFinancialYearID = " & mYearID & " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90) "
                            'mSql = mSql + "         And dtAllotmentDate <= @dtDate"
                            msql = msql + " Union All"
                            msql = msql + " SELECT fltAmount fltAmtReceived, 0 fltAmtIssued, intSourceOfFundId, intCategoryID, intFinancialYearID FROM faAllotmentLetters"
                            msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
                            msql = msql + " and intFinancialYearID = " & mYearID & " "
                            msql = msql + " AND ISNULL(tnyGroupID,0) =30"
                            msql = msql + " )B"
                     End If
                     Rec.Open msql, mCnn
                     If Not (Rec.EOF And Rec.BOF) Then
                        mRAmount = IIf(IsNull(Rec!fltAmtReceived), 0, Rec!fltAmtReceived)
                     End If
                     Rec.Close
                
                Else  'SourceOfFund=3 State Sponsored Scheme Fund
                    If chkPlan.value = vbChecked And chkNonPlan = Unchecked Then
                        msql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
                        msql = msql + " And tnyStatus not in (8,0) And intFinancialYearID = " & mYearID & " "
                        msql = msql + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                        Rec.Open msql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mRAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                        End If
                    Else
                        msql = "Select Sum(fltAmount) As Amount From faAllotmentLetters Where intSourceOfFundID = " & cmbSource.ItemData(cmbSource.ListIndex)
                        msql = msql + " And tnyStatus not in (8,0) And intSchemeID = " & txtScheme.Tag & " And intFinancialYearID = " & mYearID & " "
                        msql = msql + " AND ISNULL(tnyGroupID,0) NOT IN (30,40,90)"
                        Rec.Open msql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            mRAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                        End If
                    End If
                    Rec.Close
                End If
                
                If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then 'to find how much till Alloted
                   msql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & " and  tnyStatus<>2  And intFinancialYearID = " & mYearID & " " 'And intFundCategoryID = " & cmbCategory.ItemData(cmbCategory.ListIndex) & "
                   msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)"
                ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
                   If chkPlan.value = vbChecked And chkNonPlan = Unchecked Then
                    msql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & " and  tnyStatus<>2  And intFinancialYearID = " & mYearID & " "
                    msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN   (1)"
                   Else
                    msql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & " And intSchemeID = " & txtScheme.Tag & " and  tnyStatus<>2  And intFinancialYearID = " & mYearID & " "
                    msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN   (1)"
                   End If
                ElseIf (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) Then
                    msql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & " and  tnyStatus<>2  And intFinancialYearID = " & mYearID & "  And intFundCategoryID = " & cmbCategory.ItemData(cmbCategory.ListIndex) & ""
                    msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)"
                
                Else
                   msql = "Select Sum(fltRequestedAmt) As Amount From faAllotments Where intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & "  And  tnyStatus<>2   And intFinancialYearID = " & mYearID & " "
                   msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)"
                End If
                Rec.Open msql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mPAmount = IIf(IsNull(Rec!Amount), 0, Rec!Amount)
                End If
                Rec.Close
                
                
                If cmbSource.ItemData(cmbSource.ListIndex) = 1 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 16 Or cmbSource.ItemData(cmbSource.ListIndex) = 17 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 25 Or cmbSource.ItemData(cmbSource.ListIndex) = 26 Or _
                cmbSource.ItemData(cmbSource.ListIndex) = 27 Or cmbSource.ItemData(cmbSource.ListIndex) = 41 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 28 Or cmbSource.ItemData(cmbSource.ListIndex) = 30 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 29 Then
                    If mPAmount > mRAmount Then
                        MsgBox "Amount Exceed (The amount alloted is only Rs." & Format(mRAmount, "0.00") & ")", vbInformation
                        Exit Sub
                    ElseIf val(txtAmountRequested.Text) + val(mPAmount) > mRAmount Then
                        MsgBox "Amount Exceed (The amount alloted is only Rs." & Format(mRAmount, "0.00") & ")", vbInformation
                        Exit Sub
                    End If
                End If
                'TO SKIP VALIDATION FOR OTHER LSIG's FOR THE FINANCIAL YEAR 2012
                If (cmbSource.ItemData(cmbSource.ListIndex) = 10 Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14) And mYearID <> 2012 Then
                    If mPAmount > mRAmount Then
                        MsgBox "Amount Exceed (The amount alloted is only Rs." & Format(mRAmount, "0.00") & ")", vbInformation
                        Exit Sub
                    ElseIf val(txtAmountRequested.Text) + val(mPAmount) > mRAmount Then
                        MsgBox "Amount Exceed (The amount alloted is only Rs." & Format(mRAmount, "0.00") & ")", vbInformation
                        Exit Sub
                    End If
                End If
        
                If chkPlan.value = vbChecked Then
                    If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
                        msql = "Select tnyInstallmentNo From faAllotments "
                        msql = msql + " Where numProjectID = " & txtProjectNo.Tag
                        msql = msql + " And intImplementingOfficersID =" & txtIMPOName.Tag
                        msql = msql + " And intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex) & " And intFundCategoryID = " & cmbCategory.ItemData(cmbCategory.ListIndex)
                        msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN   (1)"
                    Else
                        msql = "Select tnyInstallmentNo From faAllotments "
                        msql = msql + " Where numProjectID = " & txtProjectNo.Tag
                        msql = msql + " And intImplementingOfficersID =" & txtIMPOName.Tag
                        msql = msql + " And intSourceID = " & cmbSource.ItemData(cmbSource.ListIndex)
                        msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN   (1)"
                    End If
                    Rec.Open msql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        txtInstalmentNo.Text = IIf(IsNull(Rec!tnyInstallmentNo), "", Rec!tnyInstallmentNo)
                        If txtInstalmentNo.Text = "" Then
                            txtInstalmentNo.Text = 1
                        Else
                            txtInstalmentNo.Text = val(txtInstalmentNo.Text) + 1
                        End If
                    Else
                        txtInstalmentNo.Text = 1
                    End If
                    Rec.Close
                End If
            
        End If
        
        '''''''''''''''''       Saving           ''''''''''''''
        With Reqn
            .tnyStage = 1
            .vchRequisition = Trim(txtRequisition.Text)
            
            If mPreviousYearMode = 1 Or CheckPreviousYearRequisitions(ReqID) = 1 Then
                .dtRequisitionDate = txtRequisitiontDate.Text
                .intFinancialYearID = mYearID 'gbFinancialYearID - 1
            Else
                .dtRequisitionDate = gbTransactionDate
                .intFinancialYearID = gbFinancialYearID
            End If
            
            .intImplementingOfficersID = val(txtIMPOName.Tag)
            .vchDesignation = Trim(txtDesig.Text)
            .vchNameofIMPO = Trim(txtIMPOName.Text)
            .vchPlace = Null
            .vchDepartment = Trim(txtDept.Text)
            .vchDDOCode = Trim(txtDDOCode.Text)
            .fltRequestedAmt = val(txtAmountRequested.Text)
            If chkPlan.value = vbChecked Then
                .tnyPlanOrNonPlan = 1
                .numProjectID = val(txtProjectNo.Tag)
                .numProjectNo = txtProjectNo.Text
                .fltProjectCost = val(txtProjCost.Text)
                .vchDPCApprovalNo = txtDPCNo.Text
                If mPreviousYearMode Then
                    .dtDPCDate = txtRequisitiontDate.Text
                Else
                    .dtDPCDate = CheckDateInMMM(txtDPCDate.Text)
                End If
                .intSourceID = cmbSource.ItemData(cmbSource.ListIndex)
                .intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
            Else
                If chkUnAuthDrw.value = 1 Then
                    .tnyPlanOrNonPlan = 2
                    '.intFinancialYearID = gbFinancialYearID - 1
                Else
                    .tnyPlanOrNonPlan = 0
                End If
                .intSourceID = cmbSource.ItemData(cmbSource.ListIndex)
                If .intSourceID > 0 Then
                    .intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
                End If
            End If
            
'''''       Added on 18 mar 2017 for joint Venture
            If cmbSource.ItemData(cmbSource.ListIndex) = 4 Or cmbSource.ItemData(cmbSource.ListIndex) = 10 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 11 Or cmbSource.ItemData(cmbSource.ListIndex) = 12 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 13 Or cmbSource.ItemData(cmbSource.ListIndex) = 14 _
                Or cmbSource.ItemData(cmbSource.ListIndex) = 2 Then
                .intTreasuryID = 0
            Else
                .intTreasuryID = val(txttreasury.Tag)
            End If
'''            .intTreasuryID = val(txttreasury.Tag) ''''Commented on 18 mar 2017 for joint Venture
            .vchTreasuryCode = txtTreasuryCode.Text 'val(txtTreasuryCode.Text)
            '.vchTreasuryName = mID(txttreasury.Text, 16)
            .vchTreasuryName = txttreasury.Text
            .vchGHeadofAccount = MaskAccHead.Text
            .vchGBudgetHead = txtStateHead.Text
            .vchGDemandNo = Trim(MaskDetailAccHead.Text) & "/" & Trim(txtStateSubHead.Text)
            
            .intFunctionaryID = val(txtFunctionary.Tag)
            .intFunctionID = val(txtFunction.Tag)
            .intAccountHeadID = val(txtAccountHeadCode.Tag)
            .vchAccountHeadCode = val(txtAccountHeadCode.Text)
            
            .intLBID = gbLocalBodyID
            
            .tnyStatus = 0
            .tnyInstallmentNo = txtInstalmentNo.Text
            
            .intSchemeID = IIf(val(txtScheme.Tag) > 0, val(txtScheme.Tag), Null)
            
            .intSubSecID = IIf(IsNull(val(txtSubSector.Tag)), 0, val(txtSubSector.Tag))
            .intMircoSectorID = IIf(IsNull(val(txtMicroSector.Tag)), 0, val(txtMicroSector.Tag))
            
            .vchNatureOfClaim = IIf(IsNull(txtNatureofClaim.Text), "", txtNatureofClaim.Text)
            
            '***********UNAUTHORIZED DRAWAL***************************
            If mLoadMode = 10 Then
                .tnyTypeID = 3
            Else
                .tnyTypeID = Null
            End If
            '*********************************************************
            
            If .intSourceID = 3 Then
                If IsNull(.intSchemeID) Then
                    If chkPlan.value = Unchecked And chkNonPlan = vbChecked Then
                        MsgBox "Please select a source from which the requisition is demanding for? (Department/Scheme/Project)", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            
            arrInput = Array(RequisitionID, .tnyStage, .vchRequisition, _
                           .dtRequisitionDate, _
                           .intImplementingOfficersID, _
                           .vchDesignation, _
                           .vchNameofIMPO, _
                           .vchPlace, _
                           .vchDepartment, _
                           .vchDDOCode, _
                           .fltRequestedAmt, _
                           .tnyPlanOrNonPlan, _
                           .numProjectID, _
                           .numProjectNo, _
                           .fltProjectCost, _
                           .vchDPCApprovalNo, _
                           .dtDPCDate, _
                           .intSourceID, _
                           .intCategoryID, _
                           .intTreasuryID, _
                           .vchTreasuryCode, _
                           .vchTreasuryName, _
                           .vchGHeadofAccount, _
                           .vchGBudgetHead, _
                           .vchGDemandNo, _
                           .intFunctionaryID, .intFunctionID, .intAccountHeadID, .vchAccountHeadCode, .intLBID, .intFinancialYearID, .tnyStatus, Null, Null, Null, Null, Null, Null, Null, Null, .tnyInstallmentNo, _
                            Null, Null, Null, Null, Null, Null, Null, Null, Null, .intSchemeID, .intSubSecID, .intMircoSectorID, .tnyTypeID, Null, .vchNatureOfClaim) ' MODIFIED ON 07.Sep.2015
                          
                          
          If mLoadMode = 20 Then   '************REQUISIION INBOX***************************
              objDb.ExecuteSP "spSaveRequisitionFromInbox", arrInput, arrOutPut, True, mCnn, adCmdStoredProc
              If IsArray(arrOutPut) Then
                txtRequisition.Text = arrOutPut(0, 0)
                txtRequisition.Tag = arrOutPut(1, 0)
              End If
              msql = "Update faRequisitionInbox set  intRequisitionNO= " & txtRequisition.Text & ",tnyStage=1 WHERE intTockenNo=" & mTokenID & "  "
              objDb.ExecuteSP msql, , , , mCnn, adCmdText
              ''Call SynTOWEB
          Else
              objDb.ExecuteSP "spSaveAllotmentRequisition", arrInput, arrOutPut, True, mCnn, adCmdStoredProc
              If IsArray(arrOutPut) Then
                txtRequisition.Text = arrOutPut(0, 0)
                txtRequisition.Tag = arrOutPut(1, 0)
              End If
          End If
      End With
      
        
        
        ' ========================================================================= '
        ' TREASURY BALANCE IS ZERO AS PER PREVIOUS PROCESS
        ' ========================================================================= '
        Call CalulateAmountIssued(cmbSource.ItemData(cmbSource.ListIndex), cmbCategory.ItemData(cmbCategory.ListIndex))
        If chkNewMode.value = 0 Then
            If fltClosingTreasuryBalance - val(txtAmtIssued) <= 0 Then  '- val(txtAmountRequested.Text)
                msql = " Update faBankSource SET tnyClosingFlag = 9 WHERE intBankID = " & val(txtTreasuryBalance.Tag)
                objDb.ExecuteSP msql, , , , mCnn, adCmdText
            End If
        End If
        ' ========================================================================= '
        ' END OF BLOCK ::: TREASURY BALANCE IS ZERO AS PER PREVIOUS PROCESS
        ' ========================================================================= '
        
    
    
    
    
        If mLoadMode <> 10 Then           '**********TO SKIP THE VALIDATION FOR UNAUTHORIZED DRAWAL*******************
            Call UpdateProjectDetials(val(txtProjectNo.Tag))
            msql = "Update faAllotments set  tnyProjectStatus=1 Where intID=" & txtRequisition.Tag & "  "
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
        End If
        If mPreviousYearMode Then
            msql = "Update faPendingTaskRequest set  intKeyID= " & txtRequisition.Tag & ",tnyStatus = 8 Where intRequestID=" & mPreviousYearRequestID & "  "
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
        End If
        
        cmdNew.Enabled = True
        cmdSave.Enabled = False
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
    Private Sub cmdNew_Click()
        Call formInitialise
        RequisitionID = ""
       
        '''For SaankhyaWeb Updation
        If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
        
            cmdSave.Enabled = True
        Else
        
            cmdSave.Enabled = False
        End If
    End Sub
    Private Sub cmdSave_Click()
        If Validations = True Then
            Call SaveRequisition
        End If
    End Sub
    Private Sub cmdSearchIMPO_Click()
        gbSearchID = -1                                         ''  Setting the Search ID to -1
        frmSearchSubsidiaryAccountHeads.SubLedgerType = 1       ''  1. Implementing Officer
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtIMPOName.SetFocus
    End Sub

    Private Sub cmdSearchScheme_Click()
        On Error GoTo Err:
                Dim msql As String
'''                frmSearchMasters.QrySP = StoredProcedure
'''                'frmSearchMasters.SQLQry = "spSelectScheme"
'''                frmSearchMasters.SQLQry = "spSelectDepSchemePro"

            
                frmSearchMasters.QrySP = Qyery
                If chkBFund.value = vbChecked Then
                    frmSearchMasters.SQLQry = "SELECT intID , vchDescription FROM   faDepSchPro WHERE tnyGroupID IN (1,2) ORDER BY vchDescription asc"
                Else
                    frmSearchMasters.SQLQry = "SELECT intID , vchDescription FROM   faDepSchPro WHERE tnyGroupID IN (3) ORDER BY vchDescription asc"
                End If
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.Show vbModal
                If gbSearchStr <> "" Then
                    txtScheme.Text = gbSearchStr
                    txtScheme.Tag = gbSearchID
                End If
                gbSearchStr = ""
                gbSearchID = -1
            Exit Sub
Err:
            MsgBox (Error$)
    End Sub
    Private Sub cmdSearchStateHead_Click()
        Dim msql As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb     As New clsDB
       
        
        Dim mArr As Variant
        Dim mCatID As Integer
        Dim mReqDate As Variant
        Dim mArrIN As Variant
        
        
 
        If val(txtAmountRequested.Text) = 0 Then
            MsgBox "ENTER THE AMOUNT ", vbInformation
            txtAmountRequested.SetFocus
            Exit Sub
        End If
        Call CheckBAnkSource
        If mPreviousYearMode = 1 Then
             If val(cmbSource.Tag) = 4 Or val(cmbSource.Tag) = 2 Or val(cmbSource.Tag) = 3 _
                Or val(cmbSource.Tag) = 5 Or val(cmbSource.Tag) = 6 Or val(cmbSource.Tag) = 7 Or val(cmbSource.Tag) = 8 Or val(cmbSource.Tag) = 9 Or val(cmbSource.Tag) = 15 Or val(cmbSource.Tag) = 19 _
                Or val(cmbSource.Tag) = 20 Or val(cmbSource.Tag) = 22 _
                Or val(cmbSource.Tag) = 23 Or val(cmbSource.Tag) = 24 Then
                
                GoTo CheckValidation::
             ElseIf val(cmbSource.Tag) = 29 Or val(cmbSource.Tag) = 30 Then
                msql = "Select * From faBankSourceChild"
                msql = msql + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intSourceOfFundID = " & val(cmbSource.Tag)
                msql = msql + " AND tnyCategoryID= " & val(cmbCategory.Tag)
            
             Else
            
                msql = "Select * From faBankSourceChild"
                msql = msql + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intSourceOfFundID = " & val(cmbSource.Tag)
               
             End If
             If objDb.SetConnection(mCnn) Then
               Rec.Open msql, mCnn
               If Not (Rec.EOF And Rec.BOF) Then
                   If IsNull(Rec!tnyClosingFlag) Then
                        mNewProcessFlag = False
                   ElseIf Rec!tnyClosingFlag = 9 Then
                        mNewProcessFlag = True
                   Else
                        mNewProcessFlag = False
                   End If
               Else
               
                    MsgBox "Treasury Balances are not Verified!"
                    cmdSave.Enabled = False
                    Exit Sub
               End If
               
            Else
                MsgBox "Connection Failed", vbCritical
                Exit Sub
            End If
        Else
        
            '''Added On 30/jan/2017  To get New Budget code for the following SourceFund 2,3,4,5,6,7,8,9,15,19,20,22,23,24
            If val(cmbSource.Tag) = 4 Or val(cmbSource.Tag) = 2 Or val(cmbSource.Tag) = 3 _
                Or val(cmbSource.Tag) = 5 Or val(cmbSource.Tag) = 6 Or val(cmbSource.Tag) = 7 Or val(cmbSource.Tag) = 8 Or val(cmbSource.Tag) = 9 Or val(cmbSource.Tag) = 15 Or val(cmbSource.Tag) = 19 _
                Or val(cmbSource.Tag) = 20 Or val(cmbSource.Tag) = 22 _
                Or val(cmbSource.Tag) = 23 Or val(cmbSource.Tag) = 24 Or val(cmbSource.Tag) = 26 Or val(cmbSource.Tag) = 41 Then

                mNewProcessFlag = 1
            End If
        End If
        mNewProcessFlag = True
        If mNewProcessFlag = True Then
            txttreasury.Tag = 1
        Else
            txttreasury.Tag = 0
        End If
        
CheckValidation:
        If chkBFund.value = Unchecked Then
            If mNewProcessFlag Then
                msql = "SELECT intID,vchHeadDesc COLLATE DATABASE_DEFAULT  + '(  ' +suSourceOfFund.vchSourceFundName COLLATE DATABASE_DEFAULT + ' )'AS Sou  FROM faBudgetHead "
                msql = msql + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faBudgetHead.intSourceOfFund"
                msql = msql + " Where   ISNULL(tnyNewModeFlag,0)=1 AND intLBTypeID=" & gbLBType
                If val(cmbSource.Tag) = 26 Or val(cmbSource.Tag) = 41 Then
                    msql = msql + " And faBudgetHead.intSourceOfFund= " & val(cmbSource.Tag)
                End If
                gbSearchID = -1
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.QrySP = Qyery
                frmSearchMasters.SQLQry = msql
                frmSearchMasters.Show vbModal
                txtStateHead.SetFocus
            
            Else
            
                msql = "SELECT intID,vchHeadDesc COLLATE DATABASE_DEFAULT  + '(  ' +suSourceOfFund.vchSourceFundName COLLATE DATABASE_DEFAULT + ' )'AS Sou  FROM faBudgetHead "
                msql = msql + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faBudgetHead.intSourceOfFund"
                msql = msql + " Where   ISNULL(tnyNewModeFlag,0)=0 AND intLBTypeID=" & gbLBType
                If val(cmbSource.Tag) = 26 Or val(cmbSource.Tag) = 41 Then
                    msql = msql + " And faBudgetHead.intSourceOfFund= " & val(cmbSource.Tag)
                End If
                gbSearchID = -1
                frmSearchMasters.Connection = enuSourceString.Saankhya
                frmSearchMasters.QrySP = Qyery
                frmSearchMasters.SQLQry = msql
                frmSearchMasters.Show vbModal
                txtStateHead.SetFocus
            End If
        Else
            
            msql = "SELECT intHeadOfAccountID,vchHeadOfAccountCode + '  (  '+ vchHeadOfAccount + ' ) ' FROM faHeadOfAccount Where isnull(intSourceOfFundID,0)=3 And intLBTypeID=" & gbLBType
            gbSearchID = -1
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.QrySP = Qyery
            frmSearchMasters.SQLQry = msql
            frmSearchMasters.Show vbModal
            txtStateHead.SetFocus
        End If
        
    End Sub
''''    Private Sub cmdSearchStateHead_Click()
''''        Dim mSQL As String
''''        Dim mCnn    As New ADODB.Connection
''''        Dim Rec     As New ADODB.Recordset
''''        Dim objDB     As New clsDB
''''
''''
''''        Dim mArr As Variant
''''        Dim mCatID As Integer
''''        Dim mReqDate As Variant
''''        Dim mArrIn As Variant
''''
''''        If val(txtAmountRequested.Text) = 0 Then
''''            MsgBox "ENTER THE AMOUNT ", vbInformation
''''            txtAmountRequested.SetFocus
''''            Exit Sub
''''        End If
''''        Call CheckBAnkSource
''''        If mPreviousYearMode = 1 Then
''''             If val(cmbSource.Tag) = 4 Or val(cmbSource.Tag) = 2 Or val(cmbSource.Tag) = 3 _
''''                Or val(cmbSource.Tag) = 5 Or val(cmbSource.Tag) = 6 Or val(cmbSource.Tag) = 7 Or val(cmbSource.Tag) = 8 Or val(cmbSource.Tag) = 9 Or val(cmbSource.Tag) = 15 Or val(cmbSource.Tag) = 19 _
''''                Or val(cmbSource.Tag) = 20 Or val(cmbSource.Tag) = 22 _
''''                Or val(cmbSource.Tag) = 23 Or val(cmbSource.Tag) = 24 Then
''''
''''                GoTo CHECKVALIDATION::
''''            ElseIf val(cmbSource.Tag) = 29 Or val(cmbSource.Tag) = 30 Then
''''                mSQL = "Select * From faBankSourceChild"
''''                mSQL = mSQL + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
''''                mSQL = mSQL + " Where intSourceOfFundID = " & val(cmbSource.Tag)
''''                mSQL = mSQL + " AND tnyCategoryID= " & val(cmbCategory.Tag)
''''
''''            Else
''''
''''                mSQL = "Select * From faBankSourceChild"
''''                mSQL = mSQL + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
''''                mSQL = mSQL + " Where intSourceOfFundID = " & val(cmbSource.Tag)
''''
''''            End If
''''            If objDB.SetConnection(mCnn) Then
''''               Rec.Open mSQL, mCnn
''''               If Not (Rec.EOF And Rec.BOF) Then
''''                   If IsNull(Rec!tnyClosingFlag) Then
''''                        mNewProcessFlag = False
''''                   ElseIf Rec!tnyClosingFlag = 9 Then
''''                        mNewProcessFlag = True
''''                   Else
''''                        mNewProcessFlag = False
''''                   End If
''''               Else
''''
''''                    MsgBox "Treasury Balances are not Verified!"
''''                    cmdSave.Enabled = False
''''                    Exit Sub
''''               End If
''''
''''            Else
''''                MsgBox "Connection Failed", vbCritical
''''                Exit Sub
''''            End If
''''        End If
''''        If mNewProcessFlag = True Then
''''            txttreasury.Tag = 1
''''        Else
''''            txttreasury.Tag = 0
''''        End If
''''
''''CHECKVALIDATION:
''''        If chkBFund.value = Unchecked Then
''''            If mNewProcessFlag Then
''''                mSQL = "SELECT intID,vchHeadDesc COLLATE DATABASE_DEFAULT  + '(  ' +suSourceOfFund.vchSourceFundName COLLATE DATABASE_DEFAULT + ' )'AS Sou  FROM faBudgetHead "
''''                mSQL = mSQL + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faBudgetHead.intSourceOfFund"
''''                mSQL = mSQL + " Where   ISNULL(tnyNewModeFlag,0)=1 AND intLBTypeID=" & gbLBType
''''                gbSearchID = -1
''''                frmSearchMasters.Connection = enuSourceString.Saankhya
''''                frmSearchMasters.QrySP = Qyery
''''                frmSearchMasters.SQLQry = mSQL
''''                frmSearchMasters.Show vbModal
''''                txtStateHead.SetFocus
''''
''''            Else
''''
''''                mSQL = "SELECT intID,vchHeadDesc COLLATE DATABASE_DEFAULT  + '(  ' +suSourceOfFund.vchSourceFundName COLLATE DATABASE_DEFAULT + ' )'AS Sou  FROM faBudgetHead "
''''                mSQL = mSQL + " INNER JOIN suSourceOfFund On suSourceOfFund.intSourceFundID=faBudgetHead.intSourceOfFund"
''''                mSQL = mSQL + " Where   ISNULL(tnyNewModeFlag,0)=0 AND intLBTypeID=" & gbLBType
''''                gbSearchID = -1
''''                frmSearchMasters.Connection = enuSourceString.Saankhya
''''                frmSearchMasters.QrySP = Qyery
''''                frmSearchMasters.SQLQry = mSQL
''''                frmSearchMasters.Show vbModal
''''                txtStateHead.SetFocus
''''            End If
''''        Else
''''            mSQL = "SELECT intHeadOfAccountID,vchHeadOfAccountCode + '  (  '+ vchHeadOfAccount + ' ) ' FROM faHeadOfAccount Where isnull(intSourceOfFundID,0)=3 And intLBTypeID=" & gbLBType
''''            gbSearchID = -1
''''            frmSearchMasters.Connection = enuSourceString.Saankhya
''''            frmSearchMasters.QrySP = Qyery
''''            frmSearchMasters.SQLQry = mSQL
''''            frmSearchMasters.Show vbModal
''''            txtStateHead.SetFocus
''''        End If
''''
''''    End Sub

    Private Sub cmdSearchStateSubHead_Click()
        gbSearchID = -1
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "SELECT intHeadOfAccountID,vchHeadOfAccountCode + '  (  '+ vchHeadOfAccount + ' ) ' FROM faHeadOfAccount Where intLBTypeID=0 And isnull(intSourceOFFUndID,0)=0"
        frmSearchMasters.Show vbModal
        txtStateSubHead.SetFocus
    End Sub

    Private Sub cmdSearchTreasury_Click()
        gbSearchID = -1
        frmSearchMasters.Connection = enuSourceString.DBMaster
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "SELECT intTreasuryID, chvTreasuryCode+'-'+ chvTreasury FROM GM_Treasury"
        'frmSearchMasters.SQLQry = "SELECT chvTreasuryCode,chvTreasury,intTreasuryID FROM GM_Treasury"
        frmSearchMasters.Show vbModal
        txtTreasuryCode.Tag = gbSearchCode
        'gbSearchID = -1
        txttreasury.SetFocus
    End Sub
    Private Sub cmdSearchProject_Click()
    
'''        If chkUnAuthDrw.value = vbChecked Then
'''            frmGoDetails.Show vbModal
'''            If gbSearchID <> -1 Then
'''                txtProjectNo.Text = gbSearchCode
'''                txtProjectNo.Tag = gbSearchID
'''                txtProjName.Text = gbSearchStr
'''                gbSearchCode = ""
'''                gbSearchStr = ""
'''                gbSearchID = -1
'''            End If
'''            Exit Sub
'''        End If
    
        'If val(txtAmountRequested.Text) <> 0 Then
            'frmEstimationDetails.Mode = 0
            'frmSulekhaIntegration.Show vbModal
            frmSearchProjects.PreviousYearMode = 0
            frmSearchProjects.Show vbModal
            txtProjectNo.SetFocus
        'Else
            'MsgBox "Please Enter Amount", vbInformation
            'txtAmountRequested.SetFocus
        'End If

    End Sub

    Private Sub cmdSubSector_Click()
        If mLoadMode = 10 Then
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
        End If
    End Sub

    Private Sub Form_Load()
        Dim mExtractedStatus As Integer 'To Check Source Of Fund Opening Balance Extracted Status
        Dim mMsg As String
        
        'cmdSearchTreasury.Enabled = False
        Call formInitialise
        Call FillSource
        Call FillCategory
        
        If frmRequisition.RequisitionID <> "" Then
            Call FetchRequisitionDetails
        End If
        txtRequisitiontDate.Text = DdMmmYy(gbTransactionDate)
        'PREVIOUS YEAR REQUISITION
        If mPreviousYearMode Then
           
            Call GetPreviousYearRequestDetails
        Else
            mExtractedStatus = GetStatusFlag
            If mExtractedStatus <> 2 Then
                cmdNew.Enabled = False
                cmdSave.Enabled = False
                cmdApprove.Enabled = False
                mMsg = ""
                mMsg = mMsg + " Closing Balance Of Source Of Fund is " + vbCrLf
                mMsg = mMsg + " Either Not Brought Down  Or Approved " + vbCrLf
                mMsg = mMsg + " (Utility>>Annual Financial Statements-Finalization>>)"
                MsgBox mMsg, vbInformation
            End If
        End If
        
        '**************UNAUTHORIZED DRAWAL*********************
        
        If mLoadMode = 10 Then
            chkUnAuthDrw.Visible = True
            chkUnAuthDrw.value = vbChecked
            chkUnAuthDrw_Click
            chkUnAuthDrw.Enabled = False
        End If
        
        '******************************************************
        
        
        '********************REQUISITION INBOX****************
        If mLoadMode = 20 Then
            cmdNew.Enabled = False
            Call GetRequisitionInboxDetails
        End If
        '*****************************************************
        '''For SaankhyaWeb Updation
        If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
        
            cmdSave.Enabled = True
        Else
        
            cmdSave.Enabled = False
        End If
    End Sub
    Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim msql  As String
        Dim mTrAccHeadId As Integer
        
        If objDb.SetConnection(mCnn) Then
            msql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open msql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                
                'NOTE: Checking in Previous Year
                '      IF APPROVED tnyStatus will be 0 ELSE NULL
                Rec.Close
                msql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID - 1
                Rec.Open msql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If IsNumeric(Rec!tnyStatus) Then
                        If Rec!tnyStatus = 0 Then
                            GetStatusFlag = 9
                        Else
                            GetStatusFlag = -1
                        End If
                    Else
                        GetStatusFlag = -1
                    End If
                Else
                    GetStatusFlag = -1
                End If
            End If
            If Rec.State = 1 Then
                Rec.Close
            End If
        End If
    End Function


    Private Sub Form_Unload(Cancel As Integer)
        If mPreviousYearMode Then
            mPreviousYearMode = -1
            If mLoadMode = 10 Then
                mLoadMode = -1
            End If
            Unload Me
            
        '************UNAUTHORIZED DRAWAL**********************
        ElseIf mLoadMode = 10 Then
            mLoadMode = -1
            Unload Me
            frmListOfRequisitions.LoadMode = 10
        '************END**************************************
        
        '************REQUISITION INBOX************************
        ElseIf mLoadMode = 20 Then
            mLoadMode = -1
            Unload Me
            
        '************END**************************************
        
        
        Else
            frmListOfRequisitions.Visible = True
            frmListOfRequisitions.ZOrder (0)
        End If
        Set frmRequisition = Nothing
    End Sub



    Private Sub MaskAccHead_KeyPress(KeyAscii As Integer)
         If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
                KeyAscii = 0
         End If
    End Sub
    Private Sub MaskDetailAccHead_KeyPress(KeyAscii As Integer)
'         If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
'                KeyAscii = 0
'         End If
    End Sub

    Private Sub txtAmountRequested_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub CheckBAnkSource()
        Dim msql As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb     As New clsDB
        Dim mArrIN As Variant
        
         '==================================================================
        'INSERT INTO BANK SOURCE And BANK SOURCE CHILD
        '==================================================================
        msql = ""
        msql = "DELETE FROM faBankSource"
        objDb.ExecuteSP msql, , , , mCnn, adCmdText

        msql = " INSERT INTO faBankSource"
        msql = msql + " SELECT faAccountHeads.intAccountHeadID,faAccountHeads.vchAccountHeadCode,ISNULL(A.fltAmount,0) fltAmount,2"
        msql = msql + " From"
        msql = msql + " ("
        msql = msql + " SELECT faTransactionChild.intAccountHeadID,vchAccountHeadCode, SUM(fltAmount*((tinDebitOrCreditFlag*2)-1)) fltAmount,2 tnyFlag"
        msql = msql + " From faTransactionChild"
        msql = msql + " left JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=faTransactionChild.intAccountHeadID"
        msql = msql + " INNER JOIN faTransactions ON faTransactions.intTRansactionID=faTransactionChild.intTRansactionID"
        msql = msql + " Where"
        If gbLBPanchayat = 1 Then
            msql = msql + " vchAccountHeadCode LIKE '450650%'"
        Else
            msql = msql + " vchAccountHeadCode LIKE '450650%'"
        End If
        
        msql = msql + " AND faAccountHeads.intGroupID=2"
        msql = msql + " And ( tnyStatus <> 4 Or tnyStatus is Null )"
        'msql = msql + " AND faTransactions.dtTransactionDate < '01-Apr-2015'"   ''' Commented on 27/Nov/2015 for emergency purpose
        msql = msql + " Group by faTransactionChild.intAccountHeadID,vchAccountHeadCode)"
        msql = msql + " A"
        msql = msql + " RIGHT JOIN faAccountHeads ON faAccountHeads.intAccountHeadID=A.intAccountHeadID"
        msql = msql + " Where"
        If gbLBPanchayat = 1 Then
            msql = msql + " faAccountHeads.vchAccountHeadCode LIKE '450650%'"
        Else
            msql = msql + " faAccountHeads.vchAccountHeadCode LIKE '450650%'"
        End If
        
        
        objDb.ExecuteSP msql, , , , mCnn, adCmdText
        
        
        objDb.ExecuteSP "spSetClosingFlag", , , , mCnn

        msql = ""
        msql = "DELETE FROM faBankSourceChild"
        objDb.ExecuteSP msql, , , , mCnn, adCmdText
        
        msql = ""
        msql = " SELECT * FROM faBankSource"
        If objDb.SetConnection(mCnn) Then
           Rec.Open msql, mCnn
           While Not (Rec.EOF Or Rec.BOF)
                mArrIN = Array(Rec!intBankID)
                objDb.ExecuteSP "spSaveBankSourceChild", mArrIN, , , mCnn
                Rec.MoveNext
          Wend
          Rec.Close
        End If
        

        If gbLBPanchayat = 1 Then
            msql = ""
            msql = " UPDATE faBAnkSourceChild SET tnyCategoryID=2 WHERE intBankID IN (1494)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            msql = " UPDATE faBAnkSourceChild SET tnyCategoryID=3 WHERE intBankID IN (1495)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            msql = " UPDATE faBAnkSourceChild SET tnyCategoryID=1 WHERE intBankID IN (1418,1490,1491,1492,1493)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
        
        Else
            msql = ""
            msql = "UPDATE faBAnkSourceChild SET tnyCategoryID=2 WHERE intBankID IN (1816)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            msql = "UPDATE faBAnkSourceChild SET tnyCategoryID=3 WHERE intBankID IN (1817)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
            msql = "UPDATE faBAnkSourceChild SET tnyCategoryID=1 WHERE intBankID IN (1512,1535,1539,1755,1756)"
            objDb.ExecuteSP msql, , , , mCnn, adCmdText
        End If
        '==================================================================
        msql = ""
        msql = "IF NOT EXISTS(SELECT * FROM faBAnkSourceChild WHERE intSourceOfFundID=29 AND tnyCategoryID=3)"
        msql = msql + " Begin"
        msql = msql + "     If Exists (Select * From faLBSettings where tnyLBTypeID in (3,4) )"
        msql = msql + "         INSERT INTO faBAnkSourceChild VALUES(1816,29,0,3)"
        
        msql = msql + "     Else"
        msql = msql + "     INSERT INTO faBAnkSourceChild VALUES(1494,29,0,3)"
        msql = msql + " End"
        objDb.ExecuteSP msql, , , , mCnn, adCmdText
        
        
    End Sub

    Private Sub txtAmountRequested_LostFocus() ''''''''''Added by poornima on 13/10/2011
        Dim msql As String
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDb     As New clsDB
        Dim mTreasuryID As Integer
        
        Dim mArr As Variant
        Dim mCatID As Integer
        Dim mReqDate As Variant
        Dim mArrIN As Variant
        
        If val(txtAmountRequested.Text) <= 0 Then
            txtAmountRequested.Text = ""
            Exit Sub
        End If
        
        If cmbCategory.ListIndex > -1 Then
            mCatID = cmbCategory.ItemData(cmbCategory.ListIndex)
        End If
        If IsDate(txtRequisitiontDate.Text) Then
            mReqDate = txtRequisitiontDate.Text
        Else
            mReqDate = gbTransactionDate
        End If
        mTreasuryID = -1
        
        '================================================================'
        ' 2015 MODIFICATION CONSOLIDATED FUND
        '================================================================'
        If chkBFund.value = vbUnchecked And mLoadMode <> 10 And chkNonPlan.value = Unchecked Then
            If val(cmbSource.Tag) = 0 Then
                MsgBox "Select the Project and Source of Fund Please!", vbInformation
                txtAmountRequested.Text = ""
                Exit Sub
            End If
            '5,6,7,19,20,21,22,23,4,2,3,25,26,9,8,15 Or val(cmbSource.Tag) = 25 Or val(cmbSource.Tag) = 26
            If val(cmbSource.Tag) = 4 Or val(cmbSource.Tag) = 2 Or val(cmbSource.Tag) = 3 _
                Or val(cmbSource.Tag) = 5 Or val(cmbSource.Tag) = 6 Or val(cmbSource.Tag) = 7 Or val(cmbSource.Tag) = 8 Or val(cmbSource.Tag) = 9 Or val(cmbSource.Tag) = 15 Or val(cmbSource.Tag) = 19 _
                Or val(cmbSource.Tag) = 20 Or val(cmbSource.Tag) = 22 _
                Or val(cmbSource.Tag) = 23 Or val(cmbSource.Tag) = 24 Then   '  NO AMOUNT VALIDATION  Or val(cmbSource.Tag) = 21
                    Call CheckProjAmtValidation
                    Exit Sub
            End If
        Else
            Exit Sub
        End If
        
        '==================================================================
        'INSERT INTO BANK SOURCE And BANK SOURCE CHILD
        '==================================================================
        Call CheckBAnkSource
        
        ' =================================================================
        ' CHECKING TREASURY BALANCE and NEW MODE OF ALLOTMENT -
        ' =================================================================
        If val(cmbSource.Tag) = 29 Or val(cmbSource.Tag) = 30 Then
            msql = "Select * From faBankSourceChild"
            msql = msql + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
            msql = msql + " Where intSourceOfFundID = " & val(cmbSource.Tag)
            msql = msql + " AND tnyCategoryID= " & val(cmbCategory.Tag)
            
        Else
            msql = "Select * From faBankSourceChild"
            msql = msql + " INNER JOIN faBankSource On faBankSource.intBankID = faBankSourceChild.intBankID"
            msql = msql + " Where intSourceOfFundID = " & val(cmbSource.Tag)
            'mSql = mSql + " AND tnyCategoryID= " & val(cmbCategory.Tag)
        End If
        If objDb.SetConnection(mCnn) Then
           Rec.Open msql, mCnn
           If Not (Rec.EOF And Rec.BOF) Then
               mTreasuryID = Rec!intBankID
               txtTreasuryBalance.Tag = Rec!intBankID
               fltClosingTreasuryBalance = Rec!fltClosingBalance
               If IsNull(Rec!tnyClosingFlag) Then
                    mNewProcessFlag = False
               ElseIf Rec!tnyClosingFlag = 9 Then
                    mNewProcessFlag = True
               Else
                    mNewProcessFlag = False
               End If
           Else
           
                MsgBox "Treasury Balances are not Verified!"
                cmdSave.Enabled = False
                Exit Sub
           End If
           
        Else
            MsgBox "Connection Failed", vbCritical
            Exit Sub
        End If
        
        If mNewProcessFlag Then
            chkNewMode.value = 1
        Else
            chkNewMode.value = 0
            ' =================================================================
            ' TREASURY BALANCE
            ' =================================================================
            mArr = Array(mTreasuryID, mReqDate)
            Dim mRec As New ADODB.Recordset
            Dim mtreasuryBalance As Double
            Dim AllotAmt    As Double
            Dim mBalance    As Double
            If Rec.State Then Rec.Close
            
            Set Rec = objDb.ExecuteSP("spGetLedgerBalance", mArr, , True, mCnn, adCmdStoredProc)
            If Not (Rec.EOF And Rec.BOF) Then
                txtTreasuryBalance.Text = Rec!Balance
                If Rec!Balance < val(txtAmountRequested.Text) Then
                    'MsgBox "Only Rs." & Format(Rec!Balance, "0.00") & " is Available in Treasury", vbInformation
                    
                    '''''''--------  Modified On 30 Nov 2015 By Anisha with the help of Anju R
                    mtreasuryBalance = Rec!Balance
                    
                    Select Case val(cmbSource.Tag)
                    Case 1, 10, 11, 12, 13, 14, 21, 27, 28:
                        msql = " Select sum(fltRequestedAmt) alltAmt From faAllotments"
                        msql = msql + " Left Join faPayOrder On  faAllotments.intID=faPayOrder.intAllotmentID"
                        msql = msql + " Where faAllotments.intFinancialYearID = 2015 And tnyStage = 2 And faAllotments.tnyStatus = 1 "
                        msql = msql + " And faAllotments.intSourceID in (1,21,10,11,12,13,14,27,28)"
                        msql = msql + " And faPayOrder.intVoucherID is Null"
                    Case 16, 17:
                        msql = " Select sum(fltRequestedAmt) alltAmt From faAllotments"
                        msql = msql + " Left Join faPayOrder On  faAllotments.intID=faPayOrder.intAllotmentID"
                        msql = msql + " Where faAllotments.intFinancialYearID = 2015 And tnyStage = 2 And faAllotments.tnyStatus = 1 "
                        msql = msql + " And faAllotments.intSourceID in (16,17)"
                        msql = msql + " And faPayOrder.intVoucherID is Null"
                    Case Else:
                        msql = " Select sum(fltRequestedAmt) alltAmt From faAllotments"
                        msql = msql + " Left Join faPayOrder On  faAllotments.intID=faPayOrder.intAllotmentID"
                        msql = msql + " Where faAllotments.intFinancialYearID = 2015 And tnyStage = 2 And faAllotments.tnyStatus = 1 "
                        msql = msql + " And faAllotments.intSourceID = " & val(cmbSource.Tag)
                        msql = msql + " And faPayOrder.intVoucherID is Null"
                    End Select
                    
                    Set mRec = objDb.ExecuteSP(msql, , , True, mCnn, adCmdText)
                    If Not (mRec.EOF And mRec.BOF) Then
                        AllotAmt = IIf(IsNull(mRec!alltAmt), 0, mRec!alltAmt)
                    End If
                    
                    
                    
                    mBalance = mtreasuryBalance - AllotAmt
                    If AllotAmt < 1 Then
                        MsgBox "Only Rs." & Format(Rec!Balance, "0.00") & " is Available in Treasury", vbInformation
                        txtAmountRequested.Text = Format(Rec!Balance, "0.00")
                    Else
                    
                        msql = "Only Rs." & Format(Rec!Balance, "0.00") & " is Available in Treasury" & vbNewLine
                        msql = msql + " And Pending Requestition Amount is " & Format(AllotAmt, "0.00") & vbNewLine
                        msql = msql + " So Available balance is " & Format(mBalance, "0.00")
                        MsgBox msql, vbInformation
                        txtAmountRequested.Text = Format(mBalance, "0.00")
                    End If
                    txtProjectNo.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "No sufficient Balance Available in Treasury", vbInformation
                txtProjectNo.SetFocus
                Exit Sub
            End If
        End If
        
        
        ' =================================================================
        ' CHECKING WITH TREASURY BALANCE
        ' =================================================================
        msql = ""
'''        msql = msql + " SELECT fltClosingBalance, fltAmountIssued,  (fltClosingBalance - fltAmountIssued) fltBalanceAmount,  intSourceOfFundID FROM ( "
'''        msql = msql + " SELECT fltClosingBalance, fltAmountIssued, intSourceOfFundID FROM ( "
'''        msql = msql + " SELECT Isnull(fltClosingBalance,0) fltClosingBalance, intSourceOfFundID  FROM faBankSourceChild "
'''        msql = msql + " INNER JOIN  faBankSource ON  faBankSource.intBankID = faBankSourceChild.intBankID "
'''        msql = msql + " Where intSourceOfFundID =  " & val(cmbSource.Tag)
'''        msql = msql + " ) A LEFT JOIN "
'''        msql = msql + " ( "
'''        msql = msql + " Select ISNULL(Sum(fltRequestedAmt),0) fltAmountIssued, intSourceID From faAllotments "
'''        msql = msql + " INNER JOIN faBankSourceChild ON faBankSourceChild.intSourceOfFundID = faAllotments.intSourceID "
'''        msql = msql + " INNER JOIN faBankSource ON faBankSource.intBankID = faBankSourceChild.intBankID "
'''        msql = msql + " Where intFinancialYearID = " & gbFinancialYearID
'''        msql = msql + " And tnyStage = 2 And tnyStatus = 1 And intSourceID = " & val(cmbSource.Tag)
'''        msql = msql + " Group by intSourceID "
'''        msql = msql + " ) B ON A.intSourceOfFundID = B.intSourceID "
'''        msql = msql + " ) C "
        
        
        ''''''Added by Anisha On 27 Nov 2015
         Select Case val(cmbSource.Tag)
         Case 1, 10, 11, 12, 13, 14, 21, 27, 28:
                msql = " Select sum(fltClosingBalance) fltClosingBalance,sum(fltAmountIssued) fltAmountIssued ,"
                msql = msql + " sum(fltClosingBalance) -sum(fltAmountIssued) fltBalanceAmount"
                msql = msql + " From ("
                msql = msql + " SELECT sum(Isnull(fltAmount,0)) fltClosingBalance ,0 fltAmountIssued"
                msql = msql + " FROM faBankSourceChild  INNER JOIN  faBankSource ON  faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intSourceOfFundID in (1,10,11,12,13,14,21,27,28)" '''''""''' --and faBankSource.intBankID=" & Rec!intBankID
                msql = msql + " Union All"
                msql = msql + " SELECT sum(Isnull(fltAmount,0)) fltClosingBalance ,0 fltAmountIssued"
                msql = msql + " From faAllotmentLetters"
                msql = msql + " WHERE ISNULL(tnyStatus,0) = 1 AND intSourceOfFundID in (1,10,11,12,13,14,21,27,28)"
                msql = msql + " and intFinancialYearID = 2015"
                msql = msql + " AND ISNULL(tnyGroupID,0) =30"
                msql = msql + " Union All"
                msql = msql + " Select 0 as fltClosingBalance,ISNULL(Sum(fltRequestedAmt),0) fltAmountIssued"
                msql = msql + " From faAllotments  INNER JOIN faBankSourceChild ON faBankSourceChild.intSourceOfFundID = faAllotments.intSourceID"
                msql = msql + "     AND faAllotments.intFundCategoryID = faBankSourceChild.tnyCategoryID "
                msql = msql + " INNER JOIN faBankSource ON faBankSource.intBankID = faBankSourceChild.intBankID "
                msql = msql + " Where intFinancialYearID = 2015 And tnyStage = 2 And tnyStatus = 1 And intSourceID in (1,10,11,12,13,14,21,27,28)"
                msql = msql + " AND ISNULL(intTreasuryID,0)= 0 "
                msql = msql + " )A"

        Case 16, 17:
               msql = " Select sum(fltClosingBalance) fltClosingBalance,sum(fltAmountIssued) fltAmountIssued ,"
                msql = msql + " sum(fltClosingBalance) -sum(fltAmountIssued) fltBalanceAmount"
                msql = msql + " From ("
                msql = msql + " SELECT sum(Isnull(fltAmount,0)) fltClosingBalance ,0 fltAmountIssued"
                msql = msql + " FROM faBankSourceChild  INNER JOIN  faBankSource ON  faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intSourceOfFundID in (16,17)" '''''""''' --and faBankSource.intBankID=" & Rec!intBankID
                msql = msql + " Union All"
                msql = msql + " Select 0 as fltClosingBalance,ISNULL(Sum(fltRequestedAmt),0) fltAmountIssued"
                msql = msql + " From faAllotments  INNER JOIN faBankSourceChild ON faBankSourceChild.intSourceOfFundID = faAllotments.intSourceID"
                msql = msql + " INNER JOIN faBankSource ON faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intFinancialYearID = 2015 And tnyStage = 2 And tnyStatus = 1 And intSourceID in (16,17)"
                msql = msql + " AND ISNULL(intTreasuryID,0)= 0 "
                msql = msql + " )A"
        Case Else:
             msql = " Select sum(fltClosingBalance) fltClosingBalance,sum(fltAmountIssued) fltAmountIssued ,"
                msql = msql + " sum(fltClosingBalance) -sum(fltAmountIssued) fltBalanceAmount"
                msql = msql + " From ("
                msql = msql + " SELECT sum(Isnull(fltAmount,0)) fltClosingBalance ,0 fltAmountIssued"
                msql = msql + " FROM faBankSourceChild  INNER JOIN  faBankSource ON  faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intSourceOfFundID =" & val(cmbSource.Tag) '''''""''' --and faBankSource.intBankID=" & Rec!intBankID
                msql = msql + " Union All"
                msql = msql + " Select 0 as fltClosingBalance,ISNULL(Sum(fltRequestedAmt),0) fltAmountIssued"
                msql = msql + " From faAllotments  INNER JOIN faBankSourceChild ON faBankSourceChild.intSourceOfFundID = faAllotments.intSourceID"
                msql = msql + " INNER JOIN faBankSource ON faBankSource.intBankID = faBankSourceChild.intBankID"
                msql = msql + " Where intFinancialYearID = 2015 And tnyStage = 2 And tnyStatus = 1 And intSourceID =" & val(cmbSource.Tag)
                msql = msql + " AND ISNULL(intTreasuryID,0)= 0 "
                msql = msql + " )A"
        
        End Select
        
        If mNewProcessFlag = False Then
            txttreasury.Tag = 0 ' TREASURY TRANSACTION is Processed in Previous Years
            If objDb.SetConnection(mCnn) Then
               If Rec.State Then Rec.Close
               Rec.Open msql, mCnn
               If Not (Rec.EOF And Rec.BOF) Then
                    If Abs(Rec!fltBalanceAmount) < val(txtAmountRequested.Text) Then
                        MsgBox "No Balance In Treasury", vbInformation
                        cmdSave.Enabled = False
                        Exit Sub
                    Else
                        cmdSave.Enabled = True
                    End If
               End If
               Rec.Close
            End If
        Else
            txttreasury.Tag = 1  ' NEW MODE OF TREASURY TRANSACTION
        End If
        
       
        
        ' =================================================================
        ' CHECKING APROPRIATION CONGROL REGISTER
        ' =================================================================
        If mCatID = 0 Then
            MsgBox "Please Select Category"
            Exit Sub
        End If
        If mPreviousYearMode = 1 Then
            mArr = Array(val(cmbSource.Tag), mCatID, mReqDate, gbFinancialYearID - 1)
        Else
        mArr = Array(val(cmbSource.Tag), mCatID, mReqDate)
        End If
        If objDb.SetConnection(mCnn) Then
            
            Set Rec = objDb.ExecuteSP("spCheckACRBalance", mArr, , True, mCnn, adCmdStoredProc)
            If val(cmbSource.Tag) <> 2 Then
                If Not (Rec.EOF And Rec.BOF) Then
                    
                        'MsgBox Rec!fltBalance
                        If Rec!fltBalance < val(txtAmountRequested) Then
                            MsgBox "No Balance Available in ACR [ NEW-ACR ].Only Rs." & Format(Rec!fltBalance, "0.00") & " Available", vbInformation
                            cmdSave.Enabled = False
                            Exit Sub
                        End If
                    
                Else
                    MsgBox "No Balance Available in ACR [ NEW-ACR ]", vbCritical
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
        End If
        
        cmdSave.Enabled = True
        txtProjectNo.SetFocus
        
        If mLoadMode <> 10 Then     '******************FOR UNAUTHORIZED DRAWAL************
            If val(txtAmountRequested.Text) > 0 Then
                If IsNumeric(txtProjCost.Tag) Then
                    If val(txtAmountRequested.Text) > val(txtProjCost.Tag) Then
                        msql = "Total Amount allocated for " & cmbSource.Text & " in this Project" & vbCrLf
                        msql = msql + " is Rs. " & Format(val(txtProjCost.Tag), "0.00")
                        MsgBox msql, vbInformation
                        txtAmountRequested.Text = Format(val(txtProjCost.Tag), "0.00")
                        txtAmountRequested.SetFocus
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub txtAmountRequested_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
        txtAmountRequested.Locked = True
    Else
        txtAmountRequested.Locked = False
    End If
    End Sub

    Private Sub txtAuthorizedAmt_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub



    Private Sub txtIMPOName_GotFocus()
        If gbSearchID > 0 Then
            Dim objSubLedger As New clsSubLedger
            objSubLedger.SetSubLedgerDetails (gbSearchID)
            If objSubLedger.SubsidiaryAccountHeadID Then
                txtIMPOName.Tag = IIf(IsNull(objSubLedger.SubsidiaryAccountHeadID), 0, objSubLedger.SubsidiaryAccountHeadID)
                txtIMPOName.Text = IIf(IsNull(objSubLedger.NameOfSubLedger), "", objSubLedger.NameOfSubLedger)
                txtDesig.Text = IIf(IsNull(objSubLedger.Designation), "", objSubLedger.Designation)
                txtDept.Text = objSubLedger.Department
                txtDDOCode.Text = objSubLedger.DDOCode 'objSubLedger.SubTitle
                If txtDDOCode.Text <> "" Then
                    SetFunctionary
                End If
                If txtFunctionary.Tag = "" And mLoadMode <> 10 Then
                    cmdSearchFunctionary.Visible = True
                End If
    '            txtAmountRequested.Text = objSubLedger.
            Else
                txtIMPOName.Tag = ""
                txtIMPOName.Text = ""
                txtDesig.Text = ""
                txtDept.Text = ""
                txtDDOCode.Text = ""
            End If
        End If
        gbSearchID = -1
        '---ADDED FOR GM_Treasury----
        Dim mCnnTR    As New ADODB.Connection
        Dim RecTr     As New ADODB.Recordset
        Dim objTr     As New clsDB
        Dim mSQLTR    As String
        
        objTr.CreateNewConnection mCnnTR, enuSourceString.DBMaster
        mSQLTR = "select * from GM_Treasury "
        RecTr.Open mSQLTR, mCnnTR
        If Not (RecTr.EOF And RecTr.BOF) Then
            cmdSearchTreasury.Visible = True
        End If
        RecTr.Close
        
    End Sub
    
    Private Sub SetAccountHeads(mMicroSectorID As Integer)
                    Dim mCnn    As New ADODB.Connection
                    Dim mCnPlan As New ADODB.Connection
                    
                    Dim Rec     As New ADODB.Recordset
                    Dim obj     As New clsDB
                    Dim msql    As String
                    Dim mWHERE  As String
                    
                    Dim mCapitalExpFlag As Boolean
                    Dim mMicroSectorCount As Integer
                    Dim mMicroHeads As Integer
                                                                
                    msql = msql + " SELECT faMicroSectorHeads.intCategoryID,"
                    msql = msql + " vchTransactionCategory,"
                    msql = msql + " intSubSectorID,"
                    msql = msql + " vchSubSector,"
                    msql = msql + " vchSubSectorEng,"
                    msql = msql + " faMicroSectorHeads.intMircoSectorID,"
                    msql = msql + " suMicroSectors.vchMicroSecCode,"
                    msql = msql + " suMicroSectors.vchMicroSector,"
                    msql = msql + " suMicroSectors.vchEngMicroSector,"
                    msql = msql + " faMicroSectorHeads.intAccountHeadID,"
                    msql = msql + " faAccountHeads.vchAccountHeadCode,"
                    msql = msql + " vchAccountHead,"
                    msql = msql + " faMicroSectorHeads.intFunctionID,"
                    msql = msql + " faFunctions.vchFunctionCode, vchFunction,"
                    msql = msql + " faFunctionaryFunctions.intFunctionaryID,"
                    msql = msql + " vchFunctionary,"
                    msql = msql + " faMicroSectorHeads.intTransactionTypeID ,"
                    msql = msql + " vchTransactionType"
                    msql = msql + " From faMicroSectorHeads"
                    msql = msql + " INNER JOIN suMicroSectors ON suMicroSectors.intMicroSecID = faMicroSectorHeads.intMircoSectorID"
                    msql = msql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faMicroSectorHeads.intAccountHeadID"
                    msql = msql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faMicroSectorHeads.intFunctionID"
                    msql = msql + " LEFT JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
                    msql = msql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID"
                    msql = msql + " INNER JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faMicroSectorHeads.intCategoryID"
                    msql = msql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faMicroSectorHeads.intTransactionTypeID"
                    msql = msql + " INNER JOIN faSubSector ON faSubSector.intSubSecID = suMicroSectors.intSubSecID"
                    msql = msql + " Where faMicroSectorHeads.intMircoSectorID = " & mMicroSectorID
                    msql = msql + " And faMicroSectorHeads.intCategoryID = " & val(cmbCategory.Tag)
                    If obj.SetConnection(mCnn) Then
                        Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly
                        If Not (Rec.EOF And Rec.BOF) Then
                        
                            txtSubSector.Text = Rec!vchSubSectorEng
                            txtSubSector.Tag = Rec!intSubSectorID
                            
                            txtMicroSector.Text = Rec!vchEngMicroSector
                            txtMicroSector.Tag = Rec!intMircoSectorID
                            
                            
                            cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                            If val(cmbSource.Tag) = 10 Or val(cmbSource.Tag) = 11 _
                             Or val(cmbSource.Tag) = 12 Or val(cmbSource.Tag) = 13 Or val(cmbSource.Tag) = 14 Then
                                cmbCategory.Enabled = True
                            Else
                                cmbCategory.Enabled = False
                            End If
                            
                            txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                            txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                            
                            txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                            txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                                                        
                            txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                            txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                            
                            If mLoadMode = 10 Then
                                cmdSearchFunctionary.Enabled = True
                                cmdSearchHead.Enabled = True
                                cmdSearchFunction.Enabled = True
                            Else
                                cmdSearchFunctionary.Enabled = False
                                cmdSearchHead.Enabled = False
                                cmdSearchFunction.Enabled = False
                            End If
                            
                        Else
                            cmbCategory.ListIndex = -1
                            
                            txtSubSector.Text = ""
                            txtSubSector.Tag = ""
                            
                            txtMicroSector.Text = ""
                            txtMicroSector.Tag = ""
                            
                            
                            'cmbCategory.Text = ""
                            txtFunction.Tag = ""
                            txtFunction.Text = ""
                            cmdSearchFunction.Enabled = True
                    
                            txtAccountHeadCode.Tag = ""
                            txtAccountHeadCode.Text = ""
                            txtAccountHead.Text = ""
                            'cmdSearchHead.Enabled = False
                            cmdSearchHead.Enabled = True
                            
                            txtFunctionary.Tag = ""
                            txtFunctionary.Text = ""
                            
                            cmdSearchFunction.Enabled = True
                            cmdSearchFunctionary.Enabled = True
                    
                        End If
                        Rec.Close
                    End If
    End Sub
    Private Sub SetSchemeDetails(mSchemeID As Integer)
        Dim mCnn  As New ADODB.Connection
        Dim objDb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim msql  As String
    
           
        If objDb.SetConnection(mCnn) Then
            msql = " Select * from faDepSchPro "
            msql = msql + " Where tnyMapping= " & mSchemeID & " "
            Rec.Open msql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                 txtScheme.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                 txtScheme.Tag = IIf(IsNull(Rec!intID), "", Rec!intID)
            End If
            Rec.Close
        End If
    End Sub
    Private Sub SetProjectDetails(objProject As clsProject, mSourceOfFundID As Integer)
        Dim objProj As New clsProject
        Dim objProFund As New clsProjectFund
        Dim mProjectID As Variant
        Dim mSubsectorID As Integer
        Dim mintCategoryID As Integer
        Dim mCol As Collection
        Dim mRow As Integer
        Dim mCnPlan As New ADODB.Connection
        Dim obj As New clsDB
        Dim msql As String
        Dim Rec As New ADODB.Recordset
        Dim mMicroSectorCount As Integer
        'Dim mSql As String
        
        If objProject.ProjectID > 0 And objProject.YearID > 2012 Then
               
                txtProjName.Text = objProject.ProjectNameEnglish
                txtProjectNo.Text = objProject.ProjectSerialNo
                txtProjectNo.Tag = objProject.ProjectID
                cmbCategory.Tag = objProject.ProjCatID
                mintCategoryID = objProject.ProjCatID
                If mSourceOfFundID = 26 And objProject.YearID = 2016 Then 'KLGSDP Fund Modified on 20/12/16
                    
                    msql = "SELECT vchSourceFundName,intSourceFundID from suSourceOfFund Where intSourceFundID in (26,41)"
                    PopulateList cmbSource, msql, , False, True, True, enuSourceString.Saankhya
                    cmbSource.Enabled = True
                Else
                    cmbSource.Text = objProject.FindSourceOfFund(mSourceOfFundID)
                    cmbSource.Tag = objProject.SourceOfFundID
                    cmbSource.Enabled = False
                End If
                mSubsectorID = objProject.SubSectorID
                txtDPCNo.Tag = mSubsectorID
                txtSubSector.Tag = mSubsectorID
                
                If val(objProject.SchemeID) <> 0 Then
                    chkBFund.Enabled = False
                    lblScheme.Visible = True
                    txtScheme.Visible = True
                    txtScheme.Tag = objProject.SchemeID
                    SetSchemeDetails (val(txtScheme.Tag))
                Else
                    chkBFund.Enabled = True
                    lblScheme.Visible = False
                    txtScheme.Visible = False
                    txtScheme.Tag = ""
                End If
                
                On Error GoTo SkipSulekha:
                'BLOCK [1] Note:- Calculating Project's Total Cost
                'Set mCol = objProj.GetFundDetails(CInt(gbFinancialYearID), objProject.ProjectID)
                Set mCol = objProj.GetFundDetails(CInt(objProject.YearID), objProject.ProjectID) ' CHANGED ON 11-JUN-2014 :: AIBY
                
                For mRow = 1 To mCol.count
                    Set objProFund = mCol.Item(mRow)
                    If objProFund.SourceOfFundID = mSourceOfFundID Then
                        txtProjCost.Text = objProFund.SourceWiseAmount
                        txtProjCost.Tag = objProFund.SourceWiseAmount
                        Exit For
                    End If
                Next mRow
                'END OF BLOCK [1]
                
                'BLOCK [2]
                    ' EXTRACTING MicrosectorIDs From Plan-Project
                    If obj.CreateNewConnection(mCnPlan, enuSourceString.Sulekha) Then
                        msql = " SELECT MicroSector.intMicroSecID  FROM MicroSector WHERE decProjectID = " & val(txtProjectNo.Tag)
                        msql = msql + " AND intYearID = " & objProject.YearID
                        Rec.Open msql, mCnPlan, adOpenStatic, adLockReadOnly
                        mMicroSectorCount = 0
                        If Not (Rec.BOF And Rec.EOF) Then
                            mWHERE = ""
                            While Not Rec.EOF
                                mMicroSectorCount = mMicroSectorCount + 1
                                If Len(mWHERE) > 0 Then
                                    mWHERE = mWHERE & ", " & Rec!intMicroSecID
                                    txtDPCDate.Tag = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                                    txtMicroSector.Tag = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                                Else
                                    mWHERE = Trim(str(Rec!intMicroSecID))
                                End If
                                Rec.MoveNext
                            Wend
                        End If
                        
                        Rec.Close
                    End If
                    If mMicroSectorCount = 1 Then
                        Call SetAccountHeads(val(mWHERE))
                        cmdMicroSector.Enabled = False
                        cmdSubSector.Enabled = False
                    End If
                'END OF BLOCK [2]
        End If
        Exit Sub
SkipSulekha:
        If mFundErSulekha = 1 Then
            MsgBox "Some Mistakes in Ported data from Sulekha"
        End If
    End Sub


Private Sub txtMicroSector_GotFocus()
    If gbSearchID > 0 Then
        Call SetAccountHeads(val(gbSearchID))
        gbSearchID = -1
        gbSearchCode = ""
        gbSearchStr = ""
    End If
End Sub

    Private Sub txtProjectNo_GotFocus()
        'On Error GoTo Err
        Dim objProj As New clsProject
        Dim objProFund As New clsProjectFund
        Dim mProjectID As Variant
        Dim mSourceOfFundID As Variant
        Dim mSubsectorID As Integer
        Dim mintCategoryID As Integer
        Dim mCol As Collection
        Dim mRow As Integer
        
        Dim mYearID As Integer
        
        mProjectID = gbSearchStr
        mSourceOfFundID = gbSearchID
        
        gbSearchCode = ""
        gbSearchStr = ""
        gbSearchID = -1
        
        If mPreviousYearMode = 1 Then
            mProjectID = txtProjectNo.Tag
            mSourceOfFundID = cmbSource.Tag
            mYearID = gbFinancialYearID - 1
        End If
        
        If val(mProjectID) > 0 Then
            If mPreviousYearMode = 1 Then
                objProj.SetProject mProjectID, mYearID
            Else
                objProj.SetProject mProjectID
            End If
            
            cmbCategory.ListIndex = -1
            cmbCategory.Tag = ""
            
            txtSubSector.Text = ""
            txtSubSector.Tag = ""
            cmdSubSector.Tag = ""
            
            txtMicroSector.Text = ""
            txtMicroSector.Tag = ""
            cmdMicroSector.Tag = ""
            cmdMicroSector.Enabled = True
            
            txtFunction.Text = ""
            txtFunction.Tag = ""
            txtFunctionary.Text = ""
            txtFunctionary.Tag = ""
            
            txtAccountHeadCode.Text = ""
            txtAccountHeadCode.Tag = ""
            txtAccountHead.Text = ""
            txtAccountHead.Tag = ""
            cmdSearchHead.Tag = ""
            
            txtDPCDate.Text = ""
            txtDPCDate.Tag = ""
            
            
            
            
            If objProj.ProjectID > 0 Then
                '*************************************************************'
                ' NEW PROJECTS 2013-14
                '*************************************************************'
                mYearID = objProj.YearID
                If objProj.YearID > 2012 Then
                    Call SetProjectDetails(objProj, val(mSourceOfFundID))
                    GoTo VALIDATEAMT:
                End If
                '*************************************************************'
                
                
                txtProjName.Text = objProj.ProjectNameEnglish
                txtProjectNo.Text = objProj.ProjectSerialNo
                txtProjectNo.Tag = objProj.ProjectID
                cmbCategory.Tag = objProj.ProjCatID
                mintCategoryID = objProj.ProjCatID
                
                cmbSource.Text = objProj.FindSourceOfFund(mSourceOfFundID)
                cmbSource.Enabled = False
                'txtDPCNo.Text = ""
                'txtDPCDate.Text = ""
                mSubsectorID = objProj.SubSectorID
                txtDPCNo.Tag = mSubsectorID
                txtSubSector.Tag = mSubsectorID
                
                If val(objProj.SchemeID) <> 0 Then
                    chkBFund.Enabled = False
                    lblScheme.Visible = True
                    txtScheme.Visible = True
                    txtScheme.Tag = objProj.SchemeID
                    SetSchemeDetails (val(txtScheme.Tag))
                Else
                    chkBFund.Enabled = True
                    lblScheme.Visible = False
                    txtScheme.Visible = False
                    txtScheme.Tag = ""
                End If
                
On Error GoTo SkipSulekha:
                Set mCol = objProj.GetFundDetails(CInt(gbFinancialYearID), objProj.ProjectID)
                For mRow = 1 To mCol.count
                    Set objProFund = mCol.Item(mRow)
                    If objProFund.SourceOfFundID = mSourceOfFundID Then
                        txtProjCost.Text = objProFund.SourceWiseAmount
                        txtProjCost.Tag = objProFund.SourceWiseAmount
                        Exit For
                    End If
                Next mRow
                
                
            End If
                 
            Dim mCnn    As New ADODB.Connection
            Dim mCnPlan As New ADODB.Connection
            
            Dim Rec     As New ADODB.Recordset
            Dim obj     As New clsDB
            Dim msql    As String
            Dim mWHERE  As String
            
            Dim mCapitalExpFlag As Boolean
            Dim mMicroSectorCount As Integer
            Dim mMicroHeads As Integer

            'obj.CreateNewConnection mCnn, enuSourceString.Saankhya
            'mSQL = "Select faSubSectorHeads.*,vchAccountHead,vchFunction,vchTransactionCategory,faFunctionaries.intFunctionaryID,vchFunctionary from faSubSectorHeads"
            'mSQL = mSQL + " INNER JOIN faAccountHeads on faAccountHeads.intAccountHeadID=faSubSectorHeads.intAccountHeadID"
            'mSQL = mSQL + " INNER JOIN faFunctions on faFunctions.intFunctionID=faSubSectorHeads.intFunctionID"
            'mSQL = mSQL + " INNER JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faSubSectorHeads.intCategoryID"
            'mSQL = mSQL + " INNER JOIN faFunctionaryFunctions on faFunctionaryFunctions.intFunctionID=faSubSectorHeads.intFunctionID"
            'mSQL = mSQL + " INNER JOIN faFunctionaries on faFunctionaries.intFunctionaryID=faFunctionaryFunctions.intFunctionaryID"
            'mSQL = mSQL + " Where intSubSectorID = " & mSubSectorID & " And faSubSectorHeads.intCategoryID = " & mintCategoryID & " "
            'Rec.Open mSQL, mCnn
            'If Not (Rec.EOF And Rec.BOF) Then
            'While Not (Rec.EOF)
            '    cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
            '    txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
            '    txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
            '    txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
            '    txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
            '    txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
            '    txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
            '    txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
            '    Rec.MoveNext
            'Wend
            'End If
            'Rec.Close
            
            msql = " SELECT faSubSectorHeads.intCategoryID,vchTransactionCategory, intSubSectorID, vchSubSectorCode, vchSubSector, "
            msql = msql + " faSubSectorHeads.intAccountHeadID, faAccountHeads.vchAccountHeadCode, vchAccountHead, "
            msql = msql + " faSubSectorHeads.intFunctionID, faFunctions.vchFunctionCode, vchFunction, "
            msql = msql + " faFunctionaryFunctions.intFunctionaryID, vchFunctionary, "
            msql = msql + " faSubSectorHeads.intTransactionTypeID , vchTransactionType "
            msql = msql + " FROM faSubSectorHeads "
            msql = msql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faSubSectorHeads.intAccountHeadID "
            msql = msql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faSubSectorHeads.intFunctionID "
            msql = msql + " LEFT JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
            msql = msql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID "
            msql = msql + " INNER JOIN faTransactionCategory on faTransactionCategory.intCategoryID=faSubSectorHeads.intCategoryID"
            msql = msql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faSubSectorHeads.intTransactionTypeID "
            msql = msql + " Where faSubSectorHeads.intSubSectorID = " & mSubsectorID & " And faSubSectorHeads.intCategoryID = " & val(cmbCategory.Tag)
            
            If obj.SetConnection(mCnn) Then
                Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly
                If Not (Rec.EOF And Rec.BOF) Then
                    cmbCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    cmbCategory.Enabled = False
                    
                    txtFunction.Tag = IIf(IsNull(Rec!intFunctionID), "", Rec!intFunctionID)
                    txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                    cmdSearchFunction.Enabled = False

                    txtAccountHeadCode.Tag = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                    txtAccountHeadCode.Text = IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode)
                    txtAccountHead.Text = IIf(IsNull(Rec!vchAccountHead), "", Rec!vchAccountHead)
                    cmdSearchHead.Enabled = True 'False
                    
                    txtFunctionary.Tag = IIf(IsNull(Rec!intFunctionaryID), "", Rec!intFunctionaryID)
                    txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                    cmdSearchFunctionary.Enabled = False
                Else
                    cmbCategory.ListIndex = -1
                    'cmbCategory.Text = ""
                    txtFunction.Tag = ""
                    txtFunction.Text = ""
                    cmdSearchFunction.Enabled = True

                    txtAccountHeadCode.Tag = ""
                    txtAccountHeadCode.Text = ""
                    txtAccountHead.Text = ""
'                    cmdSearchHead.Enabled = False
                    cmdSearchHead.Enabled = True
                    
                    txtFunctionary.Tag = ""
                    txtFunctionary.Text = ""
                    
                    cmdSearchFunction.Enabled = True
                    cmdSearchFunctionary.Enabled = True
            
                End If
                Rec.Close
            End If
            
            ' EXTRACTING MicrosectorIDs From Plan-Project
            If obj.CreateNewConnection(mCnPlan, enuSourceString.Sulekha) Then
                msql = "SELECT MicroSector.intMicroSecID  FROM MicroSector WHERE decProjectID = " & val(txtProjectNo.Tag)   '118600160078
                msql = msql + " AND intYearID = " & mYearID
                Rec.Open msql, mCnPlan, adOpenStatic, adLockReadOnly
                mMicroSectorCount = 0
                If Not (Rec.BOF And Rec.EOF) Then
                    While Not Rec.EOF
                        mMicroSectorCount = mMicroSectorCount + 1
                        If Len(mWHERE) > 0 Then
                            mWHERE = mWHERE & ", " & Rec!intMicroSecID
                            txtDPCDate.Tag = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                            txtMicroSector.Tag = IIf(IsNull(Rec!intMicroSecID), 0, Rec!intMicroSecID)
                        Else
                            mWHERE = Trim(str(Rec!intMicroSecID))
                        End If
                        Rec.MoveNext
                    Wend
                End If
                Rec.Close
            End If
            
            ' Finding Account Heads from MicroSectors - CAPITAL EXPENDITURE
            If mMicroSectorCount > 0 Then
            If obj.SetConnection(mCnn) Then
                msql = "SELECT Distinct intAccountHeadID FROM faMicroSectorHeads WHERE intMircoSectorID IN ( " & mWHERE & ")" '323,324,325,326,327,350)
                Rec.Open msql, mCnn, adOpenStatic, adLockReadOnly
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
            End If
            
            'mSQL = "SELECT * FROM faAccountHeads WHERE intAccountHeadID IN  (" & mWHERE & ")"
            If mMicroHeads > 0 Then
                msql = "SELECT (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where  intAccountHeadID IN  (" & mWHERE & ") Order By faAccountHeads.vchAccountHeadCode"
                cmdSearchHead.Tag = msql
            Else
                cmdSearchHead.Tag = ""
            End If
            
            If mCapitalExpFlag Then
            Select Case mMicroHeads
                Case Is = 0
                    cmdSearchHead.Enabled = False
                Case Is = 1
                    Dim objAcc As New clsAccounts
                    objAcc.SetAccountID val(mWHERE)
                    If objAcc.AccountHeadID > 0 Then
                        txtAccountHeadCode.Tag = objAcc.AccountHeadID
                        txtAccountHeadCode.Text = objAcc.AccountCode
                        txtAccountHead.Text = objAcc.AccountHead
                        cmdSearchHead.Enabled = True 'False
                    End If
                Case Else
                    txtAccountHeadCode.Tag = ""
                    txtAccountHeadCode.Text = ""
                    txtAccountHead.Text = ""
                    cmdSearchHead.Enabled = True
                    
            End Select
            End If
        End If
        
VALIDATEAMT:
        '********ADDED FOR PROJECT AMT VALIDATION**************
        Dim mCnnAmt    As New ADODB.Connection
        Dim RecAmt     As New ADODB.Recordset
        Dim objAmt     As New clsDB
        Dim mSQLAmt    As String
        Dim mAvailBalnz As Variant

        mAvailBalnz = 0
        objAmt.CreateNewConnection mCnnAmt, enuSourceString.Saankhya
        mSQLAmt = "select fltRequestedAmt from faAllotments where tnyStatus<>2 And numProjectID= " & val(txtProjectNo.Tag) & "  And intFinancialYearID=" & gbFinancialYearID & " And intSourceID=" & val(mSourceOfFundID) & " "
        mSQLAmt = mSQLAmt + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)" '(1,2)"
        RecAmt.Open mSQLAmt, mCnnAmt
        If Not (RecAmt.EOF And RecAmt.BOF) Then
            While Not (RecAmt.EOF)
                mAvailBalnz = mAvailBalnz + IIf(IsNull(RecAmt!fltRequestedAmt), 0, RecAmt!fltRequestedAmt)
                RecAmt.MoveNext
            Wend
        End If
        RecAmt.Close

        'txtProjCost.Tag = Abs(objProFund.SourceWiseAmount - mAvailBalnz)
        txtProjCost.Tag = val(txtProjCost.Tag) - mAvailBalnz
        If IsNumeric(txtProjCost.Tag) Then
            If val(txtAmountRequested.Text) > val(txtProjCost.Tag) Then
                msql = " Balance Available for " & cmbSource.Text & " in this Project" & vbCrLf  'Amount allocated
                msql = msql + " is Rs. " & Format(val(txtProjCost.Tag), "0.00")
                MsgBox msql
                On Error Resume Next
                txtAmountRequested.Text = Format(val(txtProjCost.Tag), "0.00")
                txtAmountRequested.SetFocus
                Exit Sub
            ElseIf val(txtAmountRequested) < 0 Then
                msql = " Check the Requisition Amount "
                MsgBox msql
                On Error Resume Next
                txtAmountRequested.Text = ""
                txtAmountRequested.SetFocus
            End If
        End If
        Exit Sub
SkipSulekha:
    If mFundErSulekha = 1 Then
        MsgBox "Some Mistakes in Ported data from Sulekha"
    End If
    End Sub
    Private Sub CheckProjAmtValidation()
     '********ADDED FOR PROJECT AMT VALIDATION**************
        Dim mCnnAmt    As New ADODB.Connection
        Dim RecAmt     As New ADODB.Recordset
        Dim objAmt     As New clsDB
        Dim mSQLAmt    As String
        Dim msql As String
        Dim mAvailBalnz As Variant

        mAvailBalnz = 0
        objAmt.CreateNewConnection mCnnAmt, enuSourceString.Saankhya
        mSQLAmt = "select fltRequestedAmt from faAllotments where tnyStatus<>2 And numProjectID= " & val(txtProjectNo.Tag) & "  And intFinancialYearID=" & gbFinancialYearID & " And intSourceID=" & val(cmbSource.ItemData(cmbSource.ListIndex)) & " "
        mSQLAmt = mSQLAmt + "  AND ISNULL(tnyTypeID,0) NOT IN  (1)" '(1,2)"
        RecAmt.Open mSQLAmt, mCnnAmt
        If Not (RecAmt.EOF And RecAmt.BOF) Then
            While Not (RecAmt.EOF)
                mAvailBalnz = mAvailBalnz + IIf(IsNull(RecAmt!fltRequestedAmt), 0, RecAmt!fltRequestedAmt)
                RecAmt.MoveNext
            Wend
        End If
        RecAmt.Close

        'txtProjCost.Tag = Abs(objProFund.SourceWiseAmount - mAvailBalnz)
        txtProjCost.Tag = val(txtProjCost.Text) - mAvailBalnz
        If IsNumeric(txtProjCost.Tag) Then
            If val(txtAmountRequested.Text) > val(txtProjCost.Tag) Then
                msql = " Balance Available for " & cmbSource.Text & " in this Project" & vbCrLf  'Amount allocated
                msql = msql + " is Rs. " & Format(val(txtProjCost.Tag), "0.00")
                MsgBox msql
                On Error Resume Next
                txtAmountRequested.Text = Format(val(txtProjCost.Tag), "0.00")
                txtAmountRequested.SetFocus
                Exit Sub
            ElseIf val(txtAmountRequested) < 0 Then
                msql = " Check the Requisition Amount "
                MsgBox msql
                On Error Resume Next
                txtAmountRequested.Text = ""
                txtAmountRequested.SetFocus
            End If
        End If
    End Sub



    Private Sub txtStateHead_GotFocus()
        On Error GoTo Err
        If gbSearchID <> -1 Then
            txtStateHead.Tag = gbSearchID
            'txtStateHead.Text = gbSearchStr
            If txtStateHead.Tag <> "" Then
                Dim mCnn As New ADODB.Connection
                Dim objDb As New clsDB
                Dim msql As String
                Dim Rec As New ADODB.Recordset
                
                If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                    If chkBFund.value = vbChecked Then
                        msql = "Select vchHeadOfAccountCode,vchHeadOfAccount from faHeadOfAccount  Where intHeadOfAccountID= " & txtStateHead.Tag
                        Rec.Open msql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            MaskAccHead.Text = IIf(IsNull(Rec!vchHeadOfAccountCode), "", Rec!vchHeadOfAccountCode)
                            txtStateHead.Text = IIf(IsNull(Rec!vchHeadOfAccount), "", Rec!vchHeadOfAccount)
                        End If
                        Rec.Close
                    Else
                        msql = "Select vchBudgetHead,vchHeadDesc from faBudgetHead Where intId= " & txtStateHead.Tag
                        Rec.Open msql, mCnn
                        If Not (Rec.EOF And Rec.BOF) Then
                            MaskAccHead.Text = IIf(IsNull(Rec!vchBudgetHead), "", Rec!vchBudgetHead)
                            txtStateHead.Text = IIf(IsNull(Rec!vchHeadDesc), "", Rec!vchHeadDesc)
                        End If
                        Rec.Close
                    End If
                    
                End If
            End If
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

Private Sub txtStateHead_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
        KeyAscii = 0
  End If
End Sub

    Private Sub txtStateSubHead_GotFocus()
        On Error GoTo Err
        If gbSearchID <> -1 Then
            txtStateSubHead.Tag = gbSearchID
            'txtStateSubHead.Text = gbSearchStr
            If txtStateHead.Tag <> "" Then
                Dim mCnn As New ADODB.Connection
                Dim objDb As New clsDB
                Dim msql As String
                Dim Rec As New ADODB.Recordset
                
                If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                    msql = "Select vchHeadOfAccountCode,vchHeadOfAccount from faHeadOfAccount  Where intHeadOfAccountID= " & txtStateSubHead.Tag
                    Rec.Open msql, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        MaskDetailAccHead.Text = IIf(IsNull(Rec!vchHeadOfAccountCode), "", Rec!vchHeadOfAccountCode)
                        'MaskDetailAccHead.Tag = IIf(IsNull(Rec!vchStateSubAcHeadCode), "", Rec!vchStateSubAcHeadCode)
                        txtStateSubHead.Text = IIf(IsNull(Rec!vchHeadOfAccount), "", Rec!vchHeadOfAccount)
                    End If
                    Rec.Close
                End If
            End If
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub
Private Sub txtStateSubHead_KeyPress(KeyAscii As Integer)
'      If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = 8) Then
'            KeyAscii = 0
'      End If
End Sub

    Private Sub txttreasury_GotFocus()
        Dim msql As String
        Dim mToken1 As String
           
        If gbSearchID <> -1 Then
            mToken1 = Token(gbSearchStr, "-")
            txtTreasuryCode.Text = mToken1
            txttreasury.Text = gbSearchStr
            txtTreasuryCode.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    
    Private Sub txttreasury_LostFocus()
'''''        Dim mCnnTR    As New ADODB.Connection
'''''        Dim RecTR     As New ADODB.Recordset
'''''        Dim objTR     As New clsDB
'''''        Dim mSQLTR    As String
'''''        Dim mArrIn    As Variant
'''''        objTR.CreateNewConnection mCnnTR, enuSourceString.DBMaster
'''''
'''''        If txtTreasuryCode.Text <> "" Then
'''''            mSQLTR = "select * from GM_Treasury where chvTreasuryCode= '" & txtTreasuryCode.Text & "' "
'''''            RecTR.Open mSQLTR, mCnnTR
'''''            If (RecTR.EOF And RecTR.BOF) Then
'''''                mArrIn = Array(Null, _
'''''                        txtTreasuryCode.Text, _
'''''                        txttreasury)
'''''                objTR.ExecuteSP "spSaveTreasury", mArrIn, , , mCnnTR, adCmdStoredProc
'''''            End If
'''''            RecTR.Close
'''''        End If
    End Sub
    Private Sub txtTreasuryCode_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") And KeyAscii >= Asc("0") Or KeyAscii = 8) Then
                    KeyAscii = 0
        End If
    End Sub
    Private Function UpdateProjectDetials(mProjectID As Variant)
        Dim mCnnSulekha     As New ADODB.Connection
        Dim mCnn            As New ADODB.Connection
        Dim objDb           As New clsDB
        Dim Rec             As New ADODB.Recordset
        Dim RecSulekha      As New ADODB.Recordset
        Dim mArrIN          As Variant
        Dim msql            As String
        Dim msqlSulekha     As String
        Dim mUpdateProjectDetials As Boolean
        Dim mYearID         As Integer
        
        
        If mPreviousYearMode = 1 Then
            mYearID = gbFinancialYearID - 1
        Else
            mYearID = gbFinancialYearID
        End If
        
        If (objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            msql = "Select * from suProjectDetails Where decProjectID= " & mProjectID & " And intYearID= " & mYearID
            Rec.Open msql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mUpdateProjectDetials = False
            Else
                mUpdateProjectDetials = True
            End If
            Rec.Close
        End If
        
        If mUpdateProjectDetials = True Then
            If (objDb.CreateNewConnection(mCnnSulekha, enuSourceString.Sulekha)) Then
                msqlSulekha = "Select * from ProjectDetails "
                msqlSulekha = msqlSulekha + " left join SubjectCheckList On SubjectCheckList.decProjectID=ProjectDetails.decProjectID  And SubjectCheckList.intYearID= " & mYearID
                msqlSulekha = msqlSulekha + " Where ProjectDetails.decProjectID = " & mProjectID & ""
                msqlSulekha = msqlSulekha + " And ProjectDetails.intYearID= " & mYearID
                RecSulekha.Open msqlSulekha, mCnnSulekha
                If Not (RecSulekha.EOF And RecSulekha.BOF) Then
                    mArrIN = Array(mProjectID, _
                                          gbLBID, _
                                          mYearID, _
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
                                          Null _
                                        )
                    objDb.ExecuteSP "spUpdateProjectDetails", mArrIN, , , mCnn, adCmdStoredProc
                End If
            End If
         Else
            Exit Function
         End If
    End Function
     Public Function CheckPreviousYearRequisitions(mReqID As Variant)
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mReqID <> "" Then
            If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                If mLoadMode = 10 Then
                    msql = "Select * From faPendingTaskRequest Where intTaskID = 16 AND intKeyID = " & mReqID
                Else
                    msql = "Select * From faPendingTaskRequest Where intTaskID IN (3,13) AND intKeyID = " & mReqID
                End If
                Set Rec = objDb.ExecuteSP(msql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    CheckPreviousYearRequisitions = 1
                Else
                    CheckPreviousYearRequisitions = 0
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Function
    Private Sub GetSchemeDetails(mScehmeID As Integer)
        Dim msql        As String
        Dim objDb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset

        If mScehmeID > 0 Then
            If objDb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                msql = "Select * From faDepSchPro Where intID=" & mScehmeID
                Set Rec = objDb.ExecuteSP(msql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtScheme.Tag = IIf(IsNull(Rec!intID), 0, Rec!intID)
                    txtScheme.Text = IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
                    txtScheme.Enabled = False
                    cmdSearchScheme.Enabled = False
                End If
            End If
        End If
    End Sub
    Public Sub CalulateAmountIssued(mSourceOfFundID As Integer, mCategoryID As Integer)
        Dim mCnn    As New ADODB.Connection
        Dim objDb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim msql    As String
        Dim mYearID As Integer
        
        If mPreviousYearMode = 0 Then
            mYearID = gbFinancialYearID
        Else
            mYearID = gbFinancialYearID - 1
        End If
    
        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        Select Case val(mSourceOfFundID)
        Case 1, 27, 28, 19, 21
            If val(mCategoryID) = 1 Or val(mCategoryID) = 2 Or val(mCategoryID) = 3 Then
                msql = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                msql = msql + " ("
                msql = msql + " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2"
                msql = msql + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID= " & mYearID & " "
                msql = msql + " AND tnyStage IN (1,2)"
                msql = msql + " AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                msql = msql + " AND intFundCategoryID = 1"
                msql = msql + " Union All"
                msql = msql + " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2"
                msql = msql + " AND intSourceID IN (1) AND intFinancialYearID= " & mYearID & " "
                msql = msql + " AND tnyStage IN (1,2)"
                msql = msql + " AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                msql = msql + " )A"
            End If
            
        Case 10, 11, 12, 13, 14
            If val(mCategoryID) = 1 Then
                msql = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
                msql = msql + " ("
                msql = msql + "  Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2 "
                msql = msql + " AND intSourceID IN (21,27, 28, 10, 11, 12, 13, 14,19) AND intFinancialYearID=" & mYearID & " And   tnyStage IN (1,2)"
                msql = msql + " AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                msql = msql + " AND intFundCategoryID = 1"
                msql = msql + " Union All"
                msql = msql + " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2"
                msql = msql + " AND intSourceID IN (1) AND intFinancialYearID=" & mYearID & " And   tnyStage IN (1,2)"
                msql = msql + " AND ISNULL(tnyTypeID,0) NOT IN  (1) "
                msql = msql + " )A"
                
                
            ElseIf val(mCategoryID) = 2 Then
                GoTo SCP:
            ElseIf val(mCategoryID) = 3 Then
                GoTo TSP:
            End If
         
        Case 16, 17
            msql = " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2 "
            msql = msql + " AND intSourceID IN (16,17) AND intFinancialYearID=" & mYearID & "  And   tnyStage IN (1,2)"
            msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        Case 3
            msql = "Select Sum(fltRequestedAmt) As AmountIssued From faAllotments WHERE Isnull(tnyStatus,0)  <>2  "
            msql = msql + " AND intSourceID =" & mSourceOfFundID & "  AND intSchemeID = " & val(txtScheme.Tag) & " AND intFinancialYearID=" & mYearID & " "
            msql = msql + " AND tnyStage IN (1,2)"
            msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "

        Case 10, 11, 12, 13, 14, 29
SCP:
            msql = " SELECT SUM(A.AmountIssued) AS AmountIssued FROM"
            msql = msql + " ("
            msql = msql + "  Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2 "
            msql = msql + " AND intSourceID IN (10, 11, 12, 13, 14) AND intFinancialYearID=" & mYearID & " AND tnyStage IN (1,2)"
            msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            msql = msql + " AND intFundCategoryID IN (2)"
            msql = msql + " Union ALL"
            msql = msql + " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2 "
            msql = msql + " AND intSourceID IN (29) AND intFinancialYearID=" & mYearID & " AND tnyStage IN (1,2)"
            msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            msql = msql + " )A"
          
        Case 10, 11, 12, 13, 14, 30
TSP:
            msql = " Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where Isnull(tnyStatus,0) <>2 "
            msql = msql + " AND intSourceID IN (10, 11, 12, 13, 14,30) AND intFinancialYearID=" & mYearID & " AND tnyStage IN (1,2)"
            msql = msql + "  AND ISNULL(tnyTypeID,0) NOT IN  (1) "
            msql = msql + " AND intFundCategoryID = 3"
            
        Case Else
            msql = "Select Sum(fltRequestedAmt) As AmountIssued From faAllotments Where  Isnull(tnyStatus,0) <>2 "
            msql = msql + " AND intSourceID =" & val(txtInstalmentNo.Tag) & " AND intFinancialYearID=" & mYearID & " "
            msql = msql + " AND tnyStage IN (1,2)"
            msql = msql + " AND ISNULL(tnyTypeID,0) NOT IN  (1) "
        End Select
    
        Rec.Open msql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
              txtAmtIssued.Text = IIf(IsNull(Rec!AmountIssued), "0", Rec!AmountIssued)
        End If
        Rec.Close
        
    End Sub
    Public Property Let RequisitionID(mData As Variant)
        ReqID = mData
    End Property
    
    Public Property Get RequisitionID() As Variant
        RequisitionID = ReqID
    End Property
    
    Public Property Let FundSerialNo(mData As Variant)
        mFundSlNo = mData
    End Property

    Public Property Get FundSerialNo() As Variant
        FundSerialNo = mFundSlNo
    End Property

    Public Property Let FundErSulekha(mData As Integer)
        mFundErSulekha = mData
    End Property

    Public Property Get FundErSulekha() As Integer
        FundErSulekha = mFundErSulekha
    End Property
    
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property
    Public Property Let LoadMode(mData As Integer)
        mLoadMode = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = mLoadMode
    End Property
    Public Property Let TokenID(mData As Variant)
        mTokenID = mData
    End Property
    
    Public Property Get TokenID() As Variant
        TokenID = mTokenID
    End Property
    Public Property Let ReqInboxID(mData As Long)
        mReqInboxID = mData
    End Property
    
    Public Property Get ReqInboxID() As Long
        ReqInboxID = mReqInboxID
    End Property
    


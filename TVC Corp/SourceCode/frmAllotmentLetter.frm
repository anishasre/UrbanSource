VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmAllotmentLetter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   390
      Left            =   4620
      TabIndex        =   30
      Top             =   7035
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegenerateDemand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Re-Generate Demand"
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
      Height          =   435
      Left            =   990
      TabIndex        =   63
      Top             =   7065
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancellAllotment 
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
      Height          =   390
      Left            =   8370
      TabIndex        =   62
      Top             =   7020
      Width           =   1215
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
      Enabled         =   0   'False
      Height          =   390
      Left            =   8400
      TabIndex        =   61
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10890
      TabIndex        =   34
      Top             =   7020
      Width           =   1215
   End
   Begin VB.CommandButton cmdApprove 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Approve"
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
      Height          =   390
      Left            =   9630
      TabIndex        =   33
      Top             =   7020
      Width           =   1215
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   -3000
      Top             =   7320
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Frame fmeOthers 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   -45
      TabIndex        =   54
      Top             =   6390
      Width           =   12795
      Begin VB.TextBox txtPublicGrantHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7890
         TabIndex        =   28
         Top             =   0
         Width           =   1950
      End
      Begin VB.TextBox txtPublicBudgetHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7890
         TabIndex        =   29
         Top             =   300
         Width           =   1950
      End
      Begin VB.Label lblAccHead1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head of Account in the demand for grants in the budget / public account"
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
         Left            =   1440
         TabIndex        =   56
         Top             =   0
         Width           =   6390
      End
      Begin VB.Label lblAccHead2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head of Account in the appendix IV to the detailed budget estimates"
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
         Left            =   1815
         TabIndex        =   55
         Top             =   270
         Width           =   6030
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   12435
      TabIndex        =   46
      Top             =   0
      Width           =   12435
      Begin VB.Image Image1 
         Height          =   900
         Left            =   10920
         Picture         =   "frmAllotmentLetter.frx":0000
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use this form to Record Receipt of A fund, B Fund and C Fund in the Treasury Account"
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
         Left            =   1800
         TabIndex        =   60
         Top             =   525
         Width           =   7005
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letter of Authority:"
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
         Left            =   195
         TabIndex        =   59
         Top             =   150
         Width           =   1590
      End
   End
   Begin VB.Frame fraAllotments 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDEDED&
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
      Height          =   5685
      Left            =   -15
      TabIndex        =   35
      Top             =   840
      Width           =   12465
      Begin VB.TextBox txtGONumberValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   69
         Top             =   4440
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtNatureOfClaim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   68
         Top             =   4800
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.ComboBox cmbTransactionTypes 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   4245
      End
      Begin VB.Frame fmeProject 
         BackColor       =   &H00DDEDED&
         BorderStyle     =   0  'None
         Height          =   105
         Left            =   120
         TabIndex        =   52
         Top             =   5160
         Width           =   12315
         Begin VB.CommandButton cmdProject 
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
            Left            =   9690
            TabIndex        =   27
            Top             =   240
            Width           =   300
         End
         Begin VB.TextBox txtProjectNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3105
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   2010
         End
         Begin VB.TextBox txtProjectName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5130
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   4530
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
            Left            =   2400
            TabIndex        =   53
            Top             =   240
            Width           =   630
         End
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
         Height          =   270
         Left            =   9780
         TabIndex        =   14
         Top             =   2430
         Width           =   300
      End
      Begin VB.CommandButton cmdFunction 
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
         Left            =   9780
         TabIndex        =   21
         Top             =   3765
         Width           =   300
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3465
         Width           =   6555
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3765
         Width           =   6555
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4065
         Width           =   5250
      End
      Begin VB.CommandButton cmdSearchHead 
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
         Left            =   9780
         TabIndex        =   24
         Top             =   4065
         Width           =   300
      End
      Begin VB.TextBox txtAccountHeadCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4065
         Width           =   1290
      End
      Begin VB.CommandButton cmdFunctionary 
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
         Left            =   9780
         TabIndex        =   19
         Top             =   3465
         Width           =   300
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
         Height          =   270
         Left            =   9780
         TabIndex        =   8
         Top             =   1785
         Width           =   300
      End
      Begin VB.TextBox txtScheme 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1770
         Width           =   6555
      End
      Begin VB.ComboBox cmbSource 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   4245
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
         Height          =   270
         Left            =   4860
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtCreditAccountHeadCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2130
         Width           =   1290
      End
      Begin VB.ComboBox cmbImplementingOfficer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2775
         Width           =   6555
      End
      Begin VB.TextBox txtAllotmentNo 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   0
         Top             =   225
         Width           =   1620
      End
      Begin VB.TextBox txtNameOfTreasury 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   12
         Top             =   2445
         Width           =   3405
      End
      Begin VB.TextBox txtTreasuryCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7965
         TabIndex        =   13
         Top             =   2430
         Width           =   1785
      End
      Begin VB.TextBox txtCreditAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2130
         Width           =   5250
      End
      Begin VB.CommandButton cmdCreditAccountHead 
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
         Left            =   9780
         TabIndex        =   11
         Top             =   2145
         Width           =   300
      End
      Begin VB.TextBox txtInstalmentNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   1830
      End
      Begin VB.TextBox txtAllotmentDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtAmountInFigures 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         MaxLength       =   9
         TabIndex        =   16
         Top             =   3120
         Width           =   1665
      End
      Begin VB.ComboBox cmbCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3195
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1425
         Width           =   4245
      End
      Begin MSComCtl2.DTPicker dtpAllotmentDate 
         Height          =   315
         Left            =   9750
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67043329
         CurrentDate     =   40087
      End
      Begin VB.Label lblNature 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nature Of Claim"
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
         Left            =   1680
         TabIndex        =   67
         Top             =   4800
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblGONumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GO NUMBER"
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
         Left            =   5280
         TabIndex        =   66
         Top             =   4440
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblDDOCodeValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   65
         Top             =   4455
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblDDOCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDO CODE"
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
         Left            =   2160
         TabIndex        =   64
         Top             =   4440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblMessageBox 
         Alignment       =   2  'Center
         Caption         =   "Message Box"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   330
         Left            =   0
         TabIndex        =   58
         Top             =   5310
         Visible         =   0   'False
         Width           =   12585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Types"
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
         TabIndex        =   57
         Top             =   750
         Width           =   1605
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
         Left            =   2130
         TabIndex        =   51
         Top             =   3495
         Width           =   1020
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
         Left            =   2340
         TabIndex        =   50
         Top             =   3795
         Width           =   750
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
         Left            =   1965
         TabIndex        =   49
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department Or Scheme"
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
         Left            =   1050
         TabIndex        =   48
         Top             =   1755
         Width           =   2040
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
         Left            =   2310
         TabIndex        =   47
         Top             =   1425
         Width           =   780
      End
      Begin VB.Label lblNameOfTreasury 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Treasury"
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
         Left            =   1560
         TabIndex        =   45
         Top             =   2430
         Width           =   1560
      End
      Begin VB.Label lblTreasuryCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treasury Code"
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
         Left            =   6645
         TabIndex        =   44
         Top             =   2460
         Width           =   1260
      End
      Begin VB.Label lblImplementingOfficer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Implementing Officer"
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
         Left            =   1215
         TabIndex        =   43
         Top             =   2805
         Width           =   1875
      End
      Begin VB.Label lblCreditAccountHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit A/C Head"
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
         Left            =   1635
         TabIndex        =   42
         Top             =   2145
         Width           =   1455
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
         Left            =   4920
         TabIndex        =   41
         Top             =   3135
         Width           =   1260
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
         TabIndex        =   40
         Top             =   570
         Width           =   60
      End
      Begin VB.Label lblAllotmentNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allotment No"
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
         Left            =   1965
         TabIndex        =   39
         Top             =   240
         Width           =   1125
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
         Left            =   7410
         TabIndex        =   38
         Top             =   210
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblAmountInFigures 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   2415
         TabIndex        =   37
         Top             =   3105
         Width           =   675
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
         TabIndex        =   36
         Top             =   1095
         Width           =   645
      End
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
      Height          =   390
      Left            =   3405
      TabIndex        =   31
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5865
      TabIndex        =   32
      Top             =   7050
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllotmentLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private intLoadMode As Integer              '- - - - 10 = Receipt; 20 = Payment; 50=Opening Letter of Authority/Allotment - - - -'
    Private intAllotmentID As Variant
    Private tnyAppoveStatus As Integer
    Private strAuthorityOrAllotment As Variant
    Private mCheckDemand As Variant
    Private mPDEMode As Variant
    
    Private mPreviousYearMode As Integer
    Private mPreviousYearTaskID As Integer
    Private mPreviousYearRequestID As Integer
    
    Private Function CheckPendingAllotments() As Boolean
        Dim mCnn                    As New ADODB.Connection
        Dim objDB                   As New clsDB
        Dim Rec                     As New ADODB.Recordset
        Dim mSQL                    As String
        Dim mAuthorityOrAllotment   As Integer
        Dim mAryIn                  As Variant
        
        On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
           
        '*********************************************************************************************'
        '        Function to check whether receipt is pending for any "Letter of authority"           '
        '*********************************************************************************************'
        
        If AuthorityOrAllotment = "Authority" Then
            mAuthorityOrAllotment = 1
       
            If mAuthorityOrAllotment = 1 Then
                'mAryIn = Array(mAuthorityOrAllotment, LoadMode, CheckDateInMMM(txtFromDate.Text), CheckDateInMMM(txtToDate.Text))
                Rec.CursorLocation = adUseClient
                'Set Rec = objDB.ExecuteSP("spSelectAllotmentLetters", mAryIn, , , mCnn, adCmdStoredProc)
                mSQL = "Select *,faAllotmentLetters.tnyStatus as Status,  faAllotmentLetters.fltAmount As Amount"
                mSQL = mSQL + " From faAllotmentLetters"
                mSQL = mSQL + " Left Join suSourceOfFund On suSourceOfFund.intSourceFundID = faAllotmentLetters.intSourceOfFundID"
                mSQL = mSQL + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID"
                mSQL = mSQL + " Left Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                mSQL = mSQL + " Left Join faVouchers On faVouchers.intKeyID2 = faIDemandTBL.numDemandID"
        '        If mAuthorityOrAllotment = 1 Then
                    mSQL = mSQL + " Where intSourceOfFundID In(1,4,16,17)"
        '        Else
                   ' mSQL = mSQL + " Where intSourceOfFundID In(3)"
        '        End If
                mSQL = mSQL + " And tnyGroupID =" & LoadMode
                mSQL = mSQL + " And faAllotmentLetters.tnyStatus <> 9 And faAllotmentLetters.tnyStatus <> 8"
                Rec.Open mSQL, mCnn
                If Not Rec.EOF And Rec.BOF Then
                    While Not Rec.EOF
                        If IsNull(Rec!intVoucherNo) Then
                            MsgBox "Please issue the Receipt for previous Letter of Authority", vbInformation
                            CheckPendingAllotments = False
                            Exit Function
                        Else
                            CheckPendingAllotments = True
                        End If
                        Rec.MoveNext
                    Wend
                Else
                    CheckPendingAllotments = True
                End If
            End If
        Else
            CheckPendingAllotments = True
        End If
        Exit Function
Err:
        MsgBox Err.Description
    End Function
    
    Private Sub Printdetails(ByVal intNo As String)
        On Error GoTo Err:
            Dim objAllot        As New clsAllotmentLetter
            Dim objDB           As New clsDB
            Dim Rec             As New ADODB.Recordset
            Dim mCnn            As New ADODB.Connection
            Dim mSQL            As String
            Dim frmNewRpt       As New frmRptViewer
            Dim arInput         As Variant
            Dim frmNewViewer    As New frmRptViewer
            Dim mAllotmentID    As Integer
            
            '*********************************************************************************************'
            '                   Procedure to print the Letter of Authority/Allotment                      '
            '*********************************************************************************************'
            If objDB.SetConnection(mCnn) Then
                mSQL = "Select intAllotmentID From faAllotmentLetters Where vchAllotmentNo= '" & intNo & "'"
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    mAllotmentID = Rec!intAllotmentID
                End If
                arInput = Array(mAllotmentID)
                frmNewViewer.rptFileName = App.Path & "\Reports\rptAllotmentLetter.rpt"
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.WindowState = vbMaximized
                frmNewViewer.InputParameters = arInput
                Call frmNewViewer.ShowReport
                frmNewViewer.Visible = True
                frmNewViewer.ZOrder (0)
                Unload Me
            Else
                MsgBox "Connection Failed", vbApplicationModal
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub


    Private Sub FormInitialize()
        '*********************************************************************************************'
        '                           Procedure to clear the input fields                               '
        '*********************************************************************************************'
        On Error GoTo Err:
        
            cmbTransactionTypes.Enabled = True
            cmdNew.Enabled = True
            txtAmountInFigures.Enabled = True
                    
            txtAllotmentNo.Text = ""
            txtAllotmentNo.Tag = ""
            txtAllotmentDate.Text = ""
            txtAllotmentDate.Tag = ""
            dtpAllotmentDate.value = gbTransactionDate
            txtScheme.Text = ""
            
            txtCreditAccountHeadCode.Text = ""
            txtCreditAccountHeadCode.Tag = ""
            txtCreditAccountHead.Text = ""
            
            txtTreasuryCode.Text = ""
            txtNameOfTreasury.Text = ""
            
            txtInstalmentNo.Text = ""
            txtInstalmentNo.Tag = ""
            txtAmountInFigures.Text = ""
            
            txtFunctionary.Text = ""
            txtFunction.Text = ""
            txtAccountHeadCode.Text = ""
            txtAccountHeadCode.Tag = ""
            txtAccountHead.Text = ""
            
            txtProjectNo.Text = ""
            txtProjectNo.Tag = ""
            txtProjectName.Text = ""
            
            txtPublicGrantHead.Text = ""
            txtPublicBudgetHead.Text = ""
            
            cmbCategory.ListIndex = -1
            cmbSource.ListIndex = -1
            cmbImplementingOfficer.ListIndex = -1
            cmbTransactionTypes.ListIndex = -1
            
            lblMessageBox.Visible = False
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmbCategory_Click()
        Dim objAccounts As New clsAccounts
        
        '*********************************************************************************************'
        '                           Procedure to set the Account Head                                 '
        '*********************************************************************************************'
        If cmbCategory.ListIndex > 0 Then
            If cmbCategory.ItemData(cmbCategory.ListIndex) = 1 Then
                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
                txtCreditAccountHead.Text = objAccounts.AccountHead
                If (cmbSource.ItemData(cmbSource.ListIndex) <> 10 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 11 And cmbSource.ItemData(cmbSource.ListIndex) <> 12 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 13 And cmbSource.ItemData(cmbSource.ListIndex) <> 14 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 19) Then
                        txtAccountHeadCode.Tag = gbAcHeadIDDevelopmentFundGeneralCapital
                        txtAccountHeadCode.Text = gbAcHeadCodeDevelopmentFundGeneralCapital
                        objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundGeneralCapital)
                        txtAccountHead.Text = objAccounts.AccountHead
                        cmdSearchHead.Enabled = False
                ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 19 Then
                    If gbLBPanchayat Then
                        objAccounts.SetAccounts (1097)
                        txtAccountHeadCode.Tag = 1097 'NABARD
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHead.Text = objAccounts.AccountHead
                        cmdSearchHead.Enabled = False
                    End If
                Else
                        cmdSearchHead.Enabled = True
                End If
                
            ElseIf cmbCategory.ItemData(cmbCategory.ListIndex) = 2 Then
                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount6
                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount6
                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount6)
                txtCreditAccountHead.Text = objAccounts.AccountHead
                If (cmbSource.ItemData(cmbSource.ListIndex) <> 10 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 11 And cmbSource.ItemData(cmbSource.ListIndex) <> 12 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 13 And cmbSource.ItemData(cmbSource.ListIndex) <> 14) Then
                        txtAccountHeadCode.Tag = gbAcHeadIDDevelopmentFundSCPCapital
                        txtAccountHeadCode.Text = gbAcHeadCodeDevelopmentFundSCPCapital
                        objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundSCPCapital)
                        txtAccountHead.Text = objAccounts.AccountHead
                        cmdSearchHead.Enabled = False
                Else
                        cmdSearchHead.Enabled = True
                End If
            ElseIf cmbCategory.ItemData(cmbCategory.ListIndex) = 3 Then
                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount7
                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount7
                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount7)
                txtCreditAccountHead.Text = objAccounts.AccountHead
                If (cmbSource.ItemData(cmbSource.ListIndex) <> 10 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 11 And cmbSource.ItemData(cmbSource.ListIndex) <> 12 _
                    And cmbSource.ItemData(cmbSource.ListIndex) <> 13 And cmbSource.ItemData(cmbSource.ListIndex) <> 14) Then
                        txtAccountHeadCode.Tag = gbAcHeadIDDevelopmentFundTSPCapital
                        txtAccountHeadCode.Text = gbAcHeadCodeDevelopmentFundTSPCapital
                        objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundTSPCapital)
                        txtAccountHead.Text = objAccounts.AccountHead
                        cmdSearchHead.Enabled = False
                Else
                        cmdSearchHead.Enabled = True
                End If
            End If
            If (cmbSource.ItemData(cmbSource.ListIndex) <> 10 _
                    Or cmbSource.ItemData(cmbSource.ListIndex) <> 11 Or cmbSource.ItemData(cmbSource.ListIndex) <> 12 _
                    Or cmbSource.ItemData(cmbSource.ListIndex) <> 13 Or cmbSource.ItemData(cmbSource.ListIndex) <> 14) Then

                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccountSpecialTSB
                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccountSpecialTSB
                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccountSpecialTSB)
                txtCreditAccountHead.Text = objAccounts.AccountHead
               ' cmdSearchHead.Enabled = False
            End If
              
        End If
    End Sub

    Private Sub cmbSource_Click()
        On Error GoTo Err:
        '''txtCreditAccountHeadCode.Text = ""
        '''txtCreditAccountHeadCode.Tag = ""
        '''txtCreditAccountHead.Text = ""
        '''txtTreasuryCode.Text = ""
        '''txtNameOfTreasury.Text = ""
        '''txtPublicBudgetHead.Text = ""
        '''
        '''txtScheme.Text = ""
        '''txtScheme.Tag = ""
        '''cmbCategory.ListIndex = -1
            Dim objAccounts As New clsAccounts
            
            If cmbSource.ListIndex = -1 Then Exit Sub
            If LoadMode = 20 Then
                If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
                    txtPublicBudgetHead.Text = "8448-00-102-94-(01)"
                ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 16 Or cmbSource.ItemData(cmbSource.ListIndex) = 17 Then
                    txtPublicBudgetHead.Text = "8448-00-102-95-(01)"
                End If
            End If
            
            If LoadMode = 10 Or LoadMode = 50 Then
                cmbCategory.Enabled = True
                txtScheme.Enabled = True
                cmdSearchScheme.Enabled = True
                Select Case cmbSource.ItemData(cmbSource.ListIndex)
                    Case 1, 29, 30:
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                        txtScheme.Enabled = False
                        cmdSearchScheme.Enabled = False
                    Case 4, 16, 17, 25, 26, 27, 28, 41:
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                        txtScheme.Enabled = False
                        cmdSearchScheme.Enabled = False
                    Case 3:
                        cmbCategory.ListIndex = 1
                        cmbCategory.Enabled = False
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeCash
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDCash
                        objAccounts.SetAccounts (gbAcHeadIDCash)
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        'txtCreditAccountHead.Text = "Cash"
                   Case 10, 11, 12, 13, 14:
                        cmbCategory.Enabled = True
                        txtScheme.Enabled = False
                        cmdSearchScheme.Enabled = False
                        cmdSearchHead.Enabled = True
                End Select
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmbImplementingOfficer_Click()
'        If cmbImplementingOfficer.ListIndex > -1 Then
'            txtFunctionary.Text = cmbImplementingOfficer.Text
'            txtFunctionary.Tag = cmbImplementingOfficer.ItemData(cmbImplementingOfficer.ListIndex)
'        End If
    End Sub

    Private Sub cmbTransactionTypes_Click()
        If cmbTransactionTypes.ListIndex = -1 Then Exit Sub
        If cmbTransactionTypes.ListIndex = 0 Then
            cmbSource.ListIndex = 0
            cmbCategory.ListIndex = 0
            cmbCategory.Enabled = False
            txtCreditAccountHeadCode.Tag = ""
            txtCreditAccountHeadCode.Text = ""
            txtCreditAccountHead.Text = ""
            txtAccountHeadCode.Tag = ""
            txtAccountHeadCode.Text = ""
            txtAccountHead.Text = ""
            txtFunctionary.Text = ""
            txtFunctionary.Tag = ""
            txtFunction.Text = ""
            txtFunction.Tag = ""
        End If
        cmdCreditAccountHead.Enabled = False
        Call GetDetailsOfTransactionTypes(val(cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex)))
    End Sub

    Private Sub cmdApprove_Click()
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim objDB As New clsDB
            Dim mSQL As String
            
            If mPreviousYearMode = 1 Then
            If Not IsDate(txtAllotmentDate) Then
                MsgBox "Didn't able to fetch the Transaction Date for this transactions!", vbInformation
                Exit Sub
            End If
            If mPreviousYearRequestID < 0 Then
                If intLoadMode <> 50 Then
                MsgBox "Pending Task Request ID not found!", vbInformation
                Exit Sub
                End If
            End If
            End If
            
            '*********************************************************************************************'
            '                  Procedure to approve the Letter of Authority/Allotment                     '
            '*********************************************************************************************'
            'If gbUserTypeID <> 3 Then
            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                If val(txtAllotmentNo.Tag) = 0 Then Exit Sub
                cmdSave_Click
                If objDB.SetConnection(mCnn) Then
                    If PDEMode = 1 Then  '''Only for PDE Entries.Status is set for not to display in the Current allotment list.
                        mSQL = "Update faAllotmentLetters set tnyStatus=9 where  vchAllotmentNo= '" & Trim(txtAllotmentNo.Text) & "' "
                        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                    Else
                        mSQL = "Update faAllotmentLetters Set tnyStatus = 1, numIssuingAuthority = " & gbUserID & " Where intAllotmentID = " & val(txtAllotmentNo.Tag)
                        mCnn.Execute mSQL
                    End If
'''''               **************************************************************************************************
'''''                    DEMAND GENERATION BLOCKED AS Per NEW GO 30 MAY 2015-FOR PENDING TASK OLD PROCESS WILL FOLLOW
'''''               **************************************************************************************************
                    If mPreviousYearMode = 1 Then
                        If LoadMode = 10 Then
                            If CheckDemand = 1 Then
                              mSQL = "Update faAllotmentRegister set tnyStatus=3 where vchAllotmentNo= '" & Trim(txtAllotmentNo.Text) & "' "
                              objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
    
                            Else
                                Select Case cmbSource.ItemData(cmbSource.ListIndex)
                                 'Case 1, 2, 4, 16, 17, 21, 25, 26, 27, 28, 10, 11, 12, 13, 14, 29, 30, 41: commented on 4 apr 2017
                                Case 5, 6, 19, 20, 22, 23, 4, 2, 10, 11, 12, 13, 14
                                    Call GenerateDemand(mCnn, val(txtAllotmentNo.Tag))
                                End Select
                          End If
                        End If
                        
                    Else
                        Select Case cmbSource.ItemData(cmbSource.ListIndex)
                        Case 5, 6, 19, 20, 22, 23, 4, 2, 10, 11, 12, 13, 14   'Added  10, 11, 12, 13,14 on 18 mar 2017 '''3, 25, 26,21, removed by anisha on 22 Aug 2015
                            Call GenerateDemand(mCnn, val(txtAllotmentNo.Tag))
                        End Select
                    End If
'''''               **********************************************************************
                    cmdApprove.Enabled = False
'                    cmdReject.Enabled = False
                Else
                    MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
                End If
            End If
            Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub
    Private Function CheckTransferCredit() As Boolean
        Dim mSQL                As String
        Dim mCnn                As New ADODB.Connection
        Dim Rec                 As New ADODB.Recordset
        Dim RecChild                 As New ADODB.Recordset
        Dim objDB               As New clsDB
        
        If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
            MsgBox "The Connection to Saankhya not Present", vbCritical
            Exit Function
        End If
        
        
        mSQL = "SELECT  * FROM faAllotmentLetters"
        mSQL = mSQL + " LEFT JOIN faIDemandTBL ON faIDemandTBL.numSubLedgerID=faAllotmentLetters.intAllotmentID"
        mSQL = mSQL + " Where faAllotmentLetters.intFinancialYearID =" & gbFinancialYearID & " And IsNull(faAllotmentLetters.tnyStatus, 0) = 1"
        mSQL = mSQL + " AND ISNULL(faIDemandTBL.intVoucherID,0)=0 AND vchAllotmentNo IS NOT NULL"
        mSQL = mSQL + " AND intAllotmentID=" & val(txtAllotmentNo.Tag)
        Rec.Open mSQL, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Exit Function
        Else
            If cmbSource.ItemData(cmbSource.ListIndex) <> 1 Or cmbSource.ItemData(cmbSource.ListIndex) <> 16 Or cmbSource.ItemData(cmbSource.ListIndex) <> 17 Then
                mSQL = " SELECT * FROM faAllotmentLetters WHERE "
                mSQL = mSQL + " ISNULL(tnyGroupID,0)=30 AND intSourceOfFundID=" & cmbSource.ItemData(cmbSource.ListIndex)
                
                RecChild.Open mSQL, mCnn
                If Not (RecChild.EOF And RecChild.BOF) Then
                    CheckTransferCredit = True
                Else
                    CheckTransferCredit = False
                End If
                RecChild.Close
            End If
        
        End If
        Rec.Close
    End Function
    Private Sub cmdCancellAllotment_Click()
        Dim mCnn    As New ADODB.Connection
        Dim mSQL    As String
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mStatus As Variant
        
        '*********************************************************************************************'
        '                   Procedure to cancel the Letter of Authority/Allotment                     '
        '*********************************************************************************************'
        On Error GoTo Err
          
          
        If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 119 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 120 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 121 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 122 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 123 Then
                            
        Else
            If mPreviousYearMode <> 1 Then
                If CheckTransferCredit = True Then
                     MsgBox "Transfer Credit Already Done.No Cancellation is possible", vbInformation
                     Exit Sub
                End If
            End If
        End If
        If txtAllotmentNo.Tag <> "" Then
            If MsgBox("Do you want to Cancel this Letter of Authority?", vbYesNo) = vbYes Then
                If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                    mSQL = "Select tnyStatus From faAllotmentLetters Where intAllotmentID = " & txtAllotmentNo.Tag
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
                    End If
                    Rec.Close
                    If mStatus <> "" Then
                        If (mStatus = 0) Then 'Letter of Authority/Allotment is not approved
                            mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & txtAllotmentNo.Tag
                            MsgBox "Letter of Authority Cancelled", vbInformation
                            cmdCancellAllotment.Enabled = False
                        End If
                        If (mStatus = 1) Then 'Letter of Authority/Allotment is  approved
                            If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                                mSQL = "Select faVouchers.intVoucherNo,faIDemandTBL.numDemandID,faIDemandTBL.vchDemandNo,*"
                                mSQL = mSQL + " From faAllotmentLetters"
                                mSQL = mSQL + " Inner Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                                mSQL = mSQL + " Inner Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                                mSQL = mSQL + " Where intAllotmentID = " & txtAllotmentNo.Tag
                                mSQL = mSQL + " And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                                Rec.Open mSQL, mCnn 'Checking whether Receipt is taken against the Letter of Authority
                                If Not (Rec.EOF And Rec.BOF) Then
                                    Rec.Close
                                    mSQL = "Select faVouchers.intVoucherNo,faIDemandTBL.numDemandID,faIDemandTBL.vchDemandNo,*"
                                    mSQL = mSQL + " ,faVouchers.tnyStatus Status, faVouchers.tnyCancelFlag CancelFlag"
                                    mSQL = mSQL + " From faAllotmentLetters"
                                    mSQL = mSQL + " Inner Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                                    mSQL = mSQL + " Inner Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                                    mSQL = mSQL + " Left Join faReverseEntryChild On faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                                    mSQL = mSQL + " Left Join faReverseEntry On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID And faReverseEntry.tnyStatus = 2 "
                                    mSQL = mSQL + " Where intAllotmentID = " & txtAllotmentNo.Tag
                                    mSQL = mSQL + " And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                                    Rec.Open mSQL, mCnn 'Checking whether Receipt is Reversed
                                    If Not (Rec.EOF And Rec.BOF) Then
                                        If (IIf(IsNull(Rec!Status), 0, Rec!Status) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Or IIf(IsNull(Rec!tnyReversed), 0, Rec!tnyReversed) = 1 Then
                                            mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & txtAllotmentNo.Tag 'Cancelling the Letter of Authority/Allotment
                                            MsgBox "Letter of Authority Cancelled", vbInformation
                                            cmdCancellAllotment.Enabled = False
                                        Else
                                            MsgBox "Please Reverse/Cancel the Receipt Issued", vbInformation
                                            Exit Sub
                                        End If
                                    End If
                                    Rec.Close
                                Else
                                     If Me.AuthorityOrAllotment = "Authority" Then
                                         If txtAllotmentDate.Tag <> "" Then
                                             mCnn.Execute "Update faIDemandTBL Set tnyStatus = 9 Where numDemandID = " & txtAllotmentDate.Tag 'Cancelling the Letter of Authority/Allotment
                                         End If
                                     End If
                                     mCnn.Execute "Update faAllotmentLetters Set tnyStatus = 8 Where intAllotmentID = " & txtAllotmentNo.Tag 'Cancelling the Letter of Authority/Allotment
                                     MsgBox "Letter of Authority Cancelled", vbInformation
                                     cmdCancellAllotment.Enabled = False
                                 End If
                             Else
                                MsgBox "You are not authorized to Cancel this Approved Letter of Authority", vbInformation
                             End If
                        End If
                    End If
                Else
                    MsgBox "Connection to Finance does not exist, Please contact System Administrator", vbInformation
                End If
            End If
        End If
        cmdApprove.Enabled = False
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub cmdCreditAccountHead_Click()
        On Error GoTo Err:
            Dim mIndex  As Long
            Dim mSQL    As String
            Dim Rec     As New ADODB.Recordset
            Dim mCnn    As New ADODB.Connection
            Dim objDB   As New clsDB
            
            If cmbSource.ListIndex = -1 Then
                MsgBox "Please select the Source of Fund", vbInformation
                Exit Sub
            End If
            
            'If txtCreditAccountHeadCode.Text = "450100100" Then Exit Sub
            If txtCreditAccountHeadCode.Text = gbAcHeadCodeCash Then Exit Sub
            
            '================================================================='
            '   Credit Account Heads are Selected According to Vinods Code    '
            '================================================================='
            
            '''If cmbSource.ItemData(cmbSource.ListIndex) = 1 Then
            '''    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where vchAccountHeadCode LIKE '450650100' And tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            '''    frmSearchAccountHeads.Show vbModal
            '''ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
            '''    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where vchAccountHeadCode LIKE '450100100' And tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            '''    frmSearchAccountHeads.Show vbModal
            '''ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 16 Or cmbSource.ItemData(cmbSource.ListIndex) = 17 Then
            '''    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where vchAccountHeadCode LIKE '450650200' And tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            '''    frmSearchAccountHeads.Show vbModal
            '''End If
                        
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID"
            mSQL = mSQL + " From faAccountHeads INNER JOIN faBanks ON faBanks.intAccountHeadID = faAccountHeads.intAccountHeadID"
            mSQL = mSQL + " WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID in (1,2)"
            frmSearchAccountHeads.SQLString = mSQL   ' "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE tinHiddenFlag = 0 AND faAccountHeads.intGroupID in (1,2) "
            frmSearchAccountHeads.Show vbModal
            
            '================================================================='
            '       Needs Clarificarion - Modify if Necessory                 '
            '================================================================='
            
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    txtCreditAccountHeadCode.Text = objAccHead.AccountCode
                    txtCreditAccountHead.Text = objAccHead.AccountHead
                    txtCreditAccountHeadCode.Tag = objAccHead.AccountHeadID
                    '''If txtCreditAccountHeadCode.Tag <> "" Then
                    '''    If objDb.SetConnection(mCnn) Then
                    '''        txtTreasuryCode.Text = ""
                    '''        txtNameOfTreasury.Text = ""
                    '''        mSql = "Select vchBankName,vchBankCode From faBanks"
                    '''        mSql = mSql + " Where intAccountHeadID = " & txtCreditAccountHeadCode.Tag
                    '''        Rec.Open mSql, mCnn
                    '''        If Not (Rec.EOF And Rec.BOF) Then
                    '''            txtTreasuryCode = IIf(IsNull(Rec!vchBankCode), "", Rec!vchBankCode)
                    '''            txtNameOfTreasury.Text = IIf(IsNull(Rec!vchBankName), "", Rec!vchBankName)
                    '''        End If
                    '''    Else
                    '''        MsgBox "Connection To Finance does not Exist, Please Contact your System Administrator", vbInformation
                    '''    End If
                    '''End If
                    'If txtCreditAccountHeadCode.Tag <>1504 Then
                    If txtCreditAccountHeadCode.Tag <> gbAcHeadIDCash Then
                        txtTreasuryCode.Text = txtCreditAccountHeadCode.Text
                        txtTreasuryCode.Tag = txtCreditAccountHeadCode.Tag
                        txtNameOfTreasury.Text = txtCreditAccountHead.Text
                    Else
                        txtTreasuryCode.Text = ""
                        txtTreasuryCode.Tag = ""
                        txtNameOfTreasury.Text = ""
                    End If
                End If
                gbSearchStr = ""
                gbSearchID = -1
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdFunction_Click()
        On Error GoTo Err:
            frmSearchFunction.Show vbModal
            If Not gbSearchStr = "" Then
                txtFunction.Text = gbSearchStr
                txtFunction.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmdFunctionary_Click()
        On Error GoTo Err:
            frmSearchFunctionary.Show vbModal
            If Not gbSearchStr = "" Then
                txtFunctionary.Text = gbSearchStr
                txtFunctionary.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub cmdNew_Click()
        '2014BLOCK
        Dim mExtractedStatus As Integer
        Dim mMsg As String
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
            Exit Sub
        End If
    
        If CheckPendingAllotments Then
            Call FormInitialize
            cmdSave.Enabled = True
        End If
        
      
        
            If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
        
    End Sub

    Private Sub cmdPrint_Click()
        If txtAllotmentNo.Text = "" Then Exit Sub
        Call Printdetails(txtAllotmentNo.Text)
    End Sub

    Private Sub cmdProject_Click()
        '================================================================='
        'Linking with Project - According to Vinods Code - Doubt Remaining'
        '================================================================='
        frmEstimationDetails.Mode = 1
        frmSulekhaIntegration.Show vbModal
        '================================================================='
    End Sub
    
    Private Function SaveValidation() As Boolean
        On Error GoTo Err:
            If cmbSource.ListIndex > 0 Then
                Select Case cmbSource.ItemData(cmbSource.ListIndex)
                    Case 1:
                        If cmbCategory.ListIndex = -1 Then
                            MsgBox "Please Select the Category from the Combo Box", vbInformation
                            cmbCategory.SetFocus
                            SaveValidation = False
                            Exit Function
                        End If
                    Case 3:
                        If val(txtScheme.Tag) = 0 Then
                            MsgBox "Please Select the Scheme", vbInformation
                            'cmbCategory.SetFocus
                            SaveValidation = False
                            Exit Function
                        End If
                    Case 2:
                        If val(txtScheme.Tag) = 0 Then
                            MsgBox "Please Select the Scheme", vbInformation
                            'cmbCategory.SetFocus
                            SaveValidation = False
                            Exit Function
                        End If
                    Case 10, 11, 12, 13, 14:
                        If val(txtAccountHeadCode.Tag) = 2188 _
                        Or val(txtAccountHeadCode.Tag) = 2189 _
                        Or val(txtAccountHeadCode.Tag) = 2190 _
                        Or val(txtAccountHeadCode.Tag) = 2191 _
                        Or val(txtAccountHeadCode.Tag) = 2192 Then
                            If val(txtScheme.Tag) = 0 Then
                                MsgBox "Please Select the Scheme", vbInformation
                                'cmbCategory.SetFocus
                                SaveValidation = False
                                Exit Function
                            End If
                        End If
                    Case 4:
                        If mPreviousYearMode <> 1 Then
                            If lblDDOCodeValue.Caption = "" Then
                                MsgBox "ENTER DDO Code Please", vbInformation
                                SaveValidation = False
                                Exit Function
                            End If
                            If txtGONumberValue.Text = "" Then
                                MsgBox "ENTER GO Number Please", vbInformation
                                SaveValidation = False
                                Exit Function
                            End If
                            If txtNatureOfClaim.Text = "" Then
                                MsgBox "ENTER Nature Of Claim Please", vbInformation
                                SaveValidation = False
                                Exit Function
                            End If
                        End If
                End Select

            Else
                MsgBox "Please select the Source", vbInformation
                cmbSource.SetFocus
                SaveValidation = False
                Exit Function
            End If
            ''If txtAllotmentDate.Text = "" Then
            ''    MsgBox "Please enter the Allotment Date", vbInformation
            ''    txtAllotmentDate.SetFocus
            ''    SaveValidation = False
            ''    Exit Function
            ''End If
            If txtAmountInFigures.Text = "" Then
                MsgBox "Please enter the Amount", vbInformation
                txtAmountInFigures.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If val(txtAmountInFigures.Text) < 0 Then
                MsgBox "Please check the Amount ", vbInformation
                txtAmountInFigures.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If LoadMode = 20 Then
                If cmbImplementingOfficer.ListIndex = -1 Then
                    MsgBox "Please select the Implementing Officer", vbInformation
                    cmbImplementingOfficer.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                
                If txtProjectNo.Tag = "" Then
                    MsgBox "Please enter the Project Name", vbInformation
                    cmdProject.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                If txtPublicGrantHead.Text = "" Then
                    MsgBox "Please enter the Public Grant Head", vbInformation
                    txtPublicGrantHead.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
                If txtPublicBudgetHead.Text = "" Then
                    MsgBox "Please enter the Public Grant Head", vbInformation
                    txtPublicBudgetHead.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            ElseIf LoadMode = 50 Then
                If Not IsDate(txtAllotmentDate.Text) Then
                    MsgBox "Please Enter the Allotment Received Date!", vbInformation
                    txtAllotmentDate.SetFocus
                    SaveValidation = False
                    Exit Function
                End If
            End If
            
            If txtCreditAccountHead.Text = "" Then
                MsgBox "Please enter the Credit Account Head", vbInformation
                cmdCreditAccountHead.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If txtFunctionary.Text = "" Then
                MsgBox "Please enter the Functionary", vbInformation
                cmdFunctionary.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If txtFunction.Text = "" Then
                MsgBox "Please enter the Function", vbInformation
                cmdFunction.SetFocus
                SaveValidation = False
                Exit Function
            End If
            If txtAccountHead.Text = "" Then
                MsgBox "Please enter the Account Head", vbInformation
                cmdSearchHead.SetFocus
                SaveValidation = False
                Exit Function
            End If
            SaveValidation = True
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
   
    Private Sub cmdRegenerateDemand_Click()
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL As String

        If objDB.SetConnection(mCnn) Then
            Call GenerateDemand(mCnn, val(txtAllotmentNo.Tag))
        Else
            MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
        End If
        cmdRegenerateDemand.Enabled = False
    End Sub

'''    Private Sub cmdReject_Click() 'ADDED BY MINU FOR REJECTIONS
'''        frmReject.Mode = 8
'''        'frmReject.RequestType = txtAllotmentNo.Text
'''        frmReject.RequestTypeID = txtAllotmentNo.Tag
'''        frmReject.Show vbModal
'''        cmdReject.Enabled = False
'''        cmdApprove.Enabled = False
'''    End Sub

    Private Sub cmdSave_Click()
        On Error GoTo Err:
            
            Dim mCnn                As New ADODB.Connection
            Dim Rec                 As New ADODB.Recordset
            Dim objDB               As New clsDB
            Dim mSQL                As String
            Dim mArrIN              As Variant
            Dim mArrOut             As Variant
            
            Dim mAmountReceived     As Variant
            Dim mAmountIssued       As Variant
            Dim mAllotedAmount      As Variant
            Dim mCategoryID         As Variant
            Dim mImpOfficerID       As Variant
            Dim mSchemeID           As Variant
            Dim mProjectID          As Variant
            Dim mWHERE              As String
            Dim mYearID             As Integer
            Dim mDate               As Date
            
            
            If SaveValidation = False Then Exit Sub
            
            If objDB.SetConnection(mCnn) Then
                
                '
                'NOTE:- LETTER OF AUTHORITY =>> LoadMode will be 10 and it Straight goes to BLOCK [1]
                '
                
                
                '------------------------------------------------------------------------------------'
                '                          For generating Instalment No                              '
                '------------------------------------------------------------------------------------'
                If LoadMode = 20 Then
                    
                    If txtProjectNo.Tag <> "" Then
                        If txtAllotmentNo.Tag = "" Then
                            mSQL = "Select Count(*) as Count From faAllotmentLetters Where numProjectID =" & txtProjectNo.Tag
                            Rec.Open mSQL, mCnn
                            If Not (Rec.EOF And Rec.BOF) Then
                                If Rec!count <> "" Then
                                    txtInstalmentNo.Text = (Rec!count) + 1
                                End If
                            End If
                            Rec.Close
                        Else
                            mSQL = "Select numProjectID From faAllotmentLetters Where intAllotmentID = " & txtAllotmentNo.Tag
                            Rec.Open mSQL, mCnn
                            If Not (Rec.EOF And Rec.BOF) Then
                                mProjectID = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                            End If
                            Rec.Close
                        
                            If mProjectID <> "" Then
                                If mProjectID <> txtProjectNo.Tag Then
                                    mSQL = "Select Count(*) as Count From faAllotmentLetters Where numProjectID =" & txtProjectNo.Tag
                                    Rec.Open mSQL, mCnn
                                    If Not (Rec.EOF And Rec.BOF) Then
                                        If Rec!count <> "" Then
                                            txtInstalmentNo.Text = (Rec!count) + 1
                                        End If
                                    End If
                                    Rec.Close
                                End If
                            End If
                        End If
                    End If
                    
                End If
                '------------------------------------------------------------------------------------'
                '                          For Validating Amount to be Alloted                       '
                '------------------------------------------------------------------------------------'
                If txtProjectNo.Tag <> "" Then
                    mSQL = "Select fltEstAmt From suEstimation Where decProjectID = " & txtProjectNo.Tag
                    mSQL = mSQL + " And intFundID =" & cmbCategory.ItemData(cmbCategory.ListIndex)
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        txtAmountInFigures.Tag = IIf(IsNull(Rec!fltEstAmt), "", Rec!fltEstAmt)
                    End If
                    Rec.Close
                
                    mSQL = "Select Sum(fltAmount) As Sum From faAllotmentLetters Where numProjectID = " & txtProjectNo.Tag
                    Rec.Open mSQL, mCnn
                    If Not (Rec.EOF And Rec.BOF) Then
                        mAllotedAmount = IIf(IsNull(Rec!Sum), 0, Rec!Sum)
                        If mAllotedAmount <> 0 Then
                            If txtAllotmentNo.Tag <> "" Then
                                mAllotedAmount = val(txtAmountInFigures.Text) - mAllotedAmount
                            End If
                        End If
                        If ((mAllotedAmount + val(txtAmountInFigures.Text))) > val(txtAmountInFigures.Tag) Then
                            MsgBox "Amount exceeded", vbInformation
                            Exit Sub
                        End If
                    End If
                    Rec.Close
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ''''''''''''''''''''''''''''For Calculating the Amount Received & Issued'''''''''''''''
                '''mArrIn = Array(Val(txtCreditAccountHeadCode.Tag))
                '''objDb.ExecuteSP "spGetLedgerBalance", mArrIn, mArrOut, , mCnn, adCmdStoredProc
                '''mAmountReceived = mArrOut(0, 0)
                '''mAmountIssued = mArrOut(1, 0)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''mArrIn = ""
                '''mArrOut = ""
                
                
                ':::: BLOCK [1]
                
                If mPreviousYearMode = 1 Then
                    mYearID = gbFinancialYearID - 1
                    If IsDate(txtAllotmentDate.Text) Then
                        mDate = CDate(txtAllotmentDate.Text)
                    Else
                        MsgBox "Didn't able to fetch the previous Year Transaction Date!", vbInformation
                        Exit Sub
                    End If
                Else
                    If intLoadMode = 50 Then
                        mYearID = gbFinancialYearID - 1
                        mDate = CDate(txtAllotmentDate.Text)
                    Else
                        mYearID = gbFinancialYearID
                        mDate = gbTransactionDate
                    End If
                End If
                
                
                If cmbCategory.ListIndex = -1 Then
                    mCategoryID = Null
                Else
                    mCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
                End If
                
                If cmbImplementingOfficer.ListIndex = -1 Then
                    mImpOfficerID = Null
                Else
                    mImpOfficerID = cmbImplementingOfficer.ItemData(cmbImplementingOfficer.ListIndex)
                End If
                If txtAllotmentNo.Text = "" Then
                    MsgBox "Please enter the Allotment No", vbInformation
                    txtAllotmentNo.SetFocus
                    Exit Sub
                End If
                
                If val(txtScheme.Tag) = 0 Then
                    mSchemeID = Null
                Else
                    mSchemeID = val(txtScheme.Tag)
                End If
'                If LoadMode = 50 And gbSeatGroupID = gbSeatGroupAccountsClerk Then '***
'
'                    If IsNull(mCategoryID) = True Then
'                        If Not IsNull(mSchemeId) Then
'                            mSql = "Select * from faAllotmentLetters where intSchemeID = " & mSchemeId & " AND intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & "  and tnyOPening=1 and tnyStatus=1"
'                        Else
'                            mSql = "Select * from faAllotmentLetters where intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & "  and tnyOPening=1 and tnyStatus=1"
'                        End If
'                    Else
'                        If Not IsNull(mSchemeId) Then
'                            mSql = "Select * from faAllotmentLetters where intSchemeID = " & mSchemeId & " AND intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and intcategoryID= " & cmbCategory.ItemData(cmbCategory.ListIndex) & " and tnyOPening=1 and tnyStatus=1"
'                        Else
'                            mSql = "Select * from faAllotmentLetters where intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and intcategoryID= " & cmbCategory.ItemData(cmbCategory.ListIndex) & " and tnyOPening=1 and tnyStatus=1"
'                        End If
'                    End If
'
'                    Rec.Open mSql, mCnn
'                    If Not (Rec.EOF And Rec.BOF) Then
'                        If IsNull(mCategoryID) = True Then
'                            If Not IsNull(mSchemeId) Then
'                                mSql = "Update faAllotmentLetters set fltAmount=fltAmount + " & val(txtAmountInFigures.Text) & " where intSchemeID = " & mSchemeId & " AND  intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and tnyOPening=1 "
'                            Else
'                                mSql = "Update faAllotmentLetters set fltAmount=fltAmount + " & val(txtAmountInFigures.Text) & " where intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and tnyOPening=1 "
'                            End If
'                        Else
'                           If Not IsNull(mSchemeId) Then
'                                mSql = "Update faAllotmentLetters set fltAmount=fltAmount + " & val(txtAmountInFigures.Text) & " where intSchemeID = " & mSchemeId & " AND intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and tnyOPening=1 and intcategoryID= " & cmbCategory.ItemData(cmbCategory.ListIndex) & " "
'
'                           Else
'                                mSql = "Update faAllotmentLetters set fltAmount=fltAmount + " & val(txtAmountInFigures.Text) & " where intSourceOfFundID= " & cmbSource.ItemData(cmbSource.ListIndex) & " and tnyOPening=1 and intcategoryID= " & cmbCategory.ItemData(cmbCategory.ListIndex) & " "
'                           End If
'
'                        End If
'                        objdb.ExecuteSP mSql, , , , mCnn, adCmdText
'                    Else
'                        GoTo lblSave
'                    End If
'                    Rec.Close
'                Else     '**
lblSave:
                mArrIN = Array(IIf(txtAllotmentNo.Tag = "", -1, txtAllotmentNo.Tag), _
                                txtAllotmentNo.Text, _
                                mDate, _
                                mImpOfficerID, _
                                cmbSource.ItemData(cmbSource.ListIndex), mCategoryID, _
                                mSchemeID, txtInstalmentNo.Text, _
                                txtCreditAccountHeadCode.Tag, _
                                txtTreasuryCode.Text, _
                                txtNameOfTreasury.Text, _
                                txtFunctionary.Tag, _
                                txtFunction.Tag, _
                                txtAccountHeadCode.Tag, _
                                val(txtAmountInFigures.Text), _
                                txtProjectNo.Tag, _
                                txtPublicGrantHead.Text, _
                                txtPublicBudgetHead.Text, _
                                Null, _
                                gbUserID, _
                                gbTransactionDate, _
                                Null, _
                                gbLocalBodyID, _
                                mYearID, _
                                0, _
                                0, mAmountReceived, mAmountIssued, LoadMode, cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex), IIf(LoadMode = 50, 1, 0), Null, txtGONumberValue.Text, txtNatureOfClaim.Text _
                            )
                Set Rec = objDB.ExecuteSP("spSaveAllotmentLetter", mArrIN, , , mCnn, adCmdStoredProc)
                'objDb.ExecuteSP "spSaveAllotmentLetter", mArrIn, , , mcnn, adCmdStoredProc
                
                'txtAllotmentNo.Text = mArrOut(1, 0)
                lblMessageBox.Visible = True
                
                'MsgBox Rec(0).Name
                
                lblMessageBox.Caption = "Saved Successfully!"
                cmdSave.Enabled = False
                If PDEMode = 1 Then
                    mSQL = "Update faAllotmentRegister set tnyStatus=2 where vchAllotmentNo= '" & Trim(txtAllotmentNo.Text) & "' "
                    objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
''                    mSql = "Update faAllotmentLetters set tnyStatus=9 where  vchAllotmentNo= " & Trim(txtAllotmentNo.Text) & " "
''                    objDb.ExecuteSP mSql, , , , mCnn, adCmdText
                End If
                
                If mPreviousYearMode Then
                    Dim mAllotmentID As Integer
                    If IsNumeric(Rec.Fields(0).value) Then
                        mAllotmentID = Rec.Fields(0).value
                    Else
                        mAllotmentID = -1
                    End If
                    mSQL = "Update faPendingTaskRequest set  intKeyID= " & mAllotmentID & ",tnyStatus = 8 Where intRequestID=" & mPreviousYearRequestID & "  "
                    objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
                End If
                
            Else
                MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
    MsgBox (Error$)
    End Sub

    Private Sub cmdSearchAllotmentLetter_Click()
        '================================================================='
        '  According to Vinods Code - Doubt Remaining                     '
        '================================================================='
        '''frmListOfAllotmentLetters.Mode = 1
        '''frmListOfAllotmentLetters.Show vbModal
        '''Call txtAllotmentNo_LostFocus
        '''cmdNew.Visible = True
        '''cmdSave.Visible = True
        '''cmdCancel.Visible = True
    End Sub

    Private Sub cmdSearchHead_Click()

            Dim objAccounts As New clsAccounts
            Dim mSQL        As String

            If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 112 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            '---------ADDED FOR OTHER RECEIPTS FROM LSGs---------------------
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 119 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 120 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 121 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 122 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 123 Then
                frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Left Join faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID Where  tinHiddenFlag = 0  And intTransactionTypeID =" & cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) & " Order By faAccountHeads.vchAccountHeadCode"
            '-----------------------------------------------------------------
            
            ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 111 Then
                mSQL = " SELECT (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID"
                mSQL = mSQL + " FROM faAccountHeads WHERE vchAccountHeadCode BETWEEN '320100101' AND '320700210' AND vchAccountHead LIKE '%Central%'"
                frmSearchAccountHeads.SQLString = mSQL
            Else
                frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Where  tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            End If
            frmSearchAccountHeads.Show vbModal
            
            If Len(gbSearchStr) Then
                Dim objAccHead As New clsAccounts
                objAccHead.SetAccountCode (Token(gbSearchStr, " "))
                If objAccHead.AccountHeadID > 0 Then
                    txtAccountHeadCode.Text = objAccHead.AccountCode
                    txtAccountHead.Text = objAccHead.AccountHead
                    txtAccountHeadCode.Tag = objAccHead.AccountHeadID
                    
                    'IAY
                   
                    If val(txtAccountHeadCode.Tag) = 2188 _
                    Or val(txtAccountHeadCode.Tag) = 2189 _
                    Or val(txtAccountHeadCode.Tag) = 2190 _
                    Or val(txtAccountHeadCode.Tag) = 2191 _
                    Or val(txtAccountHeadCode.Tag) = 2192 Then
                        cmdCreditAccountHead.Enabled = True
                        cmdSearchScheme.Enabled = True          'ADDED BY MINU ON 24.03.2014
                    Else
                        cmdCreditAccountHead.Enabled = False
                        If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 119 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 120 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 121 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 122 Or _
                            cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 123 Then
   ''Commented On 14/Mar/2014 For joint venture
'''                                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
'''                                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
'''                                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
'''                                txtCreditAccountHead.Text = objAccounts.AccountHead
''Modified On 14/Mar/2014 For joint venture
                                txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccountSpecialTSB
                                txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccountSpecialTSB
                                objAccounts.SetAccounts (gbAcHeadIDTreasuryAccountSpecialTSB)
                                txtCreditAccountHead.Text = objAccounts.AccountHead
                         End If
                        ' Call SetDefaultBank  ''' Commented On 14/Mar/2014 For joint venture
                    End If
                    
                    
              End If
                gbSearchStr = ""
                gbSearchID = -1
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub SetDefaultBank()
     Dim objAccounts As New clsAccounts
        If cmbCategory.ListIndex > -1 Then
        If cmbCategory.ItemData(cmbCategory.ListIndex) = 1 Then
            txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
            txtCreditAccountHead.Text = objAccounts.AccountHead
        ElseIf cmbCategory.ItemData(cmbCategory.ListIndex) = 2 Then
            txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount6
            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount6
            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount6)
            txtCreditAccountHead.Text = objAccounts.AccountHead
        ElseIf cmbCategory.ItemData(cmbCategory.ListIndex) = 3 Then
            txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount7
            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount7
            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount7)
            txtCreditAccountHead.Text = objAccounts.AccountHead
        End If
        End If
    End Sub

    Private Sub cmdSearchScheme_Click()
        On Error GoTo Err:
            Dim mSQL As String
            
            frmSearchMasters.QrySP = Qyery
            If cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
                frmSearchMasters.SQLQry = "SELECT intID , vchDescription FROM   faDepSchPro WHERE tnyGroupID IN (1,2) ORDER BY  vchDescription asc"
            Else
                frmSearchMasters.SQLQry = "SELECT intID , vchDescription FROM   faDepSchPro WHERE tnyGroupID IN (3) ORDER BY  vchDescription asc"
            End If
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            If gbSearchStr <> "" Then
                txtScheme.Text = gbSearchStr
                txtScheme.Tag = gbSearchID
            Else
                txtScheme.Text = ""
                txtScheme.Tag = ""
            End If
            gbSearchStr = ""
            gbSearchID = -1
            
            If cmbSource.ItemData(cmbSource.ListIndex) <> 3 Then
                GetDetailsOfAccountHeadWithScheme (val(txtScheme.Tag))
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    Private Sub GetDetailsOfAccountHeadWithScheme(mSchemeID As Integer)
        Dim objAccounts As New clsAccounts
        
        If mSchemeID = 41 Then
            txtAccountHeadCode.Tag = gbAcHeadIDIAY
            txtAccountHeadCode.Text = gbAcHeadCodeIAY
            objAccounts.SetAccounts (gbAcHeadIDIAY)
        ElseIf mSchemeID = 52 Then
            txtAccountHeadCode.Tag = gbAcHeadIDIAYSCP
            txtAccountHeadCode.Text = gbAcHeadCodeIAYSCP
            objAccounts.SetAccounts (gbAcHeadIDIAYSCP)
        ElseIf mSchemeID = 54 Then
            txtAccountHeadCode.Tag = gbAcHeadIDIAYTSP
            txtAccountHeadCode.Text = gbAcHeadCodeIAYTSP
            objAccounts.SetAccounts (gbAcHeadIDIAYTSP)
        Else
            cmdSearchHead.Enabled = True
        End If
        txtAccountHead.Text = objAccounts.AccountHead
        txtAccountHeadCode.Text = objAccounts.AccountCode
        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
    End Sub
    Private Function GetDetailsOfTransactionTypes(ByVal intTransactionTypeID As Integer) As Boolean
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim objDB As New clsDB
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objAccounts As New clsAccounts
            '**********************************************************************************************'
            'Function to set the Account Heads, Function, Functionary etc according to the Transaction Type'
            '**********************************************************************************************'
            If intTransactionTypeID = 0 Then Exit Function
            If objDB.SetConnection(mCnn) Then
                mSQL = " Select A.intSourceFundID as SourceID, A.vchSourceFundName Source,  "
                mSQL = mSQL + " B.intCategoryID as CatID, B.vchTransactionCategory as Category, "
                mSQL = mSQL + " C.intFunctionID as FunID, C.vchFunction as [Function], "
                mSQL = mSQL + " D.intFunctionaryID as FnryID, D.vchFunctionary as Functionary "
                mSQL = mSQL + " From faTransactionType "
                mSQL = mSQL + " Left Join suSourceOfFund A On A.intSourceFundID = faTransactionType.intSourceFundID "
                mSQL = mSQL + " Left Join faTransactionCategory B on B.intCategoryID = faTransactionType.intCategoryID "
                mSQL = mSQL + " Left Join faFunctions C On C.intFunctionID = faTransactionType.intFunctionID "
                mSQL = mSQL + " Left Join faFunctionaries D On D.intFunctionaryID = faTransactionType.intFunctionaryID "
                mSQL = mSQL + " Where faTransactionType.intTransactionTypeID = " & intTransactionTypeID
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If IsNull(Rec!SourceID) Then
                        cmbSource.ListIndex = -1
                    Else
                        If intTransactionTypeID = 166 Then
                            cmbSource.Text = "Amount Received on Selection as Best Panchayat"
                            cmbSource.Enabled = False
                        Else
                            cmbSource.Text = Rec!Source
                            cmbSource.Enabled = False
                        End If
                    End If
                    
                    If IsNull(Rec!CatID) Then
                        cmbCategory.ListIndex = -1
                    Else
                        cmbCategory.Text = Rec!Category
                        cmbCategory.Enabled = False
                    End If
                    lblDDOCode.Visible = False
                    lblDDOCodeValue.Visible = False
                    lblGONumber.Visible = False
                    txtGONumberValue.Visible = False
                    lblNature.Visible = False
                    txtNatureOfClaim.Visible = False
                    '---------------------------------------------------------------------------'
                    '       108 - Development Fund - General                                    '
                    '       109 - Development Fund - SCP                                        '
                    '       110 - Development Fund - TSP                                        '
                    '       112 - B Fund - State Sponsored Scheme Funds                         '
                    '       125 - Maintenance Fund - Road Assets                                '
                    '       126 - Maintenance Fund - Non Roads Assets                           '
                    '       168 - CFC Grant                                                     '
                    '       169 - KLGSDP Grant                                                  '
                    '       170 - Special Grant                                                 '
                    '       171 - Road Renovation Fund                                          '
                    '       111 - Centrally Sponsored Scheme Funds                              '
                    '       174 - KLGSDP Grant (state)                                                 '
                    '---------------------------------------------------------------------------'
                    If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 108 Or _
                        cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 109 Or _
                        cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 110 Then
                        
'                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
'                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
'                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
'                        txtCreditAccountHead.Text = objAccounts.AccountHead

                        'Changed by Minu on 08.12.2012
                        
                        Select Case cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex)
                            Case Is = 108:  txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
                                            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
                                            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
                                            'cmbCategory.Enabled = True
                            Case Is = 109:  txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount6
                                            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount6
                                            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount6)
                            Case Is = 110:  txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount7
                                            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount7
                                            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount7)
                        End Select
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        Select Case cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex)
                            Case Is = 109: objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundSCPCapital)
                            Case Is = 110: objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundTSPCapital)
                            Case Else: objAccounts.SetAccounts (gbAcHeadIDDevelopmentFundGeneralCapital)
                        End Select
                        
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 125 Or _
                        cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 126 Then
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount3
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount3
                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount3)
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 125 Then
                            txtAccountHeadCode.Tag = gbAcHeadIDMaintenanceFundRoadAssets
                            txtAccountHeadCode.Text = gbAcHeadCodeMaintenanceFundRoadAssets
                            objAccounts.SetAccounts (gbAcHeadIDMaintenanceFundRoadAssets)
                            txtAccountHead.Text = objAccounts.AccountHead
                            txtAccountHeadCode.Text = objAccounts.AccountCode
                            txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 126 Then
                            txtAccountHeadCode.Tag = gbAcHeadIDMaintenanceFundNonRoadAssets
                            txtAccountHeadCode.Text = gbAcHeadCodeMaintenanceFundNonRoadAssets
                            objAccounts.SetAccounts (gbAcHeadIDMaintenanceFundNonRoadAssets)
                            txtAccountHead.Text = objAccounts.AccountHead
                            txtAccountHeadCode.Text = objAccounts.AccountCode
                            txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        End If
                    
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 166 Then ' AWARD
'                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount4
'                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount4
'                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount4)
'                        txtCreditAccountHead.Text = objAccounts.AccountHead
'
                        If gbLBPanchayat Then
                            objAccounts.SetAccounts (1688)
                        Else
                            objAccounts.SetAccounts (1688)
                        End If
                        txtAccountHead.Text = objAccounts.AccountHead
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        
                        'cmdCreditAccountHead.Enabled = True
                    
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 168 Then 'CFC
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount4
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount4
                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount4)
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        
                        objAccounts.SetAccounts (gbAcHeadIDCentralFinanceCommission)
                        txtAccountHead.Text = objAccounts.AccountHead
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 169 Or _
                        cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 174 Then 'KLGSDP and KLGSDP (State) Modified on 19 Dec 2016
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount5
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount5
                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount5)
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        
                        objAccounts.SetAccounts (gbAcHeadIDKLGSDP)
                        txtAccountHead.Text = objAccounts.AccountHead
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        
                    '****************OTHER LSGIS**************************** BLOCKED on 19/nov/2015
                        
'''                     ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 119 Or _
'''                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 120 Or _
'''                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 121 Or _
'''                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 122 Or _
'''                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 123 Then
'''                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
'''                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
'''                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
'''                        txtCreditAccountHead.Text = objAccounts.AccountHead
'''
'''                        txtAccountHead.Text = ""
'''                        txtAccountHeadCode.Text = ""
'''                        txtAccountHeadCode.Tag = ""
'''                        cmdSearchHead.Enabled = True
'''
'''                        'cmbSource.Enabled = True
'''                        cmbCategory.Enabled = True
                    '************************************************************
'''                    ****************OTHER LSGIS**************************** ReOpened on 4/Mar/2017 For Joint venture
'''                        In Special TSB Account
                     ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 119 Or _
                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 120 Or _
                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 121 Or _
                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 122 Or _
                     cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 123 Then
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccountSpecialTSB
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccountSpecialTSB
                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccountSpecialTSB)
                        txtCreditAccountHead.Text = objAccounts.AccountHead

                        txtAccountHead.Text = ""
                        txtAccountHeadCode.Text = ""
                        txtAccountHeadCode.Tag = ""
                        cmdSearchHead.Enabled = True

                        cmbSource.Enabled = True
                        cmbCategory.Enabled = True
''''                    ************************************************************
                    
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 170 Or _
                        cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 171 Then
                    
                        txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount2
                        txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount2
                        objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount2)
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        If cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 170 Then
                            txtAccountHeadCode.Tag = gbAcHeadIDSpecialGrant
                            txtAccountHeadCode.Text = gbAcHeadCodeSpecialGrant
                            objAccounts.SetAccounts (gbAcHeadIDSpecialGrant)
                            txtAccountHead.Text = objAccounts.AccountHead
                            txtAccountHeadCode.Text = objAccounts.AccountCode
                            txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 171 Then
                            objAccounts.SetAccounts (gbAcHeadIDRoadRenovationGrant)
                            txtAccountHead.Text = objAccounts.AccountHead
                            txtAccountHeadCode.Text = objAccounts.AccountCode
                            txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                        End If
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 155 Then
                        If mPreviousYearMode = 1 Then
                            txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccount1
                            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccount1
                            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccount1)
                        Else
                            txtCreditAccountHeadCode.Tag = gbAcHeadIDTreasuryAccountTSB
                            txtCreditAccountHeadCode.Text = gbAcHeadCodeTreasuryAccountTSB
                            objAccounts.SetAccounts (gbAcHeadIDTreasuryAccountTSB)
                            'DDO CODE MODIFICATION ONLY FOR GENERAL PURPOSE FUND
                            
                            lblDDOCode.Visible = True
                            lblDDOCodeValue.Visible = True
                            Call FetchDDOCode
                            
                            lblGONumber.Visible = True
                            txtGONumberValue.Visible = True
                            
                            lblNature.Visible = True
                            txtNatureOfClaim.Visible = True
                            
                        End If
                        txtCreditAccountHead.Text = objAccounts.AccountHead
                        
                        txtAccountHeadCode.Tag = gbAcHeadIDGeneralPurposeFund
                        txtAccountHeadCode.Text = gbAcHeadCodeGeneralPurposeFund
                        objAccounts.SetAccounts (gbAcHeadIDGeneralPurposeFund)
                        txtAccountHead.Text = objAccounts.AccountHead
                        txtAccountHeadCode.Text = objAccounts.AccountCode
                        txtAccountHeadCode.Tag = objAccounts.AccountHeadID
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 112 Then
                        txtAccountHead.Text = ""
                        txtAccountHeadCode.Text = ""
                        txtAccountHeadCode.Tag = ""
                        cmdSearchHead.Enabled = True
                    
                    '***********Centrally Sponsored Scheme Fund**********************************'
                    ElseIf cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex) = 111 Then
                        txtAccountHead.Text = ""
                        txtAccountHeadCode.Text = ""
                        txtAccountHeadCode.Tag = ""
                        cmdCreditAccountHead.Enabled = True
                        'cmdSearchHead.Enabled = True
                        cmbCategory.Enabled = True
                        cmdSearchScheme.Enabled = True
                    End If
                    txtFunction.Text = IIf(IsNull(Rec!Function), "", Rec!Function)
                    txtFunction.Tag = IIf(IsNull(Rec!FunID), "", Rec!FunID)
                    txtFunctionary.Text = IIf(IsNull(Rec!Functionary), "", Rec!Functionary)
                    txtFunctionary.Tag = IIf(IsNull(Rec!FnryID), "", Rec!FnryID)
                End If
            Else
                MsgBox "Connection to Finance does not Exist, Please contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub cmdSearchTreasury_Click()
        Dim mSQL As String
        
        On Error GoTo Err
        If txtCreditAccountHeadCode.Text = "" Then
            MsgBox "Please Give the Credit Account Head", vbInformation
            cmdSearchTreasury.SetFocus
        End If
        If val(txtCreditAccountHeadCode.Tag) = 1504 Then
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads"
            mSQL = mSQL + " Left Join faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID "
            mSQL = mSQL + " Where vchMinorAccountHeadCode In ('450250000', '450450000', '450650000') And tinHiddenFlag = 0 Order By faAccountHeads.vchAccountHeadCode"
            frmSearchAccountHeads.SQLString = mSQL
            frmSearchAccountHeads.Show vbModal
            If gbSearchStr <> "" Then
                txtNameOfTreasury.Text = Right(gbSearchStr, (Len(gbSearchStr) - Len(Left(gbSearchStr, 9))))
                txtTreasuryCode.Text = Left(gbSearchStr, 9)
                txtTreasuryCode.Tag = gbSearchID
            End If
        End If
        gbSearchStr = ""
        gbSearchID = -1
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub dtpAllotmentDate_CloseUp()
        txtAllotmentDate.Text = CheckDateInMMM(dtpAllotmentDate.value)
    End Sub
    
    
    Private Sub PreviousYearTask()
        Dim mCnn        As New ADODB.Connection
        Dim objDB       As New clsDB
        Dim mSQL        As String
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
        Dim objTr As New clsTransactionType
        Dim mTrTypeID As Integer
        
        
        'On Error GoTo Err
        If mPreviousYearRequestID > 0 Then
            If (objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                mSQL = "Select * from faPendingTaskRequest Where intRequestID = " & mPreviousYearRequestID
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    
                    txtAllotmentNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                    txtAllotmentDate.Text = DdMmmYy(Rec!dtTransactionDate)
                    txtAmountInFigures.Text = IIf(IsNull(Rec!fltAmount), 0, Rec!fltAmount)
                    
                    mTrTypeID = IIf(IsNull(Rec!intTransactionTypeID), -1, Rec!intTransactionTypeID)
                    objTr.SetTransactionType (mTrTypeID)
                    If objTr.TransactionTypeID > 0 Then
                        cmbTransactionTypes.Text = objTr.TransactionType
                    End If
                    
                    mTaskID = IIf(IsNull(Rec!intTaskID), 0, Rec!intTaskID)
                    cmbTransactionTypes.Enabled = False
                    cmdNew.Enabled = False
                    txtAmountInFigures.Enabled = False
                    
                    If mTaskID = 4 Then
                        'cmbTransactionTypes.Text = "B Fund - State Sponsored Scheme Funds"
                        cmbTransactionTypes.Enabled = True
                    End If
                    
                End If
                Rec.Close
            End If
        End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 1200
        ''Me.Width = 9420
        ''Me.Height = 6000
        WindowsXPC1.InitIDESubClassing
    End Sub
 

 Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSQL  As String
        Dim mTrAccHeadId As Integer
        
        If objDB.SetConnection(mCnn) Then
            If mPreviousYearMode Then
                mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID - 1
            Else
                mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            End If
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                
                'NOTE: Checking in Previous Year
                '      IF APPROVED tnyStatus will be 0 ELSE NULL
                Rec.Close
                mSQL = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID - 1
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If IsNumeric(Rec!tnyStatus) Then
                        'If Rec!tnyStatus = 0 Then
                        '    GetStatusFlag = 9
                        'Else
                        '    GetStatusFlag = -1
                        'End If
                        GetStatusFlag = Rec!tnyStatus
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

    Private Sub Form_Load()
        'On Error GoTo err:
            
            Dim mExtractedStatus As Integer
            
            
            Call FormInitialize
            Call FillCombo
            
            txtAllotmentDate.Text = gbTransactionDate
            
            If LoadMode = 10 Then ' Letter of Authority
                fmeProject.Visible = False
                fmeOthers.Visible = False
                lblCreditAccountHead.Caption = "Debit A/C Head"
                lblImplementingOfficer.Visible = False
                cmbImplementingOfficer.Visible = False
                lblInstalmentNo.Visible = False
                txtInstalmentNo.Visible = False
                lblTitle.Caption = "Letter of Authority"
                Call PreviousYearTask
                'lblDescription.Caption = "Use this form to Record Receipt of A fund, B Fund and C Fund in the Treasury Account"
            ElseIf LoadMode = 50 Then
                fmeProject.Visible = False
                fmeOthers.Visible = False
                lblCreditAccountHead.Caption = "Debit A/C Head"
                lblImplementingOfficer.Visible = False
                cmbImplementingOfficer.Visible = False
                lblInstalmentNo.Visible = False
                txtInstalmentNo.Visible = False
                lblMessageBox.Visible = True
                lblMessageBox.ForeColor = vbRed
                lblMessageBox.FontSize = 14
                lblMessageBox.FontBold = True
                lblMessageBox.Caption = "Applicable Only For Financial Year 2012-13"
            Else
                fmeProject.Visible = True
                fmeOthers.Visible = True
                lblCreditAccountHead.Caption = "Credit A/C Head"
                lblImplementingOfficer.Visible = True
                cmbImplementingOfficer.Visible = True
                lblInstalmentNo.Visible = True
                txtInstalmentNo.Visible = True
                
                lblTitle.Caption = "Allotment Letter"
                'lblDescription.Caption = "Use this form to Record Allotment of B Fund in the Consolidated Fund in the Treasury"
            End If
            
            txtAllotmentNo.Tag = AllotmentID
            If txtAllotmentNo.Tag <> "" Then
                'If gbUserTypeID <> 3 Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    'fraAllotments.Enabled = False
                    fmeProject.Enabled = False
                    fmeOthers.Enabled = False
                    'cmdSave.Enabled = False
                    cmdNew.Enabled = False
                    cmdApprove.Enabled = True
                    If LoadMode = 50 Then
                        cmdSave.Enabled = True
                        cmdSave.Caption = "Edit"
                    Else
                        cmdSave.Enabled = False
                        cmbTransactionTypes.Enabled = False
                    End If
'                    cmdReject.Enabled = True
                Else
                    cmdApprove.Enabled = False
'                    cmdReject.Enabled = False
                End If
                Call ReFillDetails
                If ApproveStatus = 1 Then
                    'If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        If txtInstalmentNo.Tag <> "" Then 'txtInstalment.Tag -> intVoucherID
                            Dim mSQL    As String
                            Dim mCnn    As New ADODB.Connection
                            Dim objDB   As New clsDB
                            Dim Rec     As New ADODB.Recordset

                            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
                            mSQL = "Select faVouchers.intVoucherNo,faIDemandTBL.numDemandID,faIDemandTBL.vchDemandNo,*"
                            mSQL = mSQL + " ,faVouchers.tnyStatus VrStatus,faVouchers.tnyCancelFlag CancelFlag,faVouchers.tnyReversed ReverseStatus  From faAllotmentLetters"
                            mSQL = mSQL + " Inner Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                            mSQL = mSQL + " Inner Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                            mSQL = mSQL + " Inner Join faReverseEntryChild On faVouchers.intVoucherID = faReverseEntryChild.intVoucherID"
                            mSQL = mSQL + " Inner Join faReverseEntry On faReverseEntryChild.intRequestID = faReverseEntry.intRequestID And faReverseEntry.tnyStatus = 2"
                            mSQL = mSQL + " Where intAllotmentID = " & txtAllotmentNo.Tag
                            mSQL = mSQL + " And faIDemandTBL.numDemandID = (Select Max(numDemandID) From faIDemandTBL B Where B.numSubLedgerID = faAllotmentLetters.intAllotmentID)"
                            Rec.Open mSQL, mCnn
                            If Not (Rec.EOF And Rec.BOF) Then
                                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                                    'cmdSave.Enabled = True
                                    cmdApprove.Enabled = False
                                    cmdSave.Enabled = False
                                    cmdNew.Enabled = False
                                    lblMessageBox.Visible = True
                                    lblMessageBox.Caption = "Already Appoved Once!"
                                End If
                                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                                    cmdApprove.Enabled = True
                                    txtAllotmentDate.Tag = ""
                                End If
                   
                            Else
                                cmdApprove.Enabled = False
                                cmdSave.Enabled = False
                                cmdNew.Enabled = False
                                lblMessageBox.Visible = True
                                lblMessageBox.Caption = "Already Appoved Once!"
                            End If
                            
'''                            If (IIf(IsNull(Rec!VrStatus), 0, Rec!VrStatus) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Or (IIf(IsNull(Rec!ReverseStatus), 0, Rec!ReverseStatus) = 1) Then
'''                                cmdRegenerateDemand.Visible = True
'''                                cmdRegenerateDemand.Enabled = True
'''                            Else
'''                                cmdRegenerateDemand.Visible = False
'''                            End If
                            
                            Rec.Close
                            cmdApprove.Enabled = False
                        Else
                            If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                                cmdSave.Enabled = False
                                cmdNew.Enabled = False
                            ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                                'cmdApprove.Enabled = True
                                cmdApprove.Enabled = False
                                cmdSave.Enabled = False
                                cmdNew.Enabled = False
                                lblMessageBox.Visible = True
                                lblMessageBox.Caption = "Already Appoved Once!"
                            End If
                        End If
    '                    cmdReject.Enabled = False
                        
'                    End If
                Else
                    
                    mExtractedStatus = GetStatusFlag
                    If mExtractedStatus <> 2 Then
                        cmdApprove.Enabled = False
                    End If
                    
                    If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                        txtAllotmentDate.Tag = ""
                    End If
                End If
                
            End If
            If Me.AuthorityOrAllotment = "Allotment" Then
                lblTitle.Caption = "Letter of Allotment"
            ElseIf Me.AuthorityOrAllotment = "Authority" Then
                lblTitle.Caption = "Letter of Authority"
                lblNameOfTreasury.Visible = False
                txtNameOfTreasury.Visible = False
                lblTreasuryCode.Visible = False
                txtTreasuryCode.Visible = False
                cmdSearchTreasury.Visible = False
            ElseIf Me.AuthorityOrAllotment = "OpeningAuthority" Then
                lblTitle.Caption = "Opening Letter of Authority"
                lblNameOfTreasury.Visible = False
                txtNameOfTreasury.Visible = False
                lblTreasuryCode.Visible = False
                txtTreasuryCode.Visible = False
                cmdSearchTreasury.Visible = False
            ElseIf Me.AuthorityOrAllotment = "OpeningAllotment" Then
                lblTitle.Caption = "Opening Letter of Allotment"
            End If
            
            If mPreviousYearMode = 1 And gbFinancialYearID = 2017 And gbSaankhyaWeb = 1 Then
                cmdSave.Enabled = True
            Else
                cmdSave.Enabled = False
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Function FillCombo() As Boolean
        On Error GoTo Err:
            Dim mSQL As String
            
            '*********************************************************************************************'
            '                  Function to Fill all the Combo Boxes                                       '
            '*********************************************************************************************'
            mSQL = "Select vchFunctionary,intFunctionaryID From faFunctionaries Where vchFunctionaryCode >= 310000 Order by vchFunctionary"
            PopulateList cmbImplementingOfficer, mSQL, , True, True, True, enuSourceString.Saankhya
            
            If AuthorityOrAllotment = "Authority" Or AuthorityOrAllotment = "OpeningAuthority" Then
                mSQL = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,4,16,17,21,25,26,27,28,10,11,12,13,14,29,30,41)"
                mSQL = mSQL + " Order By vchSourceFundName"
                cmdSearchScheme.Enabled = False
                txtScheme.Enabled = False
            ElseIf AuthorityOrAllotment = "Allotment" Or AuthorityOrAllotment = "OpeningAllotment" Then
                If gbLBPanchayat = 1 Then
                    mSQL = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID IN ( 3,19) " 'NOTE: 128 Added by Aiby
                Else
                    mSQL = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID =3 "
                End If   ' on 06th June,2013 NABARD Reimbursment
            End If
            PopulateList cmbSource, mSQL, , True, True, True, enuSourceString.Saankhya
            
            mSQL = "Select vchTransactionCategory,intCategoryID From faTransactionCategory"
            PopulateList cmbCategory, mSQL, , True, True, True, enuSourceString.Saankhya
            
            If AuthorityOrAllotment = "Authority" Or AuthorityOrAllotment = "OpeningAuthority" Then   '''Added  119,120,121,122,123 On Joint Venture
               If gbLBPanchayat = 1 Then
                    mSQL = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID In(111,108,109,110,125,126,155,166,168,169,170,171,174,119,120,121,122,123) Order By vchTransactionType" ' BLOCKED OTHER LSGI's AS PER NEW GO ON 19/NOV/2015 ,119,120,121,122,123
                Else
                     mSQL = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID In(111,108,109,110,125,126,155,166,168,169,170,171,174,119,120,121,122,123) Order By vchTransactionType"
                End If
            ElseIf AuthorityOrAllotment = "Allotment" Or AuthorityOrAllotment = "OpeningAllotment" Then
                If gbLBPanchayat = 1 Then
                    mSQL = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID In(112,128) Order By vchTransactionType"
                Else
                    mSQL = "Select vchTransactionType,intTransactionTypeID From faTransactionType Where intTransactionTypeID=112 Order By vchTransactionType"
                End If
            End If
            PopulateList cmbTransactionTypes, mSQL, , True, True, True, enuSourceString.Saankhya
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Private Sub Form_Unload(Cancel As Integer)
        mPreviousYearMode = 0
        mPreviousYearRequestID = -1
        mPreviousYearTaskID = -1
    End Sub

    Private Sub lblDDOCodeValue_Click()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        Dim mDDOCode As String
        
        mDDOCode = InputBox("ENTER DDO CODE OF LSGI", "DDO CODE")
        lblDDOCodeValue.Caption = mDDOCode
        mSQL = "UPDATE faConfig SET vchDDOCode= '" & mDDOCode & "'"
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
    End Sub







    Private Sub txtAllotmentDate_LostFocus()
        If txtAllotmentDate.Text <> "" Then
            txtAllotmentDate.Text = CheckDateInMMM(txtAllotmentDate.Text)
        End If
        If intLoadMode = 50 Then
            If Len(Trim(txtAllotmentDate)) Then
                txtAllotmentDate.Text = CheckDateInMMM(txtAllotmentDate.Text)
            End If
            
            If IsDate(txtAllotmentDate.Text) Then
                'If CDate(txtTrnDate.Text) >= mPreStartDate And CDate(txtTrnDate.Text) <= mPreEndDate Then
                If CDate(txtAllotmentDate.Text) < DateAdd("yyyy", -1, gbStartingDate) Or CDate(txtAllotmentDate.Text) > DateAdd("yyyy", -1, gbEndingDate) Then
                    MsgBox "Please Enter a Date betwwen Previous financialYear", vbApplicationModal
                    txtAllotmentDate.Text = ""
                    Exit Sub
                End If
            Else
                txtAllotmentDate.Text = ""
            End If
        End If
        
    End Sub
    Private Sub FetchDDOCode()
        Dim mCnn    As New ADODB.Connection
        Dim objDB   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSQL    As String
        
        If objDB.SetConnection(mCnn) Then
            mSQL = "SELECT vchDDOCode FROM faConfig"
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                lblDDOCodeValue.Caption = IIf(IsNull(Rec!vchDDOCode), "", Rec!vchDDOCode)
            End If
            Rec.Close
        End If
    End Sub
    Private Sub ReFillDetails()
        On Error GoTo Err:
            Dim mCnn    As New ADODB.Connection
            Dim objDB   As New clsDB
            Dim Rec     As New ADODB.Recordset
            Dim mSQL    As String
            
            '*********************************************************************************************'
            '                  Procedure to fill all the details for viewing, editing or Approval         '
            '*********************************************************************************************'
            If objDB.SetConnection(mCnn) Then
                'mSql = "Select *,faAllotmentLetters.fltAmount As Amount, faAllotmentLetters.intFunctionaryID As FunctionaryID, faAllotmentLetters.intFunctionID As FunctionID, A.vchAccountHeadCode As CreditAccountHeadCode, B.vchAccountHeadCode As GrossAccountHeadCode, A.vchAccountHead As CreditAccountHead, B.vchAccountHead As GrossAccountHead,C.vchFunctionary As ImplementingOfficer,D.vchFunctionary As Functionary, suSourceOfFund.vchSourceFundName as Source  , faTransactionCategory.vchTransactionCategory as Categery, F.vchAccountHead as Scheme,F.intAccountHeadID SchemeID, G.intTransactionTypeID as TrTypeID, G.vchTransactionType as TrType,faIDemandTBL.numDemandID As DemandID,faVouchers.intVoucherID As VoucherID  From faAllotmentLetters"
                mSQL = "Select *,faAllotmentLetters.tnyStatus As AllotmentStatus,faAllotmentLetters.fltAmount As Amount, faAllotmentLetters.intFunctionaryID As FunctionaryID, faAllotmentLetters.intFunctionID As FunctionID, A.vchAccountHeadCode As CreditAccountHeadCode, B.vchAccountHeadCode As GrossAccountHeadCode, A.vchAccountHead As CreditAccountHead, B.vchAccountHead As GrossAccountHead,C.vchFunctionary As ImplementingOfficer,D.vchFunctionary As Functionary, suSourceOfFund.vchSourceFundName as Source  , faTransactionCategory.vchTransactionCategory as Categery, faDepSchPro.vchDescription as Scheme,faDepSchPro.intID SchemeID, G.intTransactionTypeID as TrTypeID, G.vchTransactionType as TrType,faIDemandTBL.numDemandID As DemandID,faVouchers.intVoucherID As VoucherID  "
                
                mSQL = mSQL + " ,faVouchers.tnyStatus VrStatus,faVouchers.tnyCancelFlag  CancelFlag,faVouchers.tnyReversed ReverseStatus From faAllotmentLetters"
                
                'If mPreviousYearTaskID = 1 Then
                '    mSql = mSql + " LEFT JOIN faPendingTaskRequest ON faPendingTaskRequest.intKeyID = faAllotmentLetters.intAllotmentID AND faPendingTaskRequest.intTaskID = 1 AND NOT faPendingTaskRequest.tnyStatus IN (0,4)"
                'Else
                '    mSql = mSql + " LEFT JOIN faPendingTaskRequest ON faPendingTaskRequest.intKeyID = faAllotmentLetters.intAllotmentID AND faPendingTaskRequest.intTaskID = 4 AND NOT faPendingTaskRequest.tnyStatus IN (0,4)"
                'End If
                
                mSQL = mSQL + " LEFT JOIN faPendingTaskRequest ON faPendingTaskRequest.intKeyID = faAllotmentLetters.intAllotmentID AND faPendingTaskRequest.intTaskID  IN (1,4) AND NOT faPendingTaskRequest.tnyStatus IN (0,4)"
                mSQL = mSQL + " Left Join suProjectDetails On faAllotmentLetters.numProjectID = suProjectDetails.decProjectID"
                mSQL = mSQL + " Left Join suSourceOfFund On faAllotmentLetters.intSourceOfFundID = suSourceOfFund.intSourceFundID"
                mSQL = mSQL + " Left Join faFunctions On faAllotmentLetters.intFunctionID = faFunctions.intFunctionID"
                mSQL = mSQL + " Left Join faFunctionaries C On faAllotmentLetters.intImplementingOfficersID = C.intFunctionaryID"
                mSQL = mSQL + " Left Join faFunctionaries D On faAllotmentLetters.intFunctionaryID = D.intFunctionaryID"
                mSQL = mSQL + " Left Join faAccountHeads A On faAllotmentLetters.intCrAccountHeadID = A.intAccountHeadID"
                mSQL = mSQL + " Left Join faAccountHeads B On faAllotmentLetters.intGrossAccountHeadID = B.intAccountHeadID"
                mSQL = mSQL + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID =faAllotmentLetters.intCategoryID "
                'mSql = mSql + " Left Join faAccountHeads F On faAllotmentLetters.intSchemeID = F.intAccountHeadID "
                mSQL = mSQL + " Left Join faDepSchPro On faAllotmentLetters.intSchemeID = faDepSchPro.intID "
                mSQL = mSQL + " Left Join faTransactionType G on G.intTransactionTypeID = faAllotmentLetters.intTransactionTypeID "
                mSQL = mSQL + " Left Join faIDemandTBL On faAllotmentLetters.intAllotmentID = faIDemandTBL.numSubLedgerID And faAllotmentLetters.intTransactionTypeID = faIDemandTBL.intTransactionTypeID"
                mSQL = mSQL + " Left Join faVouchers On faIDemandTBL.intVoucherID = faVouchers.intVoucherID"
                mSQL = mSQL + " Where intAllotmentID = '" & Trim(txtAllotmentNo.Tag) & "'"
                mSQL = mSQL + " And faAllotmentLetters.tnyStatus <> 8"
                If AuthorityOrAllotment = "Authority" Then
                    mSQL = mSQL + " And faAllotmentLetters.intSourceOfFundID In(1,2,4,16,17,21,25,26,27,28,10,11,12,13,14,29,30,41)"
                ElseIf AuthorityOrAllotment = "Allotment" Then
                    mSQL = mSQL + " And faAllotmentLetters.intSourceOfFundID In(3,19)"
                ElseIf AuthorityOrAllotment = "OpeningAuthority" Then
                    mSQL = mSQL + " And faAllotmentLetters.intSourceOfFundID In(1,2,4,16,17,21,25,26,27,28,10,11,12,13,14,29,30,41) and tnyOpening=1"
                ElseIf AuthorityOrAllotment = "OpeningAllotment" Then
                    mSQL = mSQL + " And faAllotmentLetters.intSourceOfFundID In(3)and tnyOpening=1 "
                End If
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    
                    mPreviousYearRequestID = IIf(IsNull(Rec!intRequestID), -1, Rec!intRequestID)
                    If IsNull(Rec!Source) Then
                        cmbSource.ListIndex = -1
                    Else
                        cmbSource.Text = Rec!Source
                    End If
                    
                    If IsNull(Rec!TrTypeID) Then
                        cmbTransactionTypes.ListIndex = -1
                    Else
                        cmbTransactionTypes.Text = Rec!TrType
                    End If
                    
                    
                    If IsNull(Rec!Categery) Then
                        cmbCategory.ListIndex = -1
                    Else
                        cmbCategory.Text = Rec!Categery
                    End If
                    
                    
                    
                    If IsNull(Rec!ImplementingOfficer) Then
                        cmbImplementingOfficer.ListIndex = -1
                    Else
                        cmbImplementingOfficer.Text = Rec!ImplementingOfficer
                    End If
                    txtAllotmentNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                    txtAllotmentNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    txtAllotmentDate.Text = IIf(IsNull(Rec!dtAllotmentDate), "", Rec!dtAllotmentDate)
                    txtAllotmentDate.Tag = IIf(IsNull(Rec!DemandID), "", Rec!DemandID)
                    txtInstalmentNo.Text = IIf(IsNull(Rec!intInstalmentNo), "", Rec!intInstalmentNo)
                    txtInstalmentNo.Tag = IIf(IsNull(Rec!VoucherID), "", Rec!VoucherID)
                    txtAmountInFigures.Text = IIf(IsNull(Rec!Amount), "", Rec!Amount)
                    txtScheme.Text = IIf(IsNull(Rec!Scheme), "", Rec!Scheme)
                    txtScheme.Tag = IIf(IsNull(Rec!SchemeID), "", Rec!SchemeID)
                    txtTreasuryCode.Text = IIf(IsNull(Rec!vchTreasuryCode), "", Rec!vchTreasuryCode)
                    txtNameOfTreasury.Text = IIf(IsNull(Rec!vchTreasuryName), "", Rec!vchTreasuryName)
                    txtCreditAccountHeadCode.Text = IIf(IsNull(Rec!CreditAccountHeadCode), "", Rec!CreditAccountHeadCode)
                    txtCreditAccountHeadCode.Tag = IIf(IsNull(Rec!intCrAccountHeadID), "", Rec!intCrAccountHeadID)
                    txtCreditAccountHead.Text = IIf(IsNull(Rec!CreditAccountHead), "", Rec!CreditAccountHead)
                    txtFunctionary.Text = IIf(IsNull(Rec!Functionary), "", Rec!Functionary)
                    txtFunctionary.Tag = IIf(IsNull(Rec!FunctionaryID), "", Rec!FunctionaryID)
                    txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                    txtFunction.Tag = IIf(IsNull(Rec!FunctionID), "", Rec!FunctionID)
                    txtAccountHeadCode.Text = IIf(IsNull(Rec!GrossAccountHeadCode), "", Rec!GrossAccountHeadCode)
                    txtAccountHeadCode.Tag = IIf(IsNull(Rec!intGrossAccountHeadID), "", Rec!intGrossAccountHeadID)
                    txtAccountHead.Text = IIf(IsNull(Rec!GrossAccountHead), "", Rec!GrossAccountHead)
                    txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                    txtProjectNo.Tag = IIf(IsNull(Rec!numProjectID), "", Rec!numProjectID)
                    txtProjectName.Text = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
                    txtPublicGrantHead.Text = IIf(IsNull(Rec!vchPublicGrantHead), "", Rec!vchPublicGrantHead)
                    txtPublicBudgetHead.Text = IIf(IsNull(Rec!vchPublicBudgetHead), "", Rec!vchPublicBudgetHead)
                    txtGONumberValue.Text = IIf(IsNull(Rec!vchGONumber), "", Rec!vchGONumber)
                    txtNatureOfClaim.Text = IIf(IsNull(Rec!vchNatureOfClaim), "", Rec!vchNatureOfClaim)
                    
                    cmbCategory.Enabled = False
                    cmdSearchScheme.Enabled = False
                    cmdCreditAccountHead.Enabled = False
                    If Rec!AllotmentStatus = 1 Then
                        cmdApprove.Enabled = False
                    Else
                        cmdApprove.Enabled = True
                    End If
                    
'''                    If (IIf(IsNull(Rec!VrStatus), 0, Rec!VrStatus) = 4 And IIf(IsNull(Rec!CancelFlag), 0, Rec!CancelFlag) = 1) Or (Rec!ReverseStatus = 2) Then
'''                        cmdRegenerateDemand.Visible = True
'''                        cmdRegenerateDemand.Enabled = True
'''                    Else
'''                        cmdRegenerateDemand.Visible = False
'''                    End If
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection to Finance does not exist, Please contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub txtAllotmentNo_LostFocus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objDB   As New clsDB
        Dim mSQL    As String
        
        '*********************************************************************************************'
        '                  Procedure to check duplication of Letter of Authority/Allotment            '
        '*********************************************************************************************'
        On Error GoTo Err
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        If Trim(txtAllotmentNo.Text) <> "" Then
            mSQL = "Select intAllotmentID,tnyStatus From faAllotmentLetters"
            mSQL = mSQL + " Where vchAllotmentNo = '" & txtAllotmentNo.Text & "'"
            mSQL = mSQL + " And tnyStatus <> 8"
            If AuthorityOrAllotment = "Authority" Then
                mSQL = mSQL + " And intSourceOfFundID In(1,4,16,17,25,26,27,28,10,11,12,13,14,29,30,41)"
            ElseIf AuthorityOrAllotment = "Allotment" Then
                mSQL = mSQL + " And intSourceOfFundID In(3)"
            End If
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtAllotmentNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                ReFillDetails
'                If Not (IsNull(Rec!tnyStatus)) Then
'                    If Rec!tnyStatus = 1 Then
                        cmdSave.Enabled = False
'                    Else
'                        cmdSave.Enabled = True
'                    End If
'                End If
            Else
                txtAllotmentNo.Tag = ""
            End If
        End If
        Exit Sub
Err:
        MsgBox Err.Description
    End Sub

    Private Sub txtAmountInFigures_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtAmountInFigures_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
        txtAmountInFigures.Locked = True
    Else
        txtAmountInFigures.Locked = False
    End If
    End Sub



    Private Sub txtInstalmentNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtProjectNo_LostFocus()
        '''Dim mCnn    As New ADODB.Connection
        '''Dim objDB   As New clsDB
        '''Dim Rec     As New ADODB.Recordset
        '''Dim mSQL    As String
        '''
        '''objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
        '''
        '''If txtProjectNo.Text <> "" Then
        '''    mSQL = "Select * From suProjectDetails"
        '''    mSQL = mSQL + " Where chvProjectSlNo Like  '" & txtProjectNo.Text & "%'"
        '''    Rec.Open mSQL, mCnn
        '''    If Not (Rec.EOF And Rec.BOF) Then
        '''        txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
        '''        txtProjectNo.Tag = IIf(IsNull(Rec!decProjectID), "", Rec!decProjectID)
        '''        txtProjectName.Text = IIf(IsNull(Rec!chvProjectName), "", Rec!chvProjectName)
        '''    End If
        '''    Rec.Close
        '''End If
    End Sub
    
    Private Sub GenerateDemand(ByRef mCnn As ADODB.Connection, ByVal mAllotmentID As Variant)
        On Error GoTo Err:
            Dim Demand As uDemand
            Dim DemandChild As uDemandChild
            Dim DemandAddress As uDemandAddress
            
            Dim aryIn As Variant
            Dim aryOut As Variant
            Dim objDB As New clsDB
            
            '*********************************************************************************************'
            '                  Procedure to generate the Demand on Approval of "Letter of Authority"      '
            '*********************************************************************************************'
            With Demand
                If mPreviousYearMode = 1 Then
                    If IsDate(txtAllotmentDate) Then
                        .dtDemandDate = CDate(txtAllotmentDate)
                    Else
                        .dtDemandDate = DateAdd("yyyy", -1, gbEndingDate)
                    End If
                    .dtDueDate = .dtDemandDate
                    .dtExpiryDate = DateAdd("yyyy", -1, gbEndingDate)
                    .dtInstrumentDate = .dtDemandDate
                    .intFinancialYearID = gbFinancialYearID - 1
                    .intYearID = gbFinancialYearID - 1
                Else
                    .dtDemandDate = gbTransactionDate
                    .dtDueDate = gbTransactionDate
                    .dtExpiryDate = gbEndingDate 'gbTransactionDate
                    .dtInstrumentDate = gbTransactionDate
                    .intFinancialYearID = gbFinancialYearID
                    .intYearID = gbFinancialYearID
                End If
                .dtVoucherDate = Null
                .intDoorNo = Null
                .intFunctionaryID = val(txtFunctionary.Tag)
                .intFunctionID = val(txtFunction.Tag)
                
                .intInstrumentTypeID = 6
                
                'BLOCKED ON 03-Aug-2013
                'If cmbSource.ItemData(cmbSource.ListIndex) = 10 Or _
                '   cmbSource.ItemData(cmbSource.ListIndex) = 11 Or _
                '   cmbSource.ItemData(cmbSource.ListIndex) = 12 Or _
                '   cmbSource.ItemData(cmbSource.ListIndex) = 13 Or _
                '   cmbSource.ItemData(cmbSource.ListIndex) = 14 Then
                '  .intInstrumentTypeID = Null
                'Else
                '  .intInstrumentTypeID = 6
                'End If
                '
                .intKeyID = txtCreditAccountHeadCode.Tag
                .intKeyID2 = Null
                .intLBID = gbLocalBodyID
                .intSectionID = gbSectionID
                .intSourceFundID = cmbSource.ItemData(cmbSource.ListIndex)
                .intTransactionTypeID = cmbTransactionTypes.ItemData(cmbTransactionTypes.ListIndex)
                .intVoucherID = Null
                .intWardNo = Null
                
                .numCounterID = gbCounterID
                If txtAllotmentDate.Tag <> "" Then          'numDemandID
                    .numDemandID = txtAllotmentDate.Tag
                Else
                    .numDemandID = Null
                End If
                
                .numForwardedSeatID = Null
                .numLocationID = gbLocationID
                .numSeatID = gbSeatID
                .numSubLedgerID = val(mAllotmentID)
                .numUserID = gbUserID
                .numZoneID = gbLocationID
                .tnyAccrualType = Null
                .tnyArrearFlag = Null
                .tnyDemandType = 10
                .tnyExtAppID = Null
                .tnyExtModuleID = Null
                .tnyPeriodID = Null
                .tnySend = Null
                .tnyStatus = 0
                .vchAdminNote = Null
                .vchDemandNo = Null
                .vchDoorNo2 = Null
                .vchDrawnFrom = Trim(txtNameOfTreasury.Text)
                .vchDrawnPlace = Null
                .vchInstrumentNo = Trim(txtAllotmentNo.Text)
                .dtTransactionDate = Null
                .intDemandMode = Null
                .vchRemarks = "Allotment Letter"
                
                aryIn = Array(.intLBID, .tnyExtAppID, .tnyExtModuleID, .tnyDemandType, _
                                .intTransactionTypeID, .intYearID, .tnyPeriodID, .dtDemandDate, _
                                .numSubLedgerID, .intKeyID, .intKeyID2, .vchRemarks, .tnyStatus, _
                                .intVoucherID, .dtVoucherDate, .tnyArrearFlag, .dtExpiryDate, _
                                .numDemandID, .intFinancialYearID, .numSeatID, .intSectionID, _
                                .numUserID, .numCounterID, .vchAdminNote, .vchDemandNo, _
                                .numZoneID, .intWardNo, .intDoorNo, .vchDoorNo2, _
                                .numForwardedSeatID, .dtDueDate, .intInstrumentTypeID, _
                                .vchInstrumentNo, .dtInstrumentDate, .vchDrawnFrom, _
                                .vchDrawnPlace, .tnyAccrualType, .numLocationID, _
                                .intFunctionaryID, .intFunctionID, .intSourceFundID, .dtTransactionDate, .intDemandMode)
                
                objDB.ExecuteSP "spSaveIDemandTBL", aryIn, aryOut, , mCnn, adCmdStoredProc
            End With
            
            With DemandChild
                .dtOnDate = gbTransactionDate
                .dtVoucherDate = gbTransactionDate
                .fltAmount = val(txtAmountInFigures.Text)
                .intAccountHeadID = val(txtAccountHeadCode.Tag)
                .intLBID = gbLocalBodyID
                .intVoucherID = Null
                .intYearID = Null
                .numDemandID = aryOut(0, 0)
                .snyRate = Null
                .tnyArrearFlag = Null
                .tnyPeriodID = Null
                .tnySlNo = 1
                .tnyStatus = 0
                .vchAccountHeadCode = Trim(txtAccountHeadCode.Text)
                .intTransactionTypeID = Null
                .vchRemarks = "Allotment Letter"
         
                If .numDemandID <> "" Then
                    mCnn.Execute "Delete From faIDemandChild Where numDemandID = " & .numDemandID
                End If
                aryIn = Array(.numDemandID, .intLBID, .tnySlNo, .intAccountHeadID, _
                                .vchAccountHeadCode, .fltAmount, .vchRemarks, .tnyStatus, _
                                .dtOnDate, .intYearID, .tnyPeriodID, .tnyArrearFlag, .intTransactionTypeID)
                
                objDB.ExecuteSP "spSaveIDemandChild", aryIn, , , mCnn, adCmdStoredProc
            End With
            
            With DemandAddress
                .numDemandID = aryOut(0, 0)
                If gbLBType = 3 Then
                    .vchName = "Director of Urban Affairs"
                ElseIf gbLBType = 4 Then
                    .vchName = "Secretary,LSGD"
                End If
                If .numDemandID <> "" Then
                    mCnn.Execute "Delete From faIDemandAddress Where numDemandID = " & .numDemandID
                End If
                aryIn = Array(.numDemandID, _
                gbLocalBodyID, _
                Null, _
                Null, _
                Null, _
                Null, _
                .vchName, _
                Null, Null, Null, Null, _
                Null, _
                Null, _
                Null, _
                Null, _
                Null, _
                Null, _
                Null)
                
                objDB.ExecuteSP "spSaveIDemandAddress", aryIn, , , mCnn, adCmdStoredProc
            
            
            End With
            
            If mPreviousYearMode Then
            If mPreviousYearRequestID > 0 Then
                Dim mSQL As String
                mSQL = "Update faPendingTaskRequest SET  numDemandID = " & aryOut(0, 0) & " Where intRequestID = " & mPreviousYearRequestID & "  "
                objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            End If
            
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    
    Public Property Let LoadMode(mData As Integer)
        intLoadMode = mData
    End Property
    
    Public Property Get LoadMode() As Integer
        LoadMode = intLoadMode
    End Property
    
    Public Property Let AllotmentID(mData As Variant)
        intAllotmentID = mData
    End Property
    
    Public Property Get AllotmentID() As Variant
        AllotmentID = intAllotmentID
    End Property
    
    Public Property Let ApproveStatus(mData As Integer)
        tnyAppoveStatus = mData
    End Property
    
    Public Property Get ApproveStatus() As Integer
        ApproveStatus = tnyAppoveStatus
    End Property
    
    Public Property Let AuthorityOrAllotment(mData As Variant)
        strAuthorityOrAllotment = mData
    End Property
    
    Public Property Get AuthorityOrAllotment() As Variant
        AuthorityOrAllotment = strAuthorityOrAllotment
    End Property
     Public Property Let CheckDemand(mData As Variant)
        mCheckDemand = mData
    End Property
    Public Property Get CheckDemand() As Variant
        CheckDemand = mCheckDemand
    End Property
      Public Property Let PDEMode(mPDE As Variant)
        mPDEMode = mPDE
    End Property
    Public Property Get PDEMode() As Variant
        PDEMode = mPDEMode
    End Property
    
    Public Property Let PreviousYearMode(mData As Integer)
        mPreviousYearMode = mData
    End Property

    Public Property Let PreviousYearTaskID(mData As Integer)
        mPreviousYearTaskID = mData
    End Property
    
    Public Property Let PreviousYearRequestID(mData As Integer)
        mPreviousYearRequestID = mData
    End Property


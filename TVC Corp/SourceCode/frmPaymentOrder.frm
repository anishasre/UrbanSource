VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmPaymentOrder 
   BackColor       =   &H00EDF7F7&
   Caption         =   "P a y m e n t   O r d e r"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   14715
   Icon            =   "frmPaymentOrder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTreasuryID 
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
      Left            =   13440
      TabIndex        =   118
      Top             =   1680
      Visible         =   0   'False
      Width           =   1470
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
      Height          =   285
      Left            =   11970
      Locked          =   -1  'True
      TabIndex        =   112
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdAllotmentLetterNo 
      BackColor       =   &H00F5FCFC&
      Caption         =   "..."
      Height          =   300
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1320
      Width           =   300
   End
   Begin VB.CommandButton cmdDeductionPayment 
      Caption         =   "..."
      Height          =   300
      Left            =   8640
      TabIndex        =   107
      Top             =   3510
      Width           =   360
   End
   Begin VB.TextBox txtAllotedAmt 
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
      Left            =   12480
      TabIndex        =   104
      Top             =   3330
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   1650
      TabIndex        =   90
      Top             =   8925
      Width           =   8415
      Begin VB.CommandButton cmdVerify 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Verify"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   135
         Visible         =   0   'False
         Width           =   915
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
         Left            =   4020
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   120
         Width           =   1410
      End
      Begin VB.CommandButton cmdReject 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Return"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   150
         Visible         =   0   'False
         Width           =   915
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
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "C&Lose"
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
         Left            =   6900
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdApproval 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Approve"
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
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
      Left            =   12480
      TabIndex        =   88
      Top             =   3630
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.ListBox lstMasters 
      BackColor       =   &H00E0E0E0&
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
      Left            =   14460
      TabIndex        =   68
      Top             =   2160
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4485
      Left            =   12000
      Picture         =   "frmPaymentOrder.frx":1CCA
      ScaleHeight     =   4425
      ScaleWidth      =   2625
      TabIndex        =   84
      Top             =   3930
      Width           =   2685
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   3525
         Left            =   150
         TabIndex        =   86
         Top             =   735
         Width           =   2235
      End
      Begin VB.Label lblHelpTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payment Order"
         Height          =   255
         Left            =   540
         TabIndex        =   85
         Top             =   180
         Width           =   2655
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4635
      Left            =   0
      TabIndex        =   70
      Top             =   4230
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   8176
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
      TabPicture(0)   =   "frmPaymentOrder.frx":1FD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "<Subledger>"
      TabPicture(1)   =   "frmPaymentOrder.frx":1FF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "..."
      TabPicture(2)   =   "frmPaymentOrder.frx":200C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4FCFC&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   3315
         Left            =   -74940
         TabIndex        =   78
         Top             =   60
         Width           =   11850
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F4FCFC&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   60
         TabIndex        =   71
         Top             =   60
         Width           =   11835
         Begin VB.CheckBox chkPensionContribution 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Pension Contribution"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2370
            TabIndex        =   121
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtCP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9870
            TabIndex        =   119
            Top             =   60
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CommandButton cmdGo 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   11250
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   1080
            Width           =   300
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
            Height          =   270
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   1080
            Width           =   3405
         End
         Begin VB.CheckBox chkPension 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Auto pensionContribution"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   90
            TabIndex        =   109
            Top             =   90
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox txtPensionAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7785
            TabIndex        =   108
            Top             =   45
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00F4FCFC&
            BorderStyle     =   0  'None
            Height          =   2730
            Left            =   90
            TabIndex        =   91
            Top             =   360
            Width           =   5430
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
            Begin VB.CommandButton cmdSubLederType 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   300
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   150
               Width           =   345
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
               Height          =   270
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   20
               Top             =   180
               Width           =   3540
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
               Height          =   285
               Left            =   4725
               MaxLength       =   1
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   555
               Width           =   315
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
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   22
               Top             =   555
               Width           =   2190
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
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   28
               Top             =   870
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
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   29
               Top             =   1185
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
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   30
               Top             =   1500
               Width           =   3525
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
               Height          =   285
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   31
               Top             =   1815
               Width           =   3525
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
               Height          =   285
               Left            =   3735
               MaxLength       =   1
               TabIndex        =   23
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
               Height          =   285
               Left            =   4065
               MaxLength       =   1
               TabIndex        =   24
               Top             =   555
               Width           =   315
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
               Height          =   285
               Left            =   4395
               MaxLength       =   1
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   555
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
               Height          =   285
               Left            =   1500
               MaxLength       =   50
               TabIndex        =   32
               Top             =   2130
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
               Height          =   285
               Left            =   4110
               MaxLength       =   6
               TabIndex        =   33
               Top             =   2130
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
               Height          =   285
               Left            =   1500
               MaxLength       =   15
               TabIndex        =   34
               Top             =   2445
               Width           =   2220
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
               TabIndex        =   100
               Top             =   240
               Width           =   1395
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
               TabIndex        =   99
               Top             =   1215
               Width           =   525
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
               TabIndex        =   98
               Top             =   885
               Width           =   1095
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
               TabIndex        =   97
               Top             =   570
               Width           =   1305
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
               TabIndex        =   96
               Top             =   2160
               Width           =   360
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
               TabIndex        =   95
               Top             =   1845
               Width           =   900
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
               TabIndex        =   94
               Top             =   1530
               Width           =   945
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
               TabIndex        =   93
               Top             =   2160
               Width           =   255
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
         End
         Begin VB.Frame fraProject 
            BackColor       =   &H00F4FCFC&
            BorderStyle     =   0  'None
            Height          =   1755
            Left            =   5700
            TabIndex        =   87
            Top             =   1485
            Width           =   5910
            Begin VB.CommandButton cmdImplementingOfficer 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   495
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
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   510
               Width           =   3405
            End
            Begin VB.CommandButton cmdProjectNo 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   810
               Width           =   300
            End
            Begin VB.CommandButton cmdAgreementNo 
               BackColor       =   &H00F5FCFC&
               Caption         =   "..."
               Height          =   285
               Left            =   5565
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   195
               Width           =   300
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
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   210
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
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   810
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
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1410
               Width           =   3405
            End
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
               Height          =   270
               Left            =   2115
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1110
               Width           =   3405
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
               Height          =   225
               Left            =   -30
               TabIndex        =   38
               Top             =   540
               Width           =   2100
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
               TabIndex        =   41
               Top             =   840
               Width           =   1530
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
               Left            =   1455
               TabIndex        =   46
               Top             =   1425
               Width           =   630
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
               Left            =   1200
               TabIndex        =   44
               Top             =   1125
               Width           =   885
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
               Left            =   675
               TabIndex        =   35
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F4FCFC&
            Caption         =   "Final Bill"
            Height          =   195
            Left            =   1605
            TabIndex        =   57
            Top             =   3765
            Width           =   930
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
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   375
            Width           =   3420
         End
         Begin VB.TextBox txtForward2Seat 
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
            Left            =   9180
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   3765
            Width           =   1725
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
            Height          =   315
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   375
            Width           =   3420
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
            Height          =   315
            Left            =   7785
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   735
            Width           =   3420
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
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   705
            Width           =   3420
         End
         Begin VB.CommandButton cmdSeat 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   10935
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   3750
            Width           =   315
         End
         Begin VB.CommandButton cmdSourceOfFund 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   11250
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   735
            Width           =   300
         End
         Begin VB.CommandButton cmdSubsidiaryCash 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   285
            Left            =   11250
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   360
            Width           =   300
         End
         Begin VB.CommandButton cmdAsset 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            Height          =   300
            Left            =   12375
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1665
            Width           =   345
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
            Height          =   495
            Left            =   1575
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   3240
            Width           =   9630
         End
         Begin VB.Label lblCP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "CP"
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
            Left            =   9600
            TabIndex        =   120
            Top             =   90
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblGo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Go No"
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
            Left            =   7110
            TabIndex        =   116
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label lblPension 
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Contribution Amount"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4920
            TabIndex        =   110
            Top             =   90
            Visible         =   0   'False
            Width           =   2865
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
            TabIndex        =   77
            Top             =   420
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
            TabIndex        =   76
            Top             =   705
            Width           =   480
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
            TabIndex        =   75
            Top             =   780
            Width           =   1470
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
            TabIndex        =   74
            Top             =   450
            Width           =   1560
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Forward to Seat"
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
            Left            =   7575
            TabIndex        =   49
            Top             =   3795
            Width           =   1560
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   555
            TabIndex        =   48
            Top             =   3375
            Width           =   930
         End
      End
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
      Left            =   8490
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   19
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E7F1F1&
      Height          =   1005
      Left            =   90
      TabIndex        =   67
      Top             =   45
      Width           =   11925
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
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   525
         Width           =   6075
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   525
         Width           =   270
      End
      Begin MSComCtl2.DTPicker dtpDueDate 
         Height          =   360
         Left            =   7485
         TabIndex        =   4
         Top             =   195
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   635
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   41371
      End
      Begin VB.TextBox txtDueDate 
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
         Left            =   6105
         TabIndex        =   3
         Top             =   210
         Width           =   1350
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   195
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchFunction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   510
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
         Height          =   285
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   210
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
         Height          =   285
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   525
         Width           =   2760
      End
      Begin VB.TextBox txtDated 
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
         Left            =   3555
         TabIndex        =   2
         Top             =   195
         Width           =   1620
      End
      Begin VB.TextBox txtPayOrder 
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
         Height          =   270
         Left            =   1395
         TabIndex        =   1
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
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
         Left            =   5250
         TabIndex        =   69
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCBCB&
         BackStyle       =   0  'Transparent
         Caption         =   "  Dated"
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
         Left            =   2895
         TabIndex        =   61
         Top             =   225
         Width           =   630
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
         TabIndex        =   64
         Top             =   540
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
         Left            =   7800
         TabIndex        =   63
         Top             =   255
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
         Left            =   8070
         TabIndex        =   62
         Top             =   555
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "P. Order No"
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
         Left            =   45
         TabIndex        =   0
         Top             =   225
         Width           =   1245
         WordWrap        =   -1  'True
      End
   End
   Begin WinXPC_Engine.WindowsXPC XPC 
      Left            =   14445
      Top             =   10800
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.TextBox txtDrAmount 
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
      Left            =   8490
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1125
      Width           =   1890
   End
   Begin VB.TextBox txtCrHeadCode 
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
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3060
      Width           =   1710
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
      Height          =   285
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3060
      Width           =   3855
   End
   Begin VB.CommandButton cmdCrAccountHead 
      Caption         =   "..."
      Height          =   300
      Left            =   8190
      TabIndex        =   18
      Top             =   3045
      Width           =   270
   End
   Begin VB.TextBox txtDrHeadCode 
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
      Left            =   2490
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1125
      Width           =   1710
   End
   Begin VB.TextBox txtDrAccountHead 
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
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1125
      Width           =   3945
   End
   Begin VB.CommandButton cmdDrAccountHead 
      Caption         =   "..."
      Height          =   300
      Left            =   8190
      TabIndex        =   13
      Top             =   1125
      Width           =   270
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1560
      Left            =   2415
      TabIndex        =   15
      Top             =   1470
      Width           =   8295
      _cx             =   14631
      _cy             =   2752
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   3
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPaymentOrder.frx":2028
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
      Left            =   11970
      TabIndex        =   113
      Top             =   1095
      Width           =   1905
   End
   Begin VB.Label lblDedExclude 
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction Exclude From Payment"
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
      Left            =   5715
      TabIndex        =   106
      Top             =   3555
      Width           =   2940
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Alloted Amount"
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
      Left            =   11055
      TabIndex        =   105
      Top             =   3420
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(Recoveries)"
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
      Left            =   1275
      TabIndex        =   103
      Top             =   1935
      Width           =   1095
   End
   Begin VB.Label Label18 
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
      Left            =   9990
      TabIndex        =   89
      Top             =   3645
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblUtilizedAmt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4605
      TabIndex        =   83
      Top             =   3510
      Width           =   480
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Utilized Amount:"
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
      Left            =   2925
      TabIndex        =   82
      Top             =   3525
      Width           =   1605
   End
   Begin VB.Label lblBudgetAmt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2130
      TabIndex        =   81
      Top             =   3525
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Amount:"
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
      Left            =   480
      TabIndex        =   80
      Top             =   3540
      Width           =   1545
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Payable (Acc.Head)"
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
      TabIndex        =   66
      Top             =   3090
      Width           =   2025
   End
   Begin VB.Label lblDrAccountHead 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Head (Dr)"
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
      Left            =   855
      TabIndex        =   65
      Top             =   1185
      Width           =   1590
   End
End
Attribute VB_Name = "frmPaymentOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    'tnyStatus in faPayOrderTable---- 0 Request,1 Approve, 2 Approve and Paid, 3 Verify, 5 Forward
    
    Dim mSelect As Boolean
    Dim mBudgetBalanceAmt As Variant
    Dim mvarGrossSalryID As Variant
    Dim intPayOrderID As Variant
    Dim vchPayOrderNo As Variant
    Dim intLoadMode As Integer
    Dim mViewPayOrderListFormIsLoaded As Boolean
    Dim mWaterBillPOMode As Boolean
    Dim mModuleID As Integer
    Dim mAssetID As Integer
    Dim mAssetTypeID As Integer
    Dim mHelpTips As String
    
    '-------------------------'
    Dim intModuleID As Variant
    Dim intKeyID As Variant
    Dim dtKeyDate As Variant
    '-------------------------'
    Dim mSkipMsgFlag As Boolean
    
    Dim mOldRequisition As Boolean
    
    Dim mSelectCreditHeadFlag As Boolean ' TRUE = Leave for Selection of Credit Head
                                         ' FALSE = Will fill Debit Head as Credit Head
    Dim mPendingTask        As Integer  '1 Pending Task in Previous Year (Approval),2-Preyear Pay Order,3Pay order approval (through view pay order)
    Dim mPendingTaskReqID   As Integer  ' Pending Task RequestID for PayOrder
    Dim mPendingTransactionDate As Date
    Dim mPendingTransactionType As Integer
    Dim mPendingAllotmentNo As Long
    Dim mPendingAmt As Double
    Dim mPendingExpHeadID As Integer
    Dim mUnAuthorized As Integer
    Dim mCpFlag As Boolean
    
    
    
       
    Private Sub GetImplementingOfficer(mProjectID As Variant)
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        
        On Error GoTo err
        If objdb.CreateNewConnection(mCnn, enuSourceString.Sulekha) Then
            mSql = "Select chvImplOfficerDesgEng,M_ImplOfficer.intImplOfficerID From ProjectDetails"
            mSql = mSql + " Inner Join M_ImplOfficer On M_ImplOfficer.intImplOfficerID = ProjectDetails.intImplOfficerID"
            mSql = mSql + " Where decProjectID = " & mProjectID
            mSql = mSql + " And intLBID = " & gbLocalBodyID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                txtImplementingOfficer.Text = IIf(IsNull(Rec!chvImplOfficerDesgEng), "", Rec!chvImplOfficerDesgEng)
                txtImplementingOfficer.Tag = IIf(IsNull(Rec!intImplOfficerID), "", Rec!intImplOfficerID)
            End If
            Rec.Close
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
    Private Sub SetFormControls()
        Dim mFlag As Boolean
        Dim mTransactionTypeID As Long
        
        mFlag = True
        mTransactionTypeID = val(txtTransactionType.Tag)
        '----------------------------------------------------------------------'
        ' INITIALIZE FORM CONTROL
        '----------------------------------------------------------------------'
            'Note:- Make Visible
            'txtSubCashCode.Visible = mFlag
            txtAllotmentLetterNo.Visible = mFlag
            txtAgreementNo.Visible = mFlag
            txtProjectNo.Visible = mFlag
            txtSector.Visible = mFlag
            txtCategory.Visible = mFlag
            txtImplementingOfficer.Visible = mFlag
            'txtSubsidiaryCash.Visible = mFlag
            
            cmdProjectNo.Visible = mFlag
            cmdAllotmentLetterNo.Visible = mFlag
            cmdAgreementNo.Visible = mFlag
            
            cmdImplementingOfficer.Visible = mFlag
            'cmdSubsidiaryCash.Visible = mFlag
            cmdAsset.Visible = mFlag
            
            'Note:- Enable Contorls
            'txtSubCashCode.Visible = mFlag
            txtAllotmentLetterNo.Visible = mFlag
            txtAgreementNo.Visible = mFlag
            txtProjectNo.Visible = mFlag
            txtSector.Visible = mFlag
            txtCategory.Visible = mFlag
            txtImplementingOfficer.Visible = mFlag
            'txtSubsidiaryCash.Visible = mFlag
            
            cmdProjectNo.Visible = mFlag
            cmdAllotmentLetterNo.Visible = mFlag
            cmdAgreementNo.Visible = mFlag
            
            cmdImplementingOfficer.Visible = mFlag
            'cmdSubsidiaryCash.Visible = mFlag
            cmdAsset.Visible = mFlag
            
            lblNameOfPayee.Visible = mFlag
            lblType.Visible = mFlag
            lblProjectNo.Visible = mFlag
            lblSector.Visible = mFlag
            lblCategory.Visible = mFlag
            lblAllotmentLetterNo.Visible = mFlag
            lblAgreementNo.Visible = mFlag
            lblSourceOfFund.Visible = mFlag
            lblSubsidiaryCash.Visible = mFlag
            lblImplementingOfficer.Visible = mFlag


        
        lblSubsidiaryCash.Visible = False
        txtSubsidiaryCash.Visible = False
        txtSubsidiaryCash.Enabled = False
        cmdSubsidiaryCash.Visible = False
        cmdSubsidiaryCash.Enabled = False
            
        '----------------------------------------------------------------------'
        '
        '----------------------------------------------------------------------'
        
        Select Case mTransactionTypeID
            '---------------------------------'
            ' Pay Bill                        '
            '---------------------------------'
            Case Is = gbTransactionTypePayBills
                txtSubCashCode.Visible = True
                lblSubsidiaryCash.Visible = True
                
                lblSubsidiaryCash.Visible = True
                txtSubsidiaryCash.Visible = True
                txtSubsidiaryCash.Enabled = True
                cmdSubsidiaryCash.Visible = True
                
                lblAllotmentLetterNo.Visible = False
                txtAllotmentLetterNo.Visible = False
                cmdAllotmentLetterNo.Visible = False
                
                lblAgreementNo.Visible = False
                txtAgreementNo.Visible = False
                cmdAgreementNo.Visible = False
                
                lblProjectNo.Visible = False
                txtProjectNo.Visible = False
                cmdProjectNo.Visible = False
                
                lblSector.Visible = False
                txtSector.Visible = False
                
                lblCategory.Visible = False
                txtCategory.Visible = False
                
                lblImplementingOfficer.Visible = False
                txtImplementingOfficer.Visible = False
                cmdImplementingOfficer.Visible = False
                
            '---------------------------------'
            '
            '---------------------------------'
        End Select
    End Sub
    
    Private Sub SavePaymentOrder()
        Dim PO As uPaymentOrder
        Dim POC As uPaymentOrderChild
        Dim POAdd As uPaymentOrderAddress
        Dim ObjSubLed As New clsSubLedger
        
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mPaymentOrderID As Variant
        Dim mSlNo As Integer
        Dim mLoop As Integer
        Dim vchPayOrderNo As String
        
        Dim mSql As String
        
        cmdSave.Enabled = False
        objdb.SetConnection mCnn
        'mCnn.BeginTrans
        On Error GoTo ErrRollBack:
        
        If val(txtPayOrder.Tag) <> 0 Then
            mSql = "Select tnyStatus From faPayOrder Where intPayOrderID = " & val(txtPayOrder.Tag)
            Rec.Open mSql, mCnn
            If Not (Rec.EOF Or Rec.BOF) Then
                If gbLBPanchayat Then
                    If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                        If Rec!tnyStatus = 5 Then
                            MsgBox "Sorry! This Payment Order is already Forwarded. Editing is not Permitted", vbInformation
                            Exit Sub
                        ElseIf Rec!tnyStatus = 3 Then
                            MsgBox "Sorry! This Payment Order is Verified. Editing is not Permitted", vbInformation
                            Exit Sub
                        ElseIf Rec!tnyStatus = 1 Then
                            MsgBox "Sorry! This Payment Order is Approved. Editing is not Permitted", vbInformation
                            Exit Sub
                        End If
                    End If
                Else
                    If Rec!tnyStatus <> 0 Then
                        MsgBox "Sorry! This Payment Order is already approved. Editing is not Permitted", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            mSql = ""
            If Rec.State = 1 Then Rec.Close
        End If
        
        With PO
            .intPayOrderID = IIf(txtPayOrder.Tag = "", Null, txtPayOrder.Tag)
            .vchPayOrderNo = IIf(txtPayOrder.Text = "", Null, txtPayOrder.Text)
            If mPendingTask = 0 Then
                .dtPayOrderDate = gbTransactionDate
            Else
                .dtPayOrderDate = txtDated.Text
            End If
            .dtDueDate = txtDueDate.Text
            .intFunctionaryID = val(txtFunctionary.Tag)
            .intFunctionID = val(txtFunction.Tag)
            .intTransactionTypeID = val(txtTransactionType.Tag)
            .vchBillNo = Null
            .numBillAmount = Null
            .dtBillDate = Null
            .intInstrumentTypeID = Null
            .intCashOrBankHeadID = val(txtDrHeadCode.Tag)
            .vchDescription = Trim(txtNarration.Text)
            .vchTitle = Null
            .intSubLedgerTypeID = val(txtSubLedgerType.Tag)
            .intPayToSubLedgerID = val(txtName.Tag)
            .intSubsidiaryCashBookID = val(txtSubsidiaryCash.Tag)
            .intImplementingOfficerID = val(txtImplementingOfficer.Tag)
            .numProjectNo = val(txtProjectNo.Tag)
            .intStockRegisterID = Null
            .vchStockRefNo = Null
            .intAssetTypeID = mAssetID  'AssetID
            .intAssetID = mAssetTypeID  'AssetTypeID
            .numFwdSeatID = val(txtForward2Seat.Tag)
            .intLocalBodyID = gbLocalBodyID
            .intZonalID = gbLocationID
            If mPendingTask <> 0 Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .numUserID = gbUserID
            .numSeatID = gbSeatID
            .numApprovingOfficerID = Null
            .numApprovingSeatID = Null
            .dtApprovingDate = Null
            .intSourceOfFundID = val(txtSourceofFund.Tag)
            .intAllotmentID = val(txtAllotmentLetterNo.Tag)
            .intAgreementID = val(txtAgreementNo.Tag)
            .tnyCategoryID = val(txtCategory.Tag)
            .tnySectorID = val(txtSector.Tag)
            .tnyIsFinalBill = Null
            .intVoucherID = Null
            .intVoucherNo = Null
            .dtVoucherDate = Null
            If gbLBPanchayat Then
                If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
                .tnyStatus = 0
                Else
                    .tnyStatus = 5 'Forward
                End If
            Else
                .tnyStatus = 0
            End If
            If val(txtAllotmentLetterNo.Tag) > 0 Then
                If val(txtSourceofFund.Tag) <> 4 Then
                    .intKeyID = val(txtTreasuryID.Text)
                Else
                     If val(txtTransactionType.Tag) = 1141 Or val(txtTransactionType.Tag) = 1151 Or val(txtTransactionType.Tag) = 1161 Or _
                        val(txtTransactionType.Tag) = 1171 Or val(txtTransactionType.Tag) = 1181 Or val(txtTransactionType.Tag) = 1191 Then
                        .intKeyID = val(txtTreasuryID.Text)
                     Else
                        .intKeyID = intKeyID
                     End If
                End If
            Else
                .intKeyID = intKeyID 'Section ID stores from Pay Bill -Sthapana for Pay&Allowance
            End If
            .numKeyID = Null
            .dtKeyDate = IIf(IsNull(dtKeyDate), txtDueDate.Text, dtKeyDate)
            
            .tnyCancelled = 0
            .intAppID = 115
            If .intModuleID = "" Then
                .intModuleID = intModuleID
            Else
                .intModuleID = ModuleID
            End If
            
            arrInput = Array(.intPayOrderID, .vchPayOrderNo, .dtPayOrderDate, .dtDueDate, .intFunctionaryID, _
            .intFunctionID, .intTransactionTypeID, .vchBillNo, .numBillAmount, _
            .dtBillDate, .intInstrumentTypeID, .intCashOrBankHeadID, .vchDescription, _
            .vchTitle, .intSubLedgerTypeID, .intPayToSubLedgerID, .intSubsidiaryCashBookID, _
            .intImplementingOfficerID, .numProjectNo, .intStockRegisterID, .vchStockRefNo, _
            .intAssetTypeID, .intAssetID, .numFwdSeatID, .intLocalBodyID, _
            .intZonalID, .intFinancialYearID, .numUserID, .numSeatID, _
            .numApprovingOfficerID, .numApprovingSeatID, .dtApprovingDate, .intVoucherID, .intVoucherNo, .dtVoucherDate, _
            .tnyStatus, .intKeyID, .numKeyID, .dtKeyDate, .tnyCancelled, .intAppID, .intModuleID, .intSourceOfFundID, _
            .intAllotmentID, .intAgreementID, .tnyCategoryID, .tnySectorID, _
            .tnyIsFinalBill)
            
            objdb.ExecuteSP "spSavePayOrder", arrInput, arrOutPut, , mCnn, adCmdStoredProc
               
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mPaymentOrderID = arrOutPut(0, 0)
            vchPayOrderNo = arrOutPut(1, 0)
            PayOrderID = mPaymentOrderID
            PayOrderNo = vchPayOrderNo
        Else
            GoTo ErrRollBack:
        End If
        
        
        mSql = "Delete From faPayOrderChild Where intPayOrderID = " & mPaymentOrderID
        mCnn.Execute mSql
        
        mSlNo = mSlNo + 1
        With POC
            .intPayOrderID = mPaymentOrderID
            .intSlNo = mSlNo
            .intAccountHeadID = val(txtDrHeadCode.Tag)
            .vchAccountHeadCode = Trim(txtDrHeadCode.Text)
            .numAmount = val(txtDrAmount.Text)
            .tnyCategoryFlag = 1
            .tnyDebitOrCreditFlag = 1
            .vchDescription = Null
            
            
            arrInput = Array(.intPayOrderID, _
            .intSlNo, _
            .intAccountHeadID, _
            .vchAccountHeadCode, _
            .numAmount, _
            .tnyCategoryFlag, _
            .tnyDebitOrCreditFlag, _
            .vchDescription)
                        
            objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        mSlNo = mSlNo + 1
        For mLoop = 1 To vsGrid.Rows - 1
            If val(vsGrid.TextMatrix(mLoop, 1)) > 0 And val(vsGrid.TextMatrix(mLoop, 3)) > 0 Then
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = val(vsGrid.TextMatrix(mLoop, 4))
                .vchAccountHeadCode = Trim(vsGrid.TextMatrix(mLoop, 1))
                .numAmount = val(vsGrid.TextMatrix(mLoop, 3))
                .tnyCategoryFlag = 2
                .tnyDebitOrCreditFlag = 0
                .vchDescription = Null
                .tnyExcldeFromSourceFlag = val(vsGrid.TextMatrix(mLoop, 6))
                
                arrInput = Array(.intPayOrderID, _
                .intSlNo, _
                .intAccountHeadID, _
                .vchAccountHeadCode, _
                .numAmount, _
                .tnyCategoryFlag, _
                .tnyDebitOrCreditFlag, _
                .vchDescription, _
                .tnyExcldeFromSourceFlag)
                objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
            End If
        Next
        
'        If chkPension.value = vbChecked Then
'
'        End If
        
        mSlNo = mSlNo + 1
        With POC
            .intPayOrderID = mPaymentOrderID
            .intSlNo = mSlNo
            .intAccountHeadID = val(txtCrHeadCode.Tag)
            .vchAccountHeadCode = Trim(txtCrHeadCode.Text)
            .numAmount = val(txtCrAmount)
            .tnyCategoryFlag = 3
            .tnyDebitOrCreditFlag = 0
            .vchDescription = Null
            
            arrInput = Array(.intPayOrderID, _
            .intSlNo, _
            .intAccountHeadID, _
            .vchAccountHeadCode, _
            .numAmount, _
            .tnyCategoryFlag, _
            .tnyDebitOrCreditFlag, _
            .vchDescription)
            objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        If val(txtSubsidiaryCash.Tag) <> 0 Then
            mSlNo = mSlNo + 1
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = val(txtSubCashCode.Tag)
                .vchAccountHeadCode = Trim(txtSubCashCode.Text)
                .numAmount = val(txtCrAmount)
                .tnyCategoryFlag = 4
                .tnyDebitOrCreditFlag = 0
                .vchDescription = Null
                
                arrInput = Array(.intPayOrderID, _
                .intSlNo, _
                .intAccountHeadID, _
                .vchAccountHeadCode, _
                .numAmount, _
                .tnyCategoryFlag, _
                .tnyDebitOrCreditFlag, _
                .vchDescription)
                objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
        End If
        
        '--------Added On 4-12-12
        
'        If chkPension.value = vbChecked Then
        If chkPensionContribution.Value = vbChecked Then
        Else
         If val(txtPensionAmt.Text) > 0 Then
            mSlNo = mSlNo + 1
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = 0
                .vchAccountHeadCode = 0
                .numAmount = val(txtPensionAmt)
                .tnyCategoryFlag = 5
                .tnyDebitOrCreditFlag = 0
                .vchDescription = "Pension Contribution Amount"
                
                arrInput = Array(.intPayOrderID, _
                .intSlNo, _
                .intAccountHeadID, _
                .vchAccountHeadCode, _
                .numAmount, _
                .tnyCategoryFlag, _
                .tnyDebitOrCreditFlag, _
                .vchDescription)
                objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
        End If
        End If
        'Added on 20 Nov 2017
        If val(txtCP.Text) > 0 Then
            mSlNo = mSlNo + 1
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = 0
                .vchAccountHeadCode = 0
                .numAmount = val(txtCP)
                .tnyCategoryFlag = 6
                .tnyDebitOrCreditFlag = 0
                .vchDescription = "Contributory Pension Amount"

                arrInput = Array(.intPayOrderID, _
                .intSlNo, _
                .intAccountHeadID, _
                .vchAccountHeadCode, _
                .numAmount, _
                .tnyCategoryFlag, _
                .tnyDebitOrCreditFlag, _
                .vchDescription)
                objdb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
        End If
        '------------------------
        
        With POAdd
            
            ObjSubLed.SetSubLedgerDetails (val(txtName.Tag))
            .intPayOrderID = mPaymentOrderID
            .intSubsidiaryAccountHeadID = IIf(IsNull(ObjSubLed.SubsidiaryAccountHeadID), Null, ObjSubLed.SubsidiaryAccountHeadID)
            .intSubLegerTypeID = IIf(IsNull(ObjSubLed.SubLedgerTypeID), Null, ObjSubLed.SubLedgerTypeID)
            .vchSubLedgerCode = IIf(IsNull(ObjSubLed.SubLedgerCode), Null, ObjSubLed.SubLedgerCode)
            .vchName = Trim(txtName.Text)
            .vchHouseName = Trim(txtHouse)
            .vchStreet = Trim(txtStreet)
            .vchLocalPlace = Trim(txtLocalPlace)
            .vchMainPlace = Trim(txtMainPlace)
            .vchPost = Trim(txtPost)
            .vchPinCode = Trim(txtPin)
            .vchPhone = Trim(txtPhone)
            
            arrInput = Array(.intPayOrderID, _
            .intSubsidiaryAccountHeadID, _
            .intSubLegerTypeID, _
            .vchSubLedgerCode, _
            .vchName, _
            .vchHouseName, _
            .vchStreet, _
            .vchLocalPlace, _
            .vchMainPlace, _
            .vchPost, _
            .vchPinCode, _
            .vchPhone)
            objdb.ExecuteSP "spSavePayOrderAddress", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        If mPendingTask = 2 Then
            mSql = "Update faPendingTaskRequest set tnyStatus=8, numDemandID= " & mPaymentOrderID & " Where intRequestID=" & mPendingTaskReqID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            
            mSql = "  Update faPayOrder Set intModuleID=96 Where intpayOrderId=" & mPaymentOrderID
            objdb.ExecuteSP mSql, , , , mCnn, adCmdText
        End If

        If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Or val(txtTransactionType.Tag) = gbTransactionTypeProjectExpGO Then
            If val(txtGo.Tag) > 0 Then
                mSql = "Update suGOForFunds set intPayOrderID= " & mPaymentOrderID & " Where intRefID=" & val(txtGo.Tag)
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
            End If
        End If
        'mCnn.CommitTrans
        txtPayOrder.Text = vchPayOrderNo
        If gbLBPanchayat Then
        If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
            Call SaveActivityLog
        End If
        End If
        WaterBillPOMode = True
        Exit Sub
        
ErrRollBack:
            MsgBox "Error :( !", vbInformation
            If mCnn.State Then
                'mCnn.RollbackTrans
            End If
            cmdSave.Enabled = True
    End Sub
    Private Sub SaveActivityLog()
        Dim objdb As New clsDB
        Dim mCn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        
        objdb.SetConnection mCn
        If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
         arrInput = Array(Null, _
                21, _
                txtPayOrder.Text, _
                gbTransactionDate, _
                gbSeatID, _
                gbUserID, _
                "")
                objdb.ExecuteSP "spSaveActivityLog", arrInput, , , mCn, adCmdStoredProc
        End If
    End Sub
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
    
'''    Private Sub TransactionTemplate(mTransactionTypeID As Long)
'''        Dim objAc As New clsAccounts
'''        Dim mGrossSalaryID As Long
'''        Dim objDb As New clsDb
'''        Dim Rec As New ADODB.Recordset
'''        Dim mCnn As New ADODB.Connection
'''
'''        txtDrHeadCode.Visible = True
'''        txtDrAccountHead.Visible = True
'''        txtDrAmount.Visible = True
'''        cmdDrAccountHead.Visible = True
'''        lblDrAccountHead.Visible = True
'''        cmdCrAccountHead.Enabled = True
'''        cmdDrAccountHead.Enabled = True
'''        vsGrid.Enabled = True
'''
'''        Select Case mTransactionTypeID
'''
'''            Case Is = gbTransactionTypePayBills
'''                objAc.SetAccountCode ("350110200")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                    cmdCrAccountHead.Enabled = False
'''                End If
'''                objAc.SetAccountCode ("350110100")
'''                If objAc.AccountHeadID > 0 Then
'''                    mGrossSalaryID = objAc.AccountHeadID
'''                End If
'''            Case Is = 1002 Or 1003
'''                objDb.SetConnection mCnn
'''                Rec.Open "Select faAccountHeads.intAccountHeadID,fatransactionTypeChild.vchaccountHeadCode, faaccountheads.vchAccountHead From faTransactionTypeChild Inner join faAccountheads On faTransactionTypeChild.vchAccountHeadCode=faAccountHeads.vchAccountHeadcode Where tnyNetPayFlag=1 and faTransactionTypeChild.tnyListID = " & Val(vsGrid.TextMatrix(vsGrid.Row, 5)), mCnn
'''                If Not (Rec.EOF And Rec.BOF) Then
'''                    txtCrHeadCode.Text = Rec!vchAccountHeadCode
'''                    txtCrAccountHead.Text = Rec!vchAccountHead
'''                    txtCrHeadCode.Tag = Rec!intAccountHeadID
'''                End If
'''            Case Is = 1004                        '       Contigent Bills  '
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Is = 1005                              '       Contigent Bills  - Reg Emp '
'''                Call UnViewDebit
'''                LockGrid
'''
'''                objAc.SetAccountCode ("350110600")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                    cmdCrAccountHead.Enabled = False
'''                End If
'''            Case Is = 1006                              '       Contigent Bills  - Secratery '
'''                txtDrHeadCode.Visible = False
'''                txtDrAccountHead.Visible = False
'''                txtDrAmount.Visible = False
'''                cmdDrAccountHead.Visible = False
'''                lblDrAccountHead.Visible = False
'''                txtDrHeadCode.TabStop = True
'''                txtDrAccountHead.TabStop = True
'''                txtDrAmount.TabStop = True
'''                cmdDrAccountHead.TabStop = True
'''                vsGrid.Enabled = False
'''
'''                objAc.SetAccountCode ("350110700")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                    cmdCrAccountHead.Enabled = False
'''                End If
'''
'''            Case Is = 1007              '       Arrear Pay Bill                 '
'''                objAc.SetAccountCode ("350110200")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                    cmdCrAccountHead.Enabled = False
'''                End If
'''                objAc.SetAccountCode ("350110100")
'''                If objAc.AccountHeadID > 0 Then
'''                    mGrossSalaryID = objAc.AccountHeadID
'''                End If
'''            Case Is = 1008              '       Surrender Leave Salary      '
'''                vsGrid.TabStop = True
'''                vsGrid.Enabled = False
'''
'''                objAc.SetAccountCode ("210400100")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''
'''                objAc.SetAccountCode ("350110800")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                End If
'''            Case Is = 1009               '       TA Bill            '
'''                vsGrid.TabStop = True
'''                vsGrid.Enabled = False
'''                objAc.SetAccountCode ("210200100")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''                cmdDrAccountHead.Enabled = False
'''            Case Is = 1010              '       Payment of pension to Regular Employees '
'''                vsGrid.TabStop = True
'''                vsGrid.Enabled = False
'''                objAc.SetAccountCode ("350110500")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''                cmdDrAccountHead.Enabled = False
'''            Case Is = 1011              '       Payment of pension to Contigent Employees '
'''                vsGrid.TabStop = True
'''                vsGrid.Enabled = False
'''
'''                objAc.SetAccountCode ("210400100")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''
'''                objAc.SetAccountCode ("350110800")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                End If
'''                cmdCrAccountHead.Enabled = False
'''                cmdDrAccountHead.Enabled = False
'''            Case Is = 1012
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Is = 1013
'''                objAc.SetAccountCode ("220510100")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''
'''                objAc.SetAccountCode ("350109900")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtCrHeadCode.Text = objAc.AccountCode
'''                    txtCrHeadCode.Tag = objAc.AccountHeadID
'''                    txtCrAccountHead.Text = objAc.AccountHead
'''                End If
'''                cmdCrAccountHead.Enabled = False
'''                cmdDrAccountHead.Enabled = False
'''            Case Is = 1014
'''                objAc.SetAccountCode ("230500100")
'''                If objAc.AccountHeadID > 0 Then
'''                    txtDrHeadCode.Text = objAc.AccountCode
'''                    txtDrHeadCode.Tag = objAc.AccountHeadID
'''                    txtDrAccountHead.Text = objAc.AccountHead
'''                End If
'''                Call UnViewDebit
'''            Case Is = 1015
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Is = 1016
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Is = 1017
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Is = 1018
'''                Call UnViewDebit
'''                Call LockGrid
'''            Case Else
'''        End Select
'''    End Sub
    
'''    Private Sub LockDebit()
'''        txtDrHeadCode.Enabled = False
'''        txtDrAccountHead.Enabled = False
'''        cmdDrAccountHead.Enabled = False
'''    End Sub
    
'''    Private Sub UnViewDebit()
'''        txtDrHeadCode.Visible = False
'''        txtDrAccountHead.Visible = False
'''        txtDrAmount.Visible = False
'''        cmdDrAccountHead.Visible = False
'''        lblDrAccountHead.Visible = False
'''        txtDrHeadCode.TabStop = False
'''        txtDrAccountHead.TabStop = False
'''        txtDrAmount.TabStop = False
'''        cmdDrAccountHead.TabStop = False
'''    End Sub
    
'''    Private Sub LockGrid()
'''        vsGrid.Enabled = False
'''    End Sub
       
'''    Private Sub ViewDebit()
'''        txtDrHeadCode.Visible = True
'''        txtDrAccountHead.Visible = True
'''        txtDrAmount.Visible = True
'''        cmdDrAccountHead.Visible = True
'''        lblDrAccountHead.Visible = True
'''        txtDrHeadCode.TabStop = True
'''        txtDrAccountHead.TabStop = True
'''        txtDrAmount.TabStop = True
'''        cmdDrAccountHead.TabStop = True
'''    End Sub
       
    Private Sub InitializeSelectedHeads()
        txtDrHeadCode.Text = ""
        txtDrHeadCode.Tag = ""
        txtDrAccountHead.Text = ""
        txtDrAmount.Text = ""
        txtCrHeadCode.Text = ""
        txtCrHeadCode.Tag = ""
        txtCrAccountHead.Text = ""
        txtCrAmount.Text = ""
        vsGrid.Clear 1, 1
        cmdCrAccountHead.Enabled = True
    End Sub
    
    Private Sub FormInitialize()
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
                ctrl.Tag = ""
                'Debug.Print ctrl.Name
            ElseIf TypeOf ctrl Is OptionButton Then
                ctrl.Value = False
            ElseIf TypeOf ctrl Is ComboBox Then
                If ctrl.ListCount > 0 Then ctrl.ListIndex = -1
                ctrl.Tag = ""
            'ElseIf TypeOf ctrl Is Label Then
                'Debug.Print ctrl.Name
            End If
        Next
        
       
        Check1.Value = 0
        'Note:- User Type wise Functionality should enable or Disabled
        cmdApproval.Visible = False
'        cmdReject.Visible = False
        cmdSave.Caption = "Save"
        cmdSave.Enabled = True
        dtpDueDate.Enabled = True
        
        vsGrid.Clear 1, 1
        txtDated.Text = DdMmmYy(gbTransactionDate)
        txtDueDate.Text = DdMmmYy(gbTransactionDate)
        cmdCrAccountHead.Enabled = True
        mSelect = False
        mBudgetBalanceAmt = 0
        cmdVerify.Visible = False
        intKeyID = Null
        intModuleID = Null
        ModuleID = 0
        dtKeyDate = Null
        mSkipMsgFlag = False
        mSelectCreditHeadFlag = False
        
        Call SetFormControls
        Call LockProjectType(True)
        cmdAllotmentLetterNo.Tag = 0
        mViewPayOrderListFormIsLoaded = True
        txtDrAmount.Enabled = True
        txtGo.Visible = False
        cmdGo.Visible = False
        lblGo.Visible = False
        If mPendingTask = 1 Or mPendingTask = 2 Then
            Call GetPendingTaskDetails
        End If
        ''Added on 1/Dec/2017
        lblCP.Visible = False
        txtCP.Visible = False
        chkPensionContribution.Value = Unchecked
        chkPensionContribution.Visible = False
        
        If gbLBPanchayat = 1 Then
            If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                cmdSave.Caption = "Save/Fwd"
            End If
        End If
    End Sub
    
    Private Function CalculateAmt() As Variant
        Dim mCount As Integer
        Dim mTOt As Variant
        mTOt = 0
        
        For mCount = 1 To vsGrid.Rows
            If Trim(vsGrid.TextMatrix(mCount, 1)) = "" And Trim(vsGrid.TextMatrix(mCount, 3)) = "" Then
                Exit For
            End If
            
            If vsGrid.TextMatrix(mCount, 1) <> "" Then
                mTOt = val(mTOt) + val(vsGrid.TextMatrix(mCount, 3))
                If val(vsGrid.TextMatrix(mCount, 3)) = 0 Then
                    vsGrid.RemoveItem (mCount)
                    mCount = mCount - 1
                End If
            End If
        Next
        CalculateAmt = mTOt
    
    End Function
    
'''    Private Sub FillGridCombo()
'''        Dim objDb As New clsDb
'''        Dim RecAccHead As New ADODB.Recordset
'''        Dim mItem As String
'''
'''        RecAccHead.CursorLocation = adUseClient
'''        Set RecAccHead = GetRecordSet("spGetAccHead4Payments", adOpenStatic, adLockReadOnly)
'''        While Not RecAccHead.EOF
'''            mItem = mItem + "|" + RecAccHead!vchAccountHead
'''            RecAccHead.MoveNext
'''        Wend
'''        RecAccHead.Close
'''        vsGrid.ColComboList(2) = mItem
'''    End Sub

    Private Sub cmbTransactionType_Click()
'''        If cmbTransactionType.ListIndex > -1 Then
'''            Call InitializeSelectedHeads
'''            txtTransactionType.Tag = cmbTransactionType.ItemData(cmbTransactionType.ListIndex)
'''            Call TransactionTemplate(cmbTransactionType.Tag)
'''
'''            ' NOTE:- WARRNING :: Transaction Type is Heard Coded :: Added By Aiby on 11-Oct-2009
'''            '        Project Expediture Type Transactions to Link Subsystem - Sulekha Through
'''            '        Allotment Letter and Agreement Register
'''
'''            If Val(cmbTransactionType.Tag) > 1140 And Val(cmbTransactionType.Tag) < 1192 Then
'''                fraProject.Visible = True
'''            Else
'''                fraProject.Visible = False
'''            End If
'''
'''        Else
'''            Call InitializeSelectedHeads
'''        End If
    End Sub





Private Sub chkPensionContribution_Click()
        If chkPensionContribution.Value = vbChecked Then
            If MsgBox(" Are you Sure to Avoid Pension contribution?", vbYesNo, "Saankhya") = vbYes Then
                txtPensionAmt.Enabled = True
                Exit Sub
            Else
                txtPensionAmt.Enabled = False
                chkPensionContribution.Value = vbUnchecked
            End If

        End If
End Sub

'    Private Sub chkPension_Click()
'        If chkPension.value = vbChecked Then
'            lblPension.Visible = True
'            txtPensionAmt.Visible = True
'        Else
'            lblPension.Visible = False
'            txtPensionAmt.Visible = False
'        End If
'    End Sub

   Private Sub cmdAgreementNo_Click()
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    
        frmSearchAgreements.Show vbModal
        If gbSearchID <> -1 Then
            txtAgreementNo.Text = gbSearchStr
            txtAgreementNo.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
        If txtAgreementNo.Tag <> "" Then
            mSql = "select intAssetID,intAssetHeadID from faAgreements where intAgreementID=" & txtAgreementNo.Tag & " "
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            If Not (Rec.BOF And Rec.EOF) Then
                AssetID = IIf(IsNull(Rec!intAssetID), "", Rec!intAssetID)
                AssetTypeID = IIf(IsNull(Rec!intAssetHeadID), "", Rec!intAssetHeadID)
            End If
        End If
    End Sub
    
    Private Function GetStatusFlag() As Integer
        Dim mCnn  As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec   As New ADODB.Recordset
        Dim mSql  As String
        Dim mTrAccHeadId As Integer
        
        If objdb.SetConnection(mCnn) Then
            mSql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID
            Rec.Open mSql, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                GetStatusFlag = Rec!tnyStatus
            Else
                
                'NOTE: Checking in Previous Year
                '      IF APPROVED tnyStatus will be 0 ELSE NULL
                Rec.Close
                mSql = "SELECT tnyStatus FROM faExtractAllotments WHERE intFinancialYearID = " & gbFinancialYearID - 1
                Rec.Open mSql, mCnn
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
    
    Private Sub cmdAllotmentLetterNo_Click()
        Dim mExtractedStatus As Integer
        Dim mMsg As String
         '' Added on 01 June/2017 For Saankhyaweb implementation
        If gbFinancialYearID > 2016 And gbSaankhyaWeb = 1 Then
            Exit Sub
        Else
            
             
            mExtractedStatus = GetStatusFlag
            If mExtractedStatus <> 2 Then
                'cmdNew.Enabled = False
                'cmdSave.Enabled = False
                'cmdApprove.Enabled = False
                mMsg = ""
                mMsg = mMsg + " Closing Balance Of Source Of Fund is " + vbCrLf
                mMsg = mMsg + " Either Not Brought Down  Or Approved " + vbCrLf
                mMsg = mMsg + " (Utility>>Annual Financial Statements-Finalization>>)"
                MsgBox mMsg, vbInformation
                Exit Sub
            End If
        
        
        
            'frmListOfAllotmentLetters.Mode = 0
            frmListOfAllotmentLetters.Show vbModal
            If gbSearchID <> -1 Then
            
                vsGrid.Clear 1, 1
                txtCrHeadCode.Text = ""
                txtCrHeadCode.Tag = ""
                txtCrAccountHead.Text = ""
                txtCrAccountHead.Tag = ""
                
                txtTransactionType.Text = ""
                txtTransactionType.Tag = ""
                        
                txtAllotmentLetterNo.Text = gbSearchCode
                txtAllotmentLetterNo.Tag = gbSearchID
                
                Dim objAllotment As New clsAllotmentLetter
                objAllotment.SetAllotment (txtAllotmentLetterNo.Tag)
                txtSourceofFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
                txtSourceofFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), "", objAllotment.SourceOfFundID)
                
                txtCategory.Text = IIf(IsNull(objAllotment.Category), "", objAllotment.Category)
                txtCategory.Tag = IIf(IsNull(objAllotment.CategoryID), "", objAllotment.CategoryID)
                txtTreasuryID.Text = IIf(IsNull(objAllotment.mNewModeID), "", objAllotment.mNewModeID)
                txtImplementingOfficer.Text = IIf(IsNull(objAllotment.ImplementingOfficer), "", objAllotment.ImplementingOfficer)
                txtImplementingOfficer.Tag = IIf(IsNull(objAllotment.ImplementingOfficersID), "", objAllotment.ImplementingOfficersID)
                txtAllotedAmt.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)
                cmdAllotmentLetterNo.Tag = IIf(IsNull(objAllotment.TypeID), 0, objAllotment.TypeID) 'FOR UNAUTHORIZED DRAWAL
                Dim objProject As New clsProject
                objProject.SetProject (val(gbSearchStr))
                If Not IsNull(objProject.ProjectID) Then
                    txtProjectNo.Text = IIf(IsNull(objProject.ProjectSerialNo), "", objProject.ProjectSerialNo)
                    txtProjectNo.Tag = IIf(IsNull(objProject.ProjectID), "", objProject.ProjectID)
                    'txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
                    'txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
                    txtSector = IIf(IsNull(objProject.Sector), "", objProject.Sector)
                    txtSector.Tag = IIf(IsNull(objProject.SectorTypeID), "", objProject.SectorTypeID)
                End If
                
                gbSearchID = -1
                gbSearchStr = ""
                gbSearchCode = ""
            End If
            txtAllotmentLetterNo.Enabled = True
            txtAllotmentLetterNo.SetFocus
        End If
    End Sub
Private Sub cmdApproval_Click()
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    mSql = "Select tnyStatus From faPayOrder Where vchPayOrderNo = " & val(txtPayOrder) '& " And tnyStatus = 0"
    objdb.SetConnection mCnn
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (Rec.BOF And Rec.EOF) Then
        
        If Rec!tnyStatus = 1 Then
            MsgBox "This Pay Order is already approved", vbInformation
            Exit Sub
        End If
        If gbLBPanchayat Then
            If Rec!tnyStatus = 5 Then
                MsgBox "This Pay Order is not Verified ", vbInformation
                Exit Sub
            End If
        End If
        cmdApproval.Enabled = False
        'If MakePayable(Val(txtPayOrder)) Then
            MakePayable (val(txtPayOrder))
            'If Rec.State = 1 Then Rec.Close
            'mSQL = "Update faPayOrder Set tnyStatus = 1 Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
            mSql = "Update faPayOrder Set tnyStatus = 1, numApprovingOfficerID = " & gbUserID & ", numApprovingSeatID = " & gbSeatID & ", dtApprovingDate = '" & DdMmmYy(gbTransactionDate) & "' Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
            mCnn.Execute mSql
            If mPendingTask = 1 Then
                mSql = "Update faPendingTaskRequest Set tnyStatus=8 Where intRequestID=" & mPendingTaskReqID
                mCnn.Execute mSql
            End If
'            '-----------------------
'            '---Print Payment Order
'                Dim aryIn As Variant
'                aryIn = Array(val(txtPayOrder))
'                frmViewVoucher.ArrayIn = aryIn
'                frmViewVoucher.FormName = "PrintPaymentOrder"
'                frmViewVoucher.Show vbModal
'            '------------------------
            Call FormInitialize
        'End If
    End If
End Sub

    Private Sub cmdCancel_Click()
        If mViewPayOrderListFormIsLoaded Then
            Unload Me
        Else
            Call FormInitialize
        End If
    End Sub

    Private Sub cmdCrAccountHead_Click()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        ''Select Case Val(cmbTransactionType.Tag)
        ''    Case Is = 1006
        ''        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where vchAccountHeadCode = '350110800'"
        ''    Case Else
        ''        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (3)"
        ''End Select
        If objdb.SetConnection(mCnn) = True Then
            mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
            mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
            mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & " AND tnyListID = 3 And tinHiddenFlag = 0 And faAccountHeads.intGroupID is Null"
            mSql = mSql + " Order By faTransactionTypeChild.intOrder"
            Rec.Open mSql, mCnn
            If Rec.BOF Or Rec.EOF Then
                mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinHiddenFlag = 0 And intGroupID is Null" ' Where tinType IN (3)"
            End If
            
'            If Val(txtTransactionType.Tag) = 1015 Then  '***** Contigent Bills For Project Expences ****'
'                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (2)"
'            End If
'
'            If Val(txtTransactionType.Tag) = 1016 Then  '***** Contigent Bills For Payment of Advance ****'
'                mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (4)"
'            End If
            
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.VoucherMode = 400  'Exluding  Cash And Bank Ac Heads in List All
            frmSearchAccountHeads.Show vbModal
            If gbSearchStr <> "" Then
                txtCrHeadCode.Text = Token(gbSearchStr, " ")
                txtCrAccountHead.Text = Trim(gbSearchStr)
                objAcc.SetAccountCode (txtCrHeadCode.Text)
                txtCrHeadCode.Tag = objAcc.AccountHeadID
                
                gbSearchID = -1
                gbSearchStr = ""
                txtCrAmount.SetFocus
            End If
        End If
    End Sub
    
    Private Sub cmdDeductionPayment_Click()
        frmDeductionExcludePayment.Show vbModal
    End Sub
    Private Sub RemittanceOfUnUtilizedAmount()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        If gbLBPanchayat = 1 Then
            If val(txtSourceofFund.Tag) = 1 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1065" & vbNewLine
                'mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 16 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1667" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 17 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1668" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 25 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1068" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 26 Or val(txtSourceofFund.Tag) = 41 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1612" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 27 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1617" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 28 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1618" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 29 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1066" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 30 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1067" & vbNewLine
            Else
               MsgBox "Wrong Source Of Fund Selected", vbApplicationModal
               Exit Sub
            End If
            
        Else
            If val(txtSourceofFund.Tag) = 1 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=928" & vbNewLine
                'mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 16 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=2347" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 17 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=2348" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 25 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=931" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 26 Or val(txtSourceofFund.Tag) = 41 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1762" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 27 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1763" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 28 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1764" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 29 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=929" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 30 Then
                mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=930" & vbNewLine
            Else
                MsgBox "Wrong Source Of Fund Selected", vbApplicationModal
                Exit Sub
            End If
        End If
        If objdb.SetConnection(mCnn) = True Then
            Rec.Open mSql, mCnn
            frmSearchAccountHeads.SQLString = mSql
            frmSearchAccountHeads.VoucherMode = 400
            frmSearchAccountHeads.chkListAll.Enabled = False
            frmSearchAccountHeads.Show vbModal
            
            If gbSearchStr <> "" Then
                txtDrHeadCode.Text = Token(gbSearchStr, " ")
                txtDrAccountHead.Text = Trim(gbSearchStr)
                objAcc.SetAccountCode (txtDrHeadCode.Text)
                txtDrHeadCode.Tag = objAcc.AccountHeadID
                Call txtDrHeadCode_LostFocus
                gbSearchID = -1
                gbSearchStr = ""
                txtDrAmount.SetFocus
            End If
        End If
    End Sub

        Private Sub cmdDrAccountHead_Click()
            Dim mSql As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim objdb As New clsDB
            Dim objAcc As New clsAccounts
            
            ''' SaankhyaWeb Updation
            If val(txtTransactionType.Tag) < 0 Then
                MsgBox "Please Select Transaction Type", vbApplicationModal
                Exit Sub
            End If
            If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
                If val(txtSourceofFund.Tag) < 1 Then
                    MsgBox "Please Select Source of Fund", vbApplicationModal
                    Exit Sub
                Else
                    Call RemittanceOfUnUtilizedAmount
                End If
            ElseIf val(txtTransactionType.Tag) = gbTransactionTypeProjectExpGO Then
                If val(txtSourceofFund.Tag) < 1 Then
                    If gbLBPanchayat Then
                        mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                        mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                        mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                        mSql = mSql + " And faAccountHeads.intAccountHeadID in (1065,1066,1067,1068,1612,1617,1618,1667,1668 )" & vbNewLine
                    Else
                        mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                        mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                        mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                        mSql = mSql + " And faAccountHeads.intAccountHeadID in (928,929,930,931,1762,1763,1764,2347,2348)" & vbNewLine
                    End If
                    If objdb.SetConnection(mCnn) = True Then
                        Rec.Open mSql, mCnn
                        frmSearchAccountHeads.SQLString = mSql
                        frmSearchAccountHeads.VoucherMode = 400
                        frmSearchAccountHeads.chkListAll.Enabled = False
                        frmSearchAccountHeads.Show vbModal
                        
                        If gbSearchStr <> "" Then
                            txtDrHeadCode.Text = Token(gbSearchStr, " ")
                            txtDrAccountHead.Text = Trim(gbSearchStr)
                            objAcc.SetAccountCode (txtDrHeadCode.Text)
                            txtDrHeadCode.Tag = objAcc.AccountHeadID
                            Call txtDrHeadCode_LostFocus
                            gbSearchID = -1
                            gbSearchStr = ""
                            txtDrAmount.Enabled = True
                            txtDrAmount.SetFocus
                        End If
                    End If
                        
                Else
                    Call RemittanceOfUnUtilizedAmount
                End If
            Else
                If objdb.SetConnection(mCnn) = True Then
                    mSql = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                    mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                    mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag)
                    mSql = mSql + " And faTransactionTypeChild.tinDebitOrCredit = 1 And faAccountHeads.tinHiddenFlag = 0"
                    mSql = mSql + " And faTransactionTypeChild.tnyListID = 1 And faAccountHeads.intGroupID is Null"
                    mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode"
                    'mSQL = mSQL + " Order By faTransactionTypeChild.intOrder"
                    Rec.Open mSql, mCnn
                    If Rec.BOF Or Rec.EOF Then
                        mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where intGroupID is Null And tinHiddenFlag = 0" ' Where tinType IN (3)"
                    End If
                    
                    frmSearchAccountHeads.SQLString = mSql
                    frmSearchAccountHeads.VoucherMode = 400
                    frmSearchAccountHeads.Show vbModal
                    If gbSearchStr <> "" Then
                        txtDrHeadCode.Text = Token(gbSearchStr, " ")
                        txtDrAccountHead.Text = Trim(gbSearchStr)
                        objAcc.SetAccountCode (txtDrHeadCode.Text)
                        txtDrHeadCode.Tag = objAcc.AccountHeadID
                        Call txtDrHeadCode_LostFocus
                        gbSearchID = -1
                        gbSearchStr = ""
                        txtDrAmount.Enabled = True
                        txtDrAmount.SetFocus
                    End If
                    If Rec.State = 1 Then Rec.Close
                    
                    '============================================================='
                    '   Added For Getting the Head Code for Subsidiary Cash Book  '
                    '               Modified By Cijith Sreedharan                 '
                    '============================================================='
                        txtSubCashCode.Text = txtDrHeadCode.Text
                        txtSubCashCode.Tag = txtDrHeadCode.Tag
                        If val(txtCrHeadCode.Tag) < 1 And mSelectCreditHeadFlag = False Then
                            txtCrHeadCode.Text = txtDrHeadCode.Text
                            txtCrHeadCode.Tag = txtDrHeadCode.Tag
                            txtCrAccountHead.Text = txtDrAccountHead.Text
                        End If
                    '============================================================='
                Else
                
                End If
            End If
        End Sub
    
    Private Sub cmdGO_Click()
        frmGoDetails.SuFundTr = val(txtTransactionType.Tag)
        If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
            If val(txtSourceofFund.Tag) > 0 Then
                frmGoDetails.SuFund = val(txtSourceofFund.Tag)
            Else
                MsgBox "Please Select Source of Fund"
                Exit Sub
            End If
        End If
        frmGoDetails.Show vbModal
        If gbSearchID <> -1 Then
            txtGo.Text = gbSearchStr
            txtGo.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    
'''        Private Sub cmdHeadSearch_Click()
'''            Dim mSql As String
'''            Dim objAcc As New clsAccounts
'''
'''            mSql = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads WHERE faAccountHeads.intGroupID = " & faCash
'''            frmSearchAccountHeads.SQLString = mSql
'''            frmSearchAccountHeads.Show vbModal
'''
'''
'''        End Sub

    Private Sub cmdImplementingOfficer_Click()
        frmSearchMasters.SQLQry = "Select intFunctionaryID, vchFunctionary +'[' + vchFunctionaryCode + ']' From faFunctionaries Where intFunctionaryID > 13"
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        If gbSearchID <> -1 Then
            txtImplementingOfficer.Text = gbSearchStr
            txtImplementingOfficer.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    
    Private Sub cmdNew_Click()
        If mViewPayOrderListFormIsLoaded Then
            If MsgBox("Do you want to Enter Payment Order?", vbYesNo + vbDefaultButton2) = vbYes Then
                Call FormInitialize
            End If
        Else
            Call FormInitialize
        End If
    End Sub

    Private Sub cmdProjectNo_Click()
        'frmEstimationDetails.Mode = 0
        'frmSulekhaIntegration.Show vbModal
        frmSearchProjects.Show vbModal
        'If gbSearchID <> -1 Then
            'txtProjectNo.Text = gbSearchStr
            'txtProjectNo.Tag = gbSearchID
            Call txtProjectNo_GotFocus
            gbSearchID = -1
            gbSearchStr = ""
        'End If
    End Sub
    
    Private Sub cmdReject_Click()
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim mStatus As Integer
        Dim mFwdSeat As Double
        Dim mFwdSeatGpID As Integer
        Dim mPOstatus As Integer
        If gbLBPanchayat Then
            mSql = " Select numFwdSeatID,intGroupID,faPayOrder.tnyStatus postatus,* From faPayOrder"
            mSql = mSql + " Inner Join faSeats On faSeats.numSeatID=faPayOrder.numFwdSeatID"
            mSql = mSql + " Where vchPayOrderNo = " & val(txtPayOrder)
            
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            If Not (Rec.BOF And Rec.EOF) Then
                mPOstatus = IIf(IsNull(Rec!POstatus), 0, Rec!POstatus)
                mFwdSeat = IIf(IsNull(Rec!numFwdSeatID), 0, Rec!numFwdSeatID)
                mFwdSeatGpID = IIf(IsNull(Rec!intGroupID), 0, Rec!intGroupID)
            End If
            If mPOstatus = 1 Then
                MsgBox "Approved PayOrded cannot be returned", vbApplicationModal
                cmdReject.Enabled = False
                Exit Sub
            End If
            mSql = ""
            If gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBType <> 4 And gbSeatGroupAccountsOfficer = mFwdSeatGpID Then '= 3 Then
                mStatus = 0
            ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBType <> 4 And gbSeatGroupAccountsOfficer <> mFwdSeatGpID Then '= 3 Then
                mStatus = 5
            ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                mStatus = 0
            ElseIf gbSeatGroupID = gbSeatGroupAssistantSecretary Or gbSeatGroupID = gbSeatGroupHeadClerk Then
                mStatus = 0
            End If
            cmdReject.Enabled = False
            mSql = "Update faPayOrder Set tnyStatus =" & mStatus & "Where vchPayOrderNo = " & val(txtPayOrder)
            mCnn.Execute mSql
            mCnn.Close
            cmdVerify.Enabled = False
            cmdApproval.Enabled = False
            MsgBox "The Pay order Returned to Previous Level", vbApplicationModal
        End If
    End Sub

'''    Private Sub cmdReject_Click()
'''        Dim objDb As New clsDB
'''        Dim Rec As New ADODB.Recordset
'''        Dim mCnn As New ADODB.Connection
'''        Dim mSql As String
'''        mSql = "Select * From faPayOrder Where vchPayOrderNo = " & val(txtPayOrder) & " And tnyStatus = 0"
'''        objDb.SetConnection mCnn
'''        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
'''        If Not (Rec.BOF And Rec.EOF) Then
'''            frmReject.Mode = 6
'''            frmReject.RequestTypeID = txtPayOrder.Text
'''            frmReject.Show vbModal
'''            cmdReject.Enabled = False
'''            cmdApproval.Enabled = False
'''        End If
'''    End Sub

    Private Sub cmdSave_Click()
        txtDrAmount_LostFocus
        If CheckValidation = True Then
            Call SavePaymentOrder
        End If
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
        '================================================================================================'
        '                                   Modified By Cijith On 20/11/2009                             '
        '================================================================================================'
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
        
        '================================================================================================'
    End Sub
    
    Private Sub ClearSubLedger()
        txtSubLedgerType.Text = ""
        txtSubLedgerType.Tag = ""
        txtName.Text = ""
        txtName.Tag = ""
        txtInit1.Text = ""
        txtInit2.Text = ""
        txtInit3.Text = ""
        txtInit4.Text = ""
        txtHouse.Text = ""
        txtStreet.Text = ""
        txtLocalPlace.Text = ""
        txtMainPlace.Text = ""
        txtPost.Text = ""
        txtPin.Text = ""
        txtPhone.Text = ""
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
        
        If val(cmdAllotmentLetterNo.Tag) = 3 Then
            frmSearchTransactionType.ModeOfTransaction = 3
        Else
            frmSearchTransactionType.ModeOfTransaction = 2
        End If
        frmSearchTransactionType.Show vbModal
        If mOldRequisition = False Then
            If gbSearchID < 1 Or val(txtTransactionType.Tag) <> gbSearchID Then
                Call InitializeSelectedHeads
            End If
        End If
        txtTransactionType.Text = gbSearchStr
        txtTransactionType.Tag = gbSearchID
        Call txtTransactionType_LostFocus
        
        If mOldRequisition = False Then
            Call ProjectLink(False)
            Call SetSourceOfFund
        End If
        
        If gbSearchID = 1141 Or gbSearchID = 1151 Or gbSearchID = 1161 Or gbSearchID = 1171 Or gbSearchID = 1181 Or gbSearchID = 1191 Then
            lblDedExclude.Visible = True
            cmdDeductionPayment.Visible = True
        Else
            lblDedExclude.Visible = False
            cmdDeductionPayment.Visible = False
        End If
        If gbSearchID = gbTransactionTypePayBills Then
            lblPension.Visible = True
            txtPensionAmt.Visible = True
        Else
            lblPension.Visible = False
            txtPensionAmt.Visible = False
        End If
        ' BLOCKED BY AIBY ON 10 APRIL 2013
        '        If gbSearchID = gbTransactionTypeUnUtilizedAmount Or gbTransactionTypeProjectExpGO Then
        '            cmdGo.Visible = True
        '            lblGo.Visible = True
        '            txtGo.Visible = True
        '            txtSourceOfFund.Text = ""
        '            txtSourceOfFund.Tag = -1
        '            cmdGo.Enabled = True
        '        End If

        gbSearchStr = ""
        gbSearchID = -1
        gbSearchCode = ""
    End Sub
    
    Private Sub ProjectLink(mLinkFlag As Boolean)
        Call LockProjectType(True)
        
        'Note:- Project and Non Project Payment Orders
        'If (val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192) Or _
             (val(txtTransactionType.Tag) > 1370 And val(txtTransactionType.Tag) < 1382) Then
         If mLinkFlag Then
            fraProject.Visible = True
            Check1.Visible = True
            
        Else
            txtProjectNo.Text = ""
            txtProjectNo.Tag = ""
            
            txtAgreementNo.Text = ""
            txtAgreementNo.Tag = ""
            
            txtImplementingOfficer.Text = ""
            txtImplementingOfficer.Tag = ""
            
            txtSector.Text = ""
            txtSector.Tag = ""
            
            txtCategory.Text = ""
            txtCategory.Tag = ""
            fraProject.Visible = False
            Check1.Visible = False
        End If
    End Sub
    
    Private Sub SetSourceOfFund()
        Dim objTrns As New clsTransactionType
        If val(txtTransactionType.Tag) = 0 Then Exit Sub
        objTrns.SetSourceOfFund (txtTransactionType.Tag)
        If Not IsEmpty(objTrns.SourceFundID) Then
            txtSourceofFund.Text = objTrns.SourceOfFund
            txtSourceofFund.Tag = objTrns.SourceFundID
        Else
            'txtSourceOfFund.Text = "Own Fund"
            'txtSourceOfFund.Tag = 4
            txtSourceofFund.Text = ""
            txtSourceofFund.Tag = ""
        End If
    End Sub
    
    Private Function ShowDetailsForSubCashBook() As Boolean
        On Error GoTo err:
            Dim objAcc As New clsAccounts
            Dim objSubLedger As New clsSubLedger
            If val(txtTransactionType.Tag) = 1001 Or val(txtTransactionType.Tag) = 1211 Then
                'objAcc.SetAccountID (1550)
                'txtCrHeadCode.Text = objAcc.AccountCode
                'txtCrHeadCode.Tag = objAcc.AccountHeadID
                'txtCrAccountHead.Text = objAcc.AccountHead
                'cmdSubsidiaryCash.Enabled = True
                
                txtSubLedgerType.Text = "Officials"
                txtSubLedgerType.Tag = 10
            Else
'                txtCrHeadCode.Text = ""
'                txtCrHeadCode.Tag = ""
'                txtCrAccountHead.Text = ""
'                cmdSubsidiaryCash.Enabled = False
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    
    Private Sub cmdSeat_Click()
'        frmSearchSeat.Show vbModal
'        If gbSearchID <> -1 Then
'            txtForward2Seat.Text = gbSearchStr
'            txtForward2Seat.Tag = gbSearchID
'            gbSearchStr = ""
'            gbSearchID = -1
'        End If
        Dim objdb   As New clsDB             ' Added By Poornima on 05/Nov/2011
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mCnt    As Integer
        Dim mSql    As String
'        vsSeat.Visible = True
        txtForward2Seat.Tag = ""
        txtForward2Seat.Text = ""
        'mSql = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (5,6) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        mSql = "Select chvSeatTitle, numSeatID From GL_Seats Where intGroupID in (5,6,4,16,17) And intLocalBodyID = " & gbLocalBodyID & " Order By chvSeatTitle"
        frmSearchSeat.SQLString = mSql
        frmSearchSeat.Show vbModal
        If gbSearchID > -1 Then
            txtForward2Seat.Tag = gbSearchID
            txtForward2Seat.Text = gbSearchStr
            gbSearchID = -1
            gbSearchStr = ""
        Else
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub
    
    
    Private Sub SetHeadsForRemittanceOfUnUtilizedAmount()
        Dim mSql As String
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim objAcc As New clsAccounts
        If gbLBPanchayat = 1 Then
            If val(txtSourceofFund.Tag) = 1 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1065" & vbNewLine
                'mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 16 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1667" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 17 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1668" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 25 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1068" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 26 Or val(txtSourceofFund.Tag) = 41 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1612" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 27 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1617" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 28 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1618" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 29 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1066" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 30 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1067" & vbNewLine
            Else
               MsgBox "Wrong Source Of Fund Selected", vbApplicationModal
               txtSourceofFund.Text = ""
               txtSourceofFund.Tag = -1
               txtDrHeadCode.Text = ""
                txtDrAccountHead.Text = ""
                txtDrHeadCode.Tag = -1
                txtCrAccountHead.Text = ""
                txtCrHeadCode.Text = ""
                txtCrHeadCode.Tag = -1
               Exit Sub
            End If
            
        Else
            If val(txtSourceofFund.Tag) = 1 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=928" & vbNewLine
                'mSql = mSql + " Order By faTransactionTypeChild.vchAccountHeadcode" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 16 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=2347" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 17 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=2348" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 25 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=931" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 26 Or val(txtSourceofFund.Tag) = 41 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1762" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 27 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1763" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 28 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=1764" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 29 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=929" & vbNewLine
            ElseIf val(txtSourceofFund.Tag) = 30 Then
                mSql = "Select faAccountHeads.vchAccountHeadCode as AccHeadCode ,faAccountHeads.vchAccountHead as AccHead,faAccountHeads.intAccountHeadID, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join " & vbNewLine
                mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId" & vbNewLine
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag) & vbNewLine
                mSql = mSql + " And faAccountHeads.tinHiddenFlag = 0" & vbNewLine
                mSql = mSql + " And faAccountHeads.intAccountHeadID=930" & vbNewLine
            Else
                MsgBox "Wrong Source Of Fund Selected", vbApplicationModal
                txtSourceofFund.Text = ""
                txtSourceofFund.Tag = -1
                txtDrHeadCode.Text = ""
                txtDrAccountHead.Text = ""
                txtDrHeadCode.Tag = -1
                txtCrAccountHead.Text = ""
                txtCrHeadCode.Text = ""
                txtCrHeadCode.Tag = -1
                Exit Sub
            End If
            
        End If
        If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.BOF And Rec.EOF) Then
                txtDrHeadCode.Text = Rec!AccHeadCode
                txtDrAccountHead.Text = Rec!AccHead
                txtDrHeadCode.Tag = Rec!intAccountHeadID
                txtCrAccountHead.Text = Rec!AccHead
                txtCrHeadCode.Text = Rec!AccHeadCode
                txtCrHeadCode.Tag = Rec!intAccountHeadID
                txtDrAmount.SetFocus
            End If
        End If
        
    End Sub
    
    Private Sub cmdSourceOfFund_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.Show vbModal
        'txtSourceOfFund.SetFocus
        If gbSearchID <> -1 Then
            txtSourceofFund.Text = gbSearchStr
            txtSourceofFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
            If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
                txtDrAccountHead.Text = ""
                txtDrHeadCode.Text = ""
                txtDrAccountHead.Tag = -1
                txtDrHeadCode.Tag = -1
                Call SetHeadsForRemittanceOfUnUtilizedAmount
            End If
        End If
    End Sub

    Private Sub cmdSubLederType_Click()
        '================================================================================================'
        '                                   Modified By Cijith On 20/11/2009                             '
        '================================================================================================'
        'frmSearchSubsidiaryAccountHeads.Show vbModal
        txtSubLedgerType.Text = ""
        txtSubLedgerType.Tag = ""
        Call ClearSubLedger
        frmSearchMasters.SQLQry = "Select intSubLedgerTypeID, vchSubLedgerType From faSubLedgerTypes"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.Show vbModal
        If gbSearchStr <> "" Then
            txtSubLedgerType.Text = gbSearchStr
            txtSubLedgerType.Tag = gbSearchID
        End If
        gbSearchID = -1
        gbSearchStr = ""
        '================================================================================================'
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
            
            If val(txtSubsidiaryCash.Tag) <> 0 Then
                txtPayeeType.Text = txtSubsidiaryCash.Text
                Call ShowDetailsForSubCashBook
            End If
        Exit Sub
err:
        MsgBox (Error$)
        '================================================================================================'
    End Sub

    Private Sub cmdVerify_Click()
        Dim mSql    As String
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mForwardedSeat As Double
        Dim mFwdSeat As String
        Dim mStatus As Integer
        If gbLBPanchayat Then
         mSql = " Select numFwdSeatID,intGroupID,* From faPayOrder"
         mSql = mSql + " Inner Join faSeats On faSeats.numSeatID=faPayOrder.numFwdSeatID"
         mSql = mSql + " Where vchPayOrderNo = " & val(txtPayOrder)
         
         mSql = "Select dbo.fnGetSeat(numFwdSeatID) FwdSeat,* From faPayOrder Where vchPayOrderNo= " & val(txtPayOrder)
         
         objdb.SetConnection mCnn
         Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
         If Not (Rec.BOF And Rec.EOF) Then
             mForwardedSeat = Rec!numFwdSeatID
             mFwdSeat = IIf(IsNull(Rec!FwdSeat), "", Rec!FwdSeat)
             mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
         End If
         If gbSeatGroupID = gbSeatGroupAccountsClerk Then
             If val(txtForward2Seat.Tag) < 1 Then
                 txtForward2Seat.SetFocus
                 MsgBox "Please Enter the Forward to Seat", vbInformation
                 Exit Sub
             Else
                 mSql = "Update faPayOrder Set tnyStatus = 5, numSeatID = " & gbSeatID & " ,numUserID =" & gbUserID & " ,numFwdSeatID=" & val(txtForward2Seat.Tag) & " Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
                 mCnn.Execute mSql
                 
                 MsgBox "Sucessfully Forwarded to " & txtForward2Seat.Text, vbApplicationModal
                 cmdVerify.Enabled = False
             End If
         ElseIf gbSeatGroupID = gbSeatGroupHeadClerk Or gbSeatGroupID = gbSeatGroupAssistantSecretary Then
        
             If mForwardedSeat = gbSeatID Then
                 mSql = "Update faPayOrder Set tnyStatus = 3, numFwdSeatID = " & gbSeatID & " Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
                 mCnn.Execute mSql
                 MsgBox "Successfully Verified", vbApplicationModal
                 cmdVerify.Enabled = False
             Else
                 MsgBox "This PayOrder Forwarded to Seat '" & mFwdSeat & "' Please Login as '" & mFwdSeat & "'", vbApplicationModal
             End If
         ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
             If mForwardedSeat = gbSeatID Then
                 If mStatus = 5 Then
                     mSql = "Update faPayOrder Set tnyStatus = 3, numFwdSeatID = " & gbSeatID & "  Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
                     mCnn.Execute mSql
                     MsgBox "Successfully Verified", vbApplicationModal
                     cmdVerify.Enabled = False
                 ElseIf mStatus = 3 Then
                     MsgBox "Already Verified", vbApplicationModal
                 End If
                 UserPrivillage (val(txtPayOrder.Tag))
             Else
                 MsgBox "This PayOrder Forwarded to Seat '" & mFwdSeat & "' Please Login as '" & mFwdSeat & "'", vbApplicationModal
             End If
         End If
        End If
    End Sub

    Private Sub dtpDueDate_CloseUp()
        Dim dtNewDate As Date
        dtNewDate = dtpDueDate.Value
        If dtNewDate >= gbStartingDate And dtNewDate <= gbEndingDate Then
            txtDueDate.Text = DdMmmYy(dtNewDate)
            If CDate(txtDated.Text) > CDate(txtDueDate.Text) Then
                MsgBox "Invalid Date", vbInformation
                txtDueDate = ""
                dtpDueDate.SetFocus
            End If
        Else
            MsgBox "Enter a valid Date", vbInformation
            txtDueDate.Text = ""
        End If
        txtDueDate.SetFocus
    End Sub

    Private Sub Form_Activate()
        Dim intCnt As Integer
        XPC.InitSubClassing
    End Sub
    
    Private Sub Form_Load()
        Call FormInitialize
        vsGrid.ColComboList(1) = "|..."
        
        'Note:- Filling Combo Transaction Type and Instrument Type
        
        
        'Commented bY cijith for Clarification On 28.01.2010'
        
        '''If intLoadMode = 1 Then           'Normal Save Mode    '
        '''    fraApprove.Visible = False
        '''    cmdSave.Enabled = True
        '''ElseIf intLoadMode = 2 Then       'Approving Stage     '
        '''    'fraApprove.Visible = True
        '''    cmdSave.Enabled = False
        '''End If
        
'''        If gbUserTypeID = 3 Then
'''            cmdNew.Visible = True
'''            cmdSave.Visible = True
'''            cmdApproval.Visible = False
'''        ElseIf gbUserTypeID = 2 Or gbUserTypeID = 4 Then
'''            cmdSave.Visible = False
'''            cmdNew.Visible = False
'''            cmdApproval.Visible = True
'''        ElseIf gbUserTypeID = 1 Then
'''            cmdNew.Visible = True
'''            cmdSave.Visible = True
'''            cmdApproval.Visible = True
'''        End If

        '/*Replacing UserTypeID By SeatGroupID*/'
        cmdDeductionPayment.Visible = False
        lblDedExclude.Visible = False
        ''  Modified For SaankhyaWeb Updation
        If gbFinancialYearID > 2016 And gbSaankhyaWeb = 1 Then
            lblAllotmentLetterNo.Visible = False
            txtAllotmentLetterNo.Visible = False
            cmdAllotmentLetterNo.Visible = False
        End If
        If gbSeatGroupID = gbSeatGroupAccountsOfficer And gbLBType <> 4 Then '= 3 Then
            cmdSave.Visible = False
            cmdNew.Visible = False
            cmdApproval.Visible = True
            txtSubsidiaryCash.Enabled = False
            cmdSubsidiaryCash.Enabled = False
'            cmdReject.Visible = True
        ElseIf gbSeatGroupID = gbSeatGroupAccountsSuperintended And gbLBType = 4 Then
            cmdSave.Visible = False
            cmdNew.Visible = False
            cmdApproval.Visible = True
'            cmdReject.Visible = True
''----------------------Added By Anisha On 13 Feb 2015--------------------------------------

        ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
            cmdVerify.Visible = False
            cmdApproval.Visible = False
            cmdVerify.Caption = "Forward"
    
            'cmdSave.Caption = ""
        ElseIf gbSeatGroupID = gbSeatGroupAssistantSecretary Or gbSeatGroupID = gbSeatGroupHeadClerk Then
            cmdSave.Visible = False
            cmdNew.Visible = False
            cmdVerify.Visible = True
            cmdReject.Visible = True
        ElseIf gbSeatGroupID = gbSeatGroupSecretary Then
            cmdSave.Visible = False
            cmdNew.Visible = False
            cmdVerify.Visible = True
            cmdReject.Visible = True
'--------------------------------------------------------------
        Else
            cmdNew.Visible = True
            cmdSave.Visible = True
            cmdApproval.Visible = False
            txtSubsidiaryCash.Enabled = False
            cmdSubsidiaryCash.Enabled = False
        End If
        
        
       
        If gbLBPanchayat = 1 Then
            cmdVerify.Visible = True
            If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
                cmdSeat.Enabled = False
            End If
            
            If txtPayOrder.Tag = "" Then
                cmdVerify.Visible = False
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    cmdSave.Caption = "Save/Fwd"
                End If
            Else
                cmdVerify.Visible = False
                If gbSeatGroupID = gbSeatGroupAccountsClerk Then
                    cmdSave.Caption = "Edit/Fwd"
                End If
            End If
            
        
        Else
            cmdVerify.Visible = False
            cmdReject.Visible = False
        End If
        
        
        SSTab.TabVisible(1) = False
        SSTab.TabVisible(2) = False
        
        Call ProjectLink(False)
        Call SetSourceOfFund
        If mPendingTask = 1 Or mPendingTask = 2 Then
            Call GetPendingTaskDetails
        End If
    End Sub
   
    Private Sub Form_Resize()
        Me.WindowState = 2
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
        mPendingTask = 0
    End Sub







    Private Sub lstMasters_DblClick()
        '-----------------------------------------------------------------'
        '               Added On 07/04/2009 By Cijith Sreedharan          '
        '-----------------------------------------------------------------'
        If lstMasters.Tag = 1 Then
            txtFunctionary.Text = lstMasters.Text
            txtFunctionary.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        ElseIf lstMasters.Tag = 2 Then
            txtFunction.Text = lstMasters.Text
            txtFunction.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        ElseIf lstMasters.Tag = 3 Then
            txtSubLedgerType.Text = lstMasters.Text
            txtSubLedgerType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        ElseIf lstMasters.Tag = 4 Then
            txtForward2Seat.Text = lstMasters.Text
            txtForward2Seat.Tag = Left(gbLocationID, 5) & "0" & lstMasters.ItemData(lstMasters.ListIndex)
        End If
        lstMasters.Clear
        lstMasters.Visible = False
    End Sub

    Private Sub lstMasters_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call lstMasters_DblClick
        End If
    End Sub
    
    Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
        Else
            KeyAscii = 0
        End If
    End Sub

    Private Sub txtAgreementNo_GotFocus()
        If gbSearchID > 0 Then
            Dim objAgreement As New clsAgreement
            objAgreement.SetAgreements (gbSearchID)
            gbSearchID = -1
            If objAgreement.AgreementID > 0 Then
                txtAgreementNo.Text = objAgreement.AgreementNo
                txtAgreementNo.Tag = objAgreement.AgreementID
                txtProjectNo.Text = objAgreement.ProjectSlNo
                txtProjectNo.Tag = objAgreement.ProjectID
                If objAgreement.ProjectID > 0 Then
                    
                End If
            End If
        End If
    End Sub
    
    Private Sub txtAllotmentLetterNo_GotFocus()
        If val(txtAllotmentLetterNo.Tag) > 0 Then
            'Dim objAL As New clsAllotmentLetter
            'objAL.SetAllotmentLetter (val(txtAllotmentLetterNo.Tag))
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim mID As Integer
            Dim mMicroSectorID As Integer
            
            Dim objAcc As New clsAccounts
            Dim objFunc As New clsFunction
            Dim objFunry As New clsFunctionary
            Dim objTrType As New clsTransactionType
            
            mID = val(txtAllotmentLetterNo.Tag)
            
            mSql = "SELECT * FROM faAllotments WHERE  intID = " & mID
            objdb.SetConnection mCnn
            Rec.Open mSql, mCnn
            If Not (Rec.BOF And Rec.EOF) Then
            
                txtAllotmentLetterNo.Text = Rec!vchAllotmentNo
                txtAllotmentLetterNo.Tag = Rec!intID
                txtProjectNo.Text = IIf(IsNull(Rec!vchProjectNo), "", Rec!vchProjectNo)
                txtProjectNo.Tag = IIf(IsNull(Rec!numProjectID), 0, Rec!numProjectID)
                gbProject.decProjectID = Rec!numProjectID
                gbProject.intSourceOfFundID = Rec!intSourceID
                'txtSourceOfFund.Text = ""
                txtSourceofFund.Tag = Rec!intSourceID
                txtCategory.Tag = Rec!intFundCategoryID
                gbSearchStr = IIf(IsNull(gbProject.decProjectID), 0, gbProject.decProjectID)
                mMicroSectorID = IIf(IsNull(Rec!intMircoSectorID), 0, Rec!intMircoSectorID)
                mUnAuthorized = IIf(IsNull(Rec!tnyTypeID), 0, Rec!tnyTypeID)
                
                Call ProjectLink(True)
                Call txtProjectNo_GotFocus
                objAcc.SetAccountID (Rec!intAccountHeadID)
                If objAcc.AccountHeadID > 0 Then
                    txtDrAccountHead.Text = objAcc.AccountHead
                    txtDrAccountHead.Tag = objAcc.AccountType
                    txtDrHeadCode.Text = objAcc.AccountCode
                    txtDrHeadCode.Tag = objAcc.AccountHeadID
                Else
                    txtDrAccountHead.Text = ""
                    txtDrAccountHead.Tag = ""
                    txtDrHeadCode.Text = ""
                    txtDrHeadCode.Tag = ""
                End If
                
                txtDrAmount.Text = Format(Rec!fltAuthorizedAmt, "0.00")
                txtCrAmount.Text = Format(Rec!fltAuthorizedAmt, "0.00")
                txtDrAmount.Enabled = False
                Call CalculateAmt
                
                Call SetExpenditureDetails(val(txtSector.Tag), val(txtCategory.Tag), val(mMicroSectorID), IIf(mPendingTask = 2, gbFinancialYearID - 1, gbFinancialYearID))
                Call LockProjectType(False)
                'Call ProjectLink(True)
                 If txtProjectNo.Tag <> 0 Then
                    Call ProjectLink(True)
                Else
                    Call ProjectLink(False)
                End If
            Else
                Call ProjectLink(False)
            End If

            txtAllotmentLetterNo.Enabled = False
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

 

Private Sub txtCP_KeyPress(KeyAscii As Integer)
     If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
        Else
            KeyAscii = 0
        End If
End Sub

    Private Sub txtCrAmount_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtCrHeadCode.Text = Token(gbSearchStr, " ")
            txtCrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
        End If
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
    
    Private Sub txtDated_LostFocus()
        txtDated.Text = CheckDateInMMM(txtDated.Text)
        If mPendingTask <> 1 Then
            If CDate(txtDated.Text) > gbTransactionDate Then
                MsgBox "Please Enter Valid Date"
                txtDated.SetFocus
                txtDated.Text = Format(gbTransactionDate, "DD/MMM/yyyy")
                Exit Sub
            End If
        End If
    End Sub
    
    Private Sub txtDrAccountHead_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtDrHeadCode.Text = Token(gbSearchStr, " ")
            txtDrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
            ShowBudgetBalance (gbSearchID)
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtDrAccountHead_LostFocus()
        'Call Set_GlossaryHead
    End Sub

    Private Sub txtDrAmount_GotFocus()
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtDrHeadCode.Text = Token(gbSearchStr, " ")
            txtDrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
            ShowBudgetBalance (gbSearchID)
            gbSearchID = -1
        End If
    End Sub
    
    Private Sub txtDrAmount_KeyPress(KeyAscii As Integer)
        If vsGrid.Enabled = True Then
            If KeyAscii = 13 Then vsGrid.SetFocus
            If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
            Else
                KeyAscii = 0
            End If
        End If
         
    End Sub
    
    Private Sub txtDrAmount_LostFocus()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim mAmt As Double
        
        'txtDrAmount.Text = Format(txtDrAmount.Text, "0.00")
        mAmt = Format(val(txtDrAmount.Text), "0.00")
        If (mAmt - Int(mAmt)) > 0 Then
            mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
        End If
        txtDrAmount.Text = Format(mAmt, "0.00")
        
        If mBudgetBalanceAmt < val(txtDrAmount) Then
            'MsgBox "Budget Balance is Rs. " & Format(mBudgetBalanceAmt, "0.00")
            'txtDrAmount = Format(mBudgetBalanceAmt, "0.00")
        End If
        txtCrAmount.Text = Format((val(txtDrAmount.Text) - CalculateAmt), "0.00")
                        
    End Sub

    Private Sub txtDrAmount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
        txtDrAmount.Locked = True
    Else
        txtDrAmount.Locked = False
    End If
    End Sub

    Private Sub txtDrHeadCode_GotFocus()
        mHelpTips = "Select the Budget Head for the Gross Amount" & vbLf
        lblTipText.Caption = mHelpTips
        If gbSearchStr <> "" Then
            Dim mStr As String
            txtDrHeadCode.Text = Token(gbSearchStr, " ")
            txtDrAccountHead.Text = Trim(gbSearchStr)
            gbSearchStr = ""
            ShowBudgetBalance (gbSearchID)
            gbSearchID = -1
        End If
    End Sub

    Private Sub txtDrHeadCode_LostFocus()
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objdb As New clsDB
        
        If Trim(txtDrHeadCode) <> "" Then
            cmdCrAccountHead.Enabled = True
            Dim mGroupID As Integer
            objdb.SetConnection mCnn
            mSql = " Select intGroupID  From faTransactionTypeChild Where intTransactionTypeID = " & val(txtTransactionType.Tag) & " AND vchAccountHeadCode = '" & Trim(txtDrHeadCode.Text) & "'"
            Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
            If Not (Rec.BOF And Rec.EOF) Then
                mGroupID = IIf(IsNull(Rec!intGroupID), 0, Rec!intGroupID)
            End If
            Rec.Close
            
            ' Note:-
            ' This is one of the Nasty way of writing Code! No need to get upset because some times
            ' such methodology will help  you meet your target date and time. Any way let me explain!
            ' Table:-faTransactionTypeChild
            '       tnyListId = 1 -> Gross Head 2-> Deduction Heads 3-> Net Payable Head List
            '       intGroupID can be used to group set of related head under on transaction type
            ' So here checking is
            '       if once the gross head is selected and if it have a Group ID then you can check
            '       related deduction heads and Net payable heads
            
            If mGroupID > 0 Then
                mSql = "Select * From faTransactionTypeChild Inner Join "
                mSql = mSql + " faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionTypeChild.intAccountHeadID "
                mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag)
                mSql = mSql + " AND faTransactionTypeChild.intGroupID = " & mGroupID & "AND tnyNetPayFlag = 1 And tnyListID = 3"
                
                Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
                If Not (Rec.BOF And Rec.EOF) Then
                    'Note:- Group ID wise matching Net Payable Head Found
                    txtCrHeadCode.Text = Rec!vchAccountHeadCode
                    txtCrHeadCode.Tag = Rec!intAccountHeadID
                    txtCrAccountHead.Text = Rec!vchAccountHead
                    cmdCrAccountHead.Enabled = False
                Else
                    'Note:- GroupID wise Matching NetPayable Head Not defined
                    '       Then List All NetPayble if defined
                    Rec.Close
FindAnyNetPay:
                    mSql = "Select * From faTransactionTypeChild Inner Join "
                    mSql = mSql + " faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionTypeChild.intAccountHeadID "
                    mSql = mSql + " Where intTransactionTypeID = " & val(txtTransactionType.Tag)
                    'mSQL = mSQL + " AND tnyNetPayFlag = 1 "
                    mSql = mSql + " And tnyListID = 3"
                    Rec.CursorLocation = adUseClient
                    Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
                    If Not (Rec.BOF And Rec.EOF) Then
                        If Rec.RecordCount = 1 Then
                            'Note:-Only One NetPayable Head is Defined so
                            txtCrHeadCode.Text = Rec!vchAccountHeadCode
                            txtCrHeadCode.Tag = Rec!intAccountHeadID
                            txtCrAccountHead.Text = Rec!vchAccountHead
                            cmdCrAccountHead.Enabled = False
                        Else
                            'Note:- More than one head is defined as NetPayable Item so
                            '       Let users to select items ( Nothing much here to automate)
                            mSelectCreditHeadFlag = True
                            txtCrHeadCode.Text = ""
                            txtCrHeadCode.Tag = -1      'Rec!intAccountHeadID  MODIFIED ON 25-8-2011
                            txtCrAccountHead.Text = ""
                            cmdCrAccountHead.Enabled = True
                            lblBudgetAmt.Caption = "0.00"
                            lblUtilizedAmt.Caption = "0.00"
                        End If
                    Else
                        mSelectCreditHeadFlag = False ' This fill Debit Head into Credit Head
                        txtCrHeadCode.Text = ""
                        txtCrHeadCode.Tag = ""
                        txtCrAccountHead.Text = ""
                        cmdCrAccountHead.Enabled = True
                        lblBudgetAmt.Caption = "0.00"
                        lblUtilizedAmt.Caption = "0.00"
                      
                            If val(txtCrHeadCode.Tag) < 1 Then
                                txtCrHeadCode.Text = txtDrHeadCode.Text
                                txtCrHeadCode.Tag = txtDrHeadCode.Tag
                                txtCrAccountHead.Text = txtDrAccountHead.Text
                            End If
                      
                    End If
                    Rec.Close
                End If
            Else
                'Note:-Gross head is not grouped under and classification so
                '      Flow of control is redirected to fetch all NetPayble from the list
                GoTo FindAnyNetPay:
            End If
            
            '--------------------------------------------------------------------'
            'Note:-Budget Validation
            '--------------------------------------------------------------------'
            Dim objBudj As New clsBudgetCentre
            objBudj.SetBudgetAccountHead val(txtDrHeadCode.Tag), val(txtFunctionary.Tag), val(txtFunction.Tag)
            If objBudj.BudgetCentreID > 0 Then
                lblBudgetAmt.Caption = Format(objBudj.EstimatedAmount, "0.00")
                lblUtilizedAmt.Caption = Format(objBudj.UtilisedAmount, "0.00")
            End If
            Set objBudj = Nothing
            
        End If
        lblTipText.Caption = ""
    End Sub

    Private Sub txtDueDate_LostFocus()
        Dim dtNewDate As Date
        
        If Trim(txtDueDate) <> "" Then
        dtNewDate = CheckDateInMMM(txtDueDate.Text)
            txtDueDate.Text = CheckDateInMMM(txtDueDate.Text)
            If mPendingTask = 0 Then
                If CDate(txtDated.Text) > CDate(txtDueDate.Text) Then
                    MsgBox "Invalid Date", vbInformation
                    txtDueDate = ""
                    txtDueDate.SetFocus
                End If
                If dtNewDate >= gbStartingDate And dtNewDate <= gbEndingDate Then
                txtDueDate.Text = DdMmmYy(dtNewDate)
                If CDate(txtDated.Text) > CDate(txtDueDate.Text) Then
                    MsgBox "Invalid Date", vbInformation
                    txtDueDate = ""
                    dtpDueDate.SetFocus
                    End If
                Else
                    MsgBox "Enter a valid Date", vbInformation
                    txtDueDate.Text = ""
                End If
            Else
                If CDate(txtDated.Text) > CDate(txtDueDate.Text) And CDate(txtDueDate.Text) < DateAdd("yyyy", -1, gbEndingDate) Then
                    MsgBox "Invalid Date", vbInformation
                    txtDueDate = ""
                    txtDueDate.SetFocus
                End If
                If CDate(txtDueDate.Text) > gbTransactionDate And CDate(txtDueDate.Text) <= gbEndingDate Then
                    MsgBox "Please Enter Prevous Year date", vbInformation
                    txtDueDate = ""
                    txtDueDate.SetFocus
                End If
            End If
            
            
        End If
    End Sub

    Private Sub txtForward2Seat_GotFocus()
'        If gbSearchStr <> "" Then
'            txtForward2Seat.Text = gbSearchStr
'            txtForward2Seat.Tag = gbSearchID
'            gbSearchID = -1
'            gbSearchStr = ""
'        End If
    End Sub
    Private Sub FetchDetailsUsingGO()
        Dim mSql        As String
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim objdb       As New clsDB
        
        If val(txtTransactionType.Tag) = gbTransactionTypeProjectExpGO Or val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
            If txtGo.Text <> "" Then
                mSql = " Select * From suGOForFunds "
                mSql = mSql + " Left Join suSourceOfFund On suGOForFunds.intSourceOfFundID=suSourceOfFund.intSourceFundID"
                If (objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya)) Then
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.BOF And Rec.EOF) Then
                        txtGo.Text = Rec!vchRefNo
                        txtGo.Tag = Rec!intRefID
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub txtImplementingOfficer_GotFocus()
'        If gbSearchID > 0 Then
'            txtImplementingOfficer.Text = gbSearchStr
'            txtImplementingOfficer.Tag = gbSearchID
'            gbSearchID = -1
'            gbSearchCode = ""
'            gbSearchStr = ""
'        End If
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
    Private Sub txtPensionAmt_KeyPress(KeyAscii As Integer)
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then
        Else
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub txtProjectNo_GotFocus()
        Dim mProjectID As Integer
        Dim objProj As New clsProject
        If val(gbSearchStr) > 0 Then
            If mPendingTask = 2 Then
                objProj.SetProject val(gbSearchStr), gbFinancialYearID - 1
            Else
                objProj.SetProject val(gbSearchStr)
            End If
            If objProj.ProjectID > 0 Then
                txtProjectNo.Text = objProj.ProjectSerialNo
                txtProjectNo.Tag = objProj.ProjectID
                
                'txtCategory.Text = objProj.Category
                'txtCategory.Tag = objProj.CategoryID
                
                txtSector.Text = objProj.SubSector
                txtSector.Tag = objProj.SubSectorID
                
            End If
        End If
        gbSearchStr = ""
        
'        If gbProject.decProjectID > 0 Then
'            Dim objProj As New clsProject
'            objProj.SetProject gbProject.decProjectID
'            If objProj.ProjectID > 0 Then
'                txtProjectNo = objProj.ProjectSerialNo
'                txtProjectNo.Tag = objProj.ProjectID
'
'                txtCategory.Text = objProj.Category
'                txtCategory.Tag = objProj.ProjCatID
'
'                txtSector.Text = objProj.Sector
'                txtSector.Tag = objProj.SectorTypeID
'
'                If IsNumeric(gbProject.intSourceOfFundID) Then
'                    txtSourceOfFund.Tag = gbProject.intSourceOfFundID
'                End If
'                txtSourceOfFund.Text = objProj.FindSourceOfFund(val(txtSourceOfFund.Tag))
'                txtSourceOfFund.Tag = IIf(IsNull(objProj.SourceOfFundID), -1, objProj.SourceOfFundID)
'
'                Call GetImplementingOfficer(objProj.ProjectID)
'
'            End If
'            With gbProject
'                .decProjectID = Null
'                .intLBID = Null
'                .intYearID = Null
'                .intProjectSlNo = Null
'                .chvProjectSlNo = Null
'                .chvProjectName = Null
'                .chvProjectnameEnglish = Null
'                .intProjCatID = Null
'                .chvDPCOrderNo = Null
'                .dtDPCOrderDate = Null
'                .intSectorTypeID = Null
'                .intPlanID = Null
'                .intSourceOfFundID = Null
'                .fltEstSourceAmt = Null
'            End With
'        End If
    End Sub
    
    Private Sub txtSourceOfFund_GotFocus()
'''        If gbSearchID > -1 And Trim(gbSearchStr) <> "" Then
'''            txtSourceOfFund.Text = gbSearchStr
'''            txtSourceOfFund.Tag = gbSearchID
'''            gbSearchID = -1
'''            gbSearchCode = ""
'''            gbSearchStr = ""
'''        Else
'''            txtSourceOfFund.Text = ""
'''            txtSourceOfFund.Tag = ""
'''        End If
    End Sub

    Private Sub txtSubLedgerType_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = 46 Then 'Delete Key
            txtSubLedgerType.Text = ""
            txtSubLedgerType.Tag = ""
        Else
            txtSubLedgerType.Locked = True
        End If
    End Sub

    Private Sub txtTransactionType_DblClick()
        Call cmdSearchTransactionType_Click
    End Sub

    Private Sub txtTransactionType_LostFocus()
        If val(txtTransactionType.Tag) > 0 Then
            If txtTransactionType.Tag = gbTransactionTypePayBills Then
'                chkPension.Visible = True
                txtPensionAmt.Visible = True
                lblPension.Visible = True
            End If
            Call SetFormControls
        End If
    End Sub

    Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If vsGrid.Row <> 1 Then
            If vsGrid.TextMatrix(Row - 1, 1) = "" Then
                Cancel = True
            End If
        End If
    End Sub
    
    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) = False And vsGrid.Col = 3 Then
            vsGrid.TextMatrix(vsGrid.Row, 3) = ""
            MsgBox "Enter Numeric values"
        End If
        Dim mTOt As Variant
        Dim mAmt As Double
        Dim objAc As New clsAccounts
        
        If Col = 1 Then
            objAc.SetAccountCode (Trim(vsGrid.TextMatrix(Row, 1)))
            If objAc.AccountHeadID > 0 Then
                vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
                vsGrid.TextMatrix(vsGrid.Row, 1) = objAc.AccountCode
                vsGrid.TextMatrix(vsGrid.Row, 2) = objAc.AccountHead
                
                vsGrid.Col = 3
            Else
                
                vsGrid.TextMatrix(vsGrid.Row, 4) = ""
                vsGrid.TextMatrix(vsGrid.Row, 1) = ""
                vsGrid.TextMatrix(vsGrid.Row, 2) = ""
                vsGrid.TextMatrix(vsGrid.Row, 3) = ""
                vsGrid.RemoveItem (Row)
                Call txtDrAmount_LostFocus
                vsGrid.Col = 1
                
            End If
            Exit Sub
        End If
        
        If vsGrid.Col = 3 Then
                If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) Then
                    'vsGrid.TextMatrix(vsGrid.Row, 3) = Form at(val(vsGrid.TextMatrix(vsGrid.Row, 3)), "0.00")
                    If vsGrid.TextMatrix(vsGrid.Row, 1) = "" Then
                        vsGrid.TextMatrix(vsGrid.Row, 3) = ""
                        
                    End If
                    mAmt = Format(val(vsGrid.TextMatrix(vsGrid.Row, 3)), "0.00")
                    If (mAmt - Int(mAmt)) > 0 Then
                        mAmt = mAmt + (1 - (mAmt - Int(mAmt)))
                    End If
                    If mAmt > 0 Then
                        vsGrid.TextMatrix(vsGrid.Row, 3) = Format(mAmt, "0.00")
                    Else
                        vsGrid.TextMatrix(vsGrid.Row, 3) = ""
                    End If
                    mTOt = CalculateAmt
                    If val(txtDrAmount.Text) > val(mTOt) Then
                        txtCrAmount.Text = val(txtDrAmount.Text) - val(mTOt)
                    Else
                        MsgBox "Amount Out of Range"
                        Call txtDrAmount_LostFocus
                    End If
                End If
        End If
        If vsGrid.Col = 3 Then    'Added By Poornima on 09/11/2011
            If val(vsGrid.TextMatrix(Row, 3)) < 0 Then
                vsGrid.TextMatrix(Row, 3) = ""
            End If
        End If
        
    End Sub
    
    Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        vsGrid.TextMatrix(vsGrid.Row, 3) = Format(vsGrid.TextMatrix(vsGrid.Row, 3), "0.00")
    End Sub


    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        Dim objAc As New clsAccounts
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        
        frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faTransactionTypeChild INNER JOIN faAccountHeads ON faAccountHeads.vchAccountHeadCode= faTransactionTypeChild.vchAccountHeadCode WHERE faTransactionTypeChild.intTransactionTypeID=" & val(txtTransactionType.Tag) & " And faTransactionTypeChild.tnyListID = 2 And tinHiddenFlag = 0 And faAccountHeads.intGroupID is Null Order By faAccountHeads.vchAccountHeadCode"
        frmSearchAccountHeads.VoucherMode = 201
        frmSearchAccountHeads.Show vbModal
        If gbSearchID <> -1 Then
            If vsGrid.FindRow(gbSearchID, , 4) > 0 Then
                MsgBox "This AccountHead Alrady Selected", vbInformation
                Exit Sub
            End If
            'vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
            'vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
            mSql = Token(gbSearchStr, " ")
            objAc.SetAccountCode (mSql)
            If objAc.AccountHeadID > 0 Then
                vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
                vsGrid.TextMatrix(vsGrid.Row, 1) = objAc.AccountCode
                vsGrid.TextMatrix(vsGrid.Row, 2) = objAc.AccountHead
                vsGrid.Col = 3
                If gbLBPanchayat <> 1 Then
                    If vsGrid.TextMatrix(vsGrid.Row, 1) = "350200129" Then
                        mCpFlag = True
                        txtCP.Visible = True
                        lblCP.Visible = True
                        chkPensionContribution.Visible = True
                    End If
                End If
            Else
                vsGrid.TextMatrix(vsGrid.Row, 4) = ""
                vsGrid.TextMatrix(vsGrid.Row, 1) = ""
                vsGrid.TextMatrix(vsGrid.Row, 2) = ""
                vsGrid.Col = 1
            End If
            gbSearchStr = ""
            gbSearchID = -1
        Else
            If vsGrid.TextMatrix(Row + 1, 1) = "" Then
                vsGrid.TextMatrix(Row, 0) = ""
                vsGrid.TextMatrix(Row, 1) = ""
                vsGrid.TextMatrix(Row, 2) = ""
                vsGrid.TextMatrix(Row, 3) = ""
                vsGrid.TextMatrix(Row, 4) = ""
                vsGrid.TextMatrix(Row, 5) = ""
                txtCrAmount.Text = Format((val(txtDrAmount.Text) - CalculateAmt), "0.00")
            End If
        End If

        'If Val(txtTransactionType.Tag) = 1001 Or Val(txtTransactionType.Tag) = 1007 Then
        'frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faTransactionTypeChild INNER JOIN faAccountHeads ON faAccountHeads.vchAccountHeadCode= faTransactionTypeChild.vchAccountHeadCode WHERE faTransactionTypeChild.intTransactionTypeID=1001 And tnyListID = 2"
        '
        'frmSearchAccountHeads.Show vbModal
        'vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        'vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        'objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
        'vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
        'objDb.SetConnection mCnn
        'Rec.Open "Select intGroupID From faTransactionTypeChild Where vchaccountheadcode ='" & vsGrid.TextMatrix(vsGrid.Row, 1) & "'", mCnn
        '    If Not (Rec.EOF And Rec.BOF) Then
        '         vsGrid.TextMatrix(vsGrid.Row, 5) = Rec!intGroupID
        '    End If
        '    mCnn.Close
        'vsGrid.Col = 3
        'gbSearchStr = ""
        '
        'ElseIf Val(txtTransactionType.Tag) = 1002 Or Val(txtTransactionType.Tag) = 1003 Then
        'frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where faaccountHeads.vchAccountHeadCode Between '350200201' and '350200299'"
        'frmSearchAccountHeads.Show vbModal
        'vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
        'vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
        'objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
        'vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
        'objDb.SetConnection mCnn
        'Rec.Open "Select intGroupID From fatransactiontypechild Where vchaccountheadcode ='" & vsGrid.TextMatrix(vsGrid.Row, 1) & "'", mCnn
        '    If Not (Rec.EOF And Rec.BOF) Then
        '         vsGrid.TextMatrix(vsGrid.Row, 5) = Rec!intGroupID
        '    End If
        '    mCnn.Close
        'vsGrid.Col = 3
        'gbSearchStr = ""
        'Call TransactionTemplate(cmbTransactionType.ItemData(cmbTransactionType.ListIndex))
        'Else
        'If Trim(txtDrHeadCode) <> "" Then
        '    'Note:-
        '    'Debit Head is selected
        '    Dim mGroupID As Integer
        '    mGroupID = 0
        '    objDb.SetConnection mCnn
        '    mSql = "Select * From faTransactionTypeChild Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " AND vchAccountHeadCode = '" & txtDrHeadCode.Text & "'"
        '    Rec.Open mSql, mCnn, adOpenForwardOnly, adLockOptimistic
        '    If Not (Rec.BOF And Rec.EOF) Then
        '        mGroupID = Rec!intGroupID
        '        If mGroupID = 0 Then GoTo ListAllDeductions:
        '        mSql = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        '        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        '        mSql = mSql + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        '        mSql = mSql + " And (faTransactionTypeChild.intGroupID = 0 Or faTransactionTypeChild.intGroupID = " & mGroupID & " )"
        '        mSql = mSql + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        '        frmSearchAccountHeads.SQLString = mSql
        '    Else
        'ListAllDeductions:
        '        mSql = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        '        mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        '        mSql = mSql + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        '        mSql = mSql + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        '        frmSearchAccountHeads.SQLString = mSql
        '    End If
        '    Rec.Close
        '
        'Else
        '    mSql = " Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Inner Join "
        '    mSql = mSql + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadID"
        '    mSql = mSql + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tinDebitOrCredit = 0"
        '    mSql = mSql + " And faTransactionTypeChild.tnyNetPayFlag <> 1"
        '    frmSearchAccountHeads.SQLString = mSql
        'End If
        ''frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads "
        '
            
        'End If
    End Sub


    Private Sub vsGrid_GotFocus()
'        If Trim(txtDrAmount.Text = "") Then
'            MsgBox "It is Mandatory to Enter"
'            txtDrAmount.SetFocus
'            Exit Sub
'        End If
    End Sub
    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        Dim mTOt As Variant
        If KeyAscii = 13 And vsGrid.Col = 3 And Trim(vsGrid.TextMatrix(vsGrid.Row, 3)) <> "" And IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) And val(vsGrid.TextMatrix(vsGrid.Row, 3)) > 0 Then
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.Row = vsGrid.Row + 1
            vsGrid.Col = 1
            mTOt = CalculateAmt
            
            If val(txtDrAmount.Text) > val(mTOt) Then
                txtCrAmount.Text = val(txtDrAmount.Text) - val(mTOt)
            End If
        End If
    End Sub
    Private Function CheckValidation() As Boolean
        CheckValidation = False
        If Not IsDate(txtDated) Then
            MsgBox "Please Check the Transaction Date", vbInformation
            txtDated.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtFunctionary.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Functionary", vbInformation
            cmdSearchFunctionary.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtFunction.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Function", vbInformation
            cmdSearchFunction.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtTransactionType.Tag) < 1 Then
            MsgBox "Please Select Proper Transaction Type for this Transaction", vbInformation
            txtTransactionType.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If IsDate(txtDueDate.Text) = False Then
            MsgBox "Please Give Due Date", vbInformation
            dtpDueDate.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtDrHeadCode.Tag) < 1 Then
            txtDrHeadCode.SetFocus
            MsgBox "Please Enter Debit Head", vbInformation
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtDrAmount) <= 0 Then
            txtDrAmount.SetFocus
            MsgBox "Please Enter Debit Amount", vbInformation
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtCrHeadCode.Tag) <= 0 Then
            MsgBox "Please Select The Credit Account Head", vbInformation
            cmdCrAccountHead.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtCrAmount.Text) <= 0 Then
            MsgBox "Please check the Credit Amount", vbInformation
            txtCrAmount.SetFocus
            CheckValidation = False
            Exit Function
        End If
        
        If val(txtSourceofFund.Tag) < 1 Then
            MsgBox "Please Select the Source Of Fund", vbInformation
            txtSourceofFund.SetFocus
            CheckValidation = False
            Exit Function
        End If
        If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupChiefCashier Then
            If val(txtForward2Seat.Tag) < 1 Then
                txtForward2Seat.SetFocus
                CheckValidation = False
                MsgBox "Please Enter the Forward to Seat", vbInformation
                Exit Function
            End If
        End If
        If Trim(txtName.Text) = "" Then
            'SSTab.Tab = 1
            txtName.SetFocus
            MsgBox "Please Enter the Name of Payee..", vbInformation
            CheckValidation = False
            Exit Function
        End If
        
        ''  Modified For SaankhyaWeb Updation

        If mPendingTask = 0 And gbFinancialYearID >= 2017 And gbSaankhyaWeb = 1 Then
            'Exit Function
        Else
        
            If mUnAuthorized <> 3 Then
                If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Then
                   ''Validation disabled for Web integration
                    'If val(txtProjectNo.Tag) < 1 Then
'                        MsgBox "Please Enter a Project", vbInformation
'                        'txtProjectNo.SetFocus
'                        CheckValidation = False
'                        Exit Function
                    'End If
                    
'                    If val(txtCategory.Tag) < 1 Then
'                        MsgBox "Please select a Category", vbInformation
'                        txtCategory.SetFocus
'                        CheckValidation = False
'                        Exit Function
'                    End If
'
'                    If val(txtSector.Tag) < 1 Then
'                        MsgBox "Please select a Sector", vbInformation
'                        txtSector.SetFocus
'                        CheckValidation = False
'                        Exit Function
'                    End If
'
'                    If val(txtAgreementNo.Tag) < 1 Then
'                        If Not mSkipMsgFlag Then
'                            MsgBox "Please enter the agreement No. if you  have with this bill.", vbInformation
'                            txtAgreementNo.SetFocus
'                            CheckValidation = False
'                            mSkipMsgFlag = True
'                            Exit Function
'
'                        End If
'                    End If
                End If
            End If
            
            If val(txtAllotmentLetterNo.Tag) > 0 Then
                If val(txtDrAmount.Text) > val(txtAllotedAmt.Text) Then
                    MsgBox "Please check the Alloted Amount", vbInformation
                    CheckValidation = False
                    Exit Function
                End If
            End If
            
        End If
           
        
    '    If txtSubsidiaryCash.Text <> "" Then
    '        If val(txtSubLedgerType.Tag) = 10 And txtName.Text = "" Then
    '            MsgBox "Please Select the Official for disbursing the Subsidiary Cash"
    '            txtName.SetFocus
    '            CheckValidation = False
    '            Exit Function
    '        End If
    '    End If
        
        
        '-----------------------------------------------------------------------'
        '           Modified By Anisha On 11/04/2011
        ' To Avoid Editing Pre year's PO in Current Year
        If mPendingTask = 0 Then
            If gDateValidation(CDate(txtDated.Text)) = False Then
                    MsgBox "Entered Date Does not include in This Financial Year", vbApplicationModal
                    txtDated.Locked = True
                    txtDueDate.Locked = True
                    CheckValidation = False
                    Exit Function
            End If
        End If
        '-----------------------------------------------------------------------'
        
        '-----------------------------------------------------------------------'
        '           Added By Anisha On 4/12/2012
        '           Disable Auto generation of Pension Contribution Amount
        '----*********************************----------------------------------'
        'If val(txtTransactionType.Tag) = gbTransactionTypePayBills Then
        '    If txtPensionAmt.Text = "" Then
        '        MsgBox "Please Enter Pension Contribution Amount", vbApplicationModal
        '        txtPensionAmt.SetFocus
        '        CheckValidation = False
        '        Exit Function
        '    Else
        '        If val(txtPensionAmt.Text) > val(txtCrAmount.Text) Then
        '            MsgBox "Pension Contribution Amount must be less than Net Payable", vbApplicationModal
        '            CheckValidation = False
        '            Exit Function
        '        End If
        '    End If
        'End If
        '
        
        '-----------------------------------------------------------------------'
        
        If val(txtTransactionType.Tag) = gbTransactionTypePayBills Then
            
            If gbLBPanchayat Then ' P A N C H A Y A T
                If val(txtDrHeadCode.Tag) = 325 Or _
                    val(txtDrHeadCode.Tag) = 326 Or _
                    val(txtDrHeadCode.Tag) = 328 Or _
                    val(txtDrHeadCode.Tag) = 329 Then
                    If val(txtPensionAmt.Text) < 1 Then
                        MsgBox "Please Enter The Pension Contribution Amount!", vbInformation
                        txtPensionAmt.Visible = True
                        txtPensionAmt.Enabled = True
                        lblPension.Visible = True
                        txtPensionAmt.SetFocus
                        CheckValidation = False
                        Exit Function
                    End If
                End If
            Else ' M U N I C I P A L I T Y
                If val(txtDrHeadCode.Tag) = 333 Or _
                    val(txtDrHeadCode.Tag) = 334 Or _
                    val(txtDrHeadCode.Tag) = 335 Or _
                    val(txtDrHeadCode.Tag) = 336 Or _
                    val(txtDrHeadCode.Tag) = 338 Then
                    If val(txtPensionAmt.Text) < 1 Then
                        If chkPensionContribution.Value = vbUnchecked Then
                    
                            MsgBox "Please Enter The Pension Contribution Amount!", vbInformation
                            txtPensionAmt.Visible = True
                            txtPensionAmt.Enabled = True
                            lblPension.Visible = True
                            txtPensionAmt.SetFocus
                            CheckValidation = False
                            Exit Function
                        End If
                        
                    End If
                    If mCpFlag = True Then
                        If val(txtCP.Text) < 1 Then
                            If MsgBox(" Please Enter Contributory pension Amount", vbYesNo, "Saankhya") = vbYes Then
                                txtCP.SetFocus
                                Exit Function
                            Else
                                
                            End If
                        End If
                    End If
                End If
            
            End If
        End If
        If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
            If txtGo.Text = "" Then
                MsgBox "Please Select GO Details", vbApplicationModal
                txtGo.Visible = True
                lblGo.Visible = True
                
                cmdGo.Visible = True
                txtGo.SetFocus
                CheckValidation = False
                Exit Function
            End If
        End If
        
        CheckValidation = True
    End Function

Private Sub SavePayOrder() '
    
    '-----------------------------------------------------------------------'
    '   Added ON 08/04/2009 By Cijith Sreedharan for Saving PayOrder        '
    '-----------------------------------------------------------------------'
    Dim objdb As New clsDB
    Dim objAccounts As New clsAccounts
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim aryIn As Variant
    Dim aryOut As Variant
    Dim mLoop As Integer
    
    '-----------------------------------------------------------------------'
    '              Declaring Member Variables for faPayOrder                '
    '-----------------------------------------------------------------------'
    Dim intPayOrderID As Variant
    Dim vchPayOrderNo As Variant
    Dim dtPayOrderDate As Variant
    Dim dtDueDate As Variant
    Dim intFunctionaryID As Variant
    Dim intFunctionID As Variant
    Dim intTransactionTypeID As Variant
    Dim vchBillNo As Variant
    Dim numBillAmount As Variant
    Dim dtBillDate As Variant
    Dim intInstrumentTypeID As Variant
    Dim intCashOrBankHeadID As Variant
    Dim intSubLedgerTypeID As Variant           'Used in main and address'
    Dim intPayToSubLedgerID As Variant
    Dim intSubsidiaryCashBookID As Variant
    Dim intImplementingOfficerID As Variant
    Dim numProjectNo As Variant
    Dim intStockRegisterID As Variant
    Dim vchStockRefNo As Variant
    Dim intAssetTypeID As Variant
    Dim intAssetID As Variant
    Dim numFwdSeatID As Variant
    Dim intLocalBodyID As Variant
    Dim intZonalID As Variant
    Dim intFinancialYearID As Variant
    Dim numUserID As Variant
    Dim numSeatID As Variant
    Dim numApprovingOfficerID As Variant        'Not saved now'
    Dim numApprovingSeatID As Variant           'Not saved now'
    Dim dtApprovingDate As Variant              'Not saved now'
    Dim intVoucherID As Variant                 'Not saved now'
    Dim intVoucherNo As Variant                 'Not saved now'
    Dim dtVoucherDate As Variant                'Not saved now'
    Dim tnyStatus As Variant
    Dim tnyCancelled As Variant
    '-----------------------------------------------------------------------'
    '              Declaring Member Variables for faPayOrderChild           '
    '-----------------------------------------------------------------------'
    Dim intSlNo As Variant
    Dim intAccountHeadID As Variant
    Dim vchAccountHeadCode As Variant
    Dim numAmount As Variant
    Dim tnyCategoryFlag As Variant
    Dim tnyDebitOrCreditFlag As Variant
    '-----------------------------------------------------------------------'
    '              Declaring Member Variables for faPayOrderAddress         '
    '-----------------------------------------------------------------------'
    Dim intSubsidiaryAccountHeadID  As Variant
    Dim intSubLegerTypeID As Variant
    Dim vchSubLedgerCode As Variant
    Dim vchName As Variant
    Dim vchHouseName As Variant
    Dim vchStreet As Variant
    Dim vchLocalPlace As Variant
    Dim vchMainPlace As Variant
    Dim vchPost As Variant
    Dim vchPinCode As Variant
    Dim vchPhone As Variant
    
    '-----------------------------------------------------------------------'
    '              Defining Member Variables for faPayOrder                 '
    '-----------------------------------------------------------------------'
    
    If mCnn.State = 1 Then mCnn.Close
    If objdb.SetConnection(mCnn) = True Then
On Error GoTo err:
        mCnn.BeginTrans         'Begining Transactions
        intPayOrderID = val(txtPayOrder.Tag)
        vchPayOrderNo = val(txtPayOrder.Text)
        dtPayOrderDate = txtDated.Text
        dtDueDate = txtDueDate.Text
        intFunctionaryID = val(txtFunctionary.Tag)
        intFunctionID = val(txtFunction.Tag)
        intTransactionTypeID = val(txtTransactionType.Tag)
        vchBillNo = "" 'Trim(txtBillNo.Text)
        numBillAmount = Null 'Val(txtBillAmt.Text)
        dtBillDate = Null
        intInstrumentTypeID = Null
        intCashOrBankHeadID = Null 'Val(txtCrHeadCode.Tag)
        intSubLedgerTypeID = val(txtSubLedgerType.Tag)
        intPayToSubLedgerID = val(txtName.Tag)
        intSubsidiaryCashBookID = val(txtSubsidiaryCash.Tag)
        intImplementingOfficerID = val(txtImplementingOfficer.Tag)
        numProjectNo = val(txtProjectNo.Tag)
        intStockRegisterID = Null 'Val(txtStockRegister.Tag)
        vchStockRefNo = Null 'Trim(txtReferenceNo.Text)
        intAssetTypeID = Null 'Val(txtAssetType.Tag)
        intAssetID = Null 'Val(txtAsset.Tag)
        numFwdSeatID = Null 'Val(txtForwardSeat.Tag)
        intLocalBodyID = gbLocalBodyID
        intZonalID = Null
        intFinancialYearID = gbFinancialYearID
        numUserID = gbUserID
        numSeatID = gbSeatID
        numApprovingOfficerID = Null
        numApprovingSeatID = Null
        dtApprovingDate = Null
        intVoucherID = Null
        intVoucherNo = Null
        dtVoucherDate = Null
        tnyStatus = 0
        tnyCancelled = 0
        '----------------------------------------------------------------------'
        '                       Saving Pay Order                               '
        '----------------------------------------------------------------------'
        aryIn = Array(intPayOrderID, vchPayOrderNo, dtPayOrderDate, dtDueDate, _
                        intFunctionaryID, intFunctionID, intTransactionTypeID, _
                        vchBillNo, numBillAmount, dtBillDate, _
                        intInstrumentTypeID, intCashOrBankHeadID, intSubLedgerTypeID, _
                        intPayToSubLedgerID, intSubsidiaryCashBookID, intImplementingOfficerID, _
                        numProjectNo, intStockRegisterID, vchStockRefNo, _
                        intAssetTypeID, intAssetID, numFwdSeatID, _
                        intLocalBodyID, intZonalID, intFinancialYearID, _
                        numUserID, numSeatID, numApprovingOfficerID, _
                        numApprovingSeatID, dtApprovingDate, intVoucherID, _
                        intVoucherNo, dtVoucherDate, tnyStatus, tnyCancelled)
        objdb.ExecuteSP "spSavePayOrder", aryIn, aryOut, True, mCnn, adCmdStoredProc
        If IsNumeric(aryOut(0, 0)) Then
            intPayOrderID = aryOut(0, 0)
            txtPayOrder.Text = aryOut(1, 0)
        Else
            GoTo err:
        End If
        
        '----------------------------------------------------------------------'
        '                       Saving Pay Order Child                         '
        '----------------------------------------------------------------------'
        If txtDrHeadCode.Text <> "" Then
            objAccounts.SetAccountCode (Trim(txtDrHeadCode.Text))
            intAccountHeadID = objAccounts.AccountHeadID
            aryIn = Array(intPayOrderID, _
                            1, _
                            intAccountHeadID, _
                            Trim(txtDrHeadCode.Text), _
                            val(txtDrAmount), _
                            1, _
                            1)
            objdb.ExecuteSP "spSavePayOrderChild", aryIn, , , mCnn, adCmdStoredProc
        End If
        
        If vsGrid.Enabled = True Then
            For mLoop = 1 To vsGrid.Rows - 1
                If vsGrid.TextMatrix(mLoop, 1) = "" Then Exit For
                objAccounts.SetAccountCode (Trim(vsGrid.TextMatrix(mLoop, 1)))
                intAccountHeadID = objAccounts.AccountHeadID
                aryIn = Array(intPayOrderID, _
                                mLoop + 1, _
                                intAccountHeadID, _
                                Trim(vsGrid.TextMatrix(mLoop, 1)), _
                                val(vsGrid.TextMatrix(mLoop, 3)), _
                                2, _
                                0)
                objdb.ExecuteSP "spSavePayOrderChild", aryIn, , , mCnn, adCmdStoredProc
            Next
        End If
        
        If txtCrHeadCode.Text <> "" Then
            objAccounts.SetAccountCode (Trim(txtCrHeadCode.Text))
            intAccountHeadID = objAccounts.AccountHeadID
            aryIn = Array(intPayOrderID, _
                            IIf(vsGrid.Enabled = False, 1, mLoop + 1), _
                            intAccountHeadID, _
                            Trim(txtCrHeadCode.Text), _
                            val(txtCrAmount), _
                            3, _
                            0)
            objdb.ExecuteSP "spSavePayOrderChild", aryIn, , , mCnn, adCmdStoredProc
        End If
        '----------------------------------------------------------------------'
        '                       Saving Pay Order Address                       '
        '----------------------------------------------------------------------'
        intSubsidiaryAccountHeadID = val(txtName.Tag)
        intSubLegerTypeID = val(txtSubLedgerType.Tag)
        vchSubLedgerCode = Null 'Val(txtName.Tag)
        vchName = Trim(txtName.Text)
        vchHouseName = Trim(txtHouse.Text)
        vchStreet = Trim(txtStreet.Text)
        vchLocalPlace = Trim(txtLocalPlace.Text)
        vchMainPlace = Trim(txtMainPlace.Text)
        vchPost = Trim(txtPost.Text)
        vchPinCode = Trim(txtPin.Text)
        vchPhone = Trim(txtPhone.Text)

        aryIn = Array(intPayOrderID, _
                        intSubLegerTypeID, _
                        vchSubLedgerCode, _
                        vchName, _
                        vchHouseName, _
                        vchStreet, _
                        vchLocalPlace, _
                        vchMainPlace, _
                        vchPost, _
                        vchPinCode, _
                        vchPhone)
        objdb.ExecuteSP "spPayOrderAddress", aryIn, , , mCnn, adCmdStoredProc
        mCnn.CommitTrans
        MsgBox "Pay Order Sent for Verification and Approval", vbInformation
        Exit Sub
    Else
        MsgBox "Invalid Connection to Saankhya DataBase, Please Contact your System Support", vbInformation
        GoTo err:
    End If
err:
    MsgBox (err.Number)
    If Rec.State = 1 Then Rec.Close
    If mCnn.State = 1 Then
    mCnn.RollbackTrans
    mCnn.Close
    End If
End Sub

Private Function LockForm(Optional mLock As Boolean = False)
    
    txtCrAccountHead.Enabled = mLock
    txtCrAmount.Enabled = mLock
    txtCrHeadCode.Enabled = mLock
    txtDated.Enabled = mLock
    txtDrAccountHead.Enabled = mLock
    txtDrAmount.Enabled = mLock
    txtDrAmount.Enabled = mLock
    txtDrHeadCode.Enabled = mLock
    txtDueDate.Enabled = mLock
    
    txtFunction.Enabled = mLock
    txtFunctionary.Enabled = mLock
    
    If val(txtTransactionType.Tag) = gbTransactionTypePayBills Then
        txtHouse.Enabled = mLock
        txtImplementingOfficer.Visible = mLock
        txtInit1.Enabled = mLock
        txtInit2.Enabled = mLock
        txtInit3.Enabled = mLock
        txtInit4.Enabled = mLock
        txtLocalPlace.Enabled = mLock
        txtMainPlace.Enabled = mLock
        
        txtName.Enabled = mLock
        txtNarration.Enabled = mLock
        txtPayOrder.Enabled = mLock
        txtPhone.Enabled = mLock
        txtPin.Enabled = mLock
        txtPost.Enabled = mLock
        txtProjectNo.Visible = mLock
        
        txtStreet.Enabled = mLock
        txtSubLedgerType.Enabled = mLock
        txtSubsidiaryCash.Enabled = mLock
        
        cmdAsset.Enabled = mLock
        cmdDrAccountHead.Enabled = mLock
        
        cmdImplementingOfficer.Enabled = mLock
        cmdProjectNo.Enabled = mLock
        
        cmdSearchFunction.Enabled = mLock
        cmdSearchFunctionary.Enabled = mLock
        cmdSearchName.Enabled = mLock
        
        cmdSubLederType.Enabled = mLock
        cmdSubsidiaryCash.Enabled = mLock
        cmdCrAccountHead.Enabled = mLock
        
        vsGrid.Editable = flexEDNone
        dtpDueDate.Enabled = mLock
    
        End If
        'txtTransactionType.Enabled = mLock

End Function
Private Function LockPayOrder()
    
    txtCrAccountHead.Enabled = False
    txtCrAmount.Enabled = False
    txtCrHeadCode.Enabled = False
    txtDated.Enabled = False
    txtDrAccountHead.Enabled = False
    txtDrAmount.Enabled = False
    txtDrAmount.Enabled = False
    txtDrHeadCode.Enabled = False
    txtDueDate.Enabled = False
    
    txtFunction.Enabled = False
    txtFunctionary.Enabled = False
    
    txtHouse.Enabled = False
    txtImplementingOfficer.Enabled = False
    txtInit1.Enabled = False
    txtInit2.Enabled = False
    txtInit3.Enabled = False
    txtInit4.Enabled = False
    txtLocalPlace.Enabled = False
    txtMainPlace.Enabled = False
    txtName.Enabled = False
    txtNarration.Enabled = False
    txtPayOrder.Enabled = False
    txtPhone.Enabled = False
    txtPin.Enabled = False
    txtPost.Enabled = False
    txtProjectNo.Enabled = False
    
    
    txtStreet.Enabled = False
    txtSubLedgerType.Enabled = False
    txtSubsidiaryCash.Enabled = False
    
    cmdAsset.Enabled = False
    cmdDrAccountHead.Enabled = False
    
    cmdImplementingOfficer.Enabled = False
    cmdProjectNo.Enabled = False
    
    cmdSearchFunction.Enabled = False
    cmdSearchFunctionary.Enabled = False
    cmdSearchName.Enabled = False
    
    cmdSubLederType.Enabled = False
    cmdSubsidiaryCash.Enabled = False
    cmdCrAccountHead.Enabled = False
    
    vsGrid.Editable = flexEDNone
    
    
    dtpDueDate.Enabled = False
    
    
    'txtTransactionType.Enabled = False
End Function

    Private Sub LockProjectType(mFlag As Boolean)
         ''  Modified For SaankhyaWeb Updation
        If gbFinancialYearID > 2016 And gbSaankhyaWeb = 1 Then
            lblAllotmentLetterNo.Visible = False
            txtAllotmentLetterNo.Visible = False
            cmdAllotmentLetterNo.Visible = False
        End If
        
        cmdSearchTransactionType.Enabled = mFlag
        cmdSearchFunctionary.Enabled = mFlag
        cmdSearchFunction.Enabled = mFlag
        
        cmdDrAccountHead.Enabled = mFlag
        cmdImplementingOfficer.Enabled = mFlag
        cmdProjectNo.Enabled = mFlag
        cmdSourceOfFund.Enabled = mFlag
    End Sub

    Public Function FillPayOrder(intPayOrderID As Variant)
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim objAc As New clsAccounts
            Dim mLoopCount As Integer
            Dim objProject As New clsProject
            Dim objAllotment As New clsAllotmentLetter
            
            Call GetPendingTaskDetails
            If objdb.SetConnection(mCnn) Then
                '''mSQL = "Select *, CashBook.vchTitle as SubCashBook, ImpOfficer.vchFunctionary as ImplOfficer, faPayOrderChild.intAccountHeadID as HeadID "
                '''mSQL = mSQL + " , faPayOrder.intFunctionID as FnID, faPayOrder.intFunctionaryID as FryID, suSourceOfFund.vchSourceFundName as SourceOfFund "
                mSql = "Select faPayOrder.*, faPayOrderChild.*, faPayOrderAddress.*, faFunctions.vchFunction, faFunctionaries.vchFunctionary, "
                mSql = mSql + " faTransactionType.vchTransactionType,  suSourceOfFund.vchSourceFundName, faSeats.chvSeatTitle, faSubLedgerTypes.vchSubLedgerType, " & vbNewLine
                mSql = mSql + " Payee.vchTitle as SubAccHeadTitle, Payee.vchName as PayeeName, CashBook.vchTitle as SubCashBook, faAllotmentLetters.vchAllotmentNo, " & vbNewLine
                mSql = mSql + " suProjectDetails.chvProjectSlNo, faTransactionCategory.vchTransactionCategory, ImpOfficer.vchFunctionary as ImplOfficer, " & vbNewLine
                mSql = mSql + " faPayOrderChild.intAccountHeadID as HeadID, faPayOrder.intFunctionID as FnID, faPayOrder.intFunctionaryID as FryID, CashBook.vchTitle as SubCashBook, " & vbNewLine
                mSql = mSql + " suSourceOfFund.vchSourceFundName as SourceOfFund, suProjectDetails.decProjectID, faTransactionCategory.intCategoryID, faPayOrder.vchDescription as [Desc],faAgreements.*,faPayOrder.tnyStatus Status "
                mSql = mSql + " ,suGOForFunds.intRefID GOID,suGOForFunds.vchRefNo GONo,Imp.vchTitle IMPO from faPayOrder "
                mSql = mSql + " Inner join faPayOrderAddress On faPayOrder.intPayOrderID = faPayOrderAddress.intPayOrderID " & vbNewLine
                mSql = mSql + " Inner Join faPayOrderChild On faPayOrder.intPayOrderID = faPayOrderChild.intPayOrderID " & vbNewLine
                mSql = mSql + " Inner Join faFunctions On faPayOrder.intFunctionID = faFunctions.intFunctionID " & vbNewLine
                mSql = mSql + " Inner Join faFunctionaries On faPayOrder.intFunctionaryID = faFunctionaries.intFunctionaryID " & vbNewLine
                mSql = mSql + " Inner Join faTransactionType On faPayOrder.intTransactionTypeID = faTransactionType.intTransactionTypeID " & vbNewLine
                mSql = mSql + " Left Join faFunctionaries ImpOfficer On faPayOrder.intFunctionaryID = ImpOfficer.intFunctionaryID " & vbNewLine
                mSql = mSql + " Left join suSourceOfFund On faPayOrder.intSourceOfFundID = suSourceOfFund.intSourceFundID " & vbNewLine
                mSql = mSql + " Left Join faSeats On faPayOrder.numFwdSeatID = faSeats.numSeatID " & vbNewLine
                mSql = mSql + " Left Join faSubLedgerTypes On faPayOrder.intSubLedgerTypeID = faSubLedgerTypes.intSubLedgerTypeID " & vbNewLine
                mSql = mSql + " Left Join faSubSidiaryAccountHeads Payee On faPayOrder.intPayToSubLedgerID = Payee.intSubsidiaryAccountHeadID " & vbNewLine
                mSql = mSql + " Left Join faSubSidiaryAccountHeads CashBook On faPayOrder.intSubsidiaryCashBookID = CashBook.intSubsidiaryAccountHeadID " & vbNewLine
                mSql = mSql + " Left Join faAllotmentLetters On faAllotmentLetters.intAllotmentID = faPayOrder.intAllotmentID " & vbNewLine
                mSql = mSql + " Left Join suProjectDetails On suProjectDetails.decProjectID = faPayOrder.numProjectNo  AND suProjectDetails.intYearID=faPayOrder.intFinancialYearID " & vbNewLine
                mSql = mSql + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID " & vbNewLine
                mSql = mSql + " Left Join faAgreements On faAgreements.intAgreementID=faPayOrder.intAgreementID" & vbNewLine
                mSql = mSql + " Left Join suGOForFunds On suGOForFunds.intPayOrderID=faPayOrder.intPayOrderID" & vbNewLine
                 
                mSql = mSql + " Left Join faSubSidiaryAccountHeads Imp On faPayOrder.intImplementingOfficerID = Imp.intSubsidiaryAccountHeadID" & vbNewLine
             
                mSql = mSql + " Where faPayOrder.intPayOrderID = " & intPayOrderID
                Rec.Open mSql, mCnn
                
                If Not (Rec.EOF Or Rec.BOF) Then
                
                    ' BLOCK PREVIOUS YEARS PAYMENT ORDERS
                    If mPendingTask <> 0 Then
                        If Rec!dtPayOrderDate < DateAdd("yyyy", -1, gbStartingDate) Or Rec!dtPayOrderDate > DateAdd("yyyy", -1, gbEndingDate) Then
                                MsgBox "This seems to be previous years Payment Order, plz verify", vbInformation
                                Exit Function
                            End If
                    Else
                        If Rec!dtPayOrderDate < gbStartingDate Or Rec!dtPayOrderDate > gbEndingDate Then
                            MsgBox "This seems to be previous years Payment Order, plz verify", vbInformation
                            Exit Function
                        End If
                    End If
                    ' -------------------------------------------- '
                    txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    If val(txtTransactionType.Tag) = gbTransactionTypeUnUtilizedAmount Then
                        cmdGo.Enabled = True
                        txtGo.Text = IIf(IsNull(Rec!GoNo), "", Rec!GoNo)
                        txtGo.Tag = IIf(IsNull(Rec!Goid), -1, Rec!Goid)
                    End If
                    txtPayOrder.Text = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                    txtPayOrder.Tag = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                    txtDated.Text = IIf(IsNull(Rec!dtPayOrderDate), "", CheckDateInMMM(Rec!dtPayOrderDate))
                    txtDueDate.Text = IIf(IsNull(Rec!dtDueDate), "", CheckDateInMMM(Rec!dtDueDate))
                    If mPendingTask = 1 Then
                        txtDueDate.Text = CheckDateInMMM(txtDueDate.Text)
                        txtDueDate.Enabled = False
                        dtpDueDate.Enabled = False
                    ElseIf mPendingTask = 2 Or mPendingTask = 3 Then
                        txtDueDate.Enabled = False
                        dtpDueDate.Enabled = False
                    Else
                        txtDueDate.Text = IIf(IsNull(Rec!dtDueDate), "", CheckDateInMMM(Rec!dtDueDate))
                    End If
                    txtFunction.Text = IIf(IsNull(Rec!vchFunction), "", Rec!vchFunction)
                    txtFunction.Tag = IIf(IsNull(Rec!FnID), "", Rec!FnID)
                    txtFunctionary.Text = IIf(IsNull(Rec!vchFunctionary), "", Rec!vchFunctionary)
                    txtFunctionary.Tag = IIf(IsNull(Rec!FryID), "", Rec!FryID)
                    
                    txtSubLedgerType.Text = IIf(IsNull(Rec!vchSubLedgerType), "", Rec!vchSubLedgerType)
                    txtSubLedgerType.Tag = IIf(IsNull(Rec!intSubLedgerTypeID), "", Rec!intSubLedgerTypeID)
                    
                    'txtPayeeType.Text = IIf(IsNull(Rec!intSubLedgerTypeID), "", Rec!intSubLedgerTypeID)
                                        
                    txtName.Tag = IIf(IsNull(Rec!intPayToSubLedgerID), "", Rec!intPayToSubLedgerID)
                    txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
                    txtInit1.Text = IIf(IsNull(Rec!vchInit1), "", Rec!vchInit1)
                    txtInit2.Text = IIf(IsNull(Rec!vchInit2), "", Rec!vchInit2)
                    txtInit3.Text = IIf(IsNull(Rec!vchInit3), "", Rec!vchInit3)
                    txtInit4.Text = IIf(IsNull(Rec!vchInit4), "", Rec!vchInit4)
                    txtHouse.Text = IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
                    txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
                    txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
                    txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
                    txtPost.Text = IIf(IsNull(Rec!vchPost), "", Rec!vchPost)
                    txtPin.Text = IIf(IsNull(Rec!vchPinCode), "", Rec!vchPinCode)
                    txtPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
                    
'                    txtSubsidiaryCash.Tag = IIf(IsNull(Rec!intSubsidiaryCashBookID), "", Rec!intSubsidiaryCashBookID)
'                    txtSubsidiaryCash.Text = IIf(IsNull(Rec!SubCashBook), "", Rec!SubCashBook)
                    
                    txtSourceofFund.Text = IIf(IsNull(Rec!SourceOfFund), "", Rec!SourceOfFund)
                    txtSourceofFund.Tag = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
                    
                    txtImplementingOfficer.Text = IIf(IsNull(Rec!IMPO), "", Rec!IMPO)
                    txtImplementingOfficer.Tag = IIf(IsNull(Rec!intImplementingOfficerID), "", Rec!intImplementingOfficerID)
                    
                    txtAllotmentLetterNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    txtAllotmentLetterNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                    Call txtAllotmentLetterNo_GotFocus

                        
                    objAllotment.SetAllotment (val(txtAllotmentLetterNo.Tag))
                    txtAllotedAmt.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)


                    txtAgreementNo.Text = IIf(IsNull(Rec!vchAgreementNo), "", Rec!vchAgreementNo)
                    txtAgreementNo.Tag = IIf(IsNull(Rec!intAgreementID), "", Rec!intAgreementID)
                    
                    txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                    txtProjectNo.Tag = IIf(IsNull(Rec!decProjectID), "", Rec!decProjectID)
                    
                    If val(txtProjectNo.Tag) <> 0 Then
                        If mPendingTask = 1 Or mPendingTask = 2 Or mPendingTask = 3 Then
                            objProject.SetProject txtProjectNo.Tag, gbFinancialYearID - 1
                        Else
                            objProject.SetProject (txtProjectNo.Tag)
                        End If
                        'txtCategory.Text =  IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                        txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
                        txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
                        txtSector = IIf(IsNull(objProject.Sector), "", objProject.Sector)
                        txtSector.Tag = IIf(IsNull(objProject.SectorTypeID), "", objProject.SectorTypeID)
                    End If
                    txtNarration.Text = IIf(IsNull(Rec!Desc), "", Rec!Desc)
                    
                    txtForward2Seat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                    txtForward2Seat.Tag = IIf(IsNull(Rec!numFwdSeatID), "", Rec!numFwdSeatID)
                    
                    
                    intKeyID = IIf(IsNull(Rec!intKeyID), Null, Rec!intKeyID)
                    intModuleID = IIf(IsNull(Rec!intModuleID), Null, Rec!intModuleID)
                    dtKeyDate = IIf(IsNull(Rec!dtKeyDate), Null, Rec!dtKeyDate)
                    'cmdSave.Enabled = True
                    If CDate(dtKeyDate) <> CDate(gbTransactionDate) And Rec!Status = 1 Then '***********MODIFIED BY sabeen*******************
                        cmdSave.Enabled = False
                        cmdVerify.Enabled = False
                    Else
                        cmdSave.Enabled = True
                       
                    End If
                    
                    
                    If gbUserID <> Rec!numUserID Then
                        lblTipText.Caption = "Not an authorised user for this payorder"
                        cmdSave.Enabled = False
                    End If
                    While Not Rec.EOF
                        If Rec!intSlNo = 1 And Rec!tnyCategoryFlag = 1 Then
                            objAc.SetAccountID Rec!HeadID
                            If objAc.AccountHeadID > 0 Then
                                txtDrHeadCode.Text = objAc.AccountCode
                                txtDrHeadCode.Tag = objAc.AccountHeadID
                                txtDrAccountHead.Text = objAc.AccountHead
                                txtDrAmount.Text = Format(Rec!numAmount, "0.00")
                            Else
                                MsgBox "Error: Head Not Found", vbInformation
                            End If
                        End If
            
                        If Rec!tnyCategoryFlag = 2 Then
                            objAc.SetAccountID Rec!HeadID
                            If objAc.AccountHeadID > 0 Then
                                mLoopCount = mLoopCount + 1
                                vsGrid.Cell(flexcpText, mLoopCount, 0) = mLoopCount
                                vsGrid.Cell(flexcpText, mLoopCount, 1) = objAc.AccountCode
                                vsGrid.Cell(flexcpText, mLoopCount, 2) = objAc.AccountHead
                                vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
                                vsGrid.Cell(flexcpText, mLoopCount, 4) = objAc.AccountHeadID
                            End If
                        End If
            
                        If Rec!tnyCategoryFlag = 3 Then
                            objAc.SetAccountID Rec!HeadID
                            If objAc.AccountHeadID > 0 Then
                                txtCrHeadCode.Text = objAc.AccountCode
                                txtCrHeadCode.Tag = objAc.AccountHeadID
                                txtCrAccountHead.Text = objAc.AccountHead
                                txtCrAmount.Text = Format(Rec!numAmount, "0.00")
                            Else
                                MsgBox "Error: Head Not Found", vbInformation
                            End If
                        End If
            
                        If Rec!tnyCategoryFlag = 5 Then
'                            chkPension.Visible = True
'                            chkPension.value = vbChecked
                            If Rec!Status = 1 Then
                                txtPensionAmt.Enabled = False
                            Else
                                txtPensionAmt.Enabled = True
                            End If
                            lblPension.Visible = True
                            txtPensionAmt.Visible = True
                            txtPensionAmt.Text = Format(Rec!numAmount, "0.00")
                        End If
                        If Rec!tnyCategoryFlag = 6 Then
'                            txtCP.Visible = True
'                            lblCP.Visible = True
                            If Rec!Status = 1 Then
                                txtCP.Enabled = False
                            Else
                                txtCP.Enabled = True
                            End If
                            lblCP.Visible = True
                            txtCP.Visible = True
                            txtCP.Text = Format(Rec!numAmount, "0.00")
                        End If
                        Rec.MoveNext
                    Wend
                    'Call SetSourceOfFund
                    'Call ProjectLink(False)
                    txtDrAmount_LostFocus
                End If
                If gbLBPanchayat Then
                    Call UserPrivillage(intPayOrderID)
                End If
                If mPendingTask = 1 Or mPendingTask = 2 Then
                    GetPendingTaskDetails
                End If
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact Your System Administrator", vbInformation
            End If
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Sub UserPrivillage(intPayOrderID As Variant)
        Dim mSql As String
        Dim mStatus     As Integer
        Dim objdb       As New clsDB
        Dim mCn        As New ADODB.Connection
        Dim Rec1         As New ADODB.Recordset
        If gbLBPanchayat Then
            objdb.SetConnection mCn
            mSql = "Select tnyStatus,tnyCancelled From faPayOrder Where intPayOrderID=" & intPayOrderID
            Rec1.Open mSql, mCn
            If Not (Rec1.BOF And Rec1.EOF) Then
                mStatus = Rec1!tnyStatus
            Else
            End If
            If gbSeatGroupID = gbSeatGroupAccountSectionClerk Then
                cmdApproval.Visible = False
                cmdReject.Visible = False
                cmdVerify.Visible = False
                cmdSave.Caption = "Edit"
                cmdSeat.Enabled = False
            ElseIf gbSeatGroupID = gbSeatGroupAccountsClerk Then
                cmdApproval.Visible = False
                cmdReject.Visible = False
                cmdVerify.Visible = False
                cmdVerify.Caption = "Forward"
                cmdSave.Caption = "Edit/Fwd"
            ElseIf gbSeatGroupID = gbSeatGroupHeadClerk Or gbSeatGroupID = gbSeatGroupAssistantSecretary Then
                cmdApproval.Visible = False
                cmdReject.Visible = True
                cmdVerify.Visible = True
                cmdSave.Visible = False
                txtDrAmount.Enabled = False
                txtCrAmount.Enabled = False
                cmdSeat.Enabled = False
                cmdSearchTransactionType.Enabled = False
            ElseIf gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                txtDrAmount.Enabled = False
                txtCrAmount.Enabled = False
                cmdSeat.Enabled = False
                cmdSearchTransactionType.Enabled = False
                If mStatus = 3 Then
                    cmdApproval.Visible = True
                    cmdReject.Visible = True
                    cmdSave.Visible = False
                Else
                    cmdApproval.Visible = False
                    cmdReject.Visible = True
                    cmdVerify.Visible = True
                    cmdSave.Visible = False
                End If
            End If
            mCn.Close
        End If
    End Sub
    Private Function MakePayable(mPaymentOrderNo As Double) As Boolean

        Dim mVoucher            As uVoucher
        Dim mVouChildTbl        As uVChild
        
        Dim mTranTable          As uTr
        Dim mTranChildTbl       As uTrChild
        
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mCommonDescription  As String
        
        Dim objdb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim RecChild            As New ADODB.Recordset
        Dim mCnn                As New ADODB.Connection
        Dim mSql                As String
        
        Dim mCrAmt          As Double
        Dim mGrossAmt       As Double
        Dim mNetAmt         As Double
        
        Dim mSlNo           As Integer
        Dim mDrHeadCode     As String
        Dim mDrHeadID       As Integer
        Dim mCrHeadID       As Integer
        Dim objAc           As New clsAccounts
        
        '----------------------------------------------------------------------------- '
        ' Opening PaymentOrder Table And Child Tables
        '----------------------------------------------------------------------------- '
        MakePayable = False
        objdb.SetConnection mCnn
        mSql = "Select * From faPayOrder Where vchPayOrderNo = " & mPaymentOrderNo
        Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
        If Rec.BOF And Rec.EOF Then
            MsgBox "No Pay Order Found For Generate Pay Order", vbInformation
            Exit Function
        Else
            '-- ------------------------------------------------------------ --'
            '   Transaction Type : Pay and Allowance
            '   Modified on 15-Jan-2010
            '-- ------------------------------------------------------------ --'
            If Rec!intTransactionTypeID = gbTransactionTypePayBills Then
                If gbLBType = 3 Or gbLBType = 4 Then    'FOR COPERATION AND MUNICIPALITY
                    Call GeneratePayBillJournals(Rec!vchPayOrderNo, mPendingTask)
                Else
                   
                    Call GeneratePayBillJournalsForPanchayat(Rec!vchPayOrderNo, mPendingTask)  'FOR PANCHAYAT   ADDED BY MINU FOR PANCHAYAT ON 11-01-2011
                    
                End If
                MakePayable = True
                Rec.Close
                Exit Function
            End If
            '-- ------------------------------------------------------------ --'
            
            mSql = "Select * From faPayOrderChild Where intPayOrderID = " & Rec!intPayOrderID
            RecChild.Open mSql, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
            If RecChild.BOF And RecChild.EOF Then
                MsgBox "Payment Order Details not found for this Pay Order", vbInformation
                Exit Function
            End If
        End If
        
        '-------------------------------------------------------------------------------'
        ' Note:- CHECK WHERE JOURNAL IS REQUIERD OR NOT                                 '
        '     :  If HeadID with tnyCategryFlag = 1 and HeadID with tnyCagegoryFlag = 3  '
        '     :  Then No Journal is required.                                           '
        '-------------------------------------------------------------------------------'
        RecChild.MoveFirst
        mCrAmt = 0
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 1 Then
                mDrHeadID = RecChild!intAccountHeadID ' Gross Account Head
            End If
            If RecChild!tnyCategoryFlag = 2 Then
                'mCrAmt = mCrAmt + RecChild!numAmount ' Total Deduction Amount
            End If
            
            If RecChild!tnyCategoryFlag = 3 Then
                mCrHeadID = RecChild!intAccountHeadID 'Net Payable Head
            End If
            RecChild.MoveNext
        Wend
        If mDrHeadID = mCrHeadID Then
            'No Journal will be generated Here Because
            GoTo ApprovePayOrder:
        End If
        
        
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 1 Then
                mDrHeadID = RecChild!intAccountHeadID 'Rec!intCashOrBankHeadID
                GoTo GrossHead
            End If
            RecChild.MoveNext
        Wend
        GoTo ErrNoGr:
        
GrossHead:
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = RecChild!numAmount
             mGrossAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = Rec!vchDescription
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = mDrHeadID  'gbAcHeadIDNetSalaryPayable  'Debit to Net Salary Payable
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 60 'PaymentOrder-Saankhya Module
            If mPendingTask = 0 Then
                .intFinancialYearID_26 = gbFinancialYearID
            Else
                .intFinancialYearID_26 = gbFinancialYearID - 1
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        If mCnn.State Then
            'mCnn.Close
        End If
        objdb.SetConnection mCnn
        'mCnn.BeginTrans
        'On Error GoTo ErrRollBack:
        objdb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            GoTo ErrRollBack:
        End If
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            If mPendingTask = 0 Then
                .intFinancialYearID = gbFinancialYearID
            Else
                .intFinancialYearID = gbFinancialYearID - 1
            End If
            .dtTransactionDate = Rec!dtDueDate
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = Rec!vchDescription
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
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
        
        End With
        
        objdb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        '----------------------------------------------------------------- '
        '
        '----------------------------------------------------------------- '
        '                                                                  '
        '----------------------------------------------------------------- '
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 3 Then
                mNetAmt = RecChild!numAmount
                GoTo TryDeductions:
            End If
            RecChild.MoveNext
        Wend
TryDeductions:

        'Note:-Gross Salary Payable A/c Debtor
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = mDrHeadID
            .fltAmount = mGrossAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
        End With
        
        mSlNo = 1
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 2 Or RecChild!tnyCategoryFlag = 3 Then
                    mSlNo = mSlNo
                    'Note:- Deduction Heads and Net Salary to Voucher Child
                    With mVouChildTbl
                        .intVoucherID_1 = mintVoucherID
                        .intLocalBodyID_2 = gbLocalBodyID
                        .intSlNo_3 = mSlNo
                        .intAccountHeadID_4 = RecChild!intAccountHeadID
                        .tnyDebitOrCredit_5 = 0
                        If IsDate(Rec!dtKeyDate) Then
                            .intYearID_6 = Year(Rec!dtKeyDate)
                            .tnyPeriodID_7 = Month(Rec!dtKeyDate)
                        End If
                        .tnyArrearFlag_8 = Null
                        .numDemandID_9 = Rec!intKeyID
                        .fltAmount_10 = RecChild!numAmount
                    
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
                        objdb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
                    End With
                    
                    mSlNo = mSlNo + 1
                    With mTranChildTbl
                        .intTransactionID = mintTransactionID
                        .intSerialNo = mSlNo + 1
                        .intAccountHeadID = RecChild!intAccountHeadID
                        .fltAmount = RecChild!numAmount
                        .tinDebitOrCreditFlag = 0
                        .intByAccountHeadID = mDrHeadID
                        .vchNarration = RecChild!vchDescription
                        .intFundID = gbFundID
                        
                        mCrAmt = mCrAmt + RecChild!numAmount
                        
                        arrInput = Array(.intTransactionID, _
                        .intSerialNo, _
                        .intAccountHeadID, _
                        .fltAmount, _
                        .tinDebitOrCreditFlag, _
                        .intByAccountHeadID, _
                        .vchNarration, _
                        .intFundID)
                        
                        objdb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
                    End With
            End If
            RecChild.MoveNext
        Wend
        
        If mGrossAmt <> mCrAmt Then
            GoTo ErrNoAmountNotTally:
        End If
ApprovePayOrder:
        'Note:- Changing the Status of PayOrder as Approved!
        'mSQL = "Update faPayOrder Set tnyStatus = 1, numApprovingOfficerID = " & gbUserID & ", numApprovingSeatID = " & gbSeatID & ", dtApprovingDate = " & gbTransactionDate & " Where vchPayOrderNo = " & mPaymentOrderNo
        'objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
        'mCnn.CommitTrans
        MakePayable = True
        
CleanUp:
        Exit Function
ErrNoAmountNotTally:
        MsgBox "Debit and Credit Amounts are not tally!", vbInformation
        GoTo ErrRollBack:
ErrNoHeadNotFound:
        MsgBox "Account Head not found", vbInformation
        GoTo ErrRollBack:
ErrNoGr:
    MsgBox "No Gross Amount Found to Proceed!", vbInformation
    GoTo ErrRollBack:

ErrRollBack:
    'mCnn.RollbackTrans
    Set mCnn = Nothing
    


End Function



    Private Sub SetExpenditureDetails(intSubSecID As Integer, intCatID As Integer, Optional intMicroSecID As Integer, Optional intYearID As Integer)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSql As String
        Dim objTrType As New clsTransactionType
        
        If intYearID <> gbFinancialYearID Then
            mSql = " SELECT faSubSectorHeads.intSubsectorID, faSubSectorHeads.vchSubSectorCode, faSubSectorHeads.intCategoryID, faSubSectorHeads.intAccountHeadID,"
            mSql = mSql + "     faAccountHeads.vchAccountHeadCode, vchAccountHead, faSubSectorHeads.intFunctionID, faFunctions.vchFunctionCode, vchFunction,"
            mSql = mSql + "     faSubSectorHeads.intTransactionTypeID , vchTransactionType, faFunctionaryFunctions.intFunctionaryID, vchFunctionary"
            mSql = mSql + " From faSubSectorHeads"
            mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faSubSectorHeads.intAccountHeadID"
            mSql = mSql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faSubSectorHeads.intFunctionID"
            mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faSubSectorHeads.intTransactionTypeID"
            mSql = mSql + " INNER JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
            mSql = mSql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID"
            mSql = mSql + " Where faSubSectorHeads.intSubSectorID = " & intSubSecID & " And faSubSectorHeads.intCategoryID = " & intCatID
        Else
            mSql = "        SELECT faMicroSectorHeads.intSubsectorID, faSubSector.vchSubSecCode, faMicroSectorHeads.intCategoryID, faMicroSectorHeads.intAccountHeadID,"
            mSql = mSql + "     faAccountHeads.vchAccountHeadCode, vchAccountHead, faMicroSectorHeads.intFunctionID, faFunctions.vchFunctionCode, vchFunction,"
            mSql = mSql + "     faMicroSectorHeads.intTransactionTypeID , vchTransactionType, faFunctionaryFunctions.intFunctionaryID, vchFunctionary"
            mSql = mSql + " From faMicroSectorHeads"
            mSql = mSql + " INNER JOIN faSubSector ON faSubSector.intSubSecID = faMicroSectorHeads.intSubSectorID"
            mSql = mSql + " INNER JOIN faAccountHeads ON faAccountHeads.intAccountHeadID = faMicroSectorHeads.intAccountHeadID"
            mSql = mSql + " INNER JOIN faFunctions ON faFunctions.intFunctionID = faMicroSectorHeads.intFunctionID"
            mSql = mSql + " INNER JOIN faTransactionType ON faTransactionType.intTransactionTypeID = faMicroSectorHeads.intTransactionTypeID"
            mSql = mSql + " INNER JOIN faFunctionaryFunctions ON faFunctionaryFunctions.intFunctionID = faFunctions.intFunctionID"
            mSql = mSql + " INNER JOIN faFunctionaries ON faFunctionaries.intFunctionaryID = faFunctionaryFunctions.intFunctionaryID"
            mSql = mSql + " Where faMicroSectorHeads.intMircoSectorID = " & intMicroSecID & " And faMicroSectorHeads.intCategoryID = " & intCatID
        End If
        
        If objdb.SetConnection(mCnn) Then
            Rec.Open mSql, mCnn
            If Not (Rec.BOF And Rec.EOF) Then
            
                txtTransactionType.Text = Rec!vchTransactionType
                txtTransactionType.Tag = Rec!intTransactionTypeID
                
                txtFunction.Text = Rec!vchFunction
                txtFunction.Tag = Rec!intFunctionID
                
                txtFunctionary.Text = Rec!vchFunctionary
                txtFunctionary.Tag = Rec!intFunctionaryID
                
                Select Case Rec!intCategoryID
                    Case Is = 1: txtCategory.Text = "General"
                    Case Is = 2: txtCategory.Text = "SCP"
                    Case Is = 3: txtCategory.Text = "TSP"
                End Select
                txtCategory.Tag = Rec!intCategoryID
                
                ' 1141,1151,1161
                If val(txtDrAccountHead.Tag) = 4 Then
                Select Case Rec!intCategoryID
                    Case Is = 1:
                        objTrType.SetTransactionType (1141)
                        txtTransactionType.Text = objTrType.TransactionType
                        txtTransactionType.Tag = objTrType.TransactionTypeID
                    Case Is = 2: txtCategory.Text = "SCP"
                        objTrType.SetTransactionType (1151)
                        txtTransactionType.Text = objTrType.TransactionType
                        txtTransactionType.Tag = objTrType.TransactionTypeID
                    Case Is = 3: txtCategory.Text = "TSP"
                        objTrType.SetTransactionType (1161)
                        txtTransactionType.Text = objTrType.TransactionType
                        txtTransactionType.Tag = objTrType.TransactionTypeID
                End Select
                End If
                
                mOldRequisition = False
                
                'txtDrHeadCode.Text = Rec!vchAccountHeadCode
                'txtDrHeadCode.Tag = Rec!intAccountHeadID
                'txtDrAccountHead.Text = Rec!vchAccountHead
            Else
    '            txtTransactionType.Text = ""
    '            txtTransactionType.Tag = ""
    '
    '            txtFunction.Text = ""
    '            txtFunction.Tag = ""
    '
    '            txtFunctionary.Text = ""
    '            txtFunctionary.Tag = ""
    '
    '            txtCategory.Text = ""
    '            txtCategory.Tag = ""
                
                'txtDrHeadCode.Text = ""
                'txtDrHeadCode.Tag = ""
                'txtDrAccountHead.Text = ""
                
                mOldRequisition = True
                
            End If
        Else
        
        End If
    
    End Sub

Public Property Let GrossSalaryID(mID As Integer)
    mvarGrossSalryID = mID
End Property

Public Property Get GrossSalaryID() As Integer
    GrossSalaryID = mvarGrossSalryID
End Property

Public Property Get PayOrderID() As Variant
    PayOrderID = intPayOrderID
End Property

Public Property Let PayOrderID(mData As Variant)
    intPayOrderID = mData
End Property

Public Property Get PayOrderNo() As Variant
    PayOrderNo = vchPayOrderNo
End Property

Public Property Let PayOrderNo(mData As Variant)
    vchPayOrderNo = mData
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

Public Property Let WaterBillPOMode(Data As Boolean)
    mWaterBillPOMode = Data
End Property

Public Property Get WaterBillPOMode() As Boolean
    WaterBillPOMode = mWaterBillPOMode
End Property

Public Property Let ModuleID(Data As Integer)
    mModuleID = Data
End Property

Public Property Get ModuleID() As Integer
    ModuleID = mModuleID
End Property
Public Property Let AssetID(intAssetID As Integer) 'FOR ASSETS ON 26-05-2011
    mAssetID = intAssetID
End Property
Public Property Let AssetTypeID(intAssetTypeID As Integer) 'FOR ASSETS ON 26-05-2011
    mAssetTypeID = intAssetTypeID
End Property
''''Public Sub DisplayPayOrder(intPayOrderNo As Variant)
''''    Dim mCnn As New ADODB.Connection
''''    Dim Rec As New ADODB.Recordset
''''    Dim objDb As New clsDB
''''    Dim mSQL As String
''''    Dim objAc As New clsAccounts
''''    Dim mLoopCount As Long
''''    Dim objSubLedger As New clsSubLedger
''''    Dim objInst As New clsInstruments
''''    Dim objProj As New clsProject
''''
''''    Call FormInitialize
''''
''''    '-------------------------------------------------
''''    ' NOTE:- BLOCKED BY AIBY
''''    '       I WANNA KILL THE PERSON WHO WRITTEN THIS
''''    '-------------------------------------------------
''''    '    Dim mMenu   As Control
''''    '    For Each mMenu In frmPaymentOrder.Controls
''''    '        If TypeOf mMenu Is TextBox Then
''''    '            mMenu.Locked = True
''''    '        End If
''''    '        If TypeOf mMenu Is CommandButton Then
''''    '            mMenu.Enabled = False
''''    '        End If
''''    '        If TypeOf mMenu Is VSFlexGrid Then
''''    '            mMenu.Enabled = False
''''    '        End If
''''    '    Next
''''    '--------------------------------------------------'
''''
''''    dtpDueDate.Enabled = False
''''    txtTransactionType.Enabled = False
''''    cmdApproval.Enabled = True
''''    cmdCancel.Enabled = True
''''    mLoopCount = 0
''''
''''    mSQL = " Select faPayOrder.*,faPayOrderChild.*, faPayOrderAddress.*, faFunctionaries.vchFunctionary, faFunctions.vchFunction, "
''''    mSQL = mSQL + " faTransactionType.vchTransactionType, faInstrumentTypes.vchInstrumentType, faPayOrder.vchDescription as PODesc, chvSeatTitle From faPayOrder Inner Join "
''''    mSQL = mSQL + " faPayOrderChild On faPayOrderChild.intPayOrderID = faPayOrder.intPayOrderID Left Join "
''''    mSQL = mSQL + " faPayOrderAddress On faPayOrderAddress.intPayOrderID = faPayOrder.intPayOrderID Left Join"
''''    mSQL = mSQL + " faFunctionaries On faFunctionaries.intFunctionaryID = faPayOrder.intFunctionaryID Left Join"
''''    mSQL = mSQL + " faFunctions On faFunctions.intFunctionID = faPayOrder.intFunctionID Left Join"
''''    mSQL = mSQL + " faTransactionType On faTransactionType.intTransactionTypeID = faPayOrder.intTransactionTypeID Left Join"
''''    mSQL = mSQL + " faInstrumentTypes On faInstrumentTypes.intInstrumentTypeID = faPayOrder.intInstrumentTypeID Left Join "
''''    mSQL = mSQL + " faSeats On faPayOrder.numSeatID = faSeats.numSeatID "
''''    mSQL = mSQL + " Where faPayOrder.vchPayOrderNo = " & intPayOrderNo
''''    mSQL = mSQL + " Order by tnyCategoryFlag, intSlNo"
''''
''''    objDb.SetConnection mCnn
''''    Set Rec = objDb.ExecuteSP(mSQL, , , , mCnn, adCmdText)
''''    If Not (Rec.BOF And Rec.EOF) Then
''''        txtPayOrder.Text = Rec!vchPayOrderNo
''''        txtPayOrder.Tag = Rec!intPayOrderID
''''        txtDated.Text = DdMmmYy(gbTransactionDate)
''''        If IsDate(Rec!dtDueDate) Then
''''            txtDueDate.Text = DdMmmYy(Rec!dtDueDate)
''''        End If
''''
''''        If Not IsNull(Rec!vchFunctionary) Then
''''            txtFunctionary.Text = Rec!vchFunctionary
''''            txtFunctionary.Tag = Rec!intFunctionaryID
''''        Else
''''            txtFunctionary.Text = ""
''''            txtFunctionary.Tag = ""
''''        End If
''''
''''        If Not IsNull(Rec!vchFunction) Then
''''            txtFunction.Text = Rec!vchFunction
''''            txtFunction.Tag = Rec!intFunctionID
''''        Else
''''            txtFunction.Text = ""
''''            txtFunction.Tag = ""
''''        End If
''''
''''        If Not IsNull(Rec!vchTransactionType) Then
''''            txtTransactionType.Text = Rec!vchTransactionType
''''            txtTransactionType.Tag = Rec!intTransactionTypeID
''''        Else
''''            txtTransactionType.Tag = ""
''''            txtTransactionType.Text = ""
''''        End If
''''
''''        If Not IsNull(Rec!intSourceOfFundID) Then
''''            objProj.FindSourceOfFund (Rec!intSourceOfFundID)
''''            txtSourceOfFund.Text = objProj.SourceOfFund
''''            txtSourceOfFund.Tag = objProj.SourceOfFundID
''''        End If
''''
''''        While Not Rec.EOF
''''            If Rec!intSlNo = 1 And Rec!tnyCategoryFlag = 1 Then
''''                objAc.SetAccountID Rec!intAccountHeadID
''''                If objAc.AccountHeadID > 0 Then
''''                    txtDrHeadCode.Text = objAc.AccountCode
''''                    txtDrHeadCode.Tag = objAc.AccountHeadID
''''                    txtDrAccountHead.Text = objAc.AccountHead
''''                    txtDrAmount.Text = Format(Rec!numAmount, "0.00")
''''                Else
''''                    MsgBox "Error: Head Not Found", vbInformation
''''                End If
''''            End If
''''
''''            If Rec!tnyCategoryFlag = 2 Then
''''
''''                objAc.SetAccountID Rec!intAccountHeadID
''''                If objAc.AccountHeadID > 0 Then
''''                    mLoopCount = mLoopCount + 1
''''                    vsGrid.Cell(flexcpText, mLoopCount, 0) = mLoopCount
''''                    vsGrid.Cell(flexcpText, mLoopCount, 1) = objAc.AccountCode
''''                    vsGrid.Cell(flexcpText, mLoopCount, 2) = objAc.AccountHead
''''                    vsGrid.Cell(flexcpText, mLoopCount, 3) = Rec!numAmount
''''                End If
''''            End If
''''
''''            If Rec!tnyCategoryFlag = 3 Then
''''                objAc.SetAccountID Rec!intAccountHeadID
''''                If objAc.AccountHeadID > 0 Then
''''                    txtCrHeadCode.Text = objAc.AccountCode
''''                    txtCrHeadCode.Tag = objAc.AccountHeadID
''''                    txtCrAccountHead.Text = objAc.AccountHead
''''                    txtCrAmount.Text = Format(Rec!numAmount, "0.00")
''''                Else
''''                    MsgBox "Error: Head Not Found", vbInformation
''''                End If
''''            End If
''''
''''            Rec.MoveNext
''''        Wend
''''        '------------------------------------------------------------------'
''''        'Note:- Auto Selection of  Bank Account Heads
''''        '------------------------------------------------------------------'
''''        Rec.MoveFirst
''''        Select Case Rec!intSourceOfFundID
''''            Case Is = 1  ' Development Fund (Plan Fund)
''''
''''                objAc.SetAccountCode 450650100
''''            Case Is = 16 ' Maintenance Fund (Road)
''''                objAc.SetAccountCode 450650200
''''            Case Is = 17 ' Maintenance Fund (Non-Road)
''''                objAc.SetAccountCode 450650200
''''            Case Else    ' Default Bank
''''                Dim objBank As New clsBank
''''                objBank.SetBankInfo ReadIniFile(gbSaankhyaINI, "Receipt", "DefaultBankID")
''''                If objBank.BankID > 0 Then
''''                    objAc.SetAccountID objBank.BankAccountHeadID
''''                Else
''''                    objAc.SetAccountID 0
''''                End If
''''        End Select
''''
''''        '''If objAc.AccountHeadID > 0 Then
''''        '''    txtCrHeadCode.Text = objAc.AccountCode
''''        '''    txtCrHeadCode.Tag = objAc.AccountHeadID
''''        '''    txtCrAccountHead.Text = objAc.AccountHead
''''        '''End If
''''
''''        'Note:- Net Amount Re-Calculate From Grid and Diplayed
''''        txtCrAmount.Text = Format(CalculateAmt, "0.00")
''''
''''        'Note:- Selecting InstrumentType based on Fund (Bank )
''''        Select Case Rec!intSourceOfFundID
''''            Case 1, 16, 17
''''                objInst.SetInstrumentType 7
''''            Case Else
''''                objInst.SetInstrumentType 5
''''        End Select
'''''        If objInst.InstrumentTypeID > 0 Then
'''''            txtInstrument.Text = objInst.InstrumentType
'''''            txtInstrument.Tag = objInst.InstrumentTypeID
'''''        End If
''''        'Note:- End of selecting Instrument Type
''''        If Not IsNull(Rec!vchName) Then
''''            txtName.Text = Rec!vchName
''''            txtPayee.Text = Rec!vchName
''''        Else
''''            txtName.Text = ""
''''            txtPayee.Text = ""
''''        End If
''''        On Error Resume Next
''''        If Not IsNull(Rec!vchInit1) Then
''''            txtInit1.Text = Rec!vchInit1
''''        Else
''''            txtInit1.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchInit2) Then
''''            txtInit2.Text = Rec!vchInit2
''''        Else
''''            txtInit2.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchInit3) Then
''''            txtInit3.Text = Rec!vchInit3
''''        Else
''''            txtInit3.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchInit4) Then
''''            txtInit4.Text = Rec!vchInit4
''''        Else
''''            txtInit4.Text = ""
''''        End If
''''        On Error GoTo 0
''''        If Not IsNull(Rec!vchHouseName) Then
''''            txtHouse.Text = Rec!vchHouseName
''''        Else
''''            txtHouse.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchStreet) Then
''''            txtStreet.Text = Rec!vchStreet
''''        Else
''''            txtStreet.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchLocalPlace) Then
''''            txtLocalPlace.Text = Rec!vchLocalPlace
''''        Else
''''            txtLocalPlace.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchMainPlace) Then
''''            txtMainPlace.Text = Rec!vchMainPlace
''''        Else
''''            txtMainPlace.Text = ""
''''        End If
''''
''''        If Not IsNull(Rec!vchPost) Then
''''            txtPost.Text = Rec!vchPost
''''        Else
''''            txtPost.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchPinCode) Then
''''            txtPin.Text = Rec!vchPinCode
''''        Else
''''            txtPin.Text = ""
''''        End If
''''        If Not IsNull(Rec!vchPhone) Then
''''            txtPhone.Text = Rec!vchPhone
''''        Else
''''            txtPhone.Text = ""
''''        End If
''''
''''        objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intSubsidiaryAccountHeadID), 0, Rec!intSubsidiaryAccountHeadID)
''''        If objSubLedger.SubLedgerTypeID > 0 Then
''''            txtSubLedgerType.Text = objSubLedger.SubLedgerType
''''            txtSubLedgerType.Tag = objSubLedger.SubLedgerTypeID
''''
''''            txtPayeeType.Text = objSubLedger.SubLedgerType
''''            txtPayeeType.Tag = objSubLedger.SubLedgerTypeID
''''        End If
''''
''''        txtNarration.Text = IIf(IsNull(Rec!PODesc), "", Rec!PODesc)
''''        txtForward2Seat.Tag = IIf(IsNull(Rec!numSeatID), "", Rec!numSeatID)
''''        txtForward2Seat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
''''
''''        If Not (IsNull(Rec!intSubsidiaryCashBookID) Or Rec!intSubsidiaryCashBookID = 0) Then
''''            objSubLedger.SetSubLedgerDetails IIf(IsNull(Rec!intSubsidiaryCashBookID), 0, Rec!intSubsidiaryCashBookID)
''''            txtSubsidiaryCash.Text = objSubLedger.Title
''''            txtSubsidiaryCash.Tag = Rec!intSubsidiaryCashBookID
''''        End If
''''
''''        If Not IsNull(Rec!intImplementingOfficerID) Then
''''            txtImplementingOfficer.Text = Rec!vchFunctionary
''''            txtImplementingOfficer.Tag = Rec!intImplementingOfficerID
''''        Else
''''            txtImplementingOfficer.Text = ""
''''            txtImplementingOfficer.Tag = ""
''''        End If
''''
''''        If Not IsNull(Rec!numProjectNo) Then
''''            objProj.SetProject Rec!numProjectNo
''''            If objProj.ProjectID > 0 Then
''''                txtProjectNo.Text = objProj.ProjectSerialNo
''''                txtProjectNo.Tag = objProj.ProjectID
''''                txtCategory.Text = objProj.Category
''''                txtCategory.Tag = objProj.ProjCatID
''''                txtSector.Text = objProj.Sector
''''                txtSector.Tag = objProj.SectorTypeID
''''                objProj.FindSourceOfFund Rec!intSourceOfFundID
''''                txtSourceOfFund.Text = objProj.SourceOfFund
''''                txtSourceOfFund.Tag = objProj.SourceOfFundID
''''            End If
''''        Else
''''            txtProjectNo.Text = ""
''''            txtProjectNo.Tag = ""
''''        End If
''''       '-----------------------------------------------------------'
''''       ' Recalculate                                        '
''''       '-----------------------------------------------------------'
''''       Dim mTOt As Double
''''       mTOt = CalculateAmt
''''       If Val(txtDrAmount.Text) > Val(mTOt) Then
''''            txtCrAmount.Text = Val(txtDrAmount.Text) - Val(mTOt)
''''       End If
''''
''''        '-----------------------------------------------------------'
''''        ' Locking Fields According to Transaction Types             '
''''        '-----------------------------------------------------------'
''''         Call SetFormControls
''''        '-----------------------------------------------------------'
''''
''''        If Rec!tnyStatus = 0 Then
''''            '------------------------------------------------------'
''''            ' Note:- Added On 15-Jan-2010                          '
''''            '        Only Approver/Accounts Officer can access     '
''''            '        Approve Comman                                '
''''            '------------------------------------------------------'
''''            If gbUserTypeID = UserType.Approver Or _
''''                 gbUserTypeID = UserType.Developer Or _
''''                 gbUserTypeID = UserType.AccountsOfficer Then
''''                 cmdApproval.Visible = True
''''            End If
''''        Else
''''            cmdApproval.Visible = False
''''        End If
''''
''''    End If
''''    Rec.Close
''''    Set mCnn = Nothing
''''
''''End Sub

    Private Sub UpdateDueDateForPendingTask()
        Dim mSql As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Update faPayOrder Set dtDueDate='" & Format(mPendingTransactionDate, "dd/mmm/yyyy") & "' Where intPayOrderID=" & val(txtPayOrder.Tag)
            mCnn.Execute mSql
            mCnn.Close
        End If
    End Sub
    Private Sub GetPendingTaskDetails()
        Dim mSql As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mPTaskID    As Integer
        Dim objAcc      As New clsAccounts
        Dim objTr      As New clsTransactionType
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
        
            mSql = "Select * From faPendingTaskRequest Where intRequestID=" & mPendingTaskReqID
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                mPTaskID = Rec!intTaskID
                If mPTaskID = 11 Then
                    mPendingTransactionDate = IIf(IsDate(Rec!dtTransactionDate), Rec!dtTransactionDate, DateAdd("yyyy", -1, gbEndingDate))
                    txtDueDate.Text = Format(mPendingTransactionDate, "dd/mmm/yyyy")
                Else
                    mPendingTransactionDate = IIf(IsDate(Rec!dtTransactionDate), Rec!dtTransactionDate, DateAdd("yyyy", -1, gbEndingDate))
                    txtTransactionType.Tag = Rec!intTransactionTypeID
                    objTr.SetTransactionType (val(txtTransactionType.Tag))
                    txtTransactionType.Text = objTr.TransactionType
                    txtDated.Text = Format(IIf(IsDate(Rec!dtTransactionDate), Rec!dtTransactionDate, DateAdd("yyyy", -1, gbEndingDate)), "dd/mmm/yyyy")
                    txtDueDate.Text = Format(IIf(IsDate(Rec!dtTransactionDate), Rec!dtTransactionDate, DateAdd("yyyy", -1, gbEndingDate)), "dd/mmm/yyyy")
                    dtpDueDate.Enabled = False
                    txtDrHeadCode.Text = Token(gbSearchStr, " ")
                    txtDrHeadCode.Tag = Rec!intExpenditureHead
                    objAcc.SetAccountID (val(txtDrHeadCode.Tag))
                    txtDrAccountHead.Text = objAcc.AccountHead
                    txtDrHeadCode.Text = objAcc.AccountCode
                    Call txtDrHeadCode_LostFocus
                    
                    txtDrAmount.Text = Rec!fltAmount
                    txtNarration.Text = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                    cmdDrAccountHead.Enabled = False
                    txtDrAmount.Enabled = False
                    txtDated.Enabled = False
                    txtCrAmount.Text = Rec!fltAmount
                    txtCrAmount.Enabled = True
                    
                    If val(txtTransactionType.Tag) > 1140 And val(txtTransactionType.Tag) < 1192 Or val(txtTransactionType.Tag) = 1201 Or val(txtTransactionType.Tag) = 1391 Then
                        txtAllotmentLetterNo.Text = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                        txtAllotmentLetterNo.Tag = Rec!intKeyID
                        Call txtAllotmentLetterNo_GotFocus
                        cmdAllotmentLetterNo.Enabled = False
                        Dim objAllotment As New clsAllotmentLetter
                        objAllotment.SetAllotment (txtAllotmentLetterNo.Tag)
                        txtSourceofFund.Text = IIf(IsNull(objAllotment.SourceOfFund), "", objAllotment.SourceOfFund)
                        txtSourceofFund.Tag = IIf(IsNull(objAllotment.SourceOfFundID), "", objAllotment.SourceOfFundID)
                        txtImplementingOfficer.Text = IIf(IsNull(objAllotment.ImplementingOfficer), "", objAllotment.ImplementingOfficer)
                        txtImplementingOfficer.Tag = IIf(IsNull(objAllotment.ImplementingOfficersID), "", objAllotment.ImplementingOfficersID)
                        txtAllotedAmt.Text = IIf(IsNull(objAllotment.Amount), "", objAllotment.Amount)
                        txtTreasuryID.Text = IIf(IsNull(objAllotment.mNewModeID), 0, objAllotment.mNewModeID)
                    End If
                    
                    
                    
'                    Dim objProject As New clsProject
'                    objProject.SetProject (val(txtProjectNo.Tag)), gbFinancialYearID - 1
'                    If Not IsNull(objProject.ProjectID) Then
'                        txtProjectNo.Text = IIf(IsNull(objProject.ProjectSerialNo), "", objProject.ProjectSerialNo)
'                        txtProjectNo.Tag = IIf(IsNull(objProject.ProjectID), "", objProject.ProjectID)
'                        txtCategory.Text = IIf(IsNull(objProject.Category), "", objProject.Category)
'                        txtCategory.Tag = IIf(IsNull(objProject.CategoryID), "", objProject.CategoryID)
'                        txtSector = IIf(IsNull(objProject.Sector), "", objProject.Sector)
'                        txtSector.Tag = IIf(IsNull(objProject.SectorTypeID), "", objProject.SectorTypeID)
'                    End If
                    
                End If
            End If
            Rec.Close
            If mPTaskID = 11 Then
                If CDate(txtDueDate.Text) <> mPendingTransactionDate Then
                    UpdateDueDateForPendingTask
                End If
            End If
            mCnn.Close
            
        End If
    End Sub
    
    Private Sub FinancialYearSetForPEndingTask()
        Dim mSql    As String
        Dim objdb       As New clsDB
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim Trndate     As Date
        Dim mTrnYear    As Integer
        Dim Curyear     As Integer
        Dim mStartDate   As Date
        Dim mEndDate   As Date
        If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
            mSql = "Select * From faFinancialYear Where tinCurrentFinancialYearFlag=1"
            Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
            If Not (Rec.EOF Or Rec.BOF) Then
                Curyear = Rec!intFinancialYear
            End If
            Rec.Close
            If mPendingTask = 1 Then
                mSql = "Select * From faFinancialYear Where intFinancialYear=" & Curyear - 1
                    Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.EOF Or Rec.BOF) Then
                        mStartDate = Rec!dtStartingDate
                        mEndDate = Rec!dtEndingDate
                        gbStartingDate = mStartDate
                        gbEndingDate = mEndDate
                        gbFinancialYearID = Curyear - 1
                    End If
                    Rec.Close
                 mSql = "Select * From faPendingTaskRequest Where intRequestID=" & mPendingTaskReqID
                 Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                    If Not (Rec.EOF Or Rec.BOF) Then
                        gbTransactionDate = mEndDate
                    End If
                    Rec.Close
            Else
                mSql = "Select *,GetDate() as TrnDate From faFinancialYear Where tinCurrentFinancialYearFlag=1"
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If Not (Rec.EOF Or Rec.BOF) Then
                    mTrnYear = Rec!intFinancialYear
                    Trndate = Rec!Trndate
                    gbTransactionDate = Trndate
                    gbFinancialYearID = mTrnYear
                    gbStartingDate = Rec!dtStartingDate
                    gbEndingDate = Rec!dtEndingDate
                End If
                Rec.Close
            End If
            mCnn.Close
        End If
    End Sub
    Public Property Let PendingTask(ByVal val As Integer)
        mPendingTask = val
    End Property
    Public Property Let PendingTaskReqID(mData As Integer)
        mPendingTaskReqID = mData
    End Property


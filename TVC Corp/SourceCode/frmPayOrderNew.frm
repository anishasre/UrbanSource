VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPayOrderNew 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Order"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   30
      TabIndex        =   78
      Top             =   7530
      Width           =   11910
      Begin VB.CommandButton cmdApprove 
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
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   180
         Width           =   1380
      End
      Begin VB.CommandButton cmdNew 
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
         Left            =   7110
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   180
         Width           =   1380
      End
      Begin VB.CommandButton cmdSeat 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
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
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   195
         Width           =   315
      End
      Begin VB.TextBox txtForward2Seat 
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
         Height          =   270
         Left            =   1500
         TabIndex        =   54
         Top             =   210
         Width           =   1725
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
         Left            =   9945
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   180
         Width           =   1410
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   180
         Width           =   1380
      End
      Begin VB.Label Label26 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   97
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   30
      TabIndex        =   74
      Top             =   3705
      Width           =   11865
      Begin VB.Frame fraProject 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   150
         TabIndex        =   91
         Top             =   1305
         Width           =   5625
         Begin VB.TextBox txtCategory 
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
            Height          =   270
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1065
            Width           =   3405
         End
         Begin VB.TextBox txtSector 
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
            Height          =   270
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1365
            Width           =   3405
         End
         Begin VB.TextBox txtProjectNo 
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
            Height          =   270
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   765
            Width           =   3405
         End
         Begin VB.TextBox txtAgreementNo 
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
            Height          =   270
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   465
            Width           =   3405
         End
         Begin VB.TextBox txtAllotmentLetterNo 
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
            Height          =   270
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   165
            Width           =   3405
         End
         Begin VB.CommandButton cmdAgreementNo 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
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
            Left            =   5190
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   450
            Width           =   300
         End
         Begin VB.CommandButton cmdAllotmentLetterNo 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5190
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   135
            Width           =   300
         End
         Begin VB.CommandButton cmdProjectNo 
            BackColor       =   &H00F5FCFC&
            Caption         =   "..."
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
            Left            =   5190
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   765
            Width           =   300
         End
         Begin VB.Label Label22 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1140
            TabIndex        =   96
            Top             =   1425
            Width           =   555
         End
         Begin VB.Label Label17 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   95
            Top             =   495
            Width           =   1230
         End
         Begin VB.Label Label8 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   45
            TabIndex        =   94
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label Label24 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   900
            TabIndex        =   93
            Top             =   1095
            Width           =   795
         End
         Begin VB.Label Label9 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   360
            TabIndex        =   92
            Top             =   795
            Width           =   1335
         End
      End
      Begin VB.TextBox txtNarration 
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
         Height          =   540
         Left            =   945
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   3105
         Width           =   4350
      End
      Begin VB.TextBox txtPayeeType 
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
         Height          =   270
         Left            =   7740
         TabIndex        =   40
         Top             =   1065
         Width           =   3540
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   7725
         TabIndex        =   35
         Top             =   150
         Width           =   3570
         Begin VB.OptionButton optNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No"
            Height          =   240
            Left            =   2790
            TabIndex        =   37
            Top             =   150
            Width           =   705
         End
         Begin VB.OptionButton optYes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Yes"
            Height          =   240
            Left            =   2115
            TabIndex        =   36
            Top             =   150
            Width           =   705
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Link With SubLedger"
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
            Left            =   135
            TabIndex        =   88
            Top             =   165
            Width           =   1755
         End
      End
      Begin VB.CommandButton cmdSearchName 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   300
         Left            =   11310
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1440
         Width           =   345
      End
      Begin VB.CommandButton cmdSubLederType 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   300
         Left            =   11310
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   660
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   38
         Top             =   690
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
         Left            =   10965
         MaxLength       =   1
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1455
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   41
         Top             =   1455
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   47
         Top             =   1770
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   48
         Top             =   2085
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   49
         Top             =   2400
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
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   50
         Top             =   2715
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
         Left            =   9975
         MaxLength       =   1
         TabIndex        =   42
         Top             =   1455
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
         Left            =   10305
         MaxLength       =   1
         TabIndex        =   43
         Top             =   1455
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
         Left            =   10635
         MaxLength       =   1
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1455
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
         Left            =   7740
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3030
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
         Left            =   10350
         MaxLength       =   6
         TabIndex        =   52
         Top             =   3030
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
         Left            =   7740
         MaxLength       =   30
         TabIndex        =   53
         Top             =   3345
         Width           =   2220
      End
      Begin VB.CommandButton cmdSubsidiaryCash 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   285
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   195
         Width           =   285
      End
      Begin VB.CommandButton cmdImplementingOfficer 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   285
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   960
         Width           =   300
      End
      Begin VB.CommandButton cmdSourceOfFund 
         BackColor       =   &H00F5FCFC&
         Caption         =   "..."
         Height          =   300
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   645
         Width           =   300
      End
      Begin VB.TextBox txtSourceOfFund 
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
         Height          =   270
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   675
         Width           =   3420
      End
      Begin VB.TextBox txtSubsidiaryCash 
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
         Height          =   270
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   210
         Width           =   3420
      End
      Begin VB.TextBox txtImplementingOfficer 
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
         Height          =   270
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   975
         Width           =   3420
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   90
         Top             =   3270
         Width           =   795
      End
      Begin VB.Label Label23 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7260
         TabIndex        =   89
         Top             =   1065
         Width           =   420
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
         Left            =   6285
         TabIndex        =   87
         Top             =   750
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
         Left            =   7140
         TabIndex        =   86
         Top             =   2115
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
         Left            =   6600
         TabIndex        =   85
         Top             =   1785
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
         Left            =   6390
         TabIndex        =   84
         Top             =   1470
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
         Left            =   7305
         TabIndex        =   83
         Top             =   3060
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
         Left            =   6795
         TabIndex        =   82
         Top             =   2745
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
         Left            =   6765
         TabIndex        =   81
         Top             =   2430
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
         Left            =   10050
         TabIndex        =   80
         Top             =   3060
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
         Left            =   6885
         TabIndex        =   79
         Top             =   3375
         Width           =   810
      End
      Begin VB.Label Label12 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   77
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label Label20 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   465
         TabIndex        =   76
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label Label21 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   555
         TabIndex        =   75
         Top             =   690
         Width           =   1290
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
      Left            =   10230
      TabIndex        =   19
      Top             =   3285
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.CommandButton cmdDrAccountHead 
      Caption         =   "..."
      Height          =   300
      Left            =   8085
      TabIndex        =   12
      Top             =   1020
      Width           =   270
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
      Left            =   4110
      TabIndex        =   11
      Top             =   1020
      Width           =   3945
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
      Left            =   2385
      TabIndex        =   10
      Top             =   1020
      Width           =   1710
   End
   Begin VB.CommandButton cmdCrAccountHead 
      Caption         =   "..."
      Height          =   300
      Left            =   8085
      TabIndex        =   17
      Top             =   2940
      Width           =   270
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
      Left            =   4200
      TabIndex        =   16
      Top             =   2955
      Width           =   3855
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
      Left            =   2475
      TabIndex        =   15
      Top             =   2955
      Width           =   1710
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
      Left            =   8385
      TabIndex        =   13
      Top             =   1020
      Width           =   1890
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
      Left            =   8385
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2955
      Width           =   1875
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E7F1F1&
      Height          =   1005
      Left            =   15
      TabIndex        =   8
      Top             =   -15
      Width           =   11925
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
         TabIndex        =   0
         Top             =   195
         Width           =   1575
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
         TabIndex        =   1
         Top             =   195
         Width           =   1620
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
         TabIndex        =   9
         Top             =   525
         Width           =   2760
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
         TabIndex        =   6
         Top             =   210
         Width           =   2760
      End
      Begin VB.CommandButton cmdSearchFunction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   510
         Width           =   270
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   195
         Width           =   270
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
         TabIndex        =   2
         Top             =   210
         Width           =   1350
      End
      Begin VB.CommandButton cmdSearchTransactionType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   525
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
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   525
         Width           =   6075
      End
      Begin MSComCtl2.DTPicker dtpDueDate 
         Height          =   360
         Left            =   7485
         TabIndex        =   3
         Top             =   195
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   635
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   39910
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00592525&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Order No"
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
         Left            =   195
         TabIndex        =   66
         Top             =   240
         Width           =   1155
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
         TabIndex        =   65
         Top             =   555
         Width           =   705
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
         TabIndex        =   64
         Top             =   255
         Width           =   990
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
         TabIndex        =   63
         Top             =   540
         Width           =   1230
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
         TabIndex        =   62
         Top             =   225
         Width           =   630
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
         TabIndex        =   61
         Top             =   240
         Width           =   810
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid vsGrid 
      Height          =   1560
      Left            =   2310
      TabIndex        =   14
      Top             =   1365
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   3
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPayOrderNew.frx":0000
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
      Left            =   7740
      TabIndex        =   73
      Top             =   3315
      Visible         =   0   'False
      Width           =   2475
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
      Left            =   750
      TabIndex        =   72
      Top             =   1080
      Width           =   1590
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
      Left            =   420
      TabIndex        =   71
      Top             =   2985
      Width           =   2025
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
      Left            =   2175
      TabIndex        =   70
      Top             =   3345
      Width           =   1545
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
      Left            =   3825
      TabIndex        =   69
      Top             =   3330
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
      Left            =   5295
      TabIndex        =   68
      Top             =   3330
      Width           =   1605
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
      Left            =   6975
      TabIndex        =   67
      Top             =   3315
      Width           =   480
   End
End
Attribute VB_Name = "frmPayOrderNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================='
'Due to the difficulty in modifying the existing Payment Order Form '
'added new form for Payment Order On 18/01/2010 by Cijith Sreedharan'
'==================================================================='

Option Explicit

    Private intUserTypeID As Integer

    Dim mBudgetBalanceAmt As Variant


    Private Sub cmdAgreementNo_Click()
        frmSearchAgreements.Show vbModal
        If gbSearchID = -1 Then
            txtAgreementNo.Text = gbSearchStr
            txtAgreementNo.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdAllotmentLetterNo_Click()
        'frmListOfAllotmentLetters.Mode = 0
        frmListOfAllotmentLetters.Show vbModal
        If gbSearchID <> -1 Then
            txtAllotmentLetterNo.Text = gbSearchStr
            txtAllotmentLetterNo.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub

    Private Sub cmdApprove_Click()
'''        On Error GoTo Err:
'''            Dim objDb As New clsDb
'''            Dim Rec As New ADODB.Recordset
'''            Dim mCnn As New ADODB.Connection
'''            Dim mSQL As String
'''            mSQL = "Select * From faPayOrder Where vchPayOrderNo = '" & Val(txtPayOrder) & "'"
'''            If objDb.SetConnection(mCnn) Then
'''                Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic, adCmdText
'''                If Rec!tnyStatus = 0 Then
'''                    cmdApprove.Enabled = False
'''                    Call MakePayable(Val(txtPayOrder))
'''                    If Rec.State = 1 Then Rec.Close
'''                    mSQL = "Update faPayOrder Set tnyStatus = 1 Where vchPayOrderNo = '" & txtPayOrder.Text & "'"
'''                    mCnn.Execute mSQL
'''                    If Rec.State = 1 Then Rec.Close
'''                Else
'''                    MsgBox "This Payment Order is already Approved once!", vbInformation
'''                End If
'''            Else
'''                MsgBox "Connection To Finance does not Exist, Please contact your System Administrator", vbInformation
'''            End If
'''        Exit Sub
'''Err:
'''        MsgBox (Error$)
    End Sub

    Private Sub cmdCrAccountHead_Click()
        On Error GoTo Err:
            Dim mSQL As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDb
            Dim objAcc As New clsAccounts
            ''Select Case Val(cmbTransactionType.Tag)
            ''    Case Is = 1006
            ''        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where vchAccountHeadCode = '350110800'"
            ''    Case Else
            ''        frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (3)"
            ''End Select
            If objDb.SetConnection(mCnn) Then
                mSQL = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " Order By faTransactionTypeChild.intOrder"
                Rec.Open mSQL, mCnn
                If Rec.BOF Or Rec.EOF Then
                    mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads" ' Where tinType IN (3)"
                End If
                
                If Val(txtTransactionType.Tag) = 1015 Then  '***** Contigent Bills For Project Expences ****'
                    mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (2)"
                End If
                
                If Val(txtTransactionType.Tag) = 1016 Then  '***** Contigent Bills For Payment of Advance ****'
                    mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads Where tinType IN (4)"
                End If
                
                frmSearchAccountHeads.SQLString = mSQL
                frmSearchAccountHeads.Show vbModal
                
                txtCrHeadCode.Text = Token(gbSearchStr, " ")
                txtCrAccountHead.Text = Trim(gbSearchStr)
                objAcc.SetAccountCode (txtCrHeadCode.Text)
                txtCrHeadCode.Tag = objAcc.AccountHeadID
                
                gbSearchID = -1
                gbSearchStr = ""
                txtCrAmount.SetFocus
            Else
                MsgBox "Connection To Finance does not Exist, Please contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

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
        Call FormInitialize
    End Sub

    Private Sub cmdProjectNo_Click()
        frmEstimationDetails.Mode = 0
        frmSulekhaIntegration.Show vbModal
        If gbSearchID <> -1 Then
            txtProjectNo.Text = gbSearchStr
            txtProjectNo.Tag = gbSearchID
            
            gbSearchID = -1
            gbSearchStr = ""
        End If
    End Sub
    
    Private Function CheckValidation() As Boolean
        CheckValidation = False
        If Not IsDate(txtDated) Then
            MsgBox "Please Check the Transaction Date", vbInformation
            txtDated.SetFocus
            Exit Function
        End If
        
        If Val(txtFunctionary.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Functionary", vbInformation
            cmdSearchFunctionary.SetFocus
            Exit Function
        End If
        
        If Val(txtFunction.Tag) < 1 Then
            MsgBox "Please Select Proper Budget Function", vbInformation
            cmdSearchFunction.SetFocus
            Exit Function
        End If
        
        If Val(txtTransactionType.Tag) < 1 Then
            MsgBox "Please Select Proper Transaction Type for this Transaction", vbInformation
            txtTransactionType.SetFocus
            Exit Function
        End If
        
        If IsDate(txtDueDate.Text) = False Then
            MsgBox "Please Give Due Date", vbInformation
            dtpDueDate.SetFocus
            Exit Function
        End If
        
        If Val(txtDrHeadCode.Tag) < 1 Then
            txtDrHeadCode.SetFocus
            MsgBox "Please Enter Debit Head", vbInformation
            Exit Function
        End If
        
        If Val(txtDrAmount) <= 0 Then
            txtDrAmount.SetFocus
            MsgBox "Please Enter Debit Amount", vbInformation
            Exit Function
        End If
        
        If Val(txtCrHeadCode.Tag) = 0 Then
            MsgBox "Please Select The Credit Account Head", vbInformation
            cmdCrAccountHead.SetFocus
            Exit Function
        End If
        
        If txtCrAmount.Text = "" Then
            MsgBox "Please Give the Credit Amount", vbInformation
            txtCrAmount.SetFocus
            Exit Function
        End If
        
        If Val(txtSourceOfFund.Tag) < 1 Then
            MsgBox "Please Select the Source Of Fund", vbInformation
            txtSourceOfFund.SetFocus
            Exit Function
        End If
        
        If Val(txtForward2Seat.Tag) < 1 Then
            MsgBox "Please Select Forward to Seat", vbInformation
            txtForward2Seat.SetFocus
            Exit Function
        End If
        
        If txtName.Text = "" Then
            txtName.SetFocus
            MsgBox "Please Enter the Name of Payee..", vbInformation
            Exit Function
        End If
        
        
        If Val(txtTransactionType.Tag) > 1140 And Val(txtTransactionType.Tag) < 1192 Then
            If Val(txtProjectNo.Tag) < 1 Then
                MsgBox "Please select a Project", vbInformation
                txtProjectNo.SetFocus
                Exit Function
            End If
            
            If Val(txtCategory.Tag) < 1 Then
                MsgBox "Please select a Category", vbInformation
                txtCategory.SetFocus
                Exit Function
            End If
            
            If Val(txtSector.Tag) < 1 Then
                MsgBox "Please select a Sector", vbInformation
                txtSector.SetFocus
                Exit Function
            End If
        End If
        
        
        If txtSubsidiaryCash.Text <> "" Then
            If Val(txtSubLedgerType.Tag) = 10 And txtName.Text = "" Then
                MsgBox "Please Select the Official for disbursing the Subsidiary Cash"
                txtName.SetFocus
                CheckValidation = False
                Exit Function
            End If
        End If
        
        CheckValidation = True
    End Function
    
    Private Sub SavePaymentOrder()
        Dim PO As uPaymentOrder
        Dim POC As uPaymentOrderChild
        Dim POAdd As uPaymentOrderAddress
        Dim ObjSubLed As New clsSubLedger
        
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim mPaymentOrderID As Variant
        Dim mSlNo As Integer
        Dim mLoop As Integer
        Dim vchPayOrderNo As String
        Dim mSQL As String
        
        cmdSave.Enabled = False
        objDb.SetConnection mCnn
        'mCnn.BeginTrans
        On Error GoTo ErrRollBack:
        With PO
            .intPayOrderID = IIf(txtPayOrder.Tag = "", Null, txtPayOrder.Tag)
            .vchPayOrderNo = IIf(txtPayOrder.Text = "", Null, txtPayOrder.Text)
            .dtPayOrderDate = gbTransactionDate
            .dtDueDate = txtDueDate.Text
            .intFunctionaryID = Val(txtFunctionary.Tag)
            .intFunctionID = Val(txtFunction.Tag)
            .intTransactionTypeID = Val(txtTransactionType.Tag)
            .vchBillNo = Null
            .numBillAmount = Null
            .dtBillDate = Null
            .intInstrumentTypeID = Null
            .intCashOrBankHeadID = Null
            .vchDescription = Trim(txtNarration.Text)
            .vchTitle = Null
            .intSubLedgerTypeID = Val(txtSubLedgerType.Tag)
            .intPayToSubLedgerID = Val(txtName.Tag)
            .intSubsidiaryCashBookID = Val(txtSubsidiaryCash.Tag)
            .intImplementingOfficerID = Val(txtImplementingOfficer.Tag)
            .numProjectNo = Val(txtProjectNo.Tag)
            .intStockRegisterID = Null
            .vchStockRefNo = Null
            .intAssetTypeID = Null
            .intAssetID = Null
            .numFwdSeatID = Val(txtForward2Seat.Tag)
            .intLocalBodyID = gbLocalBodyID
            .intZonalID = gbLocationID
            .intFinancialYearID = gbFinancialYearID
            .numUserID = gbUserID
            .numSeatID = gbSeatID
            .numApprovingOfficerID = Null
            .numApprovingSeatID = Null
            .dtApprovingDate = Null
            .intSourceOfFundID = Val(txtSourceOfFund.Tag)
            .intAllotmentID = Val(txtAllotmentLetterNo.Tag)
            .intAgreementID = Val(txtAgreementNo.Tag)
            .tnyCategoryID = Val(txtCategory.Tag)
            .tnySectorID = Val(txtSector.Tag)
            .tnyIsFinalBill = Null
            .intVoucherID = Null
            .intVoucherNo = Null
            .dtVoucherDate = Null
            .tnyStatus = 0
            .tnyCancelled = 0
            .intAppID = 115
            .intModuleID = 2
            
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
            
            objDb.ExecuteSP "spSavePayOrder", arrInput, arrOutPut, , mCnn, adCmdStoredProc
               
        End With
        
        If IsNumeric(arrOutPut(0, 0)) Then
            mPaymentOrderID = arrOutPut(0, 0)
            vchPayOrderNo = arrOutPut(1, 0)
        Else
            GoTo ErrRollBack:
        End If
        
        mSQL = "Delete From faPayOrderChild Where intPayOrderID = " & mPaymentOrderID
        mCnn.Execute mSQL
        
        mSlNo = mSlNo + 1
        With POC
            .intPayOrderID = mPaymentOrderID
            .intSlNo = mSlNo
            .intAccountHeadID = Val(txtDrHeadCode.Tag)
            .vchAccountHeadCode = Trim(txtDrHeadCode.Text)
            .numAmount = Val(txtDrAmount.Text)
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
                        
            objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        mSlNo = mSlNo + 1
        For mLoop = 1 To vsGrid.Rows - 1
            If Val(vsGrid.TextMatrix(mLoop, 1)) > 0 And Val(vsGrid.TextMatrix(mLoop, 3)) > 0 Then
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = Val(vsGrid.TextMatrix(mLoop, 4))
                .vchAccountHeadCode = Trim(vsGrid.TextMatrix(mLoop, 1))
                .numAmount = Val(vsGrid.TextMatrix(mLoop, 3))
                .tnyCategoryFlag = 2
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
                objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
            End If
        Next
        
        mSlNo = mSlNo + 1
        With POC
            .intPayOrderID = mPaymentOrderID
            .intSlNo = mSlNo
            .intAccountHeadID = Val(txtCrHeadCode.Tag)
            .vchAccountHeadCode = Trim(txtCrHeadCode.Text)
            .numAmount = Val(txtCrAmount)
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
            objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
        End With
        
        If Val(txtSubsidiaryCash.Tag) <> 0 Then
            mSlNo = mSlNo + 1
            With POC
                .intPayOrderID = mPaymentOrderID
                .intSlNo = mSlNo
                .intAccountHeadID = Val(txtSubCashCode.Tag)
                .vchAccountHeadCode = Trim(txtSubCashCode.Text)
                .numAmount = Val(txtCrAmount)
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
                objDb.ExecuteSP "spSavePayOrderChild", arrInput, , , mCnn, adCmdStoredProc
            End With
        End If
        
        With POAdd
            
            ObjSubLed.SetSubLedgerDetails (Val(txtName.Tag))
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
            objDb.ExecuteSP "spSavePayOrderAddress", arrInput, , , mCnn, adCmdStoredProc
           
        End With
        
        'mCnn.CommitTrans
        txtPayOrder.Text = vchPayOrderNo
        cmdSave.Enabled = False
        Exit Sub
        
ErrRollBack:
        MsgBox (Error$)
        cmdSave.Enabled = True
    End Sub
    
    
    Private Sub cmdSave_Click()
        If CheckValidation Then
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
        frmSearchSubsidiaryAccountHeads.SubLedgerType = Val(txtSubLedgerType.Tag)
        frmSearchSubsidiaryAccountHeads.Show vbModal
        If gbSearchStr <> "" Then
            txtName.Tag = gbSearchID
            objSubLedger.SetSubLedgerDetails (Val(txtName.Tag))
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
        
        '================================================================================================'
    End Sub

    Private Sub cmdSearchTransactionType_Click()
        On Error GoTo Err:
            gbSearchStr = ""
            gbSearchID = -1
            gbSearchCode = ""
            
            frmSearchTransactionType.ModeOfTransaction = 2
            frmSearchTransactionType.Show vbModal
            
            If gbSearchID <> -1 Then
                txtTransactionType.Text = gbSearchStr
                txtTransactionType.Tag = gbSearchID
            End If
            
            If Val(txtTransactionType.Tag) > 1140 And Val(txtTransactionType.Tag) < 1192 Then
                fraProject.Visible = True
            Else
                fraProject.Visible = False
            End If
            
            Dim objTrns As New clsTransactionType
            objTrns.SetSourceOfFund (txtTransactionType.Tag)
            If Not IsEmpty(objTrns.SourceFundID) Then
                txtSourceOfFund.Text = objTrns.SourceOfFund
                txtSourceOfFund.Tag = objTrns.SourceFundID
            Else
                txtSourceOfFund.Text = "Own Fund"
                txtSourceOfFund.Tag = 4
            End If
            
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
            
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub cmdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID <> -1 Then
            txtForward2Seat.Text = gbSearchStr
            txtForward2Seat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    
    Private Sub cmdSourceOfFund_Click()
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intSourceFundID, vchSourceFundName From suSourceOfFund"
        frmSearchMasters.Show vbModal
        'txtSourceOfFund.SetFocus
        If gbSearchID <> -1 Then
            txtSourceOfFund.Text = gbSearchStr
            txtSourceOfFund.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    
    Private Sub cmdSubLederType_Click()
        '================================================================================================'
        '                                   Modified By Cijith On 20/11/2009                             '
        '================================================================================================'
        'frmSearchSubsidiaryAccountHeads.Show vbModal
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
    
    Private Sub cmdSubsidiaryCash_Click()
        '================================================================================================'
        '               Added By Cijith On 20/11/2009   For Integrating Subsidiary Cash Book             '
        '================================================================================================'
        On Error GoTo Err:
            frmSearchSubsidiaryAccountHeads.SubLedgerType = 12
            frmSearchSubsidiaryAccountHeads.Show vbModal
            If gbSearchStr <> "" Then
                txtSubsidiaryCash.Text = gbSearchStr
                txtSubsidiaryCash.Tag = gbSearchID
            End If
            gbSearchID = -1
            gbSearchStr = ""
            
            If Val(txtSubsidiaryCash.Tag) <> 0 Then
                txtPayeeType.Text = txtSubsidiaryCash.Text
                Call ShowDetailsForSubCashBook
            End If
        Exit Sub
Err:
        MsgBox (Error$)
        '================================================================================================'
    End Sub

    Private Function ShowDetailsForSubCashBook() As Boolean
        On Error GoTo Err:
            Dim objAcc As New clsAccounts
            Dim objSubLedger As New clsSubLedger
            If Val(txtTransactionType.Tag) = 1001 Or Val(txtTransactionType.Tag) = 1211 Then
                objAcc.SetAccountID (1550)
                txtCrHeadCode.Text = objAcc.AccountCode
                txtCrHeadCode.Tag = objAcc.AccountHeadID
                txtCrAccountHead.Text = objAcc.AccountHead
                cmdSubsidiaryCash.Enabled = True
                
                txtSubLedgerType.Text = "Subsidiary Cash Book"
                txtSubLedgerType.Tag = 10
            Else
                '''txtCrHeadCode.Text = ""
                '''txtCrHeadCode.Tag = ""
                '''txtCrAccountHead.Text = ""
                '''cmdSubsidiaryCash.Enabled = False
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
End Sub

    Private Sub Form_Load()
        vsGrid.ColComboList(1) = "|..."
        Call FormInitialize
    End Sub

    Private Sub FormInitialize()
        On Error GoTo Err:
            Dim ctrl As Control
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
            'Check1.Value = 0
            'Note:- User Type wise Functionality should enable or Disabled
            'cmdApproval.Visible = False
            
            If gbUserTypeID = 3 Then
                cmdSave.Visible = True
                cmdApprove.Visible = False
            ElseIf gbUserID = 2 Or gbUserTypeID = 4 Then
                cmdSave.Visible = False
                cmdApprove.Visible = True
            Else
                cmdSave.Visible = True
                cmdApprove.Visible = True
            End If
            
            
            cmdSave.Enabled = True
            
            vsGrid.Clear 1, 1
            
            optNo.Value = True
            optYes.Value = False
            'Call optNo_Click
            
            txtDated.Text = DdMmmYy(gbTransactionDate)
            txtDueDate.Text = DdMmmYy(gbTransactionDate)
            cmdCrAccountHead.Enabled = True
'            mSelect = False
'            mBudgetBalanceAmt = 0
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub optNo_Click()
        If optYes.Value = True Then
            cmdSubLederType.Enabled = True
            cmdSearchName.Enabled = True
        Else
            cmdSubLederType.Enabled = False
            cmdSearchName.Enabled = False
        End If
    End Sub

    Private Sub optYes_Click()
        If optYes.Value = True Then
            cmdSubLederType.Enabled = True
            cmdSearchName.Enabled = True
        Else
            cmdSubLederType.Enabled = False
            cmdSearchName.Enabled = False
        End If
    End Sub

    Private Sub txtDated_Change()
        txtDated.Text = CheckDateInMMM(txtDated.Text)
    End Sub

    Private Sub txtDueDate_LostFocus()
        If Trim(txtDueDate) <> "" Then
            txtDueDate.Text = CheckDateInMMM(txtDueDate.Text)
            If CDate(txtDated.Text) > CDate(txtDueDate.Text) Then
                MsgBox "Invalid Date", vbInformation
                txtDueDate = ""
                txtDueDate.SetFocus
            End If
        End If
    End Sub
    
    Private Sub cmdDrAccountHead_Click()
        On Error GoTo Err:
            Dim mSQL As String
            Dim Rec As New ADODB.Recordset
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDb
            Dim objAcc As New clsAccounts
            
            If objDb.SetConnection(mCnn) = True Then
                mSQL = "Select (faAccountHeads.vchAccountHeadCode + '  ' + faAccountHeads.vchAccountHead) as AccHead, faAccountHeads.intMajorAccountHeadID From faAccountHeads Inner Join "
                mSQL = mSQL + " faTransactionTypeChild On faTransactionTypeChild.intAccountHeadID = faAccountHeads.intAccountHeadId"
                mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag)
                mSQL = mSQL + " And faTransactionTypeChild.tinDebitOrCredit = 1 "
                mSQL = mSQL + " And faTransactionTypeChild.tnyListID = 1 "
                mSQL = mSQL + " Order By faTransactionTypeChild.intOrder"
                Rec.Open mSQL, mCnn
                If Rec.BOF Or Rec.EOF Then
                    mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faAccountHeads" ' Where tinType IN (3)"
                End If
                
                frmSearchAccountHeads.SQLString = mSQL
                frmSearchAccountHeads.Show vbModal
                txtDrHeadCode.Text = Token(gbSearchStr, " ")
                txtDrAccountHead.Text = Trim(gbSearchStr)
                objAcc.SetAccountCode (txtDrHeadCode.Text)
                txtDrHeadCode.Tag = objAcc.AccountHeadID
                
                Call txtDrHeadCode_LostFocus
                
                gbSearchID = -1
                gbSearchStr = ""
                txtDrAmount.SetFocus
                If Rec.State = 1 Then Rec.Close
                
                '============================================================='
                '   Added For Getting the Head Code for Subsidiary Cash Book  '
                '               Modified By Cijith Sreedharan                 '
                '============================================================='
                    txtSubCashCode.Text = txtDrHeadCode.Text
                    txtSubCashCode.Tag = txtDrHeadCode.Tag
                    If txtCrHeadCode.Text = "" Then
                        txtCrHeadCode.Text = txtDrHeadCode.Text
                        txtCrHeadCode.Tag = txtDrHeadCode.Tag
                        txtCrAccountHead.Text = txtDrAccountHead.Text
                    End If
                '============================================================='
            Else
                MsgBox "Connection to Finance does not Exist, Please Contact your System Administrator", vbInformation
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub
    
    Private Sub txtDrHeadCode_LostFocus()
        On Error GoTo Err:
        
            Dim mSQL As String
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim objDb As New clsDb
            
            If Trim(txtDrHeadCode) <> "" Then
                cmdCrAccountHead.Enabled = True
                Dim mGroupID As Integer
                objDb.SetConnection mCnn
                mSQL = " Select intGroupID  From faTransactionTypeChild Where intTransactionTypeID = " & Val(txtTransactionType.Tag) & " AND vchAccountHeadCode = '" & Trim(txtDrHeadCode.Text) & "'"
                Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic
                If Not (Rec.BOF And Rec.EOF) Then
                    mGroupID = Rec!intGroupID
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
                    mSQL = "Select * From faTransactionTypeChild Inner Join "
                    mSQL = mSQL + " faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionTypeChild.intAccountHeadID "
                    mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag)
                    mSQL = mSQL + " AND faTransactionTypeChild.intGroupID = " & mGroupID & "AND tnyNetPayFlag = 1 And tnyListID = 3"
                    
                    Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic
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
                        mSQL = "Select * From faTransactionTypeChild Inner Join "
                        mSQL = mSQL + " faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionTypeChild.intAccountHeadID "
                        mSQL = mSQL + " Where intTransactionTypeID = " & Val(txtTransactionType.Tag)
                        mSQL = mSQL + " AND tnyNetPayFlag = 1 And tnyListID = 3"
                        Rec.CursorLocation = adUseClient
                        Rec.Open mSQL, mCnn, adOpenForwardOnly, adLockOptimistic
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
                                txtCrHeadCode.Text = ""
                                txtCrHeadCode.Tag = ""
                                txtCrAccountHead.Text = ""
                                cmdCrAccountHead.Enabled = True
                                lblBudgetAmt.Caption = "0.00"
                                lblUtilizedAmt.Caption = "0.00"
                            End If
                        Else
                            txtCrHeadCode.Text = ""
                            txtCrHeadCode.Tag = ""
                            txtCrAccountHead.Text = ""
                            cmdCrAccountHead.Enabled = True
                            lblBudgetAmt.Caption = "0.00"
                            lblUtilizedAmt.Caption = "0.00"
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
                objBudj.SetBudgetAccountHead Val(txtDrHeadCode.Tag), Val(txtFunctionary.Tag), Val(txtFunction.Tag)
                If objBudj.BudgetCentreID > 0 Then
                    lblBudgetAmt.Caption = Format(objBudj.EstimatedAmount, "0.00")
                    lblUtilizedAmt.Caption = Format(objBudj.UtilisedAmount, "0.00")
                End If
                Set objBudj = Nothing
                
            End If
            'lblTipText.Caption = ""
        Exit Sub
Err:
        MsgBox (Error$)
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
        Dim objDb As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        txtDrAmount.Text = Format(txtDrAmount.Text, "0.00")
        If mBudgetBalanceAmt < Val(txtDrAmount) Then
            'MsgBox "Budget Balance is Rs. " & Format(mBudgetBalanceAmt, "0.00")
            'txtDrAmount = Format(mBudgetBalanceAmt, "0.00")
        End If
        txtCrAmount.Text = Format(txtDrAmount.Text, "0.00")
    End Sub
    
    Private Sub txtDrAmount_GotFocus()
        ShowBudgetBalance (gbSearchID)
    End Sub
    
    Private Sub ShowBudgetBalance(mAcHeadID As Long)
        On Error GoTo Err:
            Dim objAc As New clsAccounts
            Dim objBudjet As New clsBudgetCentre
            Dim mFunctionaryID As Variant
            Dim mFunctionID As Variant
            Dim mBudgetAmt As Variant
            Dim mUtilizedAmt As Variant
            
            mBudgetBalanceAmt = 0
            If Val(txtFunctionary.Tag) > 0 Then
                mFunctionaryID = txtFunctionary.Tag
            End If
            If Val(txtFunction.Tag) > 0 Then
                mFunctionID = txtFunctionary.Tag
            End If
            
            mBudgetAmt = objBudjet.GetBudgetAmount(mFunctionID, mFunctionaryID, mAcHeadID)
            'lblBudget.Caption = "Budget Amount :" & Format(mBudgetAmt, "0.00")
            
            mUtilizedAmt = objAc.GetLedgerBalance(mAcHeadID, , mFunctionaryID, mFunctionID)
            'lblBudget.Caption = lblBudget.Caption & "/   Budget Utilized : " & Format(mUtilizedAmt, "0.00")
            mBudgetBalanceAmt = Format(mBudgetAmt, "0.00") - Format(mUtilizedAmt, "0.00")
        Exit Sub
Err:
        MsgBox (Error$)
         
    End Sub

    Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) = False And vsGrid.Col = 3 Then
            vsGrid.TextMatrix(vsGrid.Row, 3) = ""
            MsgBox "Enter Numeric values"
        End If
        Dim mTOt As Variant
        If vsGrid.Col = 3 Then
                If IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) Then
                    vsGrid.TextMatrix(vsGrid.Row, 3) = Format(Val(vsGrid.TextMatrix(vsGrid.Row, 3)), "0.00")
                    
                    mTOt = CalculateAmt
                    If Val(txtDrAmount.Text) > Val(mTOt) Then
                        txtCrAmount.Text = Val(txtDrAmount.Text) - Val(mTOt)
                    Else
                        MsgBox "Amount Out of Range"
                    End If
                End If
        End If
    End Sub

    Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        vsGrid.TextMatrix(vsGrid.Row, 3) = Format(vsGrid.TextMatrix(vsGrid.Row, 3), "0.00")
    End Sub

    Private Sub vsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
        On Error GoTo Err:
            Dim objAc As New clsAccounts
            Dim objDb As New clsDb
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            
            frmSearchAccountHeads.SQLString = "Select (faAccountHeads.vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, faAccountHeads.intAccountHeadID From faTransactionTypeChild INNER JOIN faAccountHeads ON faAccountHeads.vchAccountHeadCode= faTransactionTypeChild.vchAccountHeadCode WHERE faTransactionTypeChild.intTransactionTypeID=" & Val(txtTransactionType.Tag) & " And faTransactionTypeChild.tnyListID = 2 Order By faAccountHeads.vchAccountHeadCode"
            frmSearchAccountHeads.Show vbModal
            If gbSearchID <> -1 Then
                vsGrid.TextMatrix(vsGrid.Row, 1) = Token(gbSearchStr, " ")
                vsGrid.TextMatrix(vsGrid.Row, 2) = Trim(gbSearchStr)
                objAc.SetAccountCode (vsGrid.TextMatrix(vsGrid.Row, 1))
                vsGrid.TextMatrix(vsGrid.Row, 4) = objAc.AccountHeadID
                vsGrid.Col = 3
                gbSearchStr = ""
                gbSearchID = -1
            End If
        Exit Sub
Err:
        MsgBox (Error$)
    End Sub

    Private Sub vsGrid_KeyPress(KeyAscii As Integer)
        Dim mTOt As Variant
        If KeyAscii = 13 And vsGrid.Col = 3 And Trim(vsGrid.TextMatrix(vsGrid.Row, 3)) <> "" And IsNumeric(vsGrid.TextMatrix(vsGrid.Row, 3)) And Val(vsGrid.TextMatrix(vsGrid.Row, 3)) > 0 Then
            vsGrid.Rows = vsGrid.Rows + 1
            vsGrid.Row = vsGrid.Row + 1
            vsGrid.Col = 1
            mTOt = CalculateAmt
            
            If Val(txtDrAmount.Text) > Val(mTOt) Then
                txtCrAmount.Text = Val(txtDrAmount.Text) - Val(mTOt)
            End If
        End If
    End Sub
    
    Private Function CalculateAmt() As Variant
        Dim mCount As Integer
        Dim mTOt As Variant
        mTOt = 0
        
        For mCount = 1 To vsGrid.Rows
            If Trim(vsGrid.TextMatrix(mCount, 3)) = "" Then Exit For
            mTOt = Val(mTOt) + Val(vsGrid.TextMatrix(mCount, 3))
        Next
        CalculateAmt = mTOt
    End Function
    
    Public Function FillPayOrder(intPayOrderID As Variant)
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            Dim objDb As New clsDb
            Dim objAc As New clsAccounts
            Dim mLoopCount As Integer
            
            If objDb.SetConnection(mCnn) Then
                '''mSQL = "Select *, CashBook.vchTitle as SubCashBook, ImpOfficer.vchFunctionary as ImplOfficer, faPayOrderChild.intAccountHeadID as HeadID "
                '''mSQL = mSQL + " , faPayOrder.intFunctionID as FnID, faPayOrder.intFunctionaryID as FryID, suSourceOfFund.vchSourceFundName as SourceOfFund "
                mSQL = "Select faPayOrder.*, faPayOrderChild.*, faPayOrderAddress.*, faFunctions.vchFunction, faFunctionaries.vchFunctionary, "
                mSQL = mSQL + " faTransactionType.vchTransactionType,  suSourceOfFund.vchSourceFundName, faSeats.chvSeatTitle, faSubLedgerTypes.vchSubLedgerType, "
                mSQL = mSQL + " Payee.vchTitle as SubAccHeadTitle, Payee.vchName as PayeeName, CashBook.vchTitle as SubCashBook, faAllotmentLetters.vchAllotmentNo, "
                mSQL = mSQL + " suProjectDetails.chvProjectSlNo, faTransactionCategory.vchTransactionCategory, ImpOfficer.vchFunctionary as ImplOfficer, "
                mSQL = mSQL + " faPayOrderChild.intAccountHeadID as HeadID, faPayOrder.intFunctionID as FnID, faPayOrder.intFunctionaryID as FryID, CashBook.vchTitle as SubCashBook, "
                mSQL = mSQL + " suSourceOfFund.vchSourceFundName as SourceOfFund, suProjectDetails.decProjectID, faTransactionCategory.intCategoryID, faPayOrder.vchDescription as [Desc] "
                mSQL = mSQL + " from faPayOrder "
                mSQL = mSQL + " Inner join faPayOrderAddress On faPayOrder.intPayOrderID = faPayOrderAddress.intPayOrderID "
                mSQL = mSQL + " Inner Join faPayOrderChild On faPayOrder.intPayOrderID = faPayOrderChild.intPayOrderID "
                mSQL = mSQL + " Inner Join faFunctions On faPayOrder.intFunctionID = faFunctions.intFunctionID "
                mSQL = mSQL + " Inner Join faFunctionaries On faPayOrder.intFunctionaryID = faFunctionaries.intFunctionaryID "
                mSQL = mSQL + " Inner Join faTransactionType On faPayOrder.intTransactionTypeID = faTransactionType.intTransactionTypeID "
                mSQL = mSQL + " Left Join faFunctionaries ImpOfficer On faPayOrder.intFunctionaryID = ImpOfficer.intFunctionaryID "
                mSQL = mSQL + " Left join suSourceOfFund On faPayOrder.intSourceOfFundID = suSourceOfFund.intSourceFundID "
                mSQL = mSQL + " Left Join faSeats On faPayOrder.numFwdSeatID = faSeats.numSeatID "
                mSQL = mSQL + " Left Join faSubLedgerTypes On faPayOrder.intPayOrderID = faSubLedgerTypes.intSubLedgerTypeID "
                mSQL = mSQL + " Left Join faSubSidiaryAccountHeads Payee On faPayOrder.intPayToSubLedgerID = Payee.intSubsidiaryAccountHeadID "
                mSQL = mSQL + " Left Join faSubSidiaryAccountHeads CashBook On faPayOrder.intSubsidiaryCashBookID = CashBook.intSubsidiaryAccountHeadID "
                mSQL = mSQL + " Left Join faAllotmentLetters On faAllotmentLetters.intAllotmentID = faPayOrder.intAllotmentID "
                mSQL = mSQL + " Left Join suProjectDetails On suProjectDetails.decProjectID = faAllotmentLetters.numProjectID "
                mSQL = mSQL + " Left Join faTransactionCategory On faTransactionCategory.intCategoryID = faAllotmentLetters.intCategoryID "
                mSQL = mSQL + " Where faPayOrder.intPayOrderID = " & intPayOrderID
                Rec.Open mSQL, mCnn
                
                If Not (Rec.EOF Or Rec.BOF) Then
                    txtTransactionType.Text = IIf(IsNull(Rec!vchTransactionType), "", Rec!vchTransactionType)
                    txtTransactionType.Tag = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                    
                    txtPayOrder.Text = IIf(IsNull(Rec!vchPayOrderNo), "", Rec!vchPayOrderNo)
                    txtPayOrder.Tag = IIf(IsNull(Rec!intPayOrderID), "", Rec!intPayOrderID)
                    
                    txtDated.Text = IIf(IsNull(Rec!dtPayOrderDate), "", CheckDateInMMM(Rec!dtPayOrderDate))
                    txtDueDate.Text = IIf(IsNull(Rec!dtDueDate), "", CheckDateInMMM(Rec!dtDueDate))
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
                    
                    txtSubsidiaryCash.Tag = IIf(IsNull(Rec!intSubsidiaryCashBookID), "", Rec!intSubsidiaryCashBookID)
                    txtSubsidiaryCash.Text = IIf(IsNull(Rec!SubCashBook), "", Rec!SubCashBook)
                    
                    txtSourceOfFund.Text = IIf(IsNull(Rec!SourceOfFund), "", Rec!SourceOfFund)
                    txtSourceOfFund.Tag = IIf(IsNull(Rec!intSourceOfFundID), "", Rec!intSourceOfFundID)
                    
                    txtImplementingOfficer.Text = IIf(IsNull(Rec!ImplOfficer), "", Rec!ImplOfficer)
                    txtImplementingOfficer.Tag = IIf(IsNull(Rec!intImplementingOfficerID), "", Rec!intImplementingOfficerID)
                    
                    txtAllotmentLetterNo.Text = IIf(IsNull(Rec!vchAllotmentNo), "", Rec!vchAllotmentNo)
                    txtAllotmentLetterNo.Tag = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                    
                    'txtAgreementNo.Text = IIf(IsNull(Rec!intAllotmentID), "", Rec!intAllotmentID)
                        
                    txtProjectNo.Text = IIf(IsNull(Rec!chvProjectSlNo), "", Rec!chvProjectSlNo)
                    txtProjectNo.Tag = IIf(IsNull(Rec!decProjectID), "", Rec!decProjectID)
                    
                    txtCategory.Text = IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    txtCategory.Tag = IIf(IsNull(Rec!intCategoryID), "", Rec!intCategoryID)
                    
                    'txtSector= IIf(IsNull(Rec!vchTransactionCategory), "", Rec!vchTransactionCategory)
                    txtNarration.Text = IIf(IsNull(Rec!Desc), "", Rec!Desc)
                    
                    txtForward2Seat.Text = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
                    txtForward2Seat.Tag = IIf(IsNull(Rec!numFwdSeatID), "", Rec!numFwdSeatID)
                    
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
            
                        Rec.MoveNext
                    Wend
                End If
            Else
                MsgBox "Connection To Finance does not Exist, Please Contact Your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Public Property Get UserType() As Integer
        UserType = intUserTypeID
    End Property
    
    Public Property Let UserType(mData As Integer)
        intUserTypeID = mData
    End Property

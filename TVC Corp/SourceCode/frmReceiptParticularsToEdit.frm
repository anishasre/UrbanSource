VERSION 5.00
Begin VB.Form frmReceiptParticularsToEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receipt Details - to be Edited if Necessory"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmePersonal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Personal Information"
      Height          =   3435
      Left            =   90
      TabIndex        =   19
      Top             =   2790
      Visible         =   0   'False
      Width           =   9405
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   36
         Top             =   2745
         Width           =   2505
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   6
         TabIndex        =   35
         Top             =   2400
         Width           =   2505
      End
      Begin VB.TextBox txtPost 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   34
         Top             =   2040
         Width           =   2505
      End
      Begin VB.TextBox txtInit4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8490
         MaxLength       =   1
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtInit3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8145
         MaxLength       =   1
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtInit2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   31
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtInit1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         MaxLength       =   1
         TabIndex        =   30
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   29
         Top             =   1665
         Width           =   3540
      End
      Begin VB.TextBox txtLocalPlace 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1320
         Width           =   3540
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   27
         Top             =   945
         Width           =   3540
      End
      Begin VB.TextBox txtHouse 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   26
         Top             =   600
         Width           =   3540
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   100
         TabIndex        =   25
         Top             =   270
         Width           =   2145
      End
      Begin VB.ComboBox cmbDZone 
         Height          =   390
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   315
         Width           =   1800
      End
      Begin VB.TextBox txtDoorNo2 
         Height          =   315
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1110
         Width           =   660
      End
      Begin VB.TextBox txtDoorNo1 
         Height          =   390
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1110
         Width           =   1095
      End
      Begin VB.TextBox txtWardNo 
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
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   21
         Top             =   780
         Width           =   1770
      End
      Begin VB.TextBox txtRefNo 
         Height          =   390
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1530
         Width           =   1770
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   48
         Top             =   2820
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   47
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   46
         Top             =   2085
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   45
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   44
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   43
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House/Office"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   42
         Top             =   645
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nam&E"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3645
         TabIndex        =   41
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   40
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   39
         Top             =   1185
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ward No"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   38
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RefNo"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   37
         Top             =   1605
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   90
      TabIndex        =   16
      Top             =   4320
      Width           =   9375
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   18
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Instrument Type"
      Height          =   1485
      Left            =   120
      TabIndex        =   5
      Top             =   2790
      Width           =   9375
      Begin VB.TextBox txtDated 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   11
         Top             =   690
         Width           =   1470
      End
      Begin VB.TextBox txtInstNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2385
         TabIndex        =   10
         Top             =   690
         Width           =   1740
      End
      Begin VB.TextBox txtBank 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2385
         TabIndex        =   9
         Top             =   990
         Width           =   1740
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   990
         Width           =   1470
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Top             =   330
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   5805
         TabIndex        =   6
         Top             =   330
         Width           =   345
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dated"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   10
         Left            =   4155
         TabIndex        =   15
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inst. No"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   9
         Left            =   1650
         TabIndex        =   14
         Top             =   705
         Width           =   645
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   1860
         TabIndex        =   13
         Top             =   990
         Width           =   390
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   5
         Left            =   2910
         TabIndex        =   12
         Top             =   1020
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction Type"
      Height          =   825
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   9375
      Begin VB.CommandButton cmdSearchTrType 
         Caption         =   "..."
         Height          =   315
         Left            =   8850
         TabIndex        =   4
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txtTransactionType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Top             =   300
         Width           =   7125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Possible Change "
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   9375
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1230
         Left            =   1980
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   300
         Width           =   5085
      End
   End
End
Attribute VB_Name = "frmReceiptParticularsToEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

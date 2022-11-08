VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSoochikaInward 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "I N W A R D"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14790
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   11608.6
   ScaleMode       =   0  'User
   ScaleWidth      =   14643.57
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraInstitutionDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   210
      TabIndex        =   70
      Top             =   1530
      Width           =   7695
      Begin VB.CheckBox chkInstitution 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Caption         =   "Institution"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   30
         TabIndex        =   2
         Top             =   150
         Width           =   1305
      End
      Begin VB.TextBox txtInst 
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   3
         Top             =   150
         Width           =   2415
      End
      Begin VB.TextBox txtDesg 
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
         Left            =   5190
         MaxLength       =   50
         TabIndex        =   4
         Top             =   150
         Width           =   2325
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   71
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.TextBox txtInwardNo 
      Height          =   285
      Left            =   9960
      TabIndex        =   69
      Text            =   "0"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "&Reprint"
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
      Left            =   4590
      TabIndex        =   31
      Top             =   7650
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   14655
      TabIndex        =   59
      Top             =   30
      Width           =   14685
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   11640
         TabIndex        =   76
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblDateError 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "The date is not set in correct"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   11400
         TabIndex        =   72
         Top             =   355
         Width           =   3015
      End
      Begin VB.Label lblSeat 
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   65
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   12960
         TabIndex        =   64
         Top             =   120
         Width           =   435
      End
      Begin VB.Label lblsection 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seat:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblLoginName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   120
         TabIndex        =   62
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblLogin 
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   61
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lblLB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   6015
         TabIndex        =   60
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.TextBox txtPages 
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
      Left            =   9630
      MaxLength       =   8
      TabIndex        =   25
      Top             =   1110
      Width           =   1155
   End
   Begin VB.CheckBox chkCourtFee 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "CourtFee Stamp"
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
      Height          =   270
      Left            =   11250
      TabIndex        =   26
      Top             =   1110
      Width           =   2295
   End
   Begin VB.ComboBox cboPriority 
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
      Left            =   5580
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   2385
   End
   Begin VB.ComboBox cboInwardType 
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
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   2355
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   150
      TabIndex        =   43
      Top             =   4800
      Width           =   7935
      Begin VB.CheckBox chkByRefMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ref By Member"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   128
         Top             =   840
         Width           =   1995
      End
      Begin VB.CheckBox chkByOwner 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "By Owner"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   127
         Top             =   840
         Width           =   1965
      End
      Begin VB.ComboBox cboDept 
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
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1920
         Width           =   2805
      End
      Begin VB.TextBox txtSubject 
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
         Height          =   405
         Left            =   1530
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   390
         Width           =   6225
      End
      Begin VB.TextBox txtRefNo 
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
         Left            =   1530
         MaxLength       =   25
         TabIndex        =   20
         Top             =   1110
         Width           =   2205
      End
      Begin VB.ComboBox cboSeatID 
         Height          =   315
         Left            =   1530
         TabIndex        =   57
         Text            =   "cboSeatID"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txtDeliveryDate 
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
         Left            =   5610
         MaxLength       =   25
         TabIndex        =   24
         Top             =   2280
         Width           =   2145
      End
      Begin VB.TextBox txtRefDate 
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
         Left            =   5580
         MaxLength       =   25
         TabIndex        =   21
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtSubID 
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
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   18
         Top             =   450
         Width           =   435
      End
      Begin VB.ComboBox cboSeat 
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
         Left            =   5610
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1920
         Width           =   2205
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   78
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Subject Master"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   5010
         TabIndex        =   56
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Forward To and Delivery Date"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   0
         TabIndex        =   55
         Top             =   1530
         Width           =   7950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Delivery Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   54
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   53
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Ref. Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   52
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   49
         Top             =   487
         Width           =   885
      End
      Begin VB.Label lblOfficerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1530
         TabIndex        =   48
         Top             =   2310
         Width           =   240
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Subject and Reference Details"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   240
         Left            =   30
         TabIndex        =   47
         Top             =   120
         Width           =   7920
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4365
         TabIndex        =   46
         Top             =   1972
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Index           =   4
         Left            =   1260
         TabIndex        =   45
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   10
         Left            =   4830
         TabIndex        =   44
         Top             =   1920
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   7230
      TabIndex        =   28
      Top             =   7650
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&lose"
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
      Left            =   8550
      TabIndex        =   30
      Top             =   7650
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
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
      Height          =   375
      Left            =   5910
      TabIndex        =   29
      Top             =   7650
      Width           =   1215
   End
   Begin VB.Frame fraAppNameAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   150
      TabIndex        =   32
      Top             =   1080
      Width           =   7905
      Begin VB.TextBox txtDocumentProof 
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
         Left            =   5280
         MultiLine       =   -1  'True
         TabIndex        =   126
         Top             =   3360
         Width           =   2445
      End
      Begin VB.CheckBox chkScSt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "SC / ST"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   124
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox chkBpl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "B P L"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   123
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdTaxCheck 
         Appearance      =   0  'Flat
         Caption         =   "Tax"
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
         Left            =   3330
         TabIndex        =   122
         Top             =   1800
         Width           =   465
      End
      Begin VB.CheckBox chkInsideLB 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Inside LB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   75
         Top             =   143
         Width           =   2055
      End
      Begin VB.ComboBox cboState 
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
         Left            =   1320
         TabIndex        =   14
         Top             =   2595
         Width           =   2475
      End
      Begin VB.TextBox txtDoorNo1 
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
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1800
         Width           =   765
      End
      Begin VB.ComboBox cboDistrict 
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2595
         Width           =   2445
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   5280
         MaxLength       =   12
         TabIndex        =   17
         Top             =   3030
         Width           =   2445
      End
      Begin VB.TextBox txtPostoffice 
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
         Left            =   5280
         TabIndex        =   13
         Top             =   2205
         Width           =   2445
      End
      Begin VB.TextBox txtDoorNo2 
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
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1800
         Width           =   1185
      End
      Begin VB.TextBox txtHouseName 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1433
         Width           =   2475
      End
      Begin VB.TextBox txtSender 
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
         Left            =   2370
         TabIndex        =   6
         Top             =   1050
         Width           =   5385
      End
      Begin VB.ComboBox cboGender 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1035
         Width           =   1005
      End
      Begin VB.TextBox txtPincode 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2205
         Width           =   2475
      End
      Begin VB.TextBox txtPhoneno 
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
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   16
         Top             =   3030
         Width           =   2475
      End
      Begin VB.TextBox txtLocality 
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
         Left            =   5280
         TabIndex        =   11
         Top             =   1800
         Width           =   2445
      End
      Begin VB.TextBox txtWardNo 
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
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1433
         Width           =   2445
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   129
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label45 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Doc Proof"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4020
         TabIndex        =   125
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   " Name and Address of Applicant"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   30
         TabIndex        =   74
         Top             =   120
         Width           =   7860
      End
      Begin VB.Label lblState 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   73
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4020
         TabIndex        =   42
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4020
         TabIndex        =   41
         Top             =   2647
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post Office"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4020
         TabIndex        =   40
         Top             =   2242
         Width           =   1005
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   39
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   38
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   37
         Top             =   2242
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locality *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4020
         TabIndex        =   36
         Top             =   1837
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4020
         TabIndex        =   35
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   30
         TabIndex        =   34
         Top             =   1837
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name *"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   33
         Top             =   1087
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   8115
      TabIndex        =   27
      Top             =   1455
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   10769
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   803
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Valuables"
      TabPicture(0)   =   "frmSoochikaInward.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsValuable"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Malayalam Address"
      TabPicture(1)   =   "frmSoochikaInward.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(3)=   "Label28"
      Tab(1).Control(4)=   "Label29"
      Tab(1).Control(5)=   "Label30"
      Tab(1).Control(6)=   "Label31"
      Tab(1).Control(7)=   "Label32"
      Tab(1).Control(8)=   "cboCertGender"
      Tab(1).Control(9)=   "txtCertPincode"
      Tab(1).Control(10)=   "txtCertHouseName"
      Tab(1).Control(11)=   "txtCertLocality"
      Tab(1).Control(12)=   "txtCertWardNo"
      Tab(1).Control(13)=   "txtCertDoorNo2"
      Tab(1).Control(14)=   "txtCertDoorNo1"
      Tab(1).Control(15)=   "cboCertDist"
      Tab(1).Control(16)=   "txtCertPostOffice"
      Tab(1).Control(17)=   "txtCertName"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Checklist"
      TabPicture(2)   =   "frmSoochikaInward.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsEnclosure"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Others"
      TabPicture(3)   =   "frmSoochikaInward.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ward/Member Reference"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -74850
         TabIndex        =   106
         Top             =   4290
         Width           =   6075
         Begin VB.ComboBox cboMember 
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
            Left            =   3180
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   750
            Width           =   2775
         End
         Begin VB.ComboBox cboWard 
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
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   780
            Width           =   2835
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Member"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3210
            TabIndex        =   110
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ward"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   109
            Top             =   480
            Width           =   600
         End
      End
      Begin VB.TextBox txtCertName 
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
         Left            =   -72630
         MaxLength       =   100
         TabIndex        =   105
         Top             =   990
         Width           =   3705
      End
      Begin VB.TextBox txtCertPostOffice 
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
         Left            =   -70740
         TabIndex        =   104
         Top             =   2205
         Width           =   1815
      End
      Begin VB.ComboBox cboCertDist 
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
         Height          =   315
         Left            =   -70740
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   2625
         Width           =   1845
      End
      Begin VB.TextBox txtCertDoorNo1 
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
         Left            =   -70740
         MaxLength       =   8
         TabIndex        =   102
         Top             =   1814
         Width           =   705
      End
      Begin VB.TextBox txtCertDoorNo2 
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
         Left            =   -69990
         MaxLength       =   6
         TabIndex        =   101
         Top             =   1814
         Width           =   1065
      End
      Begin VB.TextBox txtCertWardNo 
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
         Left            =   -73530
         MaxLength       =   5
         TabIndex        =   100
         Top             =   1814
         Width           =   1665
      End
      Begin VB.TextBox txtCertLocality 
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
         Left            =   -73530
         TabIndex        =   99
         Top             =   2226
         Width           =   1665
      End
      Begin VB.TextBox txtCertHouseName 
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
         Left            =   -73530
         MaxLength       =   100
         TabIndex        =   98
         Top             =   1402
         Width           =   4605
      End
      Begin VB.TextBox txtCertPincode 
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
         Left            =   -73530
         MaxLength       =   6
         TabIndex        =   97
         Top             =   2640
         Width           =   1665
      End
      Begin VB.ComboBox cboCertGender 
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
         Height          =   315
         Left            =   -73530
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   960
         Width           =   825
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bill/Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -74880
         TabIndex        =   87
         Top             =   2190
         Width           =   6015
         Begin VB.ComboBox cboType 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   300
            Width           =   1905
         End
         Begin VB.TextBox txtAmt 
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
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   90
            Top             =   1080
            Width           =   1905
         End
         Begin VB.TextBox txtBillDescr 
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
            Left            =   1560
            TabIndex        =   89
            Top             =   1440
            Width           =   3075
         End
         Begin VB.TextBox txtBillNo 
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   88
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   95
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   94
            Top             =   1110
            Width           =   720
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   93
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bill/Receipt No."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   92
            Top             =   780
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registered Post"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   -74880
         TabIndex        =   80
         Top             =   540
         Width           =   6045
         Begin VB.TextBox txtRegPostDesg 
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
            Left            =   1500
            MaxLength       =   100
            TabIndex        =   83
            Top             =   750
            Width           =   3555
         End
         Begin VB.TextBox txtRegPostNo 
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
            Left            =   1500
            MaxLength       =   5
            TabIndex        =   82
            Top             =   1110
            Width           =   1245
         End
         Begin VB.TextBox txtRegPostToWhom 
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
            Left            =   1500
            MaxLength       =   100
            TabIndex        =   81
            Top             =   390
            Width           =   3555
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   86
            Top             =   780
            Width           =   1125
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Postal No."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   85
            Top             =   1110
            Width           =   960
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Whom"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   84
            Top             =   420
            Width           =   885
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsEnclosure 
         Height          =   4485
         Left            =   -74670
         TabIndex        =   111
         Top             =   690
         Width           =   5505
         _cx             =   9710
         _cy             =   7911
         Appearance      =   1
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaInward.frx":0070
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
         Begin VB.CheckBox chkAll 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   30
            TabIndex        =   112
            Top             =   0
            Width           =   225
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid vsValuable 
         Height          =   1035
         Left            =   60
         TabIndex        =   113
         Top             =   510
         Width           =   6225
         _cx             =   10980
         _cy             =   1826
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaInward.frx":00E4
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
         TabBehavior     =   1
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
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74790
         TabIndex        =   121
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71820
         TabIndex        =   120
         Top             =   1844
         Width           =   825
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74790
         TabIndex        =   119
         Top             =   1844
         Width           =   885
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locality"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74790
         TabIndex        =   118
         Top             =   2256
         Width           =   705
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74790
         TabIndex        =   117
         Top             =   2670
         Width           =   735
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74790
         TabIndex        =   116
         Top             =   1432
         Width           =   1200
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post Office"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71820
         TabIndex        =   115
         Top             =   2265
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71820
         TabIndex        =   114
         Top             =   2685
         Width           =   645
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   79
      Top             =   6360
      Width           =   90
   End
   Begin VB.Label lblLastinward 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11040
      TabIndex        =   77
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   " * "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1740
      TabIndex        =   68
      Top             =   810
      Width           =   255
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   " * "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   5190
      TabIndex        =   67
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Enclosures"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   300
      Left            =   8055
      TabIndex        =   66
      Top             =   780
      Width           =   6720
   End
   Begin VB.Label Label40 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No of Pages"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   8130
      TabIndex        =   58
      Top             =   1170
      Width           =   1410
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Priority "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4350
      TabIndex        =   51
      Top             =   810
      Width           =   840
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Correspondence "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   150
      TabIndex        =   50
      Top             =   810
      Width           =   1620
   End
End
Attribute VB_Name = "frmSoochikaInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intDistrID As Integer
Dim intFunID As Integer
Dim intRefID As Integer
Public lSoochikaFeildID As Variant
Dim SubjectID As Integer
Dim SeatCodeName  As String
Dim MainSubTypeID As Variant
Dim KioskID As Variant
Dim tnyTypeID As Variant

Public Property Let SevanaTypeID(intTypeID As Integer)      'ID for checking receipt exist or not
    tnyTypeID = intTypeID
End Property

Public Property Let SevanaKioskID(intKioskID As Integer)    'ID of sevana forward User
    KioskID = intKioskID
End Property

Public Property Let SevanaMainSubid(subID As Integer)       'ID of sevana Main SubID
    MainSubTypeID = subID
End Property

Private Sub cboDept_Click()
    If cboDept.ListIndex > -1 Then
        Call PopulateList(cboSeatID, "SELECT  intid,chvsection From tblSection  inner join TblUser on TblUser.fldUserID=tblSection.intCurrentUsr WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & cboDept.ItemData(cboDept.ListIndex) & " order by chvSection", , True, , True, enuSourceString.SOOCHIKA)
        Call PopulateList(cboSeat, "SELECT  chvsection,chvsection From tblSection  inner join TblUser on TblUser.fldUserID=tblSection.intCurrentUsr WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & cboDept.ItemData(cboDept.ListIndex) & " order by chvSection", , True, , True, enuSourceString.SOOCHIKA)
    End If
End Sub

Private Sub cboInwardType_Click()
    If (cboInwardType.ItemData(cboInwardType.ListIndex) >= 2 And cboInwardType.ItemData(cboInwardType.ListIndex) <= 4) Or cboInwardType.ItemData(cboInwardType.ListIndex) = 12 Then
        cboPriority.ListIndex = 1
    Else
        cboPriority.ListIndex = 4
    End If
End Sub
Private Sub cboState_Click()
    If gbLinkWithSevana = 1 Or gbLinkWithSevana = 2 Then
        PopulateList cboDistrict, "select chvName,intDistrictID from StateDistrict where intDistrictID<>0 and intstateid='" & cboState.ItemData(cboState.ListIndex) & "'", , , , True, enuSourceString.SevanaCommon
    End If
End Sub
Private Sub cboSeat_Click()
    cboSeatID.ListIndex = cboSeat.ListIndex
    getCurrentUser (cboSeatID.Text)
End Sub
Private Sub cboSeatID_Change()
    cboSeat.ListIndex = cboSeatID.ListIndex
End Sub
Private Sub cboWard_Click()
    cboMember.ListIndex = cboWard.ListIndex
End Sub
Private Sub chkAll_Click()
    Dim i As Integer
    If chkAll.Value = 1 Then
        For i = 1 To vsEnclosure.Rows - 1
            vsEnclosure.TextMatrix(i, 0) = 1
        Next i
    Else
        For i = 1 To vsEnclosure.Rows - 1
            vsEnclosure.TextMatrix(i, 0) = 0
        Next i
    End If
End Sub

Private Sub chkByOwner_Click()
    If chkByOwner.Value <> 0 Or chkByRefMember.Value <> 0 Then
        txtDeliveryDate.Text = Date
    Else
        getDeliveryDate
    End If
End Sub

Private Sub chkByRefMember_Click()
    If chkByRefMember.Value <> 0 Or chkByOwner.Value <> 0 Then
        txtDeliveryDate.Text = Date
    Else
        getDeliveryDate
    End If
End Sub

Private Sub chkCourtFee_Click()
'    If chkCourtFee.Value = 1 Then
'        SSTab1.TabVisible(1) = False
'    Else
'        SSTab1.TabVisible(1) = True
'    End If
End Sub
Private Sub chkInstitution_Click()
    If chkInstitution.Value = Checked Then
        txtInst.Enabled = True
        txtDesg.Enabled = True
        txtInst.SetFocus
        cboGender.ListIndex = 3
    Else
        txtInst.Enabled = False
        txtDesg.Enabled = False
        txtInst.Text = ""
        txtDesg.Text = ""
        cboGender.ListIndex = 0
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdNew_Click()
    ClearDetails
    cboGender.SetFocus
    Unload frmReceiptsCounter
End Sub
Private Sub cmdReprint_Click()
    If lSoochikaFeildID <> "" Then
        Ack (lSoochikaFeildID)
    Else
        MsgBox "Reprint is not possible"
    End If
End Sub
Private Sub cmdSave_Click()
    InwardMode = 0
    If lSaveValidate = True Then
        If gbLinkWithSevana = 1 Or gbLinkWithSevana = 2 Then       ' Check the sevana installation and user mapping
            Dim mCnn As New ADODB.Connection
            Dim objdb As New clsDB
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim mCnnSoochika As New ADODB.Connection
                            
            
            If txtSubID.Text <> "" Then
                If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
                    MsgBox "Connection not present", vbDefaultButton1
                    Exit Sub
                End If
                mSql = "Select isnull(intMainSubID,'0') as intMainSubID from TblSubjectcoding where intsubID= " & frmSoochikaInward.txtSubID
                Rec.Open mSql, mCnn
                If Not (Rec.BOF Or Rec.EOF) Then
                    If Rec!intMainSubID <> 0 Then
                        frmSevanaInward.Show vbModal
                        'frmReceiptsCounter.Visible = True
                        'frmReceiptsCounter.ZOrder (0)
                    Else
                        If txtInwardNo = 0 Then
                            objdb.CreateNewConnection mCnnSoochika, enuSourceString.SOOCHIKA
                            mCnnSoochika.BeginTrans
                            On Error GoTo RollData1
                                Call SaveSoochika(mCnnSoochika)
                            mCnnSoochika.CommitTrans
                            Ack (lSoochikaFeildID)
                            GoTo Clear1
                            'MsgBox "Soochika Sucessfully"
RollData1:
                            mCnnSoochika.RollbackTrans
                            MsgBox Error$, vbCritical, "SOOCHIKA ERROR"
Clear1:                            ClearDetails
                        End If
                    End If
                End If
                Rec.Close
            Else
                If txtInwardNo = 0 Then
                    objdb.CreateNewConnection mCnnSoochika, enuSourceString.SOOCHIKA
                    mCnnSoochika.BeginTrans
                            On Error GoTo RollData2
                                Call SaveSoochika(mCnnSoochika)
                            mCnnSoochika.CommitTrans
                            Ack (lSoochikaFeildID)
                            GoTo Clear2
                            'MsgBox "Soochika Sucessfully"
RollData2:
                            mCnnSoochika.RollbackTrans
                            MsgBox Error$, vbCritical, "SOOCHIKA ERROR"
Clear2:                            ClearDetails
                End If
            End If
        Else
            If txtInwardNo = 0 Then
                objdb.CreateNewConnection mCnnSoochika, enuSourceString.SOOCHIKA
                 mCnnSoochika.BeginTrans
                            On Error GoTo RollData3
                                Call SaveSoochika(mCnnSoochika)
                            mCnnSoochika.CommitTrans
                            Ack (lSoochikaFeildID)
                            GoTo Clear3
                            'MsgBox "Soochika Sucessfully"
RollData3:
                            mCnnSoochika.RollbackTrans
                            MsgBox Error$, vbCritical, "SOOCHIKA ERROR"
Clear3:                            ClearDetails
            End If
        End If
        
        'cmdSave.Enabled = False
        'ClearDetails
    End If
End Sub
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtDeliveryDate.Text = ""
End Sub
'Private Sub dtpDelivery_Change()
'    If dtpDelivery.Value > Date Then
'       dtpDelivery.Value = Date
'    End If
'End Sub
Private Sub chkInsideLB_Click()
    If chkInsideLB.Value = 1 Then
        cboState.Text = "Kerala"
        cboDistrict.Enabled = False
        cboState.Enabled = False
        'cboState.ListIndex = 30
        cboDistrict.ListIndex = gbDistID - 1
    Else
        cboDistrict.Enabled = True
        cboState.Enabled = True
    End If
End Sub

Private Sub cmdTaxCheck_Click()
    If txtWardNo.Text = "" Then
        MsgBox "Please enter ward no"
        Exit Sub
    ElseIf txtDoorNo1.Text = "" Then
        MsgBox "Please enter door no 1"
        Exit Sub
    End If
    If txtDoorNo2.Text = "" Then txtDoorNo2.Text = 0
'    frmSoochikaBuildingDetails.txtWardNo.Text = Val(txtWardNo.Text)
'    frmSoochikaBuildingDetails.txtDoorNo1.Text = txtDoorNo1.Text
'    frmSoochikaBuildingDetails.txtDoorNo2.Text = txtDoorNo2.Text
    Load frmSoochikaBuildingDetails
    frmSoochikaBuildingDetails.Show vbModal
End Sub

Private Sub Form_Activate()
    Me.Left = 0
    Me.Top = 0
    Me.Height = 8625
    
End Sub

Private Function CheckSoochikaSettings() As Boolean
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
        MsgBox "Connection not Present", vbInformation
        Unload Me
    End If
    
    CheckSoochikaSettings = True
    mSql = "select * from TblUser inner join Tblsection on Tblsection.intCurrentUSR=tblUser.FldUserid where TblUser.FldUserID=" & gbUserID & " and tblsection.intID=" & gbSeatID & " and (tblUser.fldTypeID=6  or(tbluser.FldTypeID=5 and tblUser.flgClerical=1))"
    Rec.Open mSql, mCnn
    If Rec.EOF Or Rec.BOF Then
        CheckSoochikaSettings = False
    ElseIf Rec!FldTypeID <> 6 And Rec!FldTypeID <> 5 Then
        CheckSoochikaSettings = False
    End If
    Rec.Close
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub Form_Load()
    SetSoochkaEnvironment
    If (SetNewInwardNo = True) Then
        frmSoochikaStartup.Show (1)
        Unload frmSoochikaInward
    End If
    lblLogin.Caption = gbShortname
    lblSeat.Caption = gbSeat
    lblDate.Caption = Format(Date, "DD/MM/YYYY")
    checkDate                   'Check for date for validation
    GetlastInward               'Get last inward number
    lblLB.Caption = gbLBName
    If CheckSoochikaSettings = False Then
        MsgBox "You have no Permission for accepting any Inward,please contact your System Administrator", vbInformation
        cmdSave.Enabled = False
        Exit Sub
    End If
    Call PopulateList(cboPriority, "Select chvPriority, bntPriorityId from tblPriority", , , , True, enuSourceString.SOOCHIKA)
    cboPriority.ListIndex = 4
    Call PopulateList(cboInwardType, "SELECT chvInwardType,intInwardType FROM TblInwardType where intCategory=1", , , , True, enuSourceString.SOOCHIKA)
    cboInwardType.ListIndex = 0
    getDeptID
    Call PopulateList(cboGender, "SELECT chvCode,IntGenderID FROM TblGender", , , , True, enuSourceString.SOOCHIKA)
    Call PopulateList(cboCertGender, "SELECT chvMalCode, IntGenderID From TblGender", , , , True, enuSourceString.SOOCHIKA)
    
    If gbLinkWithSevana = 1 Or gbLinkWithSevana = 2 Then
        Call PopulateList(cboState, "select ChvName,intstateID from StateDistrict where intDistrictID=0", , , , True, enuSourceString.SevanaCommon)
        Call PopulateList(cboDistrict, "select ChvName,intDistrictID from StateDistrict where intDistrictID<>0 and intstateid='32'", , , , True, enuSourceString.SevanaCommon)
    Else
        Call PopulateList(cboState, "select ChvName,intstateID from TB_State_MST where intDistrictID=0", , , , True, enuSourceString.SOOCHIKA)
        Call PopulateList(cboDistrict, "SELECT chvEngDistName, intID From dbo.TB_District_MST", , , , True, enuSourceString.SOOCHIKA)
    End If
    Call PopulateList(cboCertDist, "SELECT chvDistName, intID From dbo.TB_District_MST", , , , True, enuSourceString.SOOCHIKA)
    Call PopulateList(cboWard, "SELECT chvWardName,intId From tblWard Order by intID", , , , True, enuSourceString.SOOCHIKA)
    Call PopulateList(cboMember, "SELECT chvMember,intId From tblWard Order by intID", , , , True, enuSourceString.SOOCHIKA)
    cboState.ListIndex = 31 'Modified on 29/11/12 (sevana commin index issue)
    cboDistrict.ListIndex = gbDistID - 1
    cboCertDist.ListIndex = gbDistID - 1
    cboCertGender.ListIndex = 0
    cboGender.ListIndex = 0
    Call PopulateList(cboDept, "spSelectDepartment", , True, , True, enuSourceString.SOOCHIKA)
    'Call PopulateList(cboSeatID, "SELECT  intid,chvsection From tblSection  inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & gbDeptID & " order by intID", , True, , True, enuSourceString.SOOCHIKA)
    'Call PopulateList(cboSeat, "SELECT  chvsection,chvsection From tblSection  inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & gbDeptID & " order by intID", , True, , True, enuSourceString.SOOCHIKA)
    Call FillFlexGridCombo(vsValuable, 0, "SELECT intInstrumentType,chvInstrument From TblInstrument", adCmdText, enuSourceString.SOOCHIKA)
    gSubSetFont vsValuable, 1, 0, vsValuable.Rows - 1, 0, "Verdana"
    gSubSetFont vsValuable, 1, 4, vsValuable.Rows - 1, 4, "Verdana"
    Call PopulateList(cboType, "SelectType", , True, , , enuSourceString.SOOCHIKA)
    If (lInitialise = True) Then
        frmSoochikaStartup.Show (1)
    End If
    chkInsideLB.Value = 1
End Sub
Private Function lInitialise() As Boolean
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mCount As Integer
        
    If objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False Then
        MsgBox "Cannot Continue.., Connection not present", vbInformation, "Soochika"
        Exit Function
    End If
    mSql = "SELECT * FROM TblReason WHERE FlgReason=1"
    Rec.Open mSql, mCnn
    If (Rec.EOF And Rec.BOF) Then
        lInitialise = False
    Else
        lInitialise = True
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Function SetNewInwardNo() As Boolean
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mCount As Integer
    Dim FileID As Variant
    Dim ZonalID As String
    
    ZonalID = Right(gbnumZonalID, 1)
    FileID = gbLBID & "0" & ZonalID & Year(Date) & "000000"
        
    If objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False Then
        MsgBox "Cannot Continue.., Connection not present", vbInformation, "Soochika"
        Exit Function
    End If
    mSql = "SELECT max(year(flddateofreceipt)) as MaxYear FROM TblTappalDetails"
    Rec.Open mSql, mCnn
    If IsNull(Rec!MaxYear) = True Then
        SetNewInwardNo = True
    ElseIf Rec!MaxYear < Year(Date) Then
        mSql = "set identity_insert FileIDCreation on "
        mSql = mSql & " insert into FileIDCreation (numFileID ,dtDate)values(" & FileID & ",getdate())"
        mSql = mSql & " set identity_insert FileIDCreation off"
        mCnn.Execute mSql
    Else
        SetNewInwardNo = False
    End If
    Rec.Close
    
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Function
Private Sub GetlastInward()
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
    Set Rec = objdb.ExecuteSP("select right(Max(fldFileid),6) as InwardNo from tbltappaldetails", , , , mCnn, adCmdText)
    If Not (Rec.EOF Or Rec.BOF) Then
        lblLastinward.Caption = "Last inward is : " & Rec!InwardNo
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Public Sub ClearDetails()
    txtSender.Text = ""
    txtDoorNo1.Text = ""
    txtWardNo.Text = ""
    txtHouseName.Text = ""
    txtDoorNo2.Text = ""
    txtLocality.Text = ""
    cboDistrict.ListIndex = gbDistID - 1
    cboInwardType.ListIndex = 0
    cboPriority.ListIndex = 4
    txtPostoffice.Text = ""
    txtPincode.Text = ""
    txtPhoneno.Text = ""
    txtSubID.Text = ""
    txtSubject.Text = ""
    txtRefDate.Text = ""
    txtRefNo.Text = ""
    txtDeliveryDate.Text = ""
    cboSeat.ListIndex = -1
    cboDept.ListIndex = -1
    chkCourtFee.Value = 0
    chkAll.Value = 0
    vsEnclosure.Clear 1
    vsEnclosure.Rows = 2
    vsValuable.Clear 1
    vsValuable.Rows = 2
    cboCertGender.ListIndex = 0
    cboGender.ListIndex = 0
    txtCertDoorNo1.Text = ""
    txtCertDoorNo2.Text = ""
    txtCertHouseName.Text = ""
    txtCertLocality.Text = ""
    txtCertName.Text = ""
    txtCertPincode.Text = ""
    txtCertPostOffice.Text = ""
    txtCertWardNo.Text = ""
    txtDeliveryDate.Text = ""
    txtRegPostDesg.Text = ""
    txtRegPostNo.Text = ""
    txtRegPostToWhom.Text = ""
    cboType.ListIndex = 0
    txtBillDescr.Text = ""
    txtBillNo.Text = ""
    txtAmt.Text = ""
    txtPages.Text = ""
    txtInst.Text = ""
    txtDesg.Text = ""
    txtEmail.Text = ""
    txtInwardNo.Text = 0
    chkInstitution.Value = 0
    gSubSetFont vsValuable, 1, 0, vsValuable.Rows - 1, 0, "Verdana"
    gSubSetFont vsValuable, 1, 4, vsValuable.Rows - 1, 4, "Verdana"
    chkInsideLB.Value = 1
    chkByOwner.Value = 0
    chkByRefMember.Value = 0
    chkByOwner.Enabled = False
    chkByRefMember.Enabled = False
    GetlastInward               'Get last inward number
    txtSender.SetFocus
    MainSubTypeID = 0
    
End Sub
Public Sub checkDate()
    Dim mCnn As New ADODB.Connection
    Dim mSql As String
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim MaxDate As Variant
    
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
        MsgBox "Connection is not present", vbCritical, "Soochika"
        Exit Sub
    End If
    
    mSql = "select max(flddateofreceipt)as Maxdate from TblTappalDetails "
    Rec.Open mSql, mCnn
    If Not (Rec.BOF Or Rec.EOF) Then
        MaxDate = Rec!MaxDate
    End If

    'If lblDate < DateValue(MaxDate) Then
    If DateValue(Date) < DateValue(MaxDate) Then
        cmdSave.Enabled = False
        lblDateError.Visible = True
        MsgBox "Error on setting the current date", vbInformation, "Error Validation"
    ElseIf IsDate(Format(lblDate.Caption, "DD/MM/YYYY")) = False Then
        cmdSave.Enabled = False
        lblDateError.Visible = True
        MsgBox "Error on setting the current date format", vbInformation, "Error Validation"
    Else
        cmdSave.Enabled = True
        lblDateError.Visible = False
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub

Private Sub Label23_Click()
gbSubID = 1
    frmSoochikaSubjectMaster.Show
End Sub
Private Sub lblLogout_Click()
    Unload Me
End Sub
Private Sub txtHouseWard_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 46 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCertDoorNo1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub
Private Sub txtCertWardNo_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub
Private Sub txtDeliveryDate_LostFocus()
    If (txtDeliveryDate.Text <> "") Then
        If (gFunIsDMYDateBoolean(txtDeliveryDate.Text) = False) Then
            MsgBox "Check the date "
            txtDeliveryDate.SetFocus
        End If
        If (CDate(txtDeliveryDate.Text) < Date) Then
            MsgBox "Delivery date should not be less than current date"
            txtDeliveryDate.SetFocus
        End If
    End If
End Sub
Private Sub txtDesg_LostFocus()
    If txtSender.Text = "" Then
        txtSender.Text = txtDesg.Text
    End If
End Sub

Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtPages_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub
Private Sub txtPages_LostFocus()
    If val(txtPages.Text) >= 1 Then
        Call FillEnclosure
    End If
End Sub
Private Sub txtPhoneno_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Public Function SaveSoochika(ByRef mCnnSoochika As ADODB.Connection) As Long
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim ForwardTo As Variant
 Dim mVarrReceipt As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 Dim mCnn As New ADODB.Connection
 Dim mCnnSevana As New ADODB.Connection
 
 
 
 ReDim mVarrIn(45)
    mVarrIn(0) = 0 'FldCurrentNo.
    mVarrIn(1) = Date  'FldDateOfReceipt.
    mVarrIn(2) = txtSender.Text 'FldSenderName.
    mVarrIn(3) = txtWardNo.Text 'FldWardNo.
    mVarrIn(4) = txtDoorNo2.Text  'FldHouseNo
    mVarrIn(5) = txtLocality.Text  'FldLocality
    mVarrIn(6) = cboDistrict.ItemData(cboDistrict.ListIndex) 'FldDistrict
    mVarrIn(7) = gbnumSeatID 'bntCurrUserId.
    mVarrIn(8) = cboSeatID.Text 'cboSeat.ItemData(cboSeat.ListIndex)  'intForwardTo.
    mVarrIn(9) = cboInwardType.ItemData(cboInwardType.ListIndex)  'intInwardType.
    mVarrIn(10) = cboPriority.ItemData(cboPriority.ListIndex)  'FldPriority
    mVarrIn(11) = Date 'dtmForwardDate
    mVarrIn(12) = txtSubject.Text   'FldRemarks
    mVarrIn(13) = Null 'intAttachmentType
    mVarrIn(14) = Null 'FldManualSummary
    mVarrIn(15) = Null 'FldElectronicsSummary
    mVarrIn(16) = cboDept.ItemData(cboDept.ListIndex) 'intDept
    mVarrIn(17) = chkCourtFee.Value   'FlgCourFeeStamp
    mVarrIn(18) = val(txtPages.Text)   'intManualPage
    mVarrIn(19) = txtRefNo.Text   'FldOutsideNo
    If (txtRefDate.Text = "") Then
        mVarrIn(20) = Null   'FldRefDate
    Else
        mVarrIn(20) = txtRefDate.Text   'FldRefDate
    End If
    mVarrIn(21) = Null 'intRegPost
    mVarrIn(22) = chkInstitution.Value  'bitInstflg
    mVarrIn(23) = txtInst.Text  'fldInstName
    mVarrIn(24) = txtDesg.Text  'fldDesign
    mVarrIn(25) = txtPostoffice.Text  'FldPostOffice
    mVarrIn(26) = txtPincode.Text 'FldPin
    mVarrIn(27) = txtEmail.Text  'FldEmail
    mVarrIn(28) = txtPhoneno.Text 'FldPhone
    mVarrIn(29) = txtRegPostToWhom.Text   'fldReglttoWhom
    mVarrIn(30) = txtRegPostDesg.Text   'fldReglttoDesign
    mVarrIn(31) = txtRegPostNo.Text   'fldRegltpoNo
    mVarrIn(32) = Null 'sessionID
    mVarrIn(33) = Null 'intBillRecFlg
    mVarrIn(34) = Null 'intInsideLBFlg
    mVarrIn(35) = txtHouseName.Text  'FldHouseName
    mVarrIn(36) = Null 'intCertAddrFlg
    mVarrIn(37) = cboGender.ItemData(cboGender.ListIndex)   'intGender
    mVarrIn(38) = val(txtDoorNo1.Text) 'intDoorNo
    mVarrIn(39) = 0 'InwardFlg
    mVarrIn(40) = gbLBID   'Lb id
    mVarrIn(41) = gbnumZonalID   'ZoalID
    mVarrIn(42) = gbSuitID  'Suit
    mVarrIn(43) = txtSubID.Text   'Subject ID
    If (txtDeliveryDate.Text = "") Then
        mVarrIn(44) = Null 'Delivery Date
    Else
        mVarrIn(44) = CDate(txtDeliveryDate.Text)   'Delivery Date
    End If
    If MainSubTypeID <> 0 Then
        mVarrIn(45) = frmSevanaInward.txtSubTypeID.Text
        mVarrIn(42) = 102
    Else
        mVarrIn(42) = 105
        mVarrIn(45) = Null
    End If
    
    'objDb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
    'mCnn.BeginTrans
'    If gbLinkWithSevana = 1 Or gbLinkWithSevana = 2 Then
'        objDb.CreateNewConnection mCnnSevana, enuSourceString.SevanaRegn
'        mCnnSevana.BeginTrans
'    End If
    
    objdb.ExecuteSP "spSaveInward", mVarrIn, mVarrOut, , mCnnSoochika, adCmdStoredProc
    txtInwardNo.Text = CDbl(Right(mVarrOut(0, 0), 6))
    If IsArray(mVarrOut) Then
       lSoochikaFeildID = mVarrOut(0, 0)
       SaveValuable mCnnSoochika
       If (txtCertName.Text <> "") Then
            SaveCertificateAddress mCnnSoochika
       End If
       SaveEnclosure mCnnSoochika
       SaveKeywords mCnnSoochika
       SaveGeneralInwardDetails mCnnSoochika
       If (txtAmt.Text <> "") Then
            SaveBillReceipt mCnnSoochika
       End If
       If (txtSubID.Text <> "") Then
            If txtSubID.Text = 264 Then
                SaveRCOwner mCnnSoochika
            End If
       End If
       '==================================================='
       
       
''       If MainSubTypeID <> 0 And (gbLinkWithSevana = 1 Or gbLinkWithSevana = 2) Then
''            Call SaveSevana(mCnnSevana)   'Save SevanaInward
''            If tnyTypeID = 1 Or tnyTypeID = 2 Then
''                ReDim mVarrReceipt(6)
''                mVarrReceipt(0) = lSoochikaFeildID
''                mVarrReceipt(1) = Right(lSoochikaFeildID, 6)
''                mVarrReceipt(2) = Date
''                mVarrReceipt(3) = frmSevanaInward.txtReceiptNo.Text
''                mVarrReceipt(4) = frmSevanaInward.DTPReceiptDate.Value
''                mVarrReceipt(5) = frmSevanaInward.txtReceiptAmount.Text
''                mVarrReceipt(6) = frmSevanaInward.txtReceiptBookNo.Text
''                objDb.ExecuteSP "spSaveInwardReceipt", mVarrReceipt, , , mCnn, adCmdStoredProc
''            End If
''       End If
       
'       mCnn.CommitTrans
'       If gbLinkWithSevana = 1 Then
'        mCnnSevana.CommitTrans
'       End If

'       Ack (lSoochikaFeildID)
'       cmdSave.Enabled = False
'       Unload frmSevanaInward
    End If
        SaveSoochika = Right(lSoochikaFeildID, 6)
        Exit Function
'MsgBox "Error on Saving", vbInformation, "Information"
'    mCnn.RollbackTrans
'    mCnnSevana.RollbackTrans
'    Unload frmSevanaInward
End Function

Public Function SaveSevana(ByVal InwordNo As Long, ByVal mReceiptNo As Variant, mAmt As Double, ByRef mCnn As Connection) As Boolean
    
    Dim mVarrIn As Variant
    Dim mVarrOut As Variant
    Dim ForwardTo As Variant
    Dim mVarrReceipt As Variant
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    ReDim mVarrIn(25)
    
    ''Rec.Open "select intsevanaModuleUserID from mUserMap where intSoochikaUserID=" & gbFldUserId, mCnn
    'If Not (Rec.EOF Or Rec.BOF) Then
      '  mVarrIn(23) = Rec!intSevanaModuleUserID             'Sevana UserID
    'Else
        mVarrIn(23) = 0
   ' End If
    'Rec.Close
    mVarrIn(0) = InwordNo 'Right(lSoochikaFeildID, 6)                 'Inward No
    mVarrIn(1) = Format(Date, "DD/MM/YYYY") ' lblDate.Caption                            'Inward Date
    mVarrIn(2) = MainSubTypeID                              'Main Sub ID
    If frmSevanaInward.cboHospitals.ListIndex >= 0 Then     'Hospital
        mVarrIn(3) = frmSevanaInward.cboHospitals.ItemData(frmSevanaInward.cboHospitals.ListIndex)
    Else
        mVarrIn(3) = 0
    End If
    mVarrIn(4) = KioskID                                    'Forward To
    mVarrIn(5) = Format(frmSevanaInward.DTPApplDate.Value, "DD/MM/YYYY")         'Application Date
    If txtWardNo.Text = "" Then                             'Ward No
        mVarrIn(6) = 0
    Else
        mVarrIn(6) = txtWardNo.Text
    End If
    mVarrIn(7) = txtLocality.Text                           'Place(Locality)
    If txtDoorNo1.Text = "" Then                            'House No
        mVarrIn(8) = ""
    Else
        mVarrIn(8) = IIf(IsNull(txtDoorNo1.Text), 0, txtDoorNo1.Text) & "/" & IIf(IsNull(txtDoorNo2.Text), "", txtDoorNo2.Text) 'House Number
    End If
    mVarrIn(9) = txtHouseName.Text                          'House Name
    mVarrIn(10) = ""                                        'Street Name
    mVarrIn(11) = ""                                        'Via
    mVarrIn(12) = 0 'txtPostoffice.Text                     'Postoffice
    mVarrIn(13) = 0                                         'Village
    mVarrIn(14) = txtSender.Text                            'Name of Applicant
    mVarrIn(15) = 0                                         'Taluk
    mVarrIn(16) = cboDistrict.ItemData(cboDistrict.ListIndex) 'District
    mVarrIn(17) = 0 'cboState.ItemData(cboState.ListIndex)     'State
    mVarrIn(18) = 0                                         'Care off ID
    mVarrIn(19) = frmSevanaInward.cboSubType.ItemData(frmSevanaInward.cboSubType.ListIndex) 'SubTypeID
    If chkInsideLB.Value = 1 Then
        mVarrIn(20) = chkInsideLB.Value                      'Polocn
    Else
        mVarrIn(20) = 2
    End If
    mVarrIn(21) = ""                                        'Covering Letter
    frmSevanaInward.txtRemarks.Text = "Data entered by " & gbUserName & ". " & frmSevanaInward.txtRemarks.Text
    If frmSevanaInward.chkZonal.Value = 1 Then
        mVarrIn(22) = "Inward from Zonal office " & frmSevanaInward.txtRemarks.Text
    Else
        mVarrIn(22) = frmSevanaInward.txtRemarks.Text           'Remarks
    End If
    mVarrIn(24) = ""                                        'Careoff Name
    mVarrIn(25) = 0                                         'Inward sequential flag
    
    objdb.ExecuteSP "spSaveInwardFromSoochika", mVarrIn, mVarrOut, , mCnn, adCmdStoredProc
    
    If tnyTypeID = 1 Or tnyTypeID = 2 Then
        ReDim mVarrIn(13)
        'Rec.Open "select intid from tInward where inWno=" & Right(lSoochikaFeildID, 6) & " order by intID desc", mCnn
        '''Rec.Open "select intid from tInward where inWno=" & InwordNo & " order by intID desc", mCnn
        '''
        '''If Not (Rec.EOF Or Rec.BOF) Then
        '''    mVarrIn(0) = Rec!intID                              'IntID from tInward
        '''End If
        '''Rec.Close
        mVarrIn(0) = mVarrOut(0, 0)     'IntID from tInward
        mVarrIn(1) = Format(Date, "DD/MM/YYYY")     'Receipt No
        mVarrIn(2) = 0      'Receipt Book
        mVarrIn(3) = mReceiptNo       'Receipt No
        mVarrIn(4) = mAmt      'Receipt Amount
        If MainSubTypeID = 5 Then
            mVarrIn(5) = frmSevanaInward.txtNoCopeis.Text
        Else
            mVarrIn(5) = frmSevanaInward.txtNoOfCertificate.Text           'No of copies
        End If
        mVarrIn(6) = frmSevanaInward.txtEnglishname.Text        'English Name
        mVarrIn(7) = frmSevanaInward.txtMalayalamname.Text      'Malayalam Name
        If frmSevanaInward.cboRelationship.ListIndex > -1 Then
            mVarrIn(8) = frmSevanaInward.cboRelationship.ItemData(frmSevanaInward.cboRelationship.ListIndex) 'CFM
        Else
            mVarrIn(8) = Null
        End If
        mVarrIn(9) = frmSevanaInward.cboLanguage.ItemData(frmSevanaInward.cboLanguage.ListIndex) 'Language
        mVarrIn(10) = InwordNo                'Inward No
        mVarrIn(11) = 0 'gbFldUserId                            'Issue User
        mVarrIn(12) = frmSevanaInward.txtRegNo.Text             'Register No
        mVarrIn(13) = frmSevanaInward.txtBookNo.Text            'Book No
        
        objdb.ExecuteSP "InsertReceiptDetails", mVarrIn, , , mCnn, adCmdStoredProc
        
        SaveSevana = True
    End If
End Function

Public Sub Ack(numFileID As Variant)
    Dim vAryInRpt(1)
    vAryInRpt(0) = CStr(numFileID)
    frmCRViewer.vShowReport App.Path & "\soochika\Reports", "AckSlip.rpt", vAryInRpt
    frmCRViewer.Show 1
End Sub
Private Sub getCurrentUser(numSeatID As Variant)
    Dim mSql As String
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    mSql = "SELECT dbo.TblUser.FldShortName, dbo.TblUser.FldUserId"
    mSql = mSql + " FROM dbo.TblUser INNER JOIN dbo.TblSection ON dbo.TblUser.FldUserId = dbo.TblSection.intCurrentUSR "
    mSql = mSql + " WHERE(ISNULL(dbo.TblUser.flgStatus, 0) <> 1) AND (dbo.TblSection.intID = " & numSeatID & ")"
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
        MsgBox "Soochika Connection is not present", vbCritical, "Common"
        Exit Sub
    End If
    If numSeatID <> "" Then
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            lblOfficerName.Caption = Rec!FldShortName
        Else
            lblOfficerName.Caption = ""
        End If
        Rec.Close
    Else
        lblOfficerName.Caption = ""
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub getDeptID()
    Dim mSql As String
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim varyOut As Variant
    objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
    Set Rec = objdb.ExecuteSP("spSelectDepartment", , varyOut, , mCnn, adCmdStoredProc)
    If IsArray(varyOut) Then
        gbDeptID = varyOut(1, 0)
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub getSubject()
    Dim mSql As String
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    mSql = "SELECT     chvSubject, intDistrID,intFunID,intRefID FROM TblSubjectCoding where intSubID=" & SubjectID & " "
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
        MsgBox "Soochika Connection is not present", vbCritical, "Common"
        Exit Sub
    End If
    If CStr(SubjectID) <> "" Then
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            txtSubject.Text = Rec!chvSubject
            intDistrID = IIf(IsNull(Rec!intDistrID), 0, Rec!intDistrID)
            intFunID = IIf(IsNull(Rec!intFunID), 0, Rec!intFunID)
            intRefID = IIf(IsNull(Rec!intRefID), 0, Rec!intRefID)
        Else
            txtSubject.Text = ""
        End If
        Rec.Close
    Else
        txtSubject.Text = ""
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub txtPincode_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub getDeliveryDate()
    Dim mSql As String
    Dim Period As Integer
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim HRec As New ADODB.Recordset
    mSql = "SELECT     intDeliveryPeriod FROM TblSubjectDeliveryPeriod where intSubjectID=" & SubjectID & " "
    If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
        MsgBox "Soochika Connection is not present", vbCritical, "Common"
        Exit Sub
    End If
    If CStr(SubjectID) <> "" Then
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            Period = IIf(IsNull(Rec!intDeliveryPeriod), 0, Rec!intDeliveryPeriod)
            txtDeliveryDate.Text = DateAdd("d", Period, Date)
                mSql = "SELECT * From TB_Holiday_MST Where dtDate=convert(datetime,'" & txtDeliveryDate.Text & "',103)"
                 HRec.Open mSql, mCnn
                    If Not (HRec.EOF And HRec.BOF) Then
                        MsgBox " Delivery Date fall on a holiday please enter the next working day"
                        txtDeliveryDate.SetFocus
                    End If
                 HRec.Close
        Else
            txtDeliveryDate.Text = ""
        End If
        Rec.Close
    Else
        txtDeliveryDate.Text = ""
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Sub getSeatCode()
    Dim mSql As String
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim varyOut As Variant
    Dim mVarrIn As Variant
    Dim i As Integer
    ReDim mVarrIn(1)
    objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
    mVarrIn(0) = SubjectID
    mVarrIn(1) = val(txtWardNo.Text)
    Set Rec = objdb.ExecuteSP("sp_SeatCoding", mVarrIn, varyOut, , mCnn, adCmdStoredProc)
    If IsArray(varyOut) Then
        For i = 0 To cboDept.ListCount - 1
            If cboDept.ItemData(i) = varyOut(1, 0) Then
                cboDept.ListIndex = i
            End If
        Next
        Call PopulateList(cboSeatID, "SELECT  intid,chvsection From tblSection inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & varyOut(1, 0) & " order by intID", , True, , True, enuSourceString.SOOCHIKA)
        Call PopulateList(cboSeat, "SELECT  chvsection,chvsection From tblSection inner join TblUser on TblUser.intSection=tblSection.intid WHERE (FldTypeID=6 or FldTypeID=5) and intDeptId = " & varyOut(1, 0) & " order by intID", , True, , True, enuSourceString.SOOCHIKA)
        For i = 0 To cboSeat.ListCount - 1
            If cboSeat.List(i) = varyOut(2, 0) Then
                cboSeat.ListIndex = i
                cboSeatID.ListIndex = i
            End If
        Next
        'SeatCodeName = varyOut(2, 0)
        'cboSeat.Text = SeatCodeName
    Else
        Call PopulateList(cboDept, "spselectdepartment", , False, True, True, enuSourceString.SOOCHIKA)
        cboSeat.Clear
        cboSeatID.Clear
        lblOfficerName.Caption = ""
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub
Private Function lSaveValidate() As Boolean
    lSaveValidate = True
    If txtSender.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter Name"
        txtSender.SetFocus
        Exit Function
    ElseIf txtLocality.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter Locality"
        txtLocality.SetFocus
        Exit Function
    ElseIf cboDistrict.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select District"
        cboDistrict.SetFocus
        Exit Function
    ElseIf txtSubject.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter Subject"
        txtSubject.SetFocus
        Exit Function
    ElseIf cboSeat.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select Seat"
        cboSeat.SetFocus
        Exit Function
    ElseIf cboPriority.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select Priority"
        cboPriority.SetFocus
        Exit Function
    ElseIf cboInwardType.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select Inward Type"
        cboInwardType.SetFocus
        Exit Function
    ElseIf cboGender.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select gender in address"
        cboGender.SetFocus
        Exit Function
'    ElseIf chkBpl.Value = 1 Or chkScSt.Value = 1 Then
'        If txtDocumentProof.Text = "" Then
'            lSaveValidate = False
'            MsgBox "Please enter the document proof for the "
'        End If
    End If
End Function

Private Sub txtPincode_LostFocus()
    If Trim(txtPincode.Text) <> "" And (gbLinkWithSevana = 1 Or gbLinkWithSevana = 2) Then
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim objdb As New clsDB
        Dim Rec As New ADODB.Recordset
        
        mSql = "select chvEngPostofficeName,intID as PostID from TB_PostOfficeAll_MST where intPincode=" & txtPincode.Text
        
        If objdb.CreateNewConnection(mCnn, enuSourceString.SevanaCommon) = False Then
            MsgBox "Sevana Connection Failure", vbInformation, "Failure"
            Exit Sub
        End If
        
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            txtPostoffice.Text = Rec!chvEngpostofficeName
        Else
            MsgBox "Invalid pincode ", vbInformation, "Failure"
            txtPincode.Text = ""
        End If
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
    End If
End Sub

Private Sub txtRefDate_LostFocus()
    If (txtRefDate.Text <> "") Then
        If (gFunIsDMYDateBoolean(txtRefDate.Text) = False) Then
            MsgBox "Check the date "
            txtRefDate.SetFocus
        End If
        If (CDate(txtRefDate.Text) > Date) Then
            MsgBox "Reference date should be less than current date"
            txtRefDate.SetFocus
        End If
    End If
End Sub
Private Sub txtSender_LostFocus()
    If CheckSoochikaSettings = True Then
        cmdSave.Enabled = True
        checkDate
    End If
End Sub

Private Sub txtSubID_Change()
    If (txtSubID.Text <> "") Then
        SubjectID = CInt(txtSubID.Text)
        getSubject
        getDeliveryDate
        getSeatCode
        If txtSubID.Text = "264" Then
            chkByOwner.Enabled = True
            chkByRefMember.Enabled = True
        ElseIf txtSubID = "329" Then
            If chkBPL.Value = 0 And chkScSt.Value = 0 Then
                If (MsgBox("Whether applicant belongs to BPL or SC?ST ??", vbYesNo) = vbYes) Then
                    chkBPL.SetFocus
                End If
            End If
        Else
            chkByOwner.Enabled = False
            chkByRefMember.Enabled = False
        End If
    End If
End Sub

Private Sub txtSubID_KeyPress(KeyAscii As Integer)
 If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub
Private Sub vsEnclosure_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 2 Then
        KeyAscii = 0
    End If
End Sub
Private Sub vsValuable_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        If IsDate(vsValuable.TextMatrix(Row, Col)) = False Then
            vsValuable.TextMatrix(Row, Col) = ""
        Else
            vsValuable.TextMatrix(Row, Col) = Format(vsValuable.TextMatrix(Row, Col), "dd/mm/yyyy")
       End If
    End If
    If Col = 3 Then
        If val(vsValuable.TextMatrix(Row, Col)) > 0 Then
            vsValuable.TextMatrix(Row, Col) = val(vsValuable.TextMatrix(Row, Col))
        Else
            vsValuable.TextMatrix(Row, Col) = ""
       End If
    End If
    If vsValuable.TextMatrix(Row, 0) <> "" And vsValuable.TextMatrix(Row, 1) <> "" And vsValuable.TextMatrix(Row, 2) <> "" And vsValuable.TextMatrix(Row, 3) <> "" And vsValuable.TextMatrix(Row, 4) <> "" Then
         vsValuable.Rows = vsValuable.Rows + 1
         vsValuable.Height = vsValuable.Height + 350
    End If
    gSubSetFont vsValuable, 1, 0, vsValuable.Rows - 1, 0, "Verdana"
    gSubSetFont vsValuable, 1, 4, vsValuable.Rows - 1, 4, "Verdana"
End Sub
Private Sub SaveValuable(ByVal mCnn As ADODB.Connection)
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim ForwardTo As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
 Dim i As Integer
     ReDim mVarrIn(6)
     'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
     For i = 1 To vsValuable.Rows - 1
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = val(vsValuable.TextMatrix(i, 0)) 'InstrumentType.
        mVarrIn(2) = vsValuable.TextMatrix(i, 2) 'Date.
        mVarrIn(3) = vsValuable.TextMatrix(i, 1) 'InstrumentNo.
        mVarrIn(4) = vsValuable.TextMatrix(i, 3)  'Amount
        mVarrIn(5) = vsValuable.TextMatrix(i, 4)  'Remarks
        mVarrIn(6) = Null 'sessionId
        If vsValuable.TextMatrix(i, 0) <> "" And vsValuable.TextMatrix(i, 1) <> "" And vsValuable.TextMatrix(i, 2) <> "" And vsValuable.TextMatrix(i, 3) <> "" And vsValuable.TextMatrix(i, 4) <> "" Then
           objdb.ExecuteSP "SpSaveValuable", mVarrIn, , , mCnn, adCmdStoredProc
        End If
    Next i
End Sub
Private Sub SaveCertificateAddress(ByVal mCnn As ADODB.Connection)
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim ForwardTo As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(11)
     'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = txtCertName.Text
        mVarrIn(2) = txtCertHouseName.Text
        mVarrIn(3) = txtCertDoorNo2.Text
        mVarrIn(4) = txtWardNo.Text
        mVarrIn(5) = cboCertDist.ItemData(cboCertDist.ListIndex)
        mVarrIn(6) = txtCertLocality.Text
        mVarrIn(7) = txtCertPostOffice.Text
        mVarrIn(8) = txtCertPincode.Text
        mVarrIn(9) = Null
        mVarrIn(10) = cboCertGender.ItemData(cboCertGender.ListIndex)
        mVarrIn(11) = txtCertDoorNo1.Text
        objdb.ExecuteSP "spSaveCertificateAddress", mVarrIn, , , mCnn, adCmdStoredProc
        
End Sub
Private Sub SaveBillReceipt(ByVal mCnn As ADODB.Connection)
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(7)
     'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = cboType.ListIndex
        mVarrIn(2) = Date
        mVarrIn(3) = txtBillNo.Text
        mVarrIn(4) = val(txtAmt.Text)
        mVarrIn(5) = txtBillDescr.Text
        mVarrIn(6) = Null
        mVarrIn(7) = cboSeatID.Text
        Set Rec = objdb.ExecuteSP("SpSaveBillReceipt", mVarrIn, , , mCnn, adCmdStoredProc)
        
End Sub
Private Sub SaveKeywords(ByVal mCnn As ADODB.Connection)
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(5)
    ' objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = intDistrID
        mVarrIn(2) = intFunID
        mVarrIn(3) = intRefID
        If cboWard.ListIndex < 0 Then
            mVarrIn(4) = Null
        Else
            mVarrIn(4) = cboWard.ItemData(cboWard.ListIndex)
        End If
        If cboMember.ListIndex < 0 Then
            mVarrIn(5) = Null
        Else
            mVarrIn(5) = cboMember.ItemData(cboMember.ListIndex)
        End If
        
        objdb.ExecuteSP "spSaveInwardKeywords", mVarrIn, , , mCnn, adCmdStoredProc
        
End Sub
Private Sub SaveGeneralInwardDetails(ByVal mCnn As ADODB.Connection)
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(3)
    ' objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = chkBPL.Value
        mVarrIn(2) = chkScSt.Value
        mVarrIn(3) = txtDocumentProof.Text
        objdb.ExecuteSP "spSaveInwardGeneralDetails", mVarrIn, , , mCnn, adCmdStoredProc
End Sub

Private Sub FillEnclosure()
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim vAryIn As Variant
    Dim varyOut As Variant
    Dim i As Integer
        objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        ReDim vAryIn(0)
        vAryIn(0) = val(txtSubID.Text)
        Set Rec = objdb.ExecuteSP("SpEnclosure", vAryIn, varyOut, , mCnn, adCmdStoredProc)
        vsEnclosure.Rows = 2
        vsEnclosure.Clear 1
           If IsArray(varyOut) Then
                For i = 0 To UBound(varyOut, 2)
                If i > 0 Then
                   vsEnclosure.Rows = vsEnclosure.Rows + 1
                End If
                   vsEnclosure.TextMatrix(i + 1, 1) = varyOut(1, i)
                   vsEnclosure.TextMatrix(i + 1, 2) = varyOut(0, i)
               Next i
           End If
           gSubSetFont vsEnclosure, 1, 2, vsEnclosure.Rows - 1, 2, "Verdana"
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
End Sub
Private Sub SaveEnclosure(ByVal mCnn As ADODB.Connection)
    Dim mVarrIn As Variant
    Dim mVarrOut As Variant
    Dim ForwardTo As Variant
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    'Dim mCnn As New ADODB.Connection
    Dim i As Integer
        ReDim mVarrIn(1)
        'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        For i = 1 To vsEnclosure.Rows - 1
           mVarrIn(0) = lSoochikaFeildID 'FileId.
           mVarrIn(1) = val(vsEnclosure.TextMatrix(i, 1)) 'EncloserId
           If vsEnclosure.Cell(flexcpChecked, i, 0) = flexChecked Then
               objdb.ExecuteSP "spSaveInwardEnclosure", mVarrIn, , , mCnn, adCmdStoredProc
           End If
       Next i
End Sub
Public Sub SaveRCOwner(ByVal mCnn As ADODB.Connection)
    Dim mSql As String
    'Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim varrin As Variant
    ReDim varrin(3)
    'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        varrin(0) = lSoochikaFeildID
        varrin(1) = Right(lSoochikaFeildID, 6)
        varrin(2) = chkByOwner.Value
        varrin(3) = chkByRefMember.Value
        objdb.ExecuteSP "SPSaveRCInward", varrin, , , mCnn, adCmdStoredProc
End Sub

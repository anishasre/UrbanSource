VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmKMBR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "K M B R"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleMode       =   0  'User
   ScaleWidth      =   11145.07
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   15135
      TabIndex        =   26
      Top             =   0
      Width           =   15135
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   1800
      Top             =   6480
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
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
      Left            =   6870
      TabIndex        =   55
      Top             =   5700
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
      Left            =   8160
      TabIndex        =   54
      Top             =   5700
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
      Left            =   5580
      TabIndex        =   53
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Frame fraAppNameAddress 
      Height          =   3225
      Left            =   60
      TabIndex        =   27
      Top             =   780
      Width           =   7515
      Begin VB.TextBox txtHouseWard 
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
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   1
         Top             =   870
         Width           =   1155
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
         Left            =   5070
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2025
         Width           =   2325
      End
      Begin VB.TextBox txtMobileNo 
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
         Left            =   5070
         MaxLength       =   12
         TabIndex        =   13
         Top             =   2820
         Width           =   2325
      End
      Begin VB.TextBox txtPincode 
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
         Left            =   5070
         TabIndex        =   11
         Top             =   2430
         Width           =   2325
      End
      Begin VB.TextBox txtLocalPlace 
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
         Left            =   5070
         TabIndex        =   7
         Top             =   1635
         Width           =   2325
      End
      Begin VB.TextBox txtResAssoName 
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
         Left            =   5070
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1260
         Width           =   2325
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
         Left            =   5070
         MaxLength       =   25
         TabIndex        =   3
         Top             =   870
         Width           =   2325
      End
      Begin VB.ComboBox cmbPostOffice 
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
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2430
         Width           =   2325
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
         Left            =   1410
         MaxLength       =   12
         TabIndex        =   12
         Top             =   2820
         Width           =   2325
      End
      Begin VB.TextBox txtMainPlace 
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
         Left            =   1410
         TabIndex        =   8
         Top             =   2025
         Width           =   2325
      End
      Begin VB.TextBox txtStreet 
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
         Left            =   1410
         TabIndex        =   6
         Top             =   1635
         Width           =   2325
      End
      Begin VB.TextBox txtResAssNo 
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
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1260
         Width           =   2325
      End
      Begin VB.TextBox txtHouseNo 
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
         Left            =   2580
         MaxLength       =   6
         TabIndex        =   2
         Top             =   870
         Width           =   1155
      End
      Begin VB.TextBox txtName 
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
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   0
         Top             =   510
         Width           =   5985
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   70
         Top             =   2550
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   4410
         TabIndex        =   69
         Top             =   2100
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   4860
         TabIndex        =   68
         Top             =   930
         Width           =   90
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
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
         Left            =   1320
         TabIndex        =   67
         Top             =   960
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   66
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   30
         TabIndex        =   52
         Top             =   120
         Width           =   7440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moblie No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   40
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   39
         Top             =   2490
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   38
         Top             =   2085
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3750
         TabIndex        =   37
         Top             =   1695
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Res.Asso.Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   36
         Top             =   1290
         Width           =   1335
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   35
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   34
         Top             =   2910
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post office"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   33
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   2070
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   1710
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Res.Asso.No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward/HouseNo."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   540
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1665
      Left            =   60
      TabIndex        =   41
      Top             =   3960
      Width           =   7515
      Begin VB.ComboBox cboBuidingType 
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
         Left            =   2140
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   630
         Width           =   1815
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   630
         Width           =   2415
      End
      Begin VB.CheckBox chkOneday 
         Appearance      =   0  'Flat
         Caption         =   "For Oneday Permit"
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
         Height          =   255
         Left            =   90
         TabIndex        =   59
         Top             =   360
         Width           =   2025
      End
      Begin VB.CheckBox chkStampPaper 
         Appearance      =   0  'Flat
         Caption         =   "For Undertaking in Stamp Paper"
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
         Height          =   255
         Left            =   150
         TabIndex        =   58
         Top             =   1020
         Width           =   3105
      End
      Begin VB.CheckBox chkSitePlan 
         Appearance      =   0  'Flat
         Caption         =   "Site Plan"
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
         Height          =   255
         Left            =   3420
         TabIndex        =   57
         Top             =   1380
         Width           =   1125
      End
      Begin VB.CheckBox chkGeneral 
         Appearance      =   0  'Flat
         Caption         =   "For General Permit"
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
         Height          =   255
         Left            =   4950
         TabIndex        =   56
         Top             =   360
         Width           =   1965
      End
      Begin VB.CheckBox chkBuildingPlan 
         Appearance      =   0  'Flat
         Caption         =   "Building Plan"
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
         Height          =   255
         Left            =   4950
         TabIndex        =   17
         Top             =   1380
         Width           =   1515
      End
      Begin VB.CheckBox chkLocationPlan 
         Appearance      =   0  'Flat
         Caption         =   "Location Plan"
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
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   1380
         Width           =   1515
      End
      Begin VB.CheckBox chkLandTaxReceipt 
         Appearance      =   0  'Flat
         Caption         =   "Latest Land Tax Receipt"
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
         Height          =   255
         Left            =   4950
         TabIndex        =   15
         Top             =   1005
         Width           =   2415
      End
      Begin VB.CheckBox chkOwnership 
         Appearance      =   0  'Flat
         Caption         =   "Ownership"
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
         Height          =   255
         Left            =   3405
         TabIndex        =   14
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   9
         Left            =   3300
         TabIndex        =   77
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   7
         Left            =   4830
         TabIndex        =   76
         Top             =   1050
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         TabIndex        =   75
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   6
         Left            =   3300
         TabIndex        =   74
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   8
         Left            =   30
         TabIndex        =   73
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   5
         Left            =   30
         TabIndex        =   72
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Index           =   4
         Left            =   4800
         TabIndex        =   71
         Top             =   660
         Width           =   90
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building in SqureMeter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   63
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seat No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4065
         TabIndex        =   61
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Minimum Documents and Seat for Appllication"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   30
         TabIndex        =   42
         Top             =   120
         Width           =   7440
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4845
      Left            =   7590
      TabIndex        =   43
      Top             =   780
      Width           =   7454
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
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   3630
         Width           =   2775
      End
      Begin VB.TextBox txtAccHead 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   960
         Width           =   2325
      End
      Begin VSFlex8LCtl.VSFlexGrid fgAccHead 
         Height          =   885
         Left            =   30
         TabIndex        =   23
         Top             =   2640
         Width           =   7335
         _cx             =   12938
         _cy             =   1561
         Appearance      =   0
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKMBR.frx":0000
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
         Editable        =   0
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
      Begin VB.TextBox txtReceiptDate 
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
         Left            =   1560
         TabIndex        =   22
         Top             =   2190
         Width           =   2325
      End
      Begin VB.TextBox txtReceiptNo 
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
         Left            =   1560
         TabIndex        =   21
         Top             =   1785
         Width           =   2325
      End
      Begin VB.TextBox txtDiscription 
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
         Height          =   585
         Left            =   1560
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   4110
         Width           =   5685
      End
      Begin VB.ComboBox cboZone 
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
         TabIndex        =   24
         Top             =   3645
         Width           =   2325
      End
      Begin VB.ComboBox cboTransaction 
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
         TabIndex        =   18
         Top             =   540
         Width           =   3765
      End
      Begin VB.ComboBox cboInstrument 
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
         TabIndex        =   20
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   12
         Left            =   4530
         TabIndex        =   79
         Top             =   3720
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   11
         Left            =   510
         TabIndex        =   78
         Top             =   3780
         Width           =   90
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   64
         Top             =   3690
         Width           =   450
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   51
         Top             =   4350
         Width           =   960
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   50
         Top             =   1410
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Head"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   49
         Top             =   1005
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   48
         Top             =   3720
         Width           =   435
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   47
         Top             =   2235
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   46
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   45
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Counter Recepits"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   240
         Left            =   30
         TabIndex        =   44
         Top             =   120
         Width           =   7380
      End
   End
End
Attribute VB_Name = "frmKMBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intcboTransactionLen As Integer
Dim lSoochikaFeildID As Long
Dim lSoochikaCurrentNo As Long
Dim mVoucherID As Long
Dim intTransactionID_1 As Integer
Dim lReceiptID As Integer
Dim N1 As String
Dim N2 As String
Dim N3 As String
Dim N4 As String
Dim I1 As String
Dim I2 As String
Dim I3 As String
Dim I4 As String



Private Sub cboBuidingType_Click()
    lSubFee
End Sub

Private Sub cboDistrict_Click()
    If cboDistrict.ListIndex >= 0 Then
        Call PopulateList(cmbPostOffice, "SELECT chvPostOfficeEnglish,intPostOfficeID From GM_PostOffice left join GL_PostOffice on left(GM_PostOffice.intPINCode,3)=GL_PostOffice.intPINCode Where tnyDistrictID =" & cboDistrict.ItemData(cboDistrict.ListIndex) & "order by  chvPostOfficeEnglish", , , , True, DBMaster)
    End If
End Sub

Private Sub cboTransaction_Click()
'If (Len(cboTransaction.List(cboTransaction.ListIndex)) * 100) > 2325 Then
'   cboTransaction.Width = Len(cboTransaction.List(cboTransaction.ListIndex)) * 100
'    If cboTransaction.Width > 5850 Then
'        cboTransaction.Width = 5850
'    End If
'Else
'    cboTransaction.Width = 2325
'End If
cboTransaction.ListIndex = 51
End Sub

Private Sub chkGeneral_Click()
    If chkGeneral.value = 1 Then
        chkOneday.value = 0
        Call PopulateList(cboBuidingType, "SELECT chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=0", , , , True, KMBR)
        cboBuidingType.ListIndex = 0
        lbl(5).Caption = ""
    Else
        chkOneday.value = 1
        Call PopulateList(cboBuidingType, "SELECT chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=1", , , , True, KMBR)
        fgAccHead.TextMatrix(1, 6) = 130
        cboBuidingType.ListIndex = 0
        lbl(5).Caption = "*"
    End If
    
End Sub

Private Sub chkOneday_Click()
    If chkOneday.value = 1 Then
        chkGeneral.value = 0
        Call PopulateList(cboBuidingType, "SELECT  chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=1", , , , True, KMBR)
        fgAccHead.TextMatrix(1, 6) = 130
        cboBuidingType.ListIndex = 0
        lbl(5).Caption = "*"
    Else
        chkGeneral.value = 1
        Call PopulateList(cboBuidingType, "SELECT  chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=0", , , , True, KMBR)
        cboBuidingType.ListIndex = 0
        lbl(5).Caption = ""
    End If
End Sub

Private Sub cmbPostOffice_Click()
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
objDB.CreateNewConnection mCnn, enuSourceString.DBMaster
    If cmbPostOffice.ListIndex >= 0 Then
        Set Rec = objDB.ExecuteSP("SELECT intPINCode From GM_PostOffice WHERE intPostOfficeID =  " & cmbPostOffice.ItemData(cmbPostOffice.ListIndex), , mVarrOut, , mCnn, adCmdText)
        If IsArray(mVarrOut) Then
           txtPincode.Text = mVarrOut(0, 0)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    ClearDetails
End Sub

Private Sub cmdSave_Click()
    If lSaveValidate = True Then
        Call SaveSoochika
        txtDiscription.Text = "Application for Building Permit File No: " & cboSeat.List(cboSeat.ListIndex) & "/" & lSoochikaCurrentNo & "/" & CStr(Year(Date))
        Call SaveSaankhya
        Call SaveSanketham
'        MsgBox "Inward: " & cboSeat.List(cboSeat.ListIndex) & "/" & lSoochikaCurrentNo & "/" & Year(Date) & " Receipt No: " & txtReceiptNo.Text
        MsgBox "Inward: " & lSoochikaCurrentNo & "/" & Year(Date) & " Receipt No: " & txtReceiptNo.Text
        ClearDetails
        PrintReceipt (mVoucherID)
    End If
End Sub

Private Sub Form_Load()
    WindowsXPC.InitIDESubClassing
    FillZone
    FillTransactionTypes
    FillSeats
    chkOneday.value = 1
    Call PopulateList(cboBuidingType, "SELECT  chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=1", , , , True, KMBR)
    fgAccHead.TextMatrix(1, 6) = 130
    Call PopulateList(cboDistrict, "Select chvDistrictEnglish, tnyDistrictID From GM_District Order By chvDistrictEnglish", , , , True, DBMaster)
    Call PopulateList(cboInstrument, "SELECT chvDebitType,intDebitId  From LMSan_Debit Where intDebitId = 1504", , , , True, KMBR)
    Call PopulateList(cboWard, "SELECT  chvWardNameEnglish,intWardNo  FROM GM_Ward WHERE intLBID = " & gbLocalBodyID & " AND intWardYear = 2005 and chvWardNameEnglish is not null", , , , True, DBMaster)
    FillDetails
    frmKMBR.Width = 15225
    For Mi = 0 To 12
        lbl(Mi).ForeColor = vbBlue
    Next Mi
End Sub
Private Sub FillZone()
    Call PopulateList(cboZone, "Select chvZoneNameEnglish, numZoneID From GM_Zone Order By chvZoneNameEnglish", , True, True, True, DBMaster)
End Sub
Private Sub FillTransactionTypes()
    Dim mSQL As String
    mSQL = "Select vchTransactionType, intTransactionTypeID, intGroupID From faTransactionType Where intGroupID = 10 Order By vchTransactionType"
    Call PopulateList(cboTransaction, mSQL, , True, True, True)
'    intcboTransactionLen = 2325
'    For i = 0 To cboTransaction.ListCount - 1
'       If (Len(cboTransaction.List(i)) * 90) > intcboTransactionLen Then
'        intcboTransactionLen = Len(cboTransaction.List(i)) * 90
'       End If
'    Next i
'    cboTransaction.Width = intcboTransactionLen
End Sub
    Private Sub FillSeats()
        PopulateList cboSeat, "SELECT chvSeatTitle,numSeatID FROM GL_Seats Where intLocalBodyID = " & gbLocalBodyID & " ORDER BY chvSeatTitle", , , True, , enuSourceString.DBMaster
    End Sub

Private Sub FillDetails()
    txtAccHead.Text = "1450100100" 'intDebitCode
'    cboTransaction.Width = intcboTransactionLen
    cboInstrument.ListIndex = 0
    fgAccHead.TextMatrix(1, 0) = 1
    fgAccHead.TextMatrix(1, 1) = gbAcHeadIDOtherFee
    fgAccHead.TextMatrix(1, 2) = gbAcHeadCodeOtherFee '140409900
    fgAccHead.TextMatrix(1, 3) = "Other Fee"
    If Month(Date) < 3 Then
        fgAccHead.TextMatrix(1, 4) = Year(Date) - 1 & " - " & Year(Date)
    Else
        fgAccHead.TextMatrix(1, 4) = Year(Date) & " - " & Year(Date) + 1
    End If
    If chkOneday.value = 1 Then
        fgAccHead.TextMatrix(1, 6) = 130
        cboBuidingType.ListIndex = 0
    Else
        fgAccHead.TextMatrix(1, 6) = 130
        cboBuidingType.ListIndex = 0
    End If
    cboTransaction.ListIndex = 51
End Sub


Private Sub ClearDetails()
    txtName.Text = ""
    txtHouseNo.Text = ""
    txtHouseWard.Text = ""
    txtHouseName.Text = ""
    txtResAssNo.Text = ""
    txtResAssoName.Text = ""
    txtStreet.Text = ""
    txtMainPlace.Text = ""
    txtLocalPlace.Text = ""
    cboDistrict.ListIndex = -1
    cmbPostOffice.ListIndex = -1
    txtPincode.Text = ""
    txtPhoneno.Text = ""
    txtMobileNo.Text = ""
    chkOneday.value = 0
    chkStampPaper.value = 0
    chkOwnership.value = 0
    chkLandTaxReceipt.value = 0
    chkSitePlan.value = 0
    chkLocationPlan.value = 0
    chkBuildingPlan.value = 0
    cboSeat.ListIndex = -1
    txtReceiptNo.Text = ""
    txtReceiptDate.Text = ""
    cboZone.ListIndex = -1
    txtDiscription.Text = ""
    
End Sub

Private Sub txtHouseWard_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMobileNo_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 32) Or (KeyAscii = 8) Or (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Then
   Else
        KeyAscii = 0
   End If
End Sub

Private Sub txtName_LostFocus()

Dim i As Integer
Dim ary As Variant
Dim sp As Integer

sp = 0
N1 = ""
N2 = ""
N3 = ""
N4 = ""
I1 = ""
I2 = ""
I3 = ""
I4 = ""

    For i = 1 To Len(Trim(txtName.Text))
        If LCase(mID(Trim(txtName.Text), i, 1)) = " " Then
            sp = sp + 1
        End If
    Next i
    ReDim ary(sp)
    If Trim(txtName.Text) <> "" Then
        ary = Split(Trim(txtName.Text), " ", , vbTextCompare)
    End If
    
    For J = 0 To sp
        If N1 = "" And Len(ary(J)) > 1 Then
            N1 = ary(J)
        ElseIf I1 = "" And Len(ary(J)) = 1 Then
            I1 = ary(J)
        ElseIf N2 = "" And Len(ary(J)) > 1 Then
            N2 = ary(J)
        ElseIf I2 = "" And Len(ary(J)) = 1 Then
            I2 = ary(J)
        ElseIf N3 = "" And Len(ary(J)) > 1 Then
            N3 = ary(J)
        ElseIf I3 = "" And Len(ary(J)) = 1 Then
            I3 = ary(J)
        ElseIf N4 = "" And Len(ary(J)) > 1 Then
            N4 = ary(J)
        ElseIf I4 = "" And Len(ary(J)) = 1 Then
            I4 = ary(J)
        End If
    Next J
    txtName.Text = UCase(txtName.Text)
End Sub

Private Sub txtPhoneno_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub SaveSoochika()
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim ForwardTo As Variant
 Dim objDB As New clsDB
 Dim Rec As New ADODB.Recordset
 Dim mCnn As New ADODB.Connection
 ReDim mVarrIn(41)
    mVarrIn(0) = 0 'FldCurrentNo.
    mVarrIn(1) = Date  'FldDateOfReceipt.
    mVarrIn(2) = txtName.Text 'FldSenderName.
    mVarrIn(3) = txtHouseWard.Text 'FldWardNo.
    mVarrIn(4) = txtHouseNo 'FldHouseNo
    mVarrIn(5) = txtMainPlace 'FldLocality
    mVarrIn(6) = cboDistrict.ItemData(cboDistrict.ListIndex) 'FldDistrict
    mVarrIn(7) = gbSeatID 'bntCurrUserId.
    
    ForwardTo = "40" & CStr(gbLocalBodyID) & CStr(cboSeat.ItemData(cboSeat.ListIndex))
    mVarrIn(8) = ForwardTo 'cboSeat.ItemData(cboSeat.ListIndex)  'intForwardTo.
    mVarrIn(9) = 1 'intInwardType.
    mVarrIn(10) = 5 'FldPriority
    mVarrIn(11) = Date 'dtmForwardDate
    mVarrIn(12) = "Application for Build Prmit"  'FldRemarks
    mVarrIn(13) = Null 'intAttachmentType
    mVarrIn(14) = Null 'FldManualSummary
    mVarrIn(15) = Null 'FldElectronicsSummary
    mVarrIn(16) = 9 'intDept
    mVarrIn(17) = Null 'FlgCourFeeStamp
    mVarrIn(18) = Null 'intManualPage
    mVarrIn(19) = Null 'FldOutsideNo
    mVarrIn(20) = Null 'FldRefDate
    mVarrIn(21) = Null 'intRegPost
    mVarrIn(22) = Null 'bitInstflg
    mVarrIn(23) = Null 'fldInstName
    mVarrIn(24) = Null 'fldDesign
    mVarrIn(25) = cmbPostOffice.List(cmbPostOffice.ListIndex) 'FldPostOffice
    mVarrIn(26) = txtPincode.Text 'FldPin
    mVarrIn(27) = Null 'FldEmail
    mVarrIn(28) = txtPhoneno.Text 'FldPhone
    mVarrIn(29) = Null 'fldReglttoWhom
    mVarrIn(30) = Null 'fldReglttoDesign
    mVarrIn(31) = Null 'fldRegltpoNo
    mVarrIn(32) = Null 'sessionID
    mVarrIn(33) = Null 'intBillRecFlg
    mVarrIn(34) = Null 'intInsideLBFlg
    mVarrIn(35) = txtHouseName.Text 'FldHouseName
    mVarrIn(36) = Null 'intCertAddrFlg
    mVarrIn(37) = Null 'intGender
    mVarrIn(38) = Null 'intDoorNo
    mVarrIn(39) = 0 'InwardFlg
    mVarrIn(40) = 0 'Suit
    If chkOneday.value = 0 Then
        mVarrIn(41) = 292 'Subject Gereral Permit
    Else
        mVarrIn(41) = 296 'Subject OneDay
    End If
    
  
    
    objDB.CreateNewConnection mCnn, enuSourceString.Soochika
    Set Rec = objDB.ExecuteSP("spSaveCorpOfficeView_KMBR", mVarrIn, mVarrOut, , mCnn, adCmdStoredProc)
    If IsArray(mVarrOut) Then
       lSoochikaFeildID = mVarrOut(0, 0)
    End If
    Set Rec = objDB.ExecuteSP("SELECT FldCurrentNo From TblTappalDetails WHERE FldFileId = " & lSoochikaFeildID, , mVarrOut, , mCnn, adCmdText)
    If IsArray(mVarrOut) Then
       lSoochikaCurrentNo = mVarrOut(0, 0)
    End If

End Sub

Private Sub SaveSaankhya()
    
    Call lReceiptNo
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim arrInput As Variant
    Dim arrOutPut As Variant
    Dim mLoopCount As Long
    Dim mLoop As Long
    Dim Rec As New ADODB.Recordset
    Dim mDemandID As Variant
    
    mDrAccountHeadID = 1504
    
    If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then ' CREATED NEW CONNECTION
            
         Dim mintVoucherID_1                As Double
         '@intLocalBodyID_2  [int],
         '@intTransactionID_3    [bigint],
         Dim mintTransactionTypeID_4        As Long
         Dim mtnyVoucherTypeID_5            As Integer
         Dim mintVoucherNo_6                As Variant
         Dim mintBookNo_7                   As Long
         Dim mdtDate_8                      As Date
         Dim mfltAmount_9                   As Double
         Dim mintInstrumentTypeID_10        As Integer
         Dim mvchInstrumentNo_11            As Variant
         Dim mdtInstrumentDate_12           As Variant
         Dim mvchDescription_13             As String
         Dim mnumZoneID_14                  As Variant
         Dim mnumWardID_15                  As Double
         Dim mintDoorNoP1_16                As Long
         Dim mvchDoorNoP2_17                As String
         Dim mvchDoorNoP3_18                As String
         Dim mintUserID_19                  As Long
         Dim mintCounterID_20               As Long
         Dim mnumSubLedgerID_21             As Variant
         Dim mintKeyID1_22                  As Variant
         Dim mintKeyID2_23                  As Variant
         Dim mintExternalApplicationID_24   As Long
         Dim mintExternalModuleID_25        As Long
         Dim mintFinancialYearID_26         As Long
         
         Dim mvchBank_33                    As Variant
         Dim mvchBankPlace_34               As Variant
         Dim mintFundID_35                  As Long
         Dim mRefNo As String
         
         Dim mInwardNo  As Long
         
         mintTransactionTypeID_4 = cboTransaction.ItemData(cboTransaction.ListIndex)
         mtnyVoucherTypeID_5 = 10
         mintVoucherNo_6 = val(txtReceiptNo.Text)
         mintBookNo_7 = 0
         mdtDate_8 = gbTransactionDate
         mfltAmount_9 = val(fgAccHead.TextMatrix(1, 6))
         mintInstrumentTypeID_10 = gbInstrumentCash
         mvchInstrumentNo_11 = Null
         mdtInstrumentDate_12 = Null
         mvchDescription_13 = Trim(txtDiscription.Text)
         
         If cboZone.ListIndex > 0 Then
             mnumZoneID_14 = cboZone.ItemData(cboZone.ListIndex)
         End If
         
         mnumWardID_15 = val(txtHouseWard.Text)
         mintDoorNoP1_16 = 0
         mvchDoorNoP2_17 = Trim(txtHouseNo.Text)
         mvchDoorNoP3_18 = 0
         mintUserID_19 = gbUserID
         mintCounterID_20 = gbCounterID
         mnumSubLedgerID_21 = Null
         mintKeyID1_22 = 1504
         mintKeyID2_23 = Null
         mintExternalApplicationID_24 = 100
         mintExternalModuleID_25 = 120
         mintFinancialYearID_26 = gbFinancialYearID
         mvchBank_33 = Null
         mvchBankPlace_34 = Null
         mintFundID_35 = 1
         mRefNo = cboSeat.List(cboSeat.ListIndex) & "/" & lSoochikaFeildID & "/" & Year(Date)
         
         mInwardNo = lSoochikaCurrentNo
         '========================================='
         ' BEGIN TRANSACTION                       '
         '-----------------------------------------'
             mCnn.BeginTrans
             On Error GoTo ErrorRollBack:
         '========================================='
         
         arrInput = Array( _
         -1, _
         gbLocalBodyID, _
         Null, _
         mintTransactionTypeID_4, _
         mtnyVoucherTypeID_5, _
         mintVoucherNo_6, _
         mintBookNo_7, _
         mdtDate_8, _
         mfltAmount_9, _
         mintInstrumentTypeID_10, _
         mvchInstrumentNo_11, _
         mdtInstrumentDate_12, _
         mvchDescription_13, _
         mnumZoneID_14, _
         mnumWardID_15, _
         mintDoorNoP1_16, _
         mvchDoorNoP2_17, _
         mvchDoorNoP3_18, _
         mintUserID_19, _
         mintCounterID_20, _
         mnumSubLedgerID_21, _
         mintKeyID1_22, mintKeyID2_23, mintExternalApplicationID_24, _
         mintExternalModuleID_25, mintFinancialYearID_26, gbShiftID, 1, 0, _
         mvchBank_33, mvchBankPlace_34, mintFundID_35, gbSeatID, gbSessionID, mRefNo, Null, Null, mInwardNo)

         objDB.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCnn
         If IsNumeric(arrOutPut(0, 0)) Then
             mintVoucherID_1 = arrOutPut(0, 0)
             mVoucherID = mintVoucherID_1
             frmRecieptCancellation.mVocherIDForCancel = mVoucherID     ' Global Variable mVoucherID - Should be Declared   '
         Else
             GoTo ErrorRollBack:
         End If
         
        
         '-------------------------------------------------------'
         ' faVoucher Child
         '-------------------------------------------------------'
         'Dim mintVoucherID_1       As Double  '
         Dim mintLocalBodyID_2       As Long
         Dim mintSlNo_3              As Long
         Dim mintAccountHeadID_4     As Long
         Dim mtnyDebitOrCredit_5     As Byte
         Dim mintYearID_6            As Long
         Dim mtnyPeriodID_7          As Byte
         Dim mtnyArrearFlag_8        As Variant
         Dim mnumDemandID_9          As Double
         Dim mfltAmount_10           As Double
         
        mintLocalBodyID_2 = gbLocalBodyID
        mintSlNo_3 = 1
        mintAccountHeadID_4 = fgAccHead.TextMatrix(1, 1)
        mtnyDebitOrCredit_5 = 0
        mintYearID_6 = gbFinancialYearID
        mtnyPeriodID_7 = 3
        mtnyArrearFlag_8 = Null
        mnumDemandID_9 = 0 'Val(vsGrid.Cell(flexcpText, mLoopCount, 10))
        mfltAmount_10 = val(fgAccHead.TextMatrix(1, 6))
        
        Set arrInput = Nothing
        arrInput = Array( _
        mintVoucherID_1, _
        mintLocalBodyID_2, _
        mintSlNo_3, _
        mintAccountHeadID_4, _
        mtnyDebitOrCredit_5, _
        mintYearID_6, _
        mtnyPeriodID_7, _
        mtnyArrearFlag_8, _
        mnumDemandID_9, _
        mfltAmount_10 _
        )
        objDB.ExecuteSP "spSaveVoucherChild", arrInput, , , mCnn
        
         '-------------------------------------------------------'
         ' faVoucher Address
         '-------------------------------------------------------'
         
        
           Dim intWardNo           As Variant
           Dim intDoorNo           As Variant
           Dim vchDoorNo2          As Variant
           Dim vchName           As Variant
           Dim vchInit1            As Variant
           Dim vchInit2            As Variant
           Dim vchInit3            As Variant
           Dim vchInit4            As Variant
           
           Dim vchHouseName      As Variant
           Dim vchStreetName     As Variant
           Dim vchLocalPlace       As Variant
           Dim vchMainPlace      As Variant
           Dim vchPostOffice     As Variant
           Dim vchDistrict       As Variant
           Dim vchPinNumber      As Variant
           Dim vchPhone            As Variant
          
           
           Dim mStartingReceiptNo As Variant       ' Keeps value on every session
           Dim mGrandTotal         As Variant
           Dim mSkipFlag           As Boolean      ' To control AutoFill Text Behaviour or TransactionType Text Box
           Dim mKeyCode            As Long
           Dim mBkSpaceFlag        As Boolean
        
        
         
         
         vchName = Trim(txtName.Text)
         vchHouseName = Trim(txtHouseName.Text)
         vchInit1 = Null
         vchInit2 = Null
         vchInit3 = Null
         vchInit4 = Null
         vchStreetName = Trim(txtStreet.Text)
         vchLocalPlace = Trim(txtLocalPlace.Text)
         vchMainPlace = Trim(txtMainPlace.Text)
         vchPostOffice = cmbPostOffice.ItemData(cmbPostOffice.ListIndex)
         vchPinNumber = val(txtPincode.Text)
         vchPhone = txtPhoneno.Text
         intWardNo = txtHouseWard.Text
         intDoorNo = Null
         vchDoorNo2 = txtHouseNo.Text
         '-------------------------------------------------------'
         arrInput = Array(mintVoucherID_1, _
                 gbLocalBodyID, _
                 vchName, _
                 vchInit1, _
                 vchInit2, _
                 vchInit3, _
                 vchInit4, _
                 vchHouseName, _
                 vchStreetName, _
                 vchLocalPlace, _
                 vchMainPlace, _
                 vchPostOffice, _
                 vchDistrict, _
                 vchPinNumber, _
                 vchPhone, _
                 intWardNo, _
                 intDoorNo, _
                 vchDoorNo2)
         objDB.ExecuteSP "spSaveVoucherAddress", arrInput, , , mCnn
         
         '-------------------------------------------------------'
         ' Transactions                                          '
         '-------------------------------------------------------'
         Dim intTransactionID_1   As Double
         'Dim mintLocalBodyID_2  As Long
         Dim mintFinancialYearID_3  As Long
         Dim mdtTransactionDate_4   As Date
         Dim mintExternalApplicationID_5    As Long
         Dim mintExternalApplicationModuleID_6  As Long
         Dim mintFunctionID_7   As Variant
         Dim mintFunctionaryID_8   As Variant
         Dim mintFieldID_9 As Variant
         Dim mintFundID_10 As Variant
         Dim mintBudgetCentreID_11  As Variant
         Dim mvchNarration_12   As String
         Dim mintTransactionTypeID_13   As Long
         Dim mintVoucherNo_14   As Long
         Dim mintProcessID_15    As Variant
         Dim mintGroupID_17    As Long
         Dim mvchGroup_16   As String
         Dim mintKeyID_18   As Variant
         Dim mnumSubLedgerID_19    As Variant
         'Dim mintUserID_20  As Long
         
         intTransactionID_1 = -1
         mintLocalBodyID_2 = gbLocalBodyID
         mintFinancialYearID_3 = gbFinancialYearID
         mdtTransactionDate_4 = gbTransactionDate
         mintExternalApplicationID_5 = 100
         mintExternalApplicationModuleID_6 = 120
         mintFunctionID_7 = 14
         mintFunctionaryID_8 = 5
         mintFieldID_9 = Null
         mintFundID_10 = Null
         mintBudgetCentreID_11 = Null
         mvchNarration_12 = Trim(txtDiscription.Text)
         mintTransactionTypeID_13 = cboTransaction.ItemData(cboTransaction.ListIndex)
         mintVoucherNo_14 = mintVoucherID_1
         mintProcessID_15 = Null
         mvchGroup_16 = "R"
         mintGroupID_17 = 10
         mintKeyID_18 = Null 'mDemandID 'Added on 3-Sep-2008
         'mnumSubLedgerID_19 = mBuildingID
         'mintUserID_20 = gbUserID
         mnumSubLedgerID_19 = Null
         
         arrInput = Array( _
         intTransactionID_1, _
         mintLocalBodyID_2, _
         mintFinancialYearID_3, _
         mdtTransactionDate_4, _
         mintExternalApplicationID_5, _
         mintExternalApplicationModuleID_6, _
         mintFunctionID_7, _
         mintFunctionaryID_8, _
         mintFieldID_9, _
         mintFundID_10, _
         mintBudgetCentreID_11, _
         mvchNarration_12, _
         mintTransactionTypeID_13, _
         mintProcessID_15, _
         mvchGroup_16, _
         mintGroupID_17, _
         mintKeyID_18, _
         mnumSubLedgerID_19, _
         gbUserID, _
         mintVoucherNo_14)
         
         Set arrOutPut = Nothing
         'objDB.ExecuteSP "spSaveReceiptTransactions", arrInput, arrOutPut, , mCnn
         objDB.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCnn
         If IsNumeric(arrOutPut(0, 0)) Then
             intTransactionID_1 = arrOutPut(0, 0)
         Else
             GoTo ErrorRollBack:
         End If
         
         '-------------------------------------------------------'
         ' Transaction Child                                     '
         '-------------------------------------------------------'
         
         arrInput = Array(intTransactionID_1, _
                        -1, _
                        1504, _
                        fgAccHead.TextMatrix(1, 6), _
                        1, _
                        Null, _
                        txtDiscription, _
                        Null)
         objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
         
         arrInput = Array(intTransactionID_1, _
                        -1, _
                        fgAccHead.TextMatrix(1, 1), _
                        fgAccHead.TextMatrix(1, 6), _
                        0, _
                        1504, _
                        txtDiscription, _
                        Null)
         objDB.ExecuteSP "spSaveTransactionChild", arrInput, , , mCnn
         
         '========================================='
         ' TRANSACTION COMMITTING                  '
         '-----------------------------------------'
             mCnn.CommitTrans
             Set mCnn = Nothing
             On Error GoTo 0
         '========================================='
         
         'Call LockForm(False)
         'cmdSave.Enabled = False
         
         'Call PrintReceipt(mintVoucherID_1)
         
         'Call FormInitialize
         
    Else
        Debug.Print "Error in establishing connection with Saankhya DB"
        Exit Sub
    End If
    Exit Sub
ErrorRollBack:
    mCnn.RollbackTrans
    mVoucherID = Null
    Set mCnn = Nothing
End Sub

Private Sub SaveSanketham()
Dim mVarrIn As Variant
Dim mVarrOut As Variant
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
ReDim mVarrIn(19)
    mVarrIn(0) = lSoochikaFeildID 'intMain,
    mVarrIn(1) = gbLocalBodyID 'intLB,
    mVarrIn(2) = cboZone.ItemData(cboZone.ListIndex) 'intZone,
    mVarrIn(3) = cboWard.ItemData(cboWard.ListIndex) 'intWard,
    mVarrIn(4) = 1 'bitDocuments,
    mVarrIn(5) = chkOneday.value 'bitOneDayPermit,
    mVarrIn(6) = txtName.Text 'chvName,
    mVarrIn(7) = txtHouseWard.Text 'intDoorNo1AddressHn,
    mVarrIn(8) = txtHouseWard.Text 'intDoorNo2AddressHn,
    mVarrIn(9) = txtHouseWard.Text 'intWardNoAddress,
    mVarrIn(10) = txtHouseName.Text 'chvHouseNameAddress,
    mVarrIn(11) = txtMainPlace.Text 'chvMainPlaceAddress,
    mVarrIn(12) = cboDistrict.ItemData(cboDistrict.ListIndex) 'intDistIdAddress,
    mVarrIn(13) = cmbPostOffice.ItemData(cmbPostOffice.ListIndex) 'intPostOfficeAddress,
    mVarrIn(14) = cboSeat.List(cboSeat.ListIndex) 'chvSeatNoFlnoPart1,
    mVarrIn(15) = lSoochikaCurrentNo 'intCurrentNoFlnoPart2,
    mVarrIn(16) = gbUserID 'intUserId,
    mVarrIn(17) = 1 'intFileStatus,
    mVarrIn(18) = 0 'intProcess,
    mVarrIn(19) = val(fgAccHead.TextMatrix(1, 6)) 'Fee
    objDB.CreateNewConnection mCnn, enuSourceString.KMBR
    Set Rec = objDB.ExecuteSP("SoochikaIns1", mVarrIn, mVarrOut, , mCnn, adCmdStoredProc)
    If IsArray(mVarrOut) Then
       lSoochikaFeildID = mVarrOut(0, 0)
    End If
    Call SaveNameTC
    Call SaveAddressTC
    Call Receipt
    Call ReceiptChild
End Sub

Private Sub SaveNameTC()
Dim mVarrIn As Variant
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
ReDim mVarrIn(10)
    mVarrIn(0) = lSoochikaFeildID 'intMain,
    mVarrIn(1) = 0 'CatId ,
    mVarrIn(2) = 4 'LanguageId,
    mVarrIn(3) = N1 'Name1,
    mVarrIn(4) = N2 'Name2,
    mVarrIn(5) = N3 'Name3,
    mVarrIn(6) = N4 'Name4,
    mVarrIn(7) = I1 'Intial1,
    mVarrIn(8) = I2 'Intial2,
    mVarrIn(9) = I3 'Intial3,
    mVarrIn(10) = I4 'Intial4,
    objDB.CreateNewConnection mCnn, enuSourceString.KMBR
    Set Rec = objDB.ExecuteSP("NameIns", mVarrIn, , , mCnn, adCmdStoredProc)
  
End Sub

Private Sub SaveAddressTC()
Dim mVarrIn As Variant
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
ReDim mVarrIn(15)
    mVarrIn(0) = lSoochikaFeildID 'intMain,
    mVarrIn(1) = 0 'tnyTypeId,
    mVarrIn(2) = txtHouseWard.Text 'intHouseNoWard,
    mVarrIn(3) = txtHouseNo.Text 'chvHouseNo,
    mVarrIn(4) = cboDistrict.ItemData(cboDistrict.ListIndex) 'intDistrict,
    mVarrIn(5) = txtHouseName.Text 'chvHouseName,
    mVarrIn(6) = txtResAssNo.Text 'chvResAssocNo,
    mVarrIn(7) = txtResAssoName.Text 'chvResAssoc,
    mVarrIn(8) = txtLocalPlace.Text 'chvLandMark,
    mVarrIn(9) = txtStreet.Text 'chvStreetName,
    mVarrIn(10) = txtMainPlace.Text 'chvMainPlace,
    mVarrIn(11) = cmbPostOffice.ItemData(cmbPostOffice.ListIndex) 'intPostOfficeId,
    mVarrIn(12) = txtPincode.Text 'intPincode,
    mVarrIn(13) = txtPhoneno.Text 'nmPhoneNo,
    mVarrIn(14) = txtMobileNo.Text 'nmMobileNo,
    mVarrIn(15) = 0 'tnyCatId
    objDB.CreateNewConnection mCnn, enuSourceString.KMBR
    Set Rec = objDB.ExecuteSP("AddressIns", mVarrIn, , , mCnn, adCmdStoredProc)
End Sub

Private Sub Receipt()
Dim mVarrOut As Variant
Dim mVarrIn As Variant
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
ReDim mVarrIn(10)
    mVarrIn(0) = lSoochikaFeildID 'ReceiptId
    mVarrIn(1) = lSoochikaFeildID 'Main
    mVarrIn(2) = gbLocalBodyID 'Lb
    mVarrIn(3) = lSoochikaFeildID 'FileNo
    mVarrIn(4) = gbCounterID 'Counter
    mVarrIn(5) = intTransactionID_1 'TransationId
    mVarrIn(6) = mVoucherID 'VoucherId
    mVarrIn(7) = txtReceiptNo.Text 'ReceiptNo
    mVarrIn(8) = txtReceiptDate.Text 'dtReceipt
    mVarrIn(9) = val(fgAccHead.TextMatrix(1, 6)) 'Amount
    mVarrIn(10) = cboInstrument.ItemData(cboInstrument.ListIndex) 'CreditId
    objDB.CreateNewConnection mCnn, enuSourceString.KMBR
    Set Rec = objDB.ExecuteSP("Receipt_InsTR", mVarrIn, mVarrOut, , mCnn, adCmdStoredProc)
    If IsArray(mVarrOut) Then
       lReceiptID = mVarrOut(0, 0)
    End If
End Sub

Private Sub ReceiptChild()
Dim mVarrIn As Variant
Dim objDB As New clsDB
Dim Rec As New ADODB.Recordset
Dim mCnn As New ADODB.Connection
ReDim mVarrIn(7)
     mVarrIn(0) = 0 'intId
     mVarrIn(1) = lReceiptID 'intReceiptId
     mVarrIn(2) = gbLocalBodyID 'intLBId
     mVarrIn(3) = lSoochikaFeildID 'intMainId
     mVarrIn(4) = val(fgAccHead.TextMatrix(1, 1)) 'intDebitId
     mVarrIn(5) = Left(Trim(fgAccHead.TextMatrix(1, 6)), 4) 'fltAmount
     mVarrIn(6) = Left(Trim(fgAccHead.TextMatrix(1, 5)), 4) 'intFinancialyear
     mVarrIn(7) = Null 'intPeriod
    objDB.CreateNewConnection mCnn, enuSourceString.KMBR
    Set Rec = objDB.ExecuteSP("Receipt_InsTC", mVarrIn, , , mCnn, adCmdStoredProc)
     
End Sub


Private Sub txtPincode_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub lReceiptNo()
       Dim arrInput As Variant
        Dim arrOutPut As Variant
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mStr As String
        
        arrInput = Array(gbCounterID, 1, gbFinancialYearID)
        Set Rec = objDB.ExecuteSP("spGetNextReceiptNo", arrInput, arrOutPut, , mCnn, adCmdStoredProc)
        If IsArray(arrOutPut) Then
            txtReceiptNo.Text = arrOutPut(0, 0)
        End If
End Sub

Private Sub lSubFee()
    If cboBuidingType.ListIndex >= 0 Then
        If cboBuidingType.ItemData(cboBuidingType.ListIndex) = 1 Then
            fgAccHead.TextMatrix(1, 6) = 130
        ElseIf cboBuidingType.ItemData(cboBuidingType.ListIndex) = 2 Then
            fgAccHead.TextMatrix(1, 6) = 130
        ElseIf cboBuidingType.ItemData(cboBuidingType.ListIndex) = 3 Then
            fgAccHead.TextMatrix(1, 6) = 200
        ElseIf cboBuidingType.ItemData(cboBuidingType.ListIndex) = 4 Then
            fgAccHead.TextMatrix(1, 6) = 300
        End If
    End If
End Sub

Private Function lSaveValidate() As Boolean
    lSaveValidate = True
    If txtName.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter Name"
        Exit Function
    ElseIf txtHouseNo.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter House No."
        Exit Function
    ElseIf val(txtHouseWard.Text) = 0 Then
        lSaveValidate = False
        MsgBox "Enter House Ward No."
        Exit Function
    ElseIf txtHouseName.Text = "" Then
        lSaveValidate = False
        MsgBox "Enter House Name"
        Exit Function
    ElseIf cboDistrict.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select District"
        Exit Function
    ElseIf cmbPostOffice.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select Postoffice"
        Exit Function
'    ElseIf txtPhoneno.Text = "" Then
'        lSaveValidate = False
'        MsgBox "Enter Phone No."
'        Exit Function
    ElseIf cboSeat.ListIndex < 0 Then
        lSaveValidate = False
        MsgBox "Select Seat"
        Exit Function
    ElseIf chkStampPaper.value = 0 And chkOneday.value = 1 Then
        lSaveValidate = False
        MsgBox "Select Under Taking Stamp Paper"
        Exit Function
    ElseIf chkOwnership.value = 0 Then
        lSaveValidate = False
        MsgBox "Select Ownership"
        Exit Function
    ElseIf chkLandTaxReceipt.value = 0 Then
        lSaveValidate = False
        MsgBox "Select Land Tax Receipt"
        Exit Function
    ElseIf chkSitePlan.value = 0 Then
        lSaveValidate = False
        MsgBox "Select Site Plan"
        Exit Function
    ElseIf chkLocationPlan.value = 0 Then
        lSaveValidate = False
        MsgBox "Select Location Plan"
        Exit Function
    ElseIf chkBuildingPlan.value = 0 Then
        lSaveValidate = False
        MsgBox "Select Building Plan"
        Exit Function
    ElseIf cboZone.ListIndex <= 0 Then
        lSaveValidate = False
        MsgBox "Select Zone"
        Exit Function
     ElseIf cboWard.ListIndex <= 0 Then
        lSaveValidate = False
        MsgBox "Select Ward"
        Exit Function
    End If

End Function
Private Sub PrintReceipt(intVoucherID As Double)
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Dim mLoop As Long
        Dim mstrYear As String
        Dim mCount As Long
        Dim objCounter As New clsCounter
        Dim objUser As New clsUser
        Dim mName As String
        Dim mChequeNo As String
        
        'PrinterInit
        gbFileNO = FreeFile
        gbFileName = "C:\Report.txt"
        If Len(Dir(gbFileName)) Then
            Kill gbFileName
        End If
        
        FileInitialize
        mSQL = "Select faVouchers.fltAmount as TotalAmt, * From faVouchers Inner Join faVoucherChild "
        mSQL = mSQL + " On faVoucherChild.intVoucherID = faVouchers.intVoucherID "
        mSQL = mSQL + " Inner join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID "
        mSQL = mSQL + " Left Join faVoucherAddress On faVoucherAddress.intVoucherID = faVouchers.intVoucherID "
        mSQL = mSQL + " Where faVouchers.intVoucherID = " & intVoucherID
        objDB.SetConnection mCnn
        Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
        
        If Rec!intTransactionTypeID = gbTransactionTypePTax Then
            If Rec.RecordCount > 9 Then
                Rec.Close
                Call PrintSummaryReceiptPTax(intVoucherID)
                Exit Sub
            End If
        End If
        On Error Resume Next
        Open gbFileName For Output As #gbFileNO
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        
        Select Case Rec!intInstrumentTypeID
        
        Case Is = 1
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CASH"; Tab(76); "CASH"; gbDoubleWidthOff
        Case Is = 4
            Print #gbFileNO, Tab(31); gbDoubleWidth; "Demand Draft"; Tab(76); "Demand Draft"; gbDoubleWidthOff
            mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Is = 5
            Print #gbFileNO, Tab(31); gbDoubleWidth; "CHEQUE"; Tab(76); "CHEQUE"; gbDoubleWidthOff
            mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
            mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
        Case Else
            Print #gbFileNO,
        End Select
        
        If Not (Rec.EOF And Rec.BOF) Then
            ' Line 6
            'Print #gbFileNO, Tab(31); gbBold; IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); Tab(120); IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo); gbBoldOff
            '    Modified By Cijith on 23/04/2009 For KMBR  ----------------------------------------'
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo); Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(65); IIf(IsNull(Rec!numInwardNo), "", Rec!numInwardNo); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            '-------------------------------------------------------------------------------------'
            Print #gbFileNO, Tab(31); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate)); Tab(120); IIf(IsNull(Rec!dtDate), "", DdMmmYy(Rec!dtDate))
            
            mName = IIf(IsNull(Rec!vchName), "", Rec!vchName)
            If Not IsNull(Rec!vchInit1) Then mName = mName & " " & Rec!vchInit1
            If Not IsNull(Rec!vchInit2) Then mName = mName & " " & Rec!vchInit2
            If Not IsNull(Rec!vchInit3) Then mName = mName & " " & Rec!vchInit3
            If Not IsNull(Rec!vchInit4) Then mName = mName & " " & Rec!vchInit4
            
            Print #gbFileNO, Tab(15); Style(mName, True); Tab(65); Style(mName, True)
            
            'Changed for Sujith by Aiby - 24-Mar-2009
            
            Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName); Tab(65); IIf(IsNull(Rec!vchHouseName), "", Rec!vchHouseName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName); Tab(65); IIf(IsNull(Rec!vchStreetName), "", Rec!vchStreetName)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace); Tab(65); IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
            Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice); Tab(65); IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
            'Print #gbFileNO, Tab(15); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber); Tab(65); IIf(IsNull(Rec!vchDistrict), "", Rec!vchDistrict) & " - "; IIf(IsNull(Rec!vchPinNumber), "", Rec!vchPinNumber)
            'Print #gbFileNO, Tab(15); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone); Tab(65); "Ph : " & IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
            
            ' --------------------------------------------------------------------------------- '
            ' To Print Check Number and DD Number Printing Phone Number is Commented
            ' --------------------------------------------------------------------------------- '
            Select Case Rec!intInstrumentTypeID
            Case Is = 1
                Print #gbFileNO,
            Case Is = 4
                mChequeNo = "DD No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Is = 5
                mChequeNo = "Cheque No.:" + IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mChequeNo = mChequeNo + "\" + DdMmmYy(Rec!dtInstrumentDate)
                Print #gbFileNO, Tab(15); mChequeNo; Tab(65); mChequeNo
            Case Else
                Print #gbFileNO,
            End Select
            
            ' Line 15 Next
            'Changed its Possition- Requested by Sujith on 24-Mar-2009
            'Print #gbFileNO, Tab(15); gbBold; IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2);
            'Print #gbFileNO, Tab(65); IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo & "/"); IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo); IIf(IsNull(Rec!vchDoorNo2), "", "-" & Rec!vchDoorNo2); gbBoldOff
            
            Print #gbFileNO, "Ref.No: "; Tab(10); IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo); Tab(55); "Ref.No: "; IIf(IsNull(Rec!vchRefNo), "", Rec!vchRefNo)
            Print #gbFileNO,
            Print #gbFileNO,
            ' Line 18 Next
            Rec.MoveFirst
            While Not Rec.EOF
                mLoop = mLoop + 1
                
                '==================================================================='
                ' Counter Foil
                '==================================================================='
                Print #gbFileNO, IIf(IsNull(Rec!vchAccountHeadCode), "", Rec!vchAccountHeadCode);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(12); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(12); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(12); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(12); mstrYear;
                    
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(27); PadL(Format(Rec!fltAmount, "0.00"), 9);
                Else
                    Print #gbFileNO, Tab(37); PadL(Format(Rec!fltAmount, "0.00"), 9);
                End If
                
                '==================================================================='
                ' Receipt Area
                '==================================================================='
                Print #gbFileNO, Tab(48); PadL(CStr(mLoop), 2);
                Print #gbFileNO, Tab(56); PadR(Rec!vchAlias, 41);
                If Not IsNull(Rec!intYearID) Then
                    mstrYear = CStr(Rec!intYearID) & "-" & Right(CStr(Rec!intYearID + 1), 2)
                Else
                    mstrYear = ""
                End If
                Select Case Rec!tnyPeriodID
                    Case Is = 1: Print #gbFileNO, Tab(98); mstrYear & "/1Hf";
                    Case Is = 2: Print #gbFileNO, Tab(98); mstrYear & "/2Hf";
                    Case Is = 3: Print #gbFileNO, Tab(98); mstrYear & "/F";
                    Case Else:   Print #gbFileNO, Tab(98); mstrYear;
                End Select
                
                If Rec!intYearID < gbFinancialYearID Then
                    Print #gbFileNO, Tab(109); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Else
                    Print #gbFileNO, Tab(126); PadL(Format(Rec!fltAmount, "0.00"), 9)
                End If
                'Print #gbFileNO, Tab(26); PadL(Trim(str(mLoop)), 3); Tab(31); Rec!vchAccountHeadCode; Tab(40); PadR(IIf(IsNull(Rec!vchAlias), "", Rec!vchAlias), 20); Rec!tnyPeriodID; Tab(70); PadL(Format(Rec!fltAmount, "0.00"), 9)
                Rec.MoveNext
            Wend
            Rec.MoveFirst
            
            For mCount = mLoop + 1 To 9
                Print #gbFileNO,
            Next mCount
            If Rec!fltAdvAmtAdj > 0 Then
                Print #gbFileNO, PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 46); Tab(47); PadL("Adv.Adj(" & Format(Rec!fltAdvAmtAdj, "0.00") & ")", 89)
            Else
                Print #gbFileNO,
            End If
            Print #gbFileNO, Tab(22); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"; Tab(76); "Rnd.Off("; Format(Rec!fltRoundOff, "0.00"); ")"
                            
            Print #gbFileNO, Tab(29); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True);
            Print #gbFileNO, Tab(117); Style(PadL(Format(Rec!TotalAmt, "0.00"), 10), True)
            
            Print #gbFileNO, Tab(7); Rupees(Rec!TotalAmt);
            Print #gbFileNO, Tab(65); Rupees(Rec!TotalAmt)
            Print #gbFileNO,
            Print #gbFileNO, Tab(7); PadR(Trim(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)), 40); Tab(61); IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription)
            
            Print #gbFileNO,
            objCounter.SetCounter (Rec!intCounterID)
            If objCounter.CounterID > 0 Then
                Print #gbFileNO, Tab(11); objCounter.CounterNo;
                Print #gbFileNO, Tab(61); objCounter.CounterNo & " : " & objCounter.CounterDescription
            End If
            objUser.SetUser (Rec!intUserID)
            If objUser.UserID > -1 Then
                Print #gbFileNO, Tab(11); objUser.UserName;
                Print #gbFileNO, Tab(61); objUser.UserName
            End If
        End If
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        'Print #gbFileNO,
        
        'Print #gbFileNO, 'Chr$(27) + Chr$(12)
        
        Close #gbFileNO
        ShellPad
        Shell "Print " & gbFileName
        'Kill gbFileName
    End Sub

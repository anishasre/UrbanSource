VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmImpOfficer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Implementing Officer"
   ClientHeight    =   7215
   ClientLeft      =   240
   ClientTop       =   810
   ClientWidth     =   12870
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   450
      Left            =   6090
      TabIndex        =   51
      Top             =   6615
      Width           =   1140
   End
   Begin VB.TextBox txtDept 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   8235
      TabIndex        =   17
      Top             =   2625
      Width           =   3850
   End
   Begin VB.TextBox txtDDOCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8235
      TabIndex        =   16
      Top             =   2280
      Width           =   4260
   End
   Begin VB.TextBox txtSubTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8235
      TabIndex        =   15
      Top             =   1905
      Width           =   4260
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8235
      TabIndex        =   14
      Top             =   1530
      Width           =   4290
   End
   Begin VB.TextBox txtResPhone 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   10
      Top             =   4980
      Width           =   4380
   End
   Begin VB.TextBox txtDoorNo2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3045
      TabIndex        =   9
      Top             =   4590
      Width           =   1215
   End
   Begin VB.TextBox txtDoorNo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   4590
      Width           =   1215
   End
   Begin VB.TextBox txtWardNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   50
      Top             =   4200
      Width           =   4380
   End
   Begin VB.TextBox txtMainPlace 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   3090
      Width           =   4380
   End
   Begin VB.TextBox txtStreet 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   2340
      Width           =   4380
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   1845
      TabIndex        =   47
      Top             =   360
      Width           =   9090
      Begin VB.ComboBox cmbImpo 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmImplOfficer.frx":0000
         Left            =   3255
         List            =   "frmImplOfficer.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   270
         Width           =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   1065
         TabIndex        =   49
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   7305
      TabIndex        =   46
      Top             =   6615
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesg 
      Caption         =   "..."
      Height          =   315
      Left            =   12120
      TabIndex        =   45
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdDept 
      Caption         =   "..."
      Height          =   315
      Left            =   12105
      TabIndex        =   44
      Top             =   2625
      Width           =   375
   End
   Begin VB.TextBox txtPinCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   3840
      Width           =   4380
   End
   Begin VB.TextBox txtPostOffice 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   3480
      Width           =   4380
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
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   2730
      Width           =   4380
   End
   Begin VB.TextBox txtdtTo 
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
      Left            =   10575
      TabIndex        =   21
      Top             =   3765
      Width           =   1900
   End
   Begin VB.TextBox txtdtFrom 
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
      Left            =   8235
      TabIndex        =   20
      Top             =   3765
      Width           =   1900
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   1620
      TabIndex        =   13
      Top             =   5760
      Width           =   4380
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   5370
      Width           =   4380
   End
   Begin VB.TextBox txtOfcPhone 
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
      Left            =   8235
      TabIndex        =   22
      Top             =   4170
      Width           =   4260
   End
   Begin VB.TextBox txtOpeningBalance 
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
      Left            =   8235
      TabIndex        =   19
      Top             =   3375
      Width           =   4260
   End
   Begin VB.TextBox txtDesignation 
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
      Height          =   315
      Left            =   8235
      TabIndex        =   18
      Top             =   3000
      Width           =   3850
   End
   Begin VB.Frame Frame1 
      Caption         =   "User name and Password"
      ForeColor       =   &H000000C0&
      Height          =   1185
      Left            =   6330
      TabIndex        =   41
      Top             =   5370
      Width           =   6315
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   23
         Top             =   345
         Width           =   2430
      End
      Begin VB.TextBox txtPswd 
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
         Height          =   325
         IMEMode         =   3  'DISABLE
         Left            =   1725
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   43
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   42
         Top             =   825
         Width           =   855
      End
   End
   Begin VB.TextBox txtHouseName 
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
      Left            =   1620
      TabIndex        =   2
      Top             =   1980
      Width           =   4380
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
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   1620
      Width           =   4380
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   450
      Left            =   4665
      TabIndex        =   34
      Top             =   6615
      Width           =   1335
   End
   Begin VB.Frame fmeCodeTitle 
      Caption         =   "Other Details"
      ForeColor       =   &H000000C0&
      Height          =   4050
      Left            =   6300
      TabIndex        =   32
      Top             =   1320
      Width           =   6360
      Begin VB.TextBox txtPENNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   53
         Top             =   3240
         Width           =   4260
      End
      Begin VB.TextBox txtPANNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   52
         Top             =   3615
         Width           =   4260
      End
      Begin VB.TextBox txtDepartment 
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
         Height          =   315
         Left            =   7995
         TabIndex        =   12
         Top             =   1320
         Width           =   3885
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "PAN/GIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "PEN OF DDO/SDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "DDO Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Date From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Office Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3945
         TabIndex        =   40
         Top             =   2460
         Width           =   210
      End
   End
   Begin VB.Frame fmePersonal 
      Caption         =   "Personal Information"
      ForeColor       =   &H000000C0&
      Height          =   5235
      Left            =   180
      TabIndex        =   0
      Top             =   1320
      Width           =   6045
      Begin VB.Label Email 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   39
         Top             =   4455
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   38
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   37
         Top             =   1845
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   36
         Top             =   2550
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Post Office"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Top             =   2220
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   31
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   705
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   29
         Top             =   1455
         Width           =   1020
      End
      Begin VB.Label lblWardNo 
         AutoSize        =   -1  'True
         Caption         =   "Ward No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   2925
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   1065
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Door No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   3300
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2835
         X2              =   2745
         Y1              =   3285
         Y2              =   3555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Phone ( Res)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   25
         Top             =   3720
         Width           =   1140
      End
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   10305
      Top             =   8640
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Creating Implementing Officer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   390
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmImpOfficer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbImpo_Click()
          On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim Rec As New ADODB.Recordset
            Dim mSql As String
            Dim objdb As New clsDB
            Dim mRowCnt As Integer
            
            If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) Then
                mSql = "Select * from suImplementingOfficer Where  intImplementingOfficerID=" & cmbImpo.ItemData(cmbImpo.ListIndex)
                Rec.Open mSql, mCnn
                txtTitle.Text = IIf(IsNull(Rec!vchImplementingOfficer), "", Rec!vchImplementingOfficer)
                txtSubTitle.Text = IIf(IsNull(Rec!vchImplementingOfficerCode), "", Rec!vchImplementingOfficerCode)
            Else
                MsgBox "Connection to Sulekha does not Exist, Please Contact your System Operator", vbInformation
            End If
            Exit Sub
Err:
        MsgBox (Error$)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDept_Click()
          Dim mSql As String
            frmSearchMasters.QrySP = Qyery
            'frmSearchMasters.SQLQry = "spSelectScheme"
            frmSearchMasters.SQLQry = "Select intDepartmentID,vchDepartment from faIMPODepartments"
            frmSearchMasters.Connection = enuSourceString.Saankhya
            frmSearchMasters.Show vbModal
            If gbSearchStr <> "" Then
                txtDept.Text = gbSearchStr
                txtDept.Tag = gbSearchID
            End If
            gbSearchStr = ""
            gbSearchID = -1
End Sub

Private Sub cmdDesg_Click()
        Dim mSql As String
        frmSearchMasters.QrySP = Qyery
        frmSearchMasters.SQLQry = "Select intDesignationID,vchDesignation from faIMPODesignations"
        frmSearchMasters.Connection = enuSourceString.Saankhya
        frmSearchMasters.Show vbModal
        If gbSearchStr <> "" Then
              txtDesignation.Text = gbSearchStr
              txtDesignation.Tag = gbSearchID
        End If
        gbSearchStr = ""
        gbSearchID = -1
End Sub
Private Sub cmdSave_Click()
   Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    objdb.SetConnection mCnn
  If SaveValidation Then
        If txtName.Tag <> "" Then
            mSql = "Select tnySyncFlag from faSubSidiaryAccountHeads Where vchSubLedgerCode=" & txtName.Tag
            Rec.Open mSql, mCnn
            mSql = ""
            If Rec!tnySyncFlag > 0 Then
               mSql = "Update faSubsidiaryAccountHeads Set chbPassword=convert(varbinary,'" & txtPswd.Text & "'),tnySyncFlag=0 where vchSubLedgerCode=" & txtName.Tag
               mCnn.Execute mSql
               MsgBox "Password Changed", vbInformation
               frmImpOfficer.Visible = False
               frmImplementingOfficerList.Show
            Else
                If SaveSubLedger Then
                MsgBox "Implementing Officer Saved Successfully", vbInformation
                Call clearField
                frmImpOfficer.Visible = False
                frmImplementingOfficerList.Show
           End If
            End If
        Else
           If SaveSubLedger Then
                MsgBox "Implementing Officer Saved Successfully", vbInformation
                Call clearField
                Unload Me
                frmImplementingOfficerList.Show
           End If
        End If
     
  End If
End Sub
  Private Function SaveValidation() As Boolean
    
'                If cmbImpo.ListIndex = -1 Then
'                    MsgBox "Please Select Implementing officer", vbInformation
'                    cmbImpo.SetFocus
'                    SaveValidation = False
'                    Exit Function
'                End If
                If Trim(txtName.Text) = "" Then
                    MsgBox "Please Enter the Name", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtTitle.Text = "") Then
                    MsgBox "Please Enter  Title", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtSubTitle.Text = "") Then
                    MsgBox "Please Enter  SubTitle", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtDDOCode.Text = "") Then
                    MsgBox "Please Enter  DDO Code", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
               If Trim(txtDept.Text = "") Then
                    MsgBox "Please Enter  Department", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtDesignation.Text = "") Then
                    MsgBox "Please Enter  Designation", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtOpeningBalance.Text = "") Then
                    MsgBox "Please Enter  OpeningBalance", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                If Trim(txtUserName.Text = "") Or Trim(txtPswd.Text = "") Then
                    MsgBox "Please Enter  User Name and Password", vbInformation
                    SaveValidation = False
                    Exit Function
                End If
                SaveValidation = True
        Exit Function
  End Function
   Private Function SaveSubLedger() As Boolean
            Dim objdb As New clsDB
            Dim mCnn As New ADODB.Connection
            Dim aryIn As Variant
            Dim aryOut As Variant
            Dim mDtFrom As Variant
            Dim mDtTo As Variant
            Dim mFlag As Variant
            If txtUserName.Tag = "" Then
                mFlag = 0
            Else
                mFlag = txtUserName.Tag
            End If
          
            mDtFrom = txtdtFrom.Text
            mDtTo = txtdtTo.Text
            If objdb.SetConnection(mCnn) Then
                aryIn = Array(1, _
                    val(txtName.Tag), _
                    Trim(txtTitle.Text), _
                    Trim(txtSubTitle.Text), _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    val(txtOpeningBalance.Text), _
                    Null, _
                    txtHouseName.Tag, _
                    txtDDOCode.Text, _
                    Null, _
                    Trim(txtName.Text), _
                    Trim(txtHouseName.Text), _
                    Trim(txtStreet.Text), _
                    Trim(txtLocalPlace.Text), _
                    Trim(txtMainPlace.Text), _
                    Trim(txtPostOffice.Text), _
                    Trim(txtPinCode.Text), _
                    Trim(txtOfcPhone.Text), _
                    val(txtWardNo.Text), _
                    val(Trim(txtDoorNo1.Text)), _
                    Trim(txtDoorNo2.Text), gbFinancialYearID, txtDesignation.Text, txtDept.Text, Null, txtDesignation.Tag, txtDept.Tag, Null, _
                    Trim(txtUserName.Text), txtPswd.Text, Trim(txtOfcPhone.Text), Trim(txtMobile.Text), Trim(txtEmail.Text), mDtFrom, mDtTo, Null, Null, mFlag, gbLocalBodyID, 0, txtPENNo.Text, txtPANNo.Text)
                    
                    
                    objdb.ExecuteSP "spSaveSubSidiaryAccountHeads", aryIn, aryOut, , mCnn
                    SaveSubLedger = True
                  
         Else
                MsgBox "Connection to Finance doesnot Exist, Please contact your System Administrator", vbInformation
                SaveSubLedger = False
         End If
         Exit Function
    End Function
Private Sub cmdSuspend_Click()
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    objdb.SetConnection mCnn
    If txtName.Tag <> "" Then
        mSql = "Select tnyDeleted,tnySyncFlag from faSubSidiaryAccountHeads Where vchSubLedgerCode=" & txtName.Tag
        Rec.Open mSql, mCnn
        mSql = ""
        If Rec!tnyDeleted = 1 Then
            mSql = "Update faSubSidiaryAccountHeads Set tnyDeleted=0,tnySyncFlag=0 Where vchSubLedgerCode=" & txtName.Tag
            mCnn.Execute (mSql)
            cmdSuspend.Caption = "Suspend"
            MsgBox "Implementing Officer Activated", vbInformation
            frmImpOfficer.Visible = False
            frmImplementingOfficerList.Show
        Else
            mSql = "Update faSubSidiaryAccountHeads Set tnyDeleted=1,tnySyncFlag=0 Where vchSubLedgerCode=" & txtName.Tag
            mCnn.Execute (mSql)
            cmdSuspend.Caption = "Acivate"
            MsgBox "Implementing Officer Suspended", vbInformation
            frmImpOfficer.Visible = False
            frmImplementingOfficerList.Show
        End If
    End If
  
End Sub
Private Sub Form_Activate()
    Me.Top = 1300
    Me.Left = 0
End Sub
Public Sub Form_Load()
    Dim mSubLedgerID As Integer
    WindowsXPC1.InitIDESubClassing
    PopulateList cmbImpo, "Select vchImplementingOfficer,intImplementingOfficerID from suImplementingOfficer Where intLBTypeID = " & gbLBType & " Order By vchImplementingOfficer", , True, , True
    txtdtFrom.Text = DdMmmYy(gbTransactionDate)
    'txtdtFrom.Enabled = False
   ' cmdSuspend.Visible = False
   ' cmdSave.Left = 5000
   ' cmdCancel.Left = 6500
    If mSubLedgerID <> 0 Then
      Call fillimpOfficer(mSubLedgerID)
    End If
End Sub
Public Sub fillimpOfficer(mSubLedgerID As Integer)
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    'txtdtFrom.Enabled = True
  
    
    mSql = "SELECT intSubsidiaryAccountHeadID,intSubLedgerTypeID,vchSubLedgerCode,vchTitle,vchSubTitle,fltOpeningBalance,vchReferenceCode,vchName,vchHouseOrOffice,vchStreet,"
    mSql = mSql + "   vchLocalPlace,vchMainPlace,vchPhone,numWardNo,intDoorNo, vchDoorNo2,intLBID,tnyDeleted,vchDesignation,vchDepartment,intDepartmentID,intDesignationID,intDepartmentID,"
    mSql = mSql + " vchPostOffice,vchPinCode,chvUserName,convert(varchar,chbPassword) chbPassword,vchoffphone,vchMphone,vchEmail,dtFromDate,dtToDate,gbIMPOID,gbIMPOCode,tnySyncFlag,tnyDeleted  FROM  faSubSidiaryAccountHeads where intSubsidiaryAccountHeadID=" & mSubLedgerID
    objdb.SetConnection mCnn
    Rec.Open mSql, mCnn
    If Not (Rec.BOF And Rec.EOF) Then
         
           txtName.Tag = IIf(IsNull(Rec!vchSubLedgerCode), "", Rec!vchSubLedgerCode)
           
           txtName.Text = IIf(IsNull(Rec!vchName), "", Rec!vchName)
           txtHouseName.Text = IIf(IsNull(Rec!vchHouseOrOffice), "", Rec!vchHouseOrOffice)
           txtLocalPlace.Text = IIf(IsNull(Rec!vchLocalPlace), "", Rec!vchLocalPlace)
           txtStreet.Text = IIf(IsNull(Rec!vchStreet), "", Rec!vchStreet)
           txtMainPlace.Text = IIf(IsNull(Rec!vchMainPlace), "", Rec!vchMainPlace)
           txtWardNo.Text = IIf(IsNull(Rec!numWardNo), "", Rec!numWardNo)
           txtOfcPhone.Text = IIf(IsNull(Rec!vchPhone), "", Rec!vchPhone)
           txtDoorNo1.Text = IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
           txtDoorNo2.Text = IIf(IsNull(Rec!vchDoorNo2), "", Rec!vchDoorNo2)
           
           txtTitle.Text = IIf(IsNull(Rec!vchTitle), "", Rec!vchTitle)
           txtSubTitle.Text = IIf(IsNull(Rec!vchSubTitle), "", Rec!vchSubTitle)
           txtDDOCode.Text = IIf(IsNull(Rec!vchReferenceCode), "", Rec!vchReferenceCode)
           txtDept.Text = IIf(IsNull(Rec!vchDepartment), "", Rec!vchDepartment)
           txtDept.Tag = IIf(IsNull(Rec!intDepartmentID), "", Rec!intDepartmentID)
           txtDesignation.Text = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
           txtDesignation.Tag = IIf(IsNull(Rec!intDesignationID), "", Rec!intDesignationID)
           txtOpeningBalance.Text = IIf(IsNull(Rec!fltOpeningBalance), "", Rec!fltOpeningBalance)
           txtUserName.Text = IIf(IsNull(Rec!chvUserName), "", Rec!chvUserName)
           txtUserName.Tag = IIf(IsNull(Rec!tnySyncFlag), 0, Rec!tnySyncFlag)
           txtPswd.Text = IIf(IsNull(Rec!chbPassword), 0, Rec!chbPassword)
           
           txtMobile.Text = IIf(IsNull(Rec!vchMphone), "", Rec!vchMphone)
           txtEmail.Text = IIf(IsNull(Rec!vchEmail), "", Rec!vchEmail)
           txtdtFrom.Text = IIf(IsNull(Rec!dtFromDate), "", Rec!dtFromDate)
           txtdtTo.Text = IIf(IsNull(Rec!dtToDate), "", Rec!dtToDate)
           txtPostOffice.Text = IIf(IsNull(Rec!vchPostOffice), "", Rec!vchPostOffice)
           txtPinCode.Text = IIf(IsNull(Rec!vchPinCode), "", Rec!vchPinCode)
    
           '------------------After Ported To Web--------------------------
           If Rec!tnySyncFlag > 0 Then
                cmbImpo.Enabled = False
                cmdDept.Enabled = False
                cmdDesg.Enabled = False
                 
                txtUserName.Enabled = False
                                  
                txtName.Enabled = False
                txtHouseName.Enabled = False
                txtLocalPlace.Enabled = False
                txtStreet.Enabled = False
                txtMainPlace.Enabled = False
                txtWardNo.Enabled = False
                txtOfcPhone.Enabled = False
                txtDoorNo1.Enabled = False
                txtDoorNo2.Enabled = False
                
                txtTitle.Enabled = False
                txtSubTitle.Enabled = False
                txtDDOCode.Enabled = False
                txtDept.Enabled = False
                txtDesignation.Enabled = False
                txtOpeningBalance.Enabled = False
                txtResPhone.Enabled = False
           
                txtMobile.Enabled = False
                txtEmail.Enabled = False
                txtdtFrom.Enabled = False
                txtdtTo.Enabled = False
                txtPostOffice.Enabled = False
                txtPinCode.Enabled = False
                
               txtUserName.Enabled = True
               txtPswd.Enabled = True
               cmdSave.Enabled = True
               cmdSave.Caption = "ResetPassword"
           ElseIf Rec!tnySyncFlag = 0 Then
               txtUserName.Enabled = True
               txtPswd.Enabled = True
               cmdSave.Enabled = True
               cmdSave.Caption = "Save"
           End If
           '---------Suspend and Activate a User-------------------------------------
           If Rec!tnyDeleted = 1 Then
                cmdSuspend.Caption = "Acivate"
           Else
                cmdSuspend.Caption = "Suspend"
           End If
           '--------------------------------------------------------------------------
           'Suspend Available only When Once Saved
            If IsNull(Rec!tnySyncFlag) Then
                     cmdSuspend.Visible = False
                     cmdSave.Left = 5100
                     cmdCancel.Left = 6700
            End If
      End If
End Sub
Private Sub clearField()
           txtName.Tag = 0
           
           txtName.Text = ""
           txtHouseName.Text = ""
           txtLocalPlace.Text = ""
           txtStreet.Text = ""
           txtMainPlace.Text = ""
           txtWardNo.Text = ""
           txtOfcPhone.Text = ""
           txtDoorNo1.Text = ""
           txtDoorNo2.Text = ""
           
           txtTitle.Text = ""
           txtSubTitle.Text = ""
           txtDDOCode.Text = ""
           txtDept.Text = ""
           txtDesignation.Text = ""
           txtOpeningBalance.Text = ""
           txtUserName.Text = ""
           txtPswd.Text = ""
       
           txtMobile.Text = ""
           txtEmail.Text = ""
           txtdtFrom.Text = ""
           txtdtTo.Text = ""
           
End Sub

Private Sub txtdtFrom_LostFocus()
    txtdtFrom.Text = CheckDateInMMM(txtdtFrom.Text)
End Sub

Private Sub txtDtTo_LostFocus()
    txtdtTo.Text = CheckDateInMMM(txtdtTo.Text)
End Sub
   

VERSION 5.00
Begin VB.Form frmChequeDetails 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   4605
      Left            =   540
      TabIndex        =   0
      Top             =   330
      Width           =   9615
      Begin VB.TextBox txtBankName 
         Height          =   315
         Left            =   6570
         TabIndex        =   32
         Top             =   3930
         Width           =   2475
      End
      Begin VB.TextBox txtInstrumentDate 
         Height          =   315
         Left            =   6570
         TabIndex        =   31
         Top             =   3420
         Width           =   2475
      End
      Begin VB.TextBox txtInstrumentNo 
         Height          =   315
         Left            =   6570
         TabIndex        =   30
         Top             =   2910
         Width           =   2475
      End
      Begin VB.TextBox txtReceivedDate 
         Height          =   315
         Left            =   6570
         TabIndex        =   29
         Top             =   2400
         Width           =   2475
      End
      Begin VB.TextBox txtForwardedSeatID 
         Height          =   315
         Left            =   6570
         TabIndex        =   28
         Top             =   1890
         Width           =   2475
      End
      Begin VB.TextBox txtFileNo 
         Height          =   315
         Left            =   6570
         TabIndex        =   27
         Top             =   1380
         Width           =   2475
      End
      Begin VB.TextBox txtDoorNo 
         Height          =   315
         Left            =   6570
         TabIndex        =   26
         Top             =   870
         Width           =   2475
      End
      Begin VB.TextBox txtWardNo 
         Height          =   315
         Left            =   6570
         TabIndex        =   25
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox txtPhone 
         Height          =   315
         Left            =   1500
         TabIndex        =   16
         Top             =   3960
         Width           =   2475
      End
      Begin VB.TextBox txtPin 
         Height          =   315
         Left            =   1500
         TabIndex        =   15
         Top             =   3480
         Width           =   2475
      End
      Begin VB.TextBox txtPost 
         Height          =   315
         Left            =   1500
         TabIndex        =   14
         Top             =   2930
         Width           =   2475
      End
      Begin VB.TextBox txtMainPlace 
         Height          =   315
         Left            =   1500
         TabIndex        =   13
         Top             =   2416
         Width           =   2475
      End
      Begin VB.TextBox txtLocalPlace 
         Height          =   315
         Left            =   1500
         TabIndex        =   12
         Top             =   1902
         Width           =   2475
      End
      Begin VB.TextBox txtStreet 
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   1388
         Width           =   2475
      End
      Begin VB.TextBox txtHouseName 
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         Top             =   874
         Width           =   2445
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label lblDoorNo 
         Caption         =   "Door No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5790
         TabIndex        =   24
         Top             =   900
         Width           =   705
      End
      Begin VB.Label lblWardNo 
         Caption         =   "Ward No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5730
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblBankName 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5490
         TabIndex        =   22
         Top             =   3930
         Width           =   1065
      End
      Begin VB.Label lblInstrumentDate 
         Caption         =   "Instrument Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5070
         TabIndex        =   21
         Top             =   3420
         Width           =   1545
      End
      Begin VB.Label lblinstrumentNo 
         Caption         =   "Instrument No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5250
         TabIndex        =   20
         Top             =   2910
         Width           =   1275
      End
      Begin VB.Label lblReceivedDate 
         Caption         =   "Received Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5220
         TabIndex        =   19
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label lblForwardedSeatID 
         Caption         =   "Forwarded Seat"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5130
         TabIndex        =   18
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label lblFileNo 
         Caption         =   "File No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5910
         TabIndex        =   17
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   930
         TabIndex        =   8
         Top             =   3960
         Width           =   585
      End
      Begin VB.Label lblPin 
         Caption         =   "Pin"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   7
         Top             =   3450
         Width           =   285
      End
      Begin VB.Label lblPost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   6
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label lblMainPlace 
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         TabIndex        =   5
         Top             =   2430
         Width           =   1005
      End
      Begin VB.Label lblLocalPlace 
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1890
         Width           =   1005
      End
      Begin VB.Label lblStreet 
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   870
         TabIndex        =   3
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblHouseName 
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmChequeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmChequeDetails.Height = 5940
    frmChequeDetails.Width = 11020
End Sub

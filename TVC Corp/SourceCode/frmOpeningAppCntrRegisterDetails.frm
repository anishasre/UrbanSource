VERSION 5.00
Begin VB.Form frmOpeningAppCntrRegisterDetails 
   Caption         =   "OpeningAppCntrRegisterDetails"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   8895
      Begin VB.TextBox txtAccountHeadCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Top             =   2520
         Width           =   3735
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1200
         Width           =   3735
      End
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdSearchAccountHeads 
         Caption         =   "..."
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         Caption         =   "..."
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdSearchFunction 
         Caption         =   "..."
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRequisitionNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Functionary"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Function"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Requistion No."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   8835
      TabIndex        =   2
      Top             =   5280
      Width           =   8895
      Begin VB.CommandButton cmdApprove 
         Caption         =   "&Approve"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Appropriation Control Register Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmOpeningAppCntrRegisterDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

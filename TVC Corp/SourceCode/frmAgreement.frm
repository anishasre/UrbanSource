VERSION 5.00
Begin VB.Form frmAgreement 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "ML-TTRevathi"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgreement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtProjectNAme 
      Height          =   315
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   990
      Width           =   3690
   End
   Begin VB.TextBox txtOrderNo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   38
      Top             =   3150
      Width           =   1860
   End
   Begin VB.TextBox txtDueDateofCommencement 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   25
      Top             =   1860
      Width           =   1860
   End
   Begin VB.TextBox txtduedateofCompletion 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   24
      Top             =   2670
      Width           =   1860
   End
   Begin VB.TextBox txtPAC 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      MaxLength       =   15
      TabIndex        =   21
      Top             =   4650
      Width           =   1860
   End
   Begin VB.TextBox txtAsset 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   19
      Top             =   4155
      Width           =   1860
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5175
      TabIndex        =   18
      Top             =   4140
      Width           =   285
   End
   Begin VB.TextBox txtWorkTitle 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   17
      Top             =   3645
      Width           =   5475
   End
   Begin VB.TextBox txtAgreementNoPart2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4650
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   495
      Width           =   510
   End
   Begin VB.TextBox txtWorkDate 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   13
      Top             =   3150
      Width           =   1860
   End
   Begin VB.TextBox txtCompletionDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   12
      Top             =   2280
      Width           =   1860
   End
   Begin VB.TextBox txtComencementDate 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   11
      Top             =   1440
      Width           =   1860
   End
   Begin VB.TextBox txtProjectNo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   975
      Width           =   1860
   End
   Begin VB.CommandButton cmdProject 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8910
      TabIndex        =   9
      Top             =   990
      Width           =   285
   End
   Begin VB.TextBox txtContractors 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   7
      Top             =   5160
      Width           =   5490
   End
   Begin VB.CommandButton cmdSearchBeneficiary 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8805
      TabIndex        =   6
      Top             =   5160
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9480
      TabIndex        =   3
      Top             =   5775
      Width           =   9540
      Begin VB.CommandButton cmdApprove 
         Caption         =   "&Approve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7680
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   42
         Top             =   120
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   41
         Top             =   120
         Width           =   1440
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.TextBox txtAgreementNoPart1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   510
      Width           =   1200
   End
   Begin VB.CommandButton cmdAgreementNo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5175
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtAgreementDate 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7350
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   1470
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order No:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5970
      TabIndex        =   39
      Top             =   3165
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agreement No:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   1905
      TabIndex        =   37
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers/ Contractors/ Beneficiary Committee Convener"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   660
      TabIndex        =   36
      Top             =   4980
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project No:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   2115
      TabIndex        =   35
      Top             =   990
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Commencement of Work"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   405
      TabIndex        =   34
      Top             =   1350
      Width           =   2670
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Completion:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1215
      TabIndex        =   33
      Top             =   2295
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Work/ Supply Order"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   1725
      TabIndex        =   32
      Top             =   3030
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title of Work:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1215
      TabIndex        =   31
      Top             =   3660
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Asset:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1800
      TabIndex        =   30
      Top             =   4170
      Width           =   1470
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agreed PAC:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2130
      TabIndex        =   29
      Top             =   4665
      Width           =   1155
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date Of Completion:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   750
      TabIndex        =   28
      Top             =   2685
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date Of Commencement of Work"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   480
      TabIndex        =   27
      Top             =   1830
      Width           =   2610
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   3000
      TabIndex        =   26
      Top             =   1710
      Width           =   300
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Agreement For Work Or Supply Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asset - Head Code"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5715
      TabIndex        =   20
      Top             =   4200
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4470
      TabIndex        =   16
      Top             =   495
      Width           =   165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   2970
      TabIndex        =   14
      Top             =   2865
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   2955
      TabIndex        =   8
      Top             =   4905
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agreement Date:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   5775
      TabIndex        =   5
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   720
      Left            =   2970
      TabIndex        =   4
      Top             =   1200
      Width           =   300
   End
End
Attribute VB_Name = "frmAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            ElseIf TypeOf mCrl Is OptionButton Then
                mCrl.value = False
            ElseIf TypeOf mCrl Is ComboBox Then
                If mCrl.ListCount > 0 Then mCrl.ListIndex = 0
            ElseIf TypeOf mCrl Is ComboBox Then
                mCrl.ListIndex = -1
            End If
        Next
        
    End Sub
'''    Private Sub FillAgreements(ByVal AgreementID As Integer)
'''       Dim objAgreement As New clsAgreement
'''       Dim mAgreementNo As Variant
'''       objAgreement.SetAgreements AgreementID
'''       mAgreementNo = Split(objAgreement.AgreementNo, "/")
'''       txtAgreementNoPart1.Text = mAgreementNo(0)
'''       txtAgreementNoPart2.Text = mAgreementNo(1)
'''       txtAgreementNoPart1.Tag = objAgreement.AgreementID
'''       txtAgreementDate.Text = objAgreement.AgreementDate
'''       txtProjectNo.Tag = objAgreement.ProjectID
'''       txtProjectNAme.Text = objAgreement.ProjectName
'''       txtProjectNo.Text = objAgreement.ProjectSlNo
'''       txtDueDateofCommencement.Text = objAgreement.DueDateToStart
'''       txtComencementDate.Text = objAgreement.ActualStartedDate
'''       txtduedateofCompletion.Text = objAgreement.DueDateOfCompletion
'''       txtCompletionDate.Text = objAgreement.ActualCompletedDate
'''       txtWorkDate.Text = objAgreement.WorkDate
'''       txtWorkTitle.Text = objAgreement.WorkTitle
'''       txtPAC.Text = objAgreement.PAC
'''       txtContractors.Text = objAgreement.SubLedger
'''       txtContractors.Tag = objAgreement.SubLedgerID
'''    End Sub


    Private Sub cmdApprove_Click()
        Dim mCnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL    As String
        
        mSQL = "Update faAgreements set tnyStatus=1, numApproverID= " & gbUserID & ",dtApprovedDate='" & DdMmmYy(gbTransactionDate) & "' where intAgreementID=" & txtAgreementNoPart1.Text & "  "
        objDB.ExecuteSP mSQL, , , , mCnn, adCmdText
        cmdApprove.Enabled = False
    End Sub

'''    Private Sub cmdAgreementNo_Click()
'''        Dim mAgreementID As Integer
'''        frmSearchAgreements.Show vbModal
'''        mAgreementID = gbSearchID
'''        Call FillAgreements(mAgreementID)
'''    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
        frmListOfAgreements.FillGrid
        frmListOfAgreements.Show
    End Sub
'''    Private Sub cmdNew_Click()
'''        cmdSave.Enabled = True
'''        Call FormInitialize
'''    End Sub
    Private Sub cmdProject_Click()
        frmSulekhaIntegration.Show vbModal
        txtProjectNo.Tag = gbProject.decProjectID
        txtProjectNo.Text = gbProject.chvProjectSlNo
        txtProjectNAme.Text = gbProject.chvProjectName
    End Sub
    Private Sub cmdSave_Click()
        Dim objDB       As New clsDB
        Dim ObjSubLed   As New clsSubLedger
        Dim Rec         As New ADODB.Recordset
        Dim mCnn        As New ADODB.Connection
        Dim mSQL        As String
        Dim arrInput    As Variant
        Dim arrOutPut   As Variant
        Dim mAggmtID    As Variant
        Dim mAgreeID    As Variant
        Dim mSubLedgerTypeID    As Integer
        Dim mAgreementID   As Integer
        Dim mAgreementNo    As Variant
        If txtAgreementNoPart1.Text = "" Then
            mAgreementID = -1
            mAgreementNo = Null
        Else
            mAgreementID = val(txtAgreementNoPart1.Tag)
            mAgreementNo = val(txtAgreementNoPart1.Text) & "/" & txtAgreementNoPart2.Text
        End If
        If txtAgreementDate.Text = "" Then
            MsgBox "Please Enter Agreement Date"
            txtAgreementDate.SetFocus
            Exit Sub
        End If
        If txtProjectNo.Text = "" Then
             MsgBox "Please Select Project"
             txtProjectNo.SetFocus
             Exit Sub
        End If
        If CDate(txtComencementDate.Text) < CDate(txtAgreementDate.Text) Then
             MsgBox "Commencement Date is less than Agreement Date!!!"
             txtComencementDate.SetFocus
             'Exit Sub
        End If
        If CDate(txtDueDateofCommencement.Text) < CDate(txtAgreementDate.Text) Then
             MsgBox "Due Date for Commencement is less than Agreement Date!!"
             txtDueDateofCommencement.SetFocus
             'Exit Sub
        End If
'''        If txtCompletionDate.Text = "" Then
'''             MsgBox "Please Select Actual date of Completion"
'''             txtCompletionDate.SetFocus
'''             Exit Sub
'''        End If
        If txtduedateofCompletion.Text = "" Then
             MsgBox "Please Enter Due date of Commencement"
             txtduedateofCompletion.SetFocus
             Exit Sub
        End If
        If CDate(txtWorkDate.Text) < CDate(txtAgreementDate.Text) Then
             MsgBox "Work Date is less than Agreement Date!!"
             txtWorkDate.SetFocus
             'Exit Sub
        End If
        If txtOrderNo.Text = "" Then
             MsgBox "Please Enter Order No"
             txtOrderNo.SetFocus
             Exit Sub
        End If
        If txtWorkTitle.Text = "" Then
             MsgBox "Please Enter Work Title"
             txtWorkTitle.SetFocus
             Exit Sub
        End If
'''        If txtWorkDate.Text = "" Then
'''             MsgBox "Please Enter Work Title"
'''             txtWorkDate.SetFocus
'''             Exit Sub
'''        End If
'''        If txtAsset.Text = "" Then
'''             MsgBox "Please Select Asset "
'''             txtAsset.SetFocus
'''             Exit Sub
'''        End If
        If txtPAC.Text = "" Or val(txtPAC.Text) = 0 Then
             MsgBox "Please Enter Agreed PAC in Rs."
             txtPAC.SetFocus
             Exit Sub
        End If
        If txtContractors.Text = "" Then
             MsgBox "Please Select Suppliers/ Contractors/ Beneficiary Committee Convener"
             txtContractors.SetFocus
             Exit Sub
        Else
             ObjSubLed.SetSubLedgerDetails (val(txtContractors.Tag))
             mSubLedgerTypeID = ObjSubLed.SubLedgerTypeID
        End If
        
        
        arrInput = Array(mAgreementID, mAgreementNo, _
                        CDate(txtAgreementDate.Text), _
                        val(txtProjectNo.Tag), _
                        CDate(txtDueDateofCommencement.Text), _
                        CDate(txtComencementDate.Text), _
                        CDate(txtduedateofCompletion.Text), _
                        Null, _
                        val(txtOrderNo.Tag), _
                        val(txtOrderNo.Text), _
                        CDate(txtWorkDate.Text), _
                        txtWorkTitle.Text, _
                        CDate(txtWorkDate.Text), _
                        txtAsset.Tag, _
                        cmdSearch.Tag, _
                        txtAsset.Text, _
                        val(txtPAC.Text), _
                        mSubLedgerTypeID, _
                        val(txtContractors.Tag), _
                        gbLocalBodyID, _
                        gbFinancialYearID, _
                        gbUserID, _
                        Null, _
                        Null, _
                        0)
        objDB.ExecuteSP "spSaveAgreement", arrInput, arrOutPut, , mCnn, adCmdStoredProc
        mAgreeID = Split(arrOutPut(0, 0), "/")
        txtAgreementNoPart1.Text = mAgreeID(0)
        txtAgreementNoPart2.Text = mAgreeID(1)
        MsgBox "Agreement Saved SuccessFully", vbApplicationModal
        cmdSave.Enabled = False
    End Sub

    Private Sub cmdsearch_Click()
        frmAssets.Show vbModal
        If gbSearchID <> -1 Then
            txtAsset.Text = gbSearchStr
            txtAsset.Tag = gbSearchID
            cmdSearch.Tag = gbSearchCode
            gbSearchID = -1
            gbSearchStr = ""
            gbSearchCode = ""
        End If
    End Sub

    Private Sub cmdSearchBeneficiary_Click()
        Dim ObjSubLed   As New clsSubLedger
        Dim mSubLedgerTypeID    As Integer
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtContractors.Tag = gbSearchID
        txtContractors.Text = CStr(gbSearchCode) + CStr(gbSearchStr)
'''        ObjSubLed.SetSubLedgerDetails (val(txtContractors.Tag))
'''        mSubLedgerTypeID = ObjSubLed.SubLedgerTypeID
        txtContractors.SetFocus
    End Sub
    Private Sub Form_Load()
        Call FormInitialize
    End Sub
    Private Sub txtAgreementDate_LostFocus()
        If Trim(txtAgreementDate) <> "" Then
           
            txtAgreementDate = CheckDateInMMM(txtAgreementDate)
        Else
            txtAgreementDate.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub



    Private Sub txtComencementDate_LostFocus()
        If Trim(txtComencementDate) <> "" Then
            txtComencementDate.Text = CheckDateInMMM(txtComencementDate.Text)
        Else
            txtComencementDate.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub
    Private Sub txtCompletionDate_LostFocus()
        If Trim(txtCompletionDate) <> "" Then
            txtCompletionDate.Text = CheckDateInMMM(txtCompletionDate.Text)
        Else
            txtCompletionDate.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub
    Private Sub txtDueDateofCommencement_LostFocus()
        If Trim(txtDueDateofCommencement) <> "" Then
            txtDueDateofCommencement.Text = CheckDateInMMM(txtDueDateofCommencement.Text)
        Else
            txtDueDateofCommencement.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub
    Private Sub txtduedateofCompletion_LostFocus()
        If Trim(txtduedateofCompletion) <> "" Then
            txtduedateofCompletion.Text = CheckDateInMMM(txtduedateofCompletion)
        Else
            txtduedateofCompletion.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub
    Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
                    KeyAscii = 0
        End If
    End Sub

    Private Sub txtPAC_KeyPress(KeyAscii As Integer)
         If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
    End Sub
    Private Sub txtPAC_LostFocus()
        txtPAC.Text = Format(val(txtPAC.Text), "0.00")
    End Sub
    Private Sub txtWorkDate_LostFocus()
        If Trim(txtWorkDate) <> "" Then
            txtWorkDate.Text = CheckDateInMMM(txtWorkDate)
        Else
            txtWorkDate.Text = DdMmmYy(gbTransactionDate)
        End If
    End Sub

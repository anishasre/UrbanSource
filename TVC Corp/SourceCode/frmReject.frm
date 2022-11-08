VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmReject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ~ Reject ~"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStage 
      Height          =   285
      Left            =   1095
      TabIndex        =   11
      Top             =   615
      Width           =   3015
   End
   Begin VB.ComboBox cmbReason 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   930
      Width           =   3300
   End
   Begin VB.TextBox txtNote 
      Height          =   765
      Left            =   1080
      TabIndex        =   8
      Top             =   1290
      Width           =   3030
   End
   Begin VB.TextBox txtFwdSeat 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   2100
      Width           =   3015
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "REJECT"
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
      Left            =   1755
      TabIndex        =   5
      Top             =   2490
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5310
      TabIndex        =   4
      Top             =   2430
      Width           =   5370
      Begin WinXPC_Engine.WindowsXPC XPC 
         Left            =   4470
         Top             =   405
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
   End
   Begin VB.CommandButton cmdFwrdSeat 
      Caption         =   "..."
      Height          =   285
      Left            =   4110
      TabIndex        =   3
      Top             =   2100
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "Stage:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   555
      TabIndex        =   12
      Top             =   660
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Reason:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   405
      TabIndex        =   10
      Top             =   975
      Width           =   660
   End
   Begin VB.Label lblRequestType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   630
      TabIndex        =   6
      Top             =   45
      Width           =   3900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fwd Seat:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   270
      TabIndex        =   2
      Top             =   2145
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   645
      TabIndex        =   1
      Top             =   1260
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Approval:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   -240
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmReject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim numRequestTypeID As Variant
    Dim mRequestTypeID As Variant
    Dim mNewMode As Integer
    Private Sub cmdFwrdSeat_Click()
        frmSearchSeat.Show vbModal
        If gbSearchID <> -1 Then
            txtFwdSeat.Text = gbSearchStr
            txtFwdSeat.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
    End Sub
    Private Function SaveValidation() As Boolean
        If Trim(cmbReason.Text) = "" Then
            MsgBox "Data Invalid", vbInformation, "Saankhya"
            cmbReason.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtNote.Text) = "" Then
            MsgBox "Data Invalid", vbInformation, "Saankhya"
            txtNote.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtFwdSeat.Text) = "" Then
            MsgBox "Data Invalid", vbInformation, "Saankhya"
            txtFwdSeat.SetFocus
            SaveValidation = False
            Exit Function
        End If
        SaveValidation = True
    End Function
    Private Sub cmdReject_Click()
    
        Dim objDB As New clsDB
        Dim mcnn As New ADODB.Connection
        Dim mArrIn  As Variant
        Dim mintID As Variant
        Dim mReasonID As Variant
 
        If SaveValidation = False Then Exit Sub
        If objDB.SetConnection(mcnn) Then
            mintID = IIf(txtNote.Tag = "", -1, val(txtNote.Tag))
            mReasonID = cmbReason.ItemData(cmbReason.ListIndex)
            mArrIn = Array(mintID, gbDate, _
                               mRequestTypeID, _
                               numRequestTypeID, _
                               mReasonID, _
                               Trim(txtStage.Text), _
                               Trim(txtNote.Text), _
                               gbUserID, _
                               gbSeatID, _
                               Trim(txtFwdSeat.Tag), _
                               0 _
                            )
            objDB.ExecuteSP "spSaveRejections", mArrIn, , , mcnn, adCmdStoredProc
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
        Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
        End If
        Call FormInitialize
        Me.Hide
    End Sub
    Private Sub Form_Load()
        XPC.InitSubClassing
        'Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons order by vchReason", , , True, True)
        Call FormInitialize
        If mNewMode = 1 Then
            lblRequestType.Caption = " DEMAND"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 2 Then
            lblRequestType.Caption = " RECEIPT CANCELLATION"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 3 Then
            lblRequestType.Caption = " INTERRUPTED RECEIPT MODE"
            Call RequestTypeFn
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
        ElseIf mNewMode = 4 Then
            lblRequestType.Caption = " INTERRUPTED RECEIPT CANCELLATION"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 5 Then
            lblRequestType.Caption = " INTERRUPTED RECEIPT EDIT"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 6 Then
            lblRequestType.Caption = " PAYMENT ORDER"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason ", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 7 Then
            lblRequestType.Caption = " CONTRA ENTRY"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=20 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 8 Then
            lblRequestType.Caption = " LETTER OF ALLOTMENT/AUTHORITY"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 9 Then
            lblRequestType.Caption = " REQUISITION"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=10 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 10 Then
            lblRequestType.Caption = " REVERSE ENTRY"   'DISCUSSION
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=55 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 11 Then
            lblRequestType.Caption = " REQUEST FOR PREVIOUS YEAR'S TRANSACTION"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intType=30 order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 12 Then
            lblRequestType.Caption = " APPROPRIATION CONTROL REGISTER-PDE"
            Call PopulateList(cmbReason, "SELECT vchReason,intReasonID FROM faReasons where intReasonID  like '1%' order by vchReason", , , True, True)
            Call RequestTypeFn
        ElseIf mNewMode = 13 Then
            lblRequestType.Caption = " PAYMENT ORDER CANCELLATION"  'DISCUSSION
            Call PopulateList(cmbReason, "SELECT * FROM faReasons where intType=10 order by vchReason ", , , True, True)
            Call RequestTypeFn
        End If
    End Sub
    Private Sub FormInitialize()
       txtNote.Text = ""
       txtFwdSeat.Text = ""
       lblRequestType.Caption = ""
       txtStage.Text = ""
       cmbReason.ListIndex = -1
    End Sub
    Public Property Let RequestTypeID(mdata As Variant)
        numRequestTypeID = mdata
    End Property
    Public Property Get RequestTypeID() As Variant
        RequestTypeID = numRequestTypeID
    End Property
    Public Property Let Mode(mMode As Variant)
        mNewMode = mMode
    End Property
    Public Property Get Mode() As Variant
        Mode = mNewMode
    End Property
    Private Sub RequestTypeFn()
        Dim mcnn  As New ADODB.Connection
        Dim objDB As New clsDB
        Dim mSQL  As String
        Dim Rec   As ADODB.Recordset
        
        objDB.CreateNewConnection mcnn, enuSourceString.Saankhya
        mSQL = "SELECT intRequestTypeID, vchRequestType From faRequestType WHERE tnyGroupID = " & mNewMode & " "
        Set Rec = objDB.ExecuteSP(mSQL, , , , mcnn, adCmdText)
        If Not (Rec.EOF Or Rec.BOF) Then
            mRequestTypeID = Rec!intRequestTypeID
        End If
       Rec.Close
       
    End Sub

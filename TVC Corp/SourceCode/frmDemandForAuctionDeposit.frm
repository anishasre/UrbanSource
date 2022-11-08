VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmDemandForAuctionDeposit 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Generation For Auction Deposit"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBidderList 
      Caption         =   "&Bidder List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6300
      TabIndex        =   21
      Top             =   6120
      Width           =   1260
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   11250
      Top             =   6420
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.ListBox lstMasters 
      Height          =   2310
      Left            =   -2610
      TabIndex        =   22
      Top             =   6345
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2805
      Left            =   30
      TabIndex        =   33
      Top             =   3105
      Width           =   11310
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3870
         TabIndex        =   17
         Top             =   1590
         Width           =   3930
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   5715
         TabIndex        =   16
         Top             =   1170
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   635
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   39684
      End
      Begin VB.TextBox txtInstrumetDate 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3855
         TabIndex        =   15
         Top             =   1170
         Width           =   1845
      End
      Begin VB.TextBox txtInstrumetNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3855
         TabIndex        =   14
         Top             =   750
         Width           =   3930
      End
      Begin VB.CommandButton cmdSearchInstrumetType 
         Caption         =   "..."
         Height          =   300
         Left            =   7815
         TabIndex        =   13
         Top             =   345
         Width           =   270
      End
      Begin VB.TextBox txtInstrumentType 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3855
         TabIndex        =   12
         Top             =   345
         Width           =   3930
      End
      Begin VB.TextBox txtDepositAmount 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3870
         TabIndex        =   18
         Top             =   1995
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   225
         Left            =   2745
         TabIndex        =   38
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Date"
         Height          =   225
         Left            =   2400
         TabIndex        =   37
         Top             =   1215
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument No"
         Height          =   225
         Left            =   2550
         TabIndex        =   36
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Type"
         Height          =   225
         Left            =   2430
         TabIndex        =   35
         Top             =   405
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Deposit Amount"
         Height          =   225
         Left            =   1770
         TabIndex        =   34
         Top             =   2055
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bidder Details"
      Height          =   2835
      Left            =   5790
      TabIndex        =   28
      Top             =   210
      Width           =   5565
      Begin VB.TextBox txtBidderName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         TabIndex        =   6
         Top             =   300
         Width           =   3930
      End
      Begin VB.TextBox txtAddress1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         TabIndex        =   7
         Top             =   720
         Width           =   3930
      End
      Begin VB.TextBox txtAddress2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         TabIndex        =   8
         Top             =   1125
         Width           =   3930
      End
      Begin VB.TextBox txtAddress3 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         TabIndex        =   9
         Top             =   1530
         Width           =   3930
      End
      Begin VB.TextBox txtPhone1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   10
         Top             =   2010
         Width           =   2115
      End
      Begin VB.TextBox txtPhone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   11
         Top             =   2370
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bidder Name"
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Left            =   630
         TabIndex        =   31
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   225
         Left            =   780
         TabIndex        =   30
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         Height          =   225
         Left            =   780
         TabIndex        =   29
         Top             =   2445
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Auction Details"
      Height          =   2850
      Left            =   45
      TabIndex        =   23
      Top             =   195
      Width           =   5730
      Begin VB.TextBox txtAuctionType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1275
         TabIndex        =   0
         Top             =   480
         Width           =   3930
      End
      Begin VB.TextBox txtFormNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1275
         TabIndex        =   5
         Top             =   1515
         Width           =   1860
      End
      Begin VB.CommandButton cmdSearchAuctionType 
         Caption         =   "..."
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
         Left            =   5235
         TabIndex        =   1
         Top             =   495
         Width           =   285
      End
      Begin VB.TextBox txtAuctionNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   1275
         TabIndex        =   4
         Top             =   1185
         Width           =   1860
      End
      Begin VB.TextBox txtAuctionTitle 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1275
         TabIndex        =   2
         Top             =   810
         Width           =   3930
      End
      Begin VB.CommandButton cmdSearchAuctionTitle 
         Caption         =   "..."
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
         Left            =   5235
         TabIndex        =   3
         Top             =   825
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Type"
         Height          =   225
         Left            =   165
         TabIndex        =   27
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction No"
         Height          =   225
         Left            =   300
         TabIndex        =   26
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auction Title"
         Height          =   225
         Left            =   225
         TabIndex        =   25
         Top             =   885
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form No"
         Height          =   225
         Left            =   435
         TabIndex        =   24
         Top             =   1575
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5190
      TabIndex        =   20
      Top             =   6120
      Width           =   1080
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4080
      TabIndex        =   19
      Top             =   6120
      Width           =   1080
   End
End
Attribute VB_Name = "frmDemandForAuctionDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FormInitialize()
    '           Personal Information            '
    txtBidderName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddress3.Text = ""
    txtPhone1.Text = ""
    txtPhone2.Text = ""
    '           Bidding Details                 '
    txtAuctionType.Text = ""
    txtAuctionTitle.Text = ""
    txtAuctionNo.Text = ""
    txtDepositAmount.Text = ""
End Sub

Private Sub cmdBidderList_Click()
    If txtFormNo.Text <> "" Then
        frmAuctionBidderList.Visible = True
        frmAuctionBidderList.ZOrder (0)
    Else
        MsgBox "Please Select the Auction Type & Title", vbCritical
    End If
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Are you Sure want to proceed?", vbYesNo) = vbYes Then
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.RecordSet
        Dim aryIn As Variant
        Dim AryOut As Variant
        Dim strQry As String
            objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
        '           To Save the Address                 '
        aryIn = Array(Null, _
                    txtBidderName.Text, _
                    Null, _
                    txtAddress1.Text, _
                    txtAddress2.Text, _
                    txtAddress3.Text, _
                    Null, _
                    Null, _
                    txtPhone1.Text, _
                    txtPhone2.Text, _
                    Null _
                    )
        objDB.ExecuteSP "spSaveAddressBook", aryIn, AryOut, , mCnn, adCmdStoredProc
        '           To Save the Auction Deposit         '
        aryIn = Array(txtAuctionTitle.Tag, _
                    AryOut(0, 0), _
                    txtAuctionNo.Text, _
                    txtDepositAmount.Text, _
                    txtFormNo.Text, _
                    txtInstrumentType.Text, _
                    txtInstrumetNo.Text, _
                    txtInstrumetDate.Text, _
                    txtBankName.Text, _
                    Null, _
                    Null, _
                    Date, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    Null, _
                    0 _
                    )
        objDB.ExecuteSP "spSaveAuctionDeposit", aryIn, , , mCnn, adCmdStoredProc
        MsgBox "Added Successfully", vbInformation
    End If
End Sub

Private Sub cmdSearchAuctionType_Click()
    lstMasters.Left = 5595
    lstMasters.Top = 630
    ListMasterFill (1)
End Sub

Private Sub cmdSearchAuctionTitle_Click()
    lstMasters.Left = 5595
    lstMasters.Top = 990
    ListMasterFill (2)
End Sub

Private Sub ListMasterFill(mListId As Integer)
    Dim mSQL As String
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    lstMasters.Tag = mListId
    lstMasters.Visible = True
    Select Case mListId
        Case 1:
            mSQL = "Select vchCollectionType,intCollectionTypeID from smMSTCollectionTypes Where tnyInAuctionList = 1 Order by intCollectionTypeID"
            Call PopulateList(lstMasters, mSQL, , , , True)
        Case 2:
            mSQL = "Select ( vchAuctionNo +'     '+ vchAuctionTitle) ,intAuctionID from smAuctionTitles Where intAuctionTypeID = " & txtAuctionType.Tag & " Order by vchAuctionNo"
            Call PopulateList(lstMasters, mSQL, , , , True)
        Case 3:
            mSQL = "Select vchInstrumentType, intInstrumentTypeID From faInstrumentTypes"
            Call PopulateList(lstMasters, mSQL, "", , , True, enuSourceString.Saankhya)
    End Select
End Sub
Private Sub cmdSearchInstrumetType_Click()
    lstMasters.Left = 8175
    lstMasters.Top = 3480
    ListMasterFill (3)
End Sub
Private Sub DTPicker1_CloseUp()
        txtInstrumetDate.Text = DdMmmYy(DTPicker1.Value)
    End Sub
    
Private Sub DTPicker1_DropDown()
    If IsDate(txtInstrumetDate) Then
        DTPicker1.Value = txtInstrumetDate.Text
    End If
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitIDESubClassing
End Sub

Private Sub lstMasters_DblClick()
    Dim mSQL As String
    Dim mChaCnt As Integer
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec As New ADODB.RecordSet
        objDB.CreateNewConnection mCnn, enuSourceString.iSaankhyaMasters
    Select Case lstMasters.Tag
        Case 1:
            txtAuctionType.Text = lstMasters.Text
            txtAuctionType.Tag = lstMasters.ItemData(lstMasters.ListIndex)
        Case 2:
            mChaCnt = InStr(lstMasters.Text, " ")
            txtAuctionTitle.Text = Right(lstMasters.Text, mChaCnt)
            txtAuctionNo.Text = Left(lstMasters.Text, mChaCnt)
            txtAuctionTitle.Tag = lstMasters.ItemData(lstMasters.ListIndex)
            mSQL = "Select vchfileNo from smAuctionTitles where vchAuctionNo = '" & txtAuctionNo.Text & "'"
            Rec.Open mSQL, mCnn
            txtFormNo.Text = IIf(IsNull(Rec!vchFileNo), "", Rec!vchFileNo)
            Rec.Close
        Case 3:
            txtInstrumentType.Text = lstMasters.Text
    End Select
    lstMasters.Visible = False
End Sub



VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmKMBRIntegration 
   BackColor       =   &H00FFFFFF&
   Caption         =   "P e r m i t s   F r o m   T o w n   P l a n n i n g"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   2610
      Left            =   210
      TabIndex        =   34
      Top             =   -45
      Width           =   8190
      Begin VB.TextBox txtDoorNoSub 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6855
         TabIndex        =   56
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtDoorNo 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5835
         TabIndex        =   54
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtWard 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5835
         TabIndex        =   52
         Top             =   675
         Width           =   555
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4005
         TabIndex        =   50
         Top             =   1830
         Width           =   930
      End
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   48
         Top             =   2130
         Width           =   3735
      End
      Begin VB.TextBox txtPost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   46
         Top             =   1830
         Width           =   2460
      End
      Begin VB.TextBox txtMainPlace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   44
         Top             =   1530
         Width           =   3735
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   42
         Top             =   1230
         Width           =   3735
      End
      Begin VB.TextBox txtHouseName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   40
         Top             =   930
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   38
         Top             =   615
         Width           =   3735
      End
      Begin VB.TextBox txtDemandNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         TabIndex        =   36
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6765
         TabIndex        =   55
         Top             =   900
         Width           =   75
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Door No"
         Height          =   195
         Left            =   5085
         TabIndex        =   53
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         Height          =   195
         Left            =   5325
         TabIndex        =   51
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin"
         Height          =   195
         Left            =   3720
         TabIndex        =   49
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   195
         Left            =   375
         TabIndex        =   47
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         Height          =   195
         Left            =   825
         TabIndex        =   45
         Top             =   1860
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         Height          =   195
         Left            =   285
         TabIndex        =   43
         Top             =   1545
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
         Height          =   195
         Left            =   735
         TabIndex        =   41
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Name"
         Height          =   195
         Left            =   105
         TabIndex        =   39
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   690
         TabIndex        =   37
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demand No"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   285
         Width           =   1005
      End
   End
   Begin VB.TextBox txtPincode 
      Height          =   285
      Left            =   1035
      TabIndex        =   32
      Top             =   5535
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstWardID 
      Height          =   255
      Left            =   15
      TabIndex        =   25
      Top             =   5535
      Visible         =   0   'False
      Width           =   1005
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   8475
      Top             =   5190
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit"
      Height          =   345
      Left            =   2730
      TabIndex        =   15
      Top             =   5370
      Width           =   1710
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4485
      TabIndex        =   14
      Top             =   5370
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   210
      TabIndex        =   0
      Top             =   2595
      Width           =   8175
      Begin VB.ComboBox cmbScheme 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtWardNo 
         Height          =   255
         Left            =   5340
         TabIndex        =   33
         Top             =   645
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.ComboBox cmbDistrict 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2100
         Width           =   2550
      End
      Begin VB.ComboBox cmbPostOffice 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5145
         TabIndex        =   27
         Text            =   "cmbPostOffice"
         Top             =   2100
         Width           =   2550
      End
      Begin VB.ComboBox cmbWard 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   615
         Width           =   1815
      End
      Begin VB.ComboBox cmbBuidingType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkOwnership 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ownership"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4140
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkLandTaxReceipt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Latest Land Tax Receipt"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5685
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkLocationPlan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Location Plan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   885
         TabIndex        =   8
         Top             =   1695
         Width           =   1515
      End
      Begin VB.CheckBox chkBuildingPlan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Building Plan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5685
         TabIndex        =   7
         Top             =   1695
         Width           =   1515
      End
      Begin VB.CheckBox chkSitePlan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Site Plan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4140
         TabIndex        =   6
         Top             =   1695
         Width           =   1125
      End
      Begin VB.CheckBox chkStampPaper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "For Undertaking in Stamp Paper"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   885
         TabIndex        =   5
         Top             =   1335
         Width           =   3105
      End
      Begin VB.OptionButton optGeneralPermit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "For General Permit"
         Height          =   270
         Left            =   870
         TabIndex        =   2
         Top             =   585
         Width           =   2025
      End
      Begin VB.OptionButton optOneDayPermit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "For Oneday Permit"
         Height          =   360
         Left            =   870
         TabIndex        =   1
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scheme"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   1020
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   3990
         TabIndex        =   30
         Top             =   2175
         Width           =   105
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   105
         TabIndex        =   29
         Top             =   2160
         Width           =   105
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post office"
         Height          =   195
         Left            =   4155
         TabIndex        =   28
         Top             =   2145
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2130
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building Ward"
         Height          =   195
         Left            =   3870
         TabIndex        =   24
         Top             =   675
         Width           =   1185
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
         Left            =   3750
         TabIndex        =   23
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   5535
         TabIndex        =   21
         Top             =   1755
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   5535
         TabIndex        =   20
         Top             =   1380
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   3990
         TabIndex        =   19
         Top             =   1725
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   3990
         TabIndex        =   18
         Top             =   1380
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   735
         TabIndex        =   17
         Top             =   1755
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   195
         Left            =   735
         TabIndex        =   16
         Top             =   1395
         Width           =   105
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
         Index           =   4
         Left            =   3735
         TabIndex        =   13
         Top             =   345
         Width           =   90
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building in SqureMeter"
         Height          =   195
         Left            =   3855
         TabIndex        =   12
         Top             =   300
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmKMBRIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
    Dim mLinkFlag As Boolean
    Dim mTransactionType As Integer
    
    Private Sub cmbDistrict_Click()
        Call PopulateList(cmbPostOffice, "SELECT chvPostOfficeEnglish,intPostOfficeID From GM_PostOffice left join GL_PostOffice on left(GM_PostOffice.intPINCode,3)=GL_PostOffice.intPINCode Where tnyDistrictID =" & cmbDistrict.ItemData(cmbDistrict.ListIndex) & "order by  chvPostOfficeEnglish", , , , True, DBMaster)
        cmbPostOffice.ListIndex = 0
    End Sub
    
    Private Sub cmbPostOffice_Click()
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mVarrOut As Variant
        objDB.CreateNewConnection mCnn, enuSourceString.DBMaster
        If cmbPostOffice.ListIndex >= 0 Then
        
            Set Rec = objDB.ExecuteSP("SELECT intPINCode From GM_PostOffice WHERE intPostOfficeID =  " & cmbPostOffice.ItemData(cmbPostOffice.ListIndex), , mVarrOut, , mCnn, adCmdText)
            If IsArray(mVarrOut) Then
               txtPincode.Text = mVarrOut(0, 0)
            End If
        End If
    End Sub
    
    Private Sub cmbWard_Click()
        If cmbWard.ListIndex > -1 Then
                txtWardNo.Text = cmbWard.ItemData(cmbWard.ListIndex)
            End If
    End Sub
    
    Private Sub cmdCancel_Click()
        'frmReceiptsCounter.txtTransactionType.Text = ""
        frmReceiptsCounter.PermitType = -1
        frmReceiptsCounter.BuildingType = -1
        frmReceiptsCounter.KMBRAccess = -1
        Unload Me
    End Sub
    
    Private Sub cmdSubmit_Click()
        If cmbScheme.ItemData(cmbScheme.ListIndex) = 2 Then
            
            Exit Sub
        End If
        If Validation = True Then
            If optGeneralPermit.value = True Then
                frmReceiptsCounter.PermitType = 0
            ElseIf optOneDayPermit.value = True Then
                frmReceiptsCounter.PermitType = 1
            End If
            If mLinkFlag Then
                frmReceiptsCounter.chkLinkDemand.value = 1
                frmReceiptsCounter.txtDemandPrefix.Text = Trim(txtDemandNo)
            Else
                If frmReceiptsCounter.chkLinkDemand.value = 1 Then
                    frmReceiptsCounter.chkLinkDemand.value = 0
                End If
            End If
            frmReceiptsCounter.BuildingType = cmbBuidingType.ItemData(cmbBuidingType.ListIndex)
            frmReceiptsCounter.KMBRAccess = 1
            'frmReceiptsCounter.BuildingWard = GetSelectedTextFromComboListBox(lstWardID, cmbWard)
            frmReceiptsCounter.BuildingWard = cmbWard.ItemData(cmbWard.ListIndex)
            frmReceiptsCounter.txtMainPlace.Text = cmbDistrict.Text
            frmReceiptsCounter.txtMainPlace.Tag = cmbDistrict.ItemData(cmbDistrict.ListIndex)
            frmReceiptsCounter.txtPost.Text = cmbPostOffice.Text
            frmReceiptsCounter.txtPost.Tag = cmbPostOffice.ItemData(cmbPostOffice.ListIndex)
            frmReceiptsCounter.txtPin.Text = Trim(txtPincode.Text)
            'frmReceiptsCounter.txtWardNo.Text = cmbWard.itemData(cmbWard.ListIndex)
            
            frmReceiptsCounter.cmbZone.Text = gbLocation
            frmReceiptsCounter.cmbDZone.Text = gbLocation '   Added   '
            frmReceiptsCounter.txtWardNo.Text = txtWard.Text 'cmbWard.ItemData(cmbWard.ListIndex) '   Added   '
            If cmbWard.ListIndex > -1 Then
                frmReceiptsCounter.txtWard.Text = cmbWard.Text
                frmReceiptsCounter.txtWard.Tag = cmbWard.ItemData(cmbWard.ListIndex)
            End If
            
            frmReceiptsCounter.txtPin.Text = txtPin.Text
            frmReceiptsCounter.txtPhone.Text = txtPhoneNo.Text
            frmReceiptsCounter.txtPost.Text = txtPost.Text
            frmReceiptsCounter.txtLocalPlace.Text = txtPlace.Text
            frmReceiptsCounter.txtMainPlace.Text = txtMainPlace.Text
            frmReceiptsCounter.txtHouse.Text = txtHouseName.Text
            frmReceiptsCounter.txtHouseNo2.Text = ""
            frmReceiptsCounter.txtDoorNo1.Text = txtDoorNo.Text   '   Added   '
            frmReceiptsCounter.txtDoorNo2.Text = txtDoorNoSub.Text   '   Added   '
            frmReceiptsCounter.txtName.Text = txtName.Text
            
            If cmbDistrict.ListIndex > 0 Then
                frmReceiptsCounter.txtMainPlace.Tag = cmbDistrict.ItemData(cmbDistrict.ListIndex)
            End If
            
            Unload Me
        End If
    End Sub
    
    Private Function Validation() As Boolean
        If optGeneralPermit.value = False And optOneDayPermit.value = False Then
            MsgBox "Please Select Either One Day / General", vbInformation
            Validation = False
            optOneDayPermit.SetFocus
            Exit Function
        End If
        If cmbBuidingType.ListIndex = -1 Then
            MsgBox "Please Give the Building Type", vbInformation
            Validation = False
            cmbBuidingType.SetFocus
            Exit Function
        End If
        If chkBuildingPlan.value <> 1 Then
            MsgBox "Please Check Building Plan Details is with the Party", vbInformation
            Validation = False
            chkBuildingPlan.SetFocus
            Exit Function
        End If
        If chkLandTaxReceipt.value <> 1 Then
            MsgBox "Please Check Latest Land Tax Receipt is with the Party", vbInformation
            Validation = False
            chkLandTaxReceipt.SetFocus
            Exit Function
        End If
        If chkLocationPlan.value <> 1 Then
            MsgBox "Please Check Location Plan is with the Party", vbInformation
            Validation = False
            chkLocationPlan.SetFocus
            Exit Function
        End If
        If chkOwnership.value <> 1 Then
            MsgBox "Please Check Ownership Details is with the Party", vbInformation
            Validation = False
            chkOwnership.SetFocus
            Exit Function
        End If
        If chkSitePlan.value <> 1 Then
            MsgBox "Please Check Site Plan is with the Party", vbInformation
            Validation = False
            chkSitePlan.SetFocus
            Exit Function
        End If
        
        If optOneDayPermit.value = True Then
            If chkStampPaper.value <> 1 Then
                MsgBox "Please Check Site Plan is with the Party", vbInformation
                Validation = False
                chkStampPaper.SetFocus
                Exit Function
            End If
        End If
        
        If cmbWard.ListIndex = -1 Then
            MsgBox "Please Select the Building Ward", vbInformation
            Validation = False
            cmbWard.SetFocus
            Exit Function
        End If
        If cmbDistrict.ListIndex = -1 Then
            MsgBox "Please Select District", vbInformation
            Validation = False
            cmbDistrict.SetFocus
            Exit Function
        End If
        If cmbPostOffice.ListIndex = -1 Then
            MsgBox "Please Select Post Office", vbInformation
            Validation = False
            cmbPostOffice.SetFocus
            Exit Function
        End If
        
        Validation = True
    End Function
    
    Private Sub Form_Load()
        WindowsXPC1.InitIDESubClassing
        Label1.Visible = True
        mLinkFlag = False
        Call PopulateList(cmbBuidingType, "SELECT chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=0", , , , True, enuSourceString.KMBR)
        Call PopulateList(cmbDistrict, "Select chvDistrictEnglish, tnyDistrictID From GM_District Order By chvDistrictEnglish", , , , True, DBMaster)
        'Call PopulateList(cmbScheme, "Select vchScheme, intSchemeID FROM TB_Scheme_MST ORDER BY vchScheme", , True, True, True, enuSourceString.KMBR)
        Call PopulateList(cmbScheme, "Select vchScheme, intSchemeID FROM TB_Scheme_MST ORDER BY vchScheme", "Regular ", , , True, enuSourceString.KMBR)
'        cmbScheme.Text = "Regular "
 
        Call FillWard
        Call FormInitialize
    End Sub
    Private Sub FormInitialize()
        Dim mCnn    As New ADODB.Connection
        Dim mRec    As New ADODB.Recordset
        Dim objDB   As New clsDB
        
        chkStampPaper.value = vbChecked
        chkOwnership.value = vbChecked
        chkLandTaxReceipt.value = vbChecked
        chkLocationPlan.value = vbChecked
        chkSitePlan.value = vbChecked
        chkBuildingPlan.value = vbChecked
        If objDB.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
            Set mRec = objDB.ExecuteSP("SELECT chvDistrictEnglish From GM_District WHERE tnyDistrictID =  " & gbDistID, , , , mCnn, adCmdText)
            If Not (mRec.EOF And mRec.BOF) Then
                cmbDistrict.Text = IIf(IsNull(mRec!chvDistrictEnglish), "", mRec!chvDistrictEnglish)
            End If
        End If
'        cmbDistrict.Text = gbDistrict
        optGeneralPermit.value = True
        cmbWard.ListIndex = 0
    End Sub
   
    Private Sub optGeneralPermit_Click()
        Call PopulateList(cmbBuidingType, "SELECT chvBuildingType,intFee FROM Fee_LM ", , , , True, enuSourceString.KMBR)
        cmbBuidingType.ListIndex = 0
        Label1.Visible = False
    End Sub
    
    Private Sub optOneDayPermit_Click()
        Call PopulateList(cmbBuidingType, "SELECT  chvBuildingType,intFee FROM Fee_LM WHERE intPermitType=1", , , , True, enuSourceString.KMBR)
        cmbBuidingType.ListIndex = 0
        Label1.Visible = True
    End Sub
    
    Private Sub FillWard()
        Dim Rec As New ADODB.Recordset
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim objDB As New clsDB
        Dim aryIn As Variant
        If objDB.CreateNewConnection(mCnn, enuSourceString.KMBR) Then
            'aryIn = Array(gbLocalBodyID, 2005)
            Set Rec = objDB.ExecuteSP("spWardMapping", , , , mCnn, adCmdStoredProc)
            While Not (Rec.EOF Or Rec.BOF)
                cmbWard.AddItem IIf(IsNull(Rec!chvWardNameEnglish), "", Rec!chvWardNameEnglish)
                cmbWard.ItemData(cmbWard.NewIndex) = IIf(IsNull(Rec!Inward), "", Rec!Inward)
                Rec.MoveNext
            Wend
           ' cmbWard.ItemData(cmbWard.ListIndex) = 1
        Else
            MsgBox "Connecton to KMBR DataBase cannot be Established, Contact your System Administrator", vbInformation
        End If
        
        
    '''
    '''    If objDb.CreateNewConnection(mCnn, enuSourceString.DBMaster) = True Then
    '''        mSql = "Select chvWardNameEnglish, numWardID, intWardNo From GM_Ward Where tnyWardType = 2 and intLBID = " & gbLocalBodyID & " and chvWardNameEnglish is not Null Order By numWardID"
    '''        Rec.Open mSql, mCnn
    '''        While Not (Rec.EOF Or Rec.BOF)
    '''            cmbWard.AddItem IIf(IsNull(Rec!chvWardNameEnglish), "", Rec!chvWardNameEnglish)
    '''            cmbWard.itemData(cmbWard.NewIndex) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
    '''            lstWardID.AddItem IIf(IsNull(Rec!numWardID), "", Rec!numWardID)
    '''            lstWardID.itemData(lstWardID.NewIndex) = IIf(IsNull(Rec!intWardNo), "", Rec!intWardNo)
    '''            Rec.MoveNext
    '''        Wend
    '''    Else
    '''        MsgBox "Connecton to DB_Masters cannot be Established, Contact your System Administrator", vbInformation
    '''        Exit Sub
    '''    End If
    
    '''    Call PopulateList(cmbWard, "SELECT  chvWardNameEnglish,intWardNo  FROM GM_Ward WHERE intLBID = " & gbLocalBodyID & " AND intWardYear = 2000 and chvWardNameEnglish is not null", , , , True, DBMaster)
    
        
    End Sub
    
     Private Function GetSelectedTextFromComboListBox(lstBox As ListBox, cmbBox As ComboBox) As Variant
            Dim mIteration As Integer
            If cmbBox.ListIndex = -1 Then
                GetSelectedTextFromComboListBox = ""
                Exit Function
            End If
            For mIteration = 0 To lstBox.ListCount - 1
                If cmbWard.ItemData(cmbBox.ListIndex) = lstBox.ItemData(mIteration) Then
                    lstBox.ListIndex = mIteration
                    GetSelectedTextFromComboListBox = lstBox.Text
                    Exit Function
                End If
            Next
            GetSelectedTextFromComboListBox = Nothing
        End Function
    
    Private Sub txtDemandNo_LostFocus()
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim objDB       As New clsDB
        Dim mSql        As String
        Dim mArrIn      As Variant
        Call LockControll(False)
        If objDB.CreateNewConnection(mCnn, enuSourceString.KMBR) = False Then
            MsgBox "The Connection to KMBR not present", vbInformation
            Exit Sub
        End If
        mArrIn = Array(Trim(txtDemandNo.Text))
        Set Rec = objDB.ExecuteSP("ToFrontOffice", mArrIn, , , mCnn)
        If Not (Rec.EOF And Rec.BOF) Then
            txtName.Text = Rec!Name
            txtWard.Text = Rec!Ward
            txtDoorNo.Text = Rec!HouseNo
            txtDoorNoSub.Text = Rec!DoorNo
            cmbDistrict.Text = GetTextOfCombo(cmbDistrict, Rec!DistId)
            txtHouseName.Text = Rec!HouseName
            txtPlace.Text = Rec!Street
            txtMainPlace.Text = Rec!MainPlace
            txtPost.Text = GetTextOfCombo(cmbPostOffice, Rec!PostId)
            txtPost.Tag = Rec!PostId
            txtPin.Text = Rec!PinCode
            txtPhoneNo.Text = Rec!Phone
            mLinkFlag = True
            Call LockControll(True)
        End If
    End Sub
    
    Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then Call PressTabKey
    End Sub

    Private Sub txtPin_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then Call PressTabKey
    End Sub



    Private Sub txtWard_KeyPress(KeyAscii As Integer)
        If Not (KeyAscii <= Asc("9") Or KeyAscii <= Asc("0")) Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then Call PressTabKey
    End Sub
    Private Sub txtWardNo_Change()
        Dim mCount As Integer
            cmbWard.ListIndex = -1
            For mCount = 0 To cmbWard.ListCount - 1
                If val(txtWardNo.Text) = cmbWard.ItemData(mCount) Then
                    cmbWard.ListIndex = mCount
                    Exit For
                End If
            Next
    End Sub
    
    Private Function GetTextOfCombo(Cmb As ComboBox, mSearchID As Long) As String
        Dim mCount As Long
        For mCount = 0 To Cmb.ListCount - 1
            If Cmb.ItemData(mCount) = mSearchID Then
                GetTextOfCombo = Trim(Cmb.List(mCount))
                Exit For
            End If
        Next mCount
    End Function

    Private Sub LockControll(LockOrUnLock As Boolean)
        txtName.Locked = LockOrUnLock
        txtHouseName.Locked = LockOrUnLock
        txtPlace.Locked = LockOrUnLock
        txtMainPlace.Locked = LockOrUnLock
        txtPost.Locked = LockOrUnLock
        txtPhoneNo.Locked = LockOrUnLock
        txtPin.Locked = LockOrUnLock
        txtWard.Locked = LockOrUnLock
        txtDoorNo.Locked = LockOrUnLock
        'cmbDistrict.Locked = LockOrUnLock
        'cmbPostOffice.Locked = LockOrUnLock
    End Sub


    Private Sub SaveSoochi()
'        If mTransactionType = gbTransactionTypeApplicationForPermitKMBR Then
'            If mKMBRFlag = True Then
'                Dim mCnnKMBR As New ADODB.Connection
'                If objDB.CreateNewConnection(mCnnKMBR, enuSourceString.KMBR) = True Then
'                    mCnnKMBR.BeginTrans
'                    If SaveSanketham(lSoochikaCurrentNo, mCnnKMBR) = True Then
'
'                        mCnnSoochika.CommitTrans
'                        mCnnKMBR.CommitTrans
'
'                    Else
'                        GoTo ErrorRollBack:
'                    End If
'                End If
'            End If
'            End If
'        End If
    End Sub

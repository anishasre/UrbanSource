VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSoochikaMiscellaneous 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Miscellaneous Reports"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chktappal 
      BackColor       =   &H80000009&
      Caption         =   "All Tappal"
      Height          =   375
      Left            =   3015
      TabIndex        =   21
      Top             =   585
      Width           =   1050
   End
   Begin VB.CheckBox chkSeatwise 
      BackColor       =   &H80000009&
      Caption         =   "Seat wise"
      Height          =   330
      Left            =   1845
      TabIndex        =   20
      Top             =   585
      Width           =   1050
   End
   Begin VB.CheckBox chkCurrent 
      BackColor       =   &H8000000E&
      Caption         =   "Current No"
      Height          =   330
      Left            =   630
      TabIndex        =   19
      Top             =   585
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   8325
      Begin VB.ComboBox cmbseatid 
         Height          =   315
         Left            =   4155
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   3165
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbseat 
         Height          =   315
         Left            =   2160
         TabIndex        =   28
         Text            =   "Seat"
         Top             =   3165
         Width           =   1905
      End
      Begin VB.TextBox txtyear 
         Height          =   390
         Left            =   2160
         TabIndex        =   24
         Top             =   2685
         Width           =   2025
      End
      Begin VB.TextBox txtInwardNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   15
         Top             =   2160
         Width           =   2010
      End
      Begin VB.CommandButton btnClose 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4140
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   4350
         Left            =   5175
         TabIndex        =   7
         Top             =   495
         Width           =   2865
         Begin VB.OptionButton optDRSeatwise 
            BackColor       =   &H80000005&
            Caption         =   "DistributionRegisterSeatwise"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   26
            Top             =   585
            Width           =   2655
         End
         Begin VB.OptionButton optserviceact 
            BackColor       =   &H80000009&
            Caption         =   "Service Act Register"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   3870
            Width           =   2655
         End
         Begin VB.OptionButton optMalInw 
            BackColor       =   &H80000009&
            Caption         =   "Malayalam Inward"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   22
            Top             =   3465
            Width           =   2655
         End
         Begin VB.OptionButton optSecR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Security Register"
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
            Height          =   330
            Left            =   90
            TabIndex        =   18
            Top             =   3105
            Width           =   2640
         End
         Begin VB.OptionButton optInward 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Inward Register"
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
            Left            =   90
            TabIndex        =   17
            Top             =   1035
            Width           =   2640
         End
         Begin VB.OptionButton optAck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Acknowledgement Slip"
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
            Left            =   90
            TabIndex        =   14
            Top             =   2745
            Width           =   2640
         End
         Begin VB.OptionButton optReceipt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Receipt Details"
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
            Left            =   90
            TabIndex        =   12
            Top             =   2430
            Width           =   2640
         End
         Begin VB.OptionButton optBill 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bill Details"
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
            Left            =   90
            TabIndex        =   11
            Top             =   2070
            Width           =   2640
         End
         Begin VB.OptionButton optPost 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Registered Post"
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
            Left            =   90
            TabIndex        =   10
            Top             =   1755
            Width           =   2640
         End
         Begin VB.OptionButton optRTI 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "RTI Register"
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
            Left            =   90
            TabIndex        =   9
            Top             =   1395
            Width           =   2640
         End
         Begin VB.OptionButton optDR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Distribution Register"
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
            Left            =   90
            TabIndex        =   8
            Top             =   180
            Width           =   2640
         End
      End
      Begin VB.CommandButton btnPrint 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4155
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   345
         Left            =   2190
         TabIndex        =   1
         Top             =   1050
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62586881
         CurrentDate     =   40021
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   345
         Left            =   2190
         TabIndex        =   4
         Top             =   1650
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62586881
         CurrentDate     =   40021
      End
      Begin VB.Label lblseat 
         BackColor       =   &H80000005&
         Caption         =   "Seat"
         Height          =   375
         Left            =   615
         TabIndex        =   27
         Top             =   3210
         Width           =   1215
      End
      Begin VB.Label lblyear 
         BackColor       =   &H80000005&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   615
         TabIndex        =   23
         Top             =   2745
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inward No."
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
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   2265
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   30
         TabIndex        =   6
         Top             =   120
         Width           =   8265
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date To"
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
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date From"
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
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1035
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmSoochikaMiscellaneous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
     Dim vAryInRpt(2)
    Dim vAryInRpt1(3)
    Dim vAryInRpt2
    Dim vAryInRpt3(3)
If optDR.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)

    If chkCurrent.value = vbChecked Then
    'frmCRViewer.vShowReport App.Path & "\soochika\Reports", "rptDistributionRegister_CurrentNo.rpt", vAryInRpt
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptDistributionRegister_CurrentNo.rpt", vAryInRpt
    frmCRViewer.Show 1
    End If
    
    If chkSeatwise.value = vbChecked Then
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptDistributionRegister_Section.rpt", vAryInRpt
    frmCRViewer.Show 1
    End If
    
    If chktappal.value = vbChecked Then
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptDistributionRegisterAll_CurrentNo.rpt", vAryInRpt
    frmCRViewer.Show 1
    End If
    
    ElseIf optDRSeatwise.value = True Then
    'CHANGED
        'If gbLBID = 167 Then
        If (cmbseat.ListIndex < 0) Then
        MsgBox "Select the Seat"
          Exit Sub
        End If
            vAryInRpt3(0) = CStr(dtpFrom.value)
            vAryInRpt3(1) = CStr(dtpTo.value)
            vAryInRpt3(2) = cmbseatid.Text
            'frmCRViewer.vShowReport App.Path & "\soochika\Reports", "rptDistributionRegister_Seat.rpt", vAryInRpt3
            
            'frmCRViewer.Show 1
            frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptDistributionRegister_Seat.rpt", vAryInRpt3
    frmCRViewer.Show 1
       ' End If
    
ElseIf optRTI.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptRightToInformationRegister.rpt", vAryInRpt
    frmCRViewer.Show 1
ElseIf optPost.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptInwardRegisteredPostRegister.rpt", vAryInRpt
    frmCRViewer.Show 1
ElseIf optBill.value = True Then
    vAryInRpt1(0) = CStr(dtpFrom.value)
    vAryInRpt1(1) = CStr(dtpTo.value)
    vAryInRpt1(2) = 1
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptInwardBillReceiptRegister.rpt", vAryInRpt1
    frmCRViewer.Show 1
ElseIf optReceipt.value = True Then
    vAryInRpt1(0) = CStr(dtpFrom.value)
    vAryInRpt1(1) = CStr(dtpTo.value)
    vAryInRpt1(2) = 2
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptInwardBillReceiptRegister.rpt", vAryInRpt1
    frmCRViewer.Show 1
    
ElseIf optInward.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptInwardRegister.rpt", vAryInRpt
    frmCRViewer.Show 1
    'add VP
ElseIf optSecR.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptSecurityRegister.rpt", vAryInRpt
    frmCRViewer.Show 1
     
   'changed --MalInward by soumya V S as on 14.05.14
     
  ElseIf optMalInw.value = True Then
    vAryInRpt(0) = CStr(dtpFrom.value)
    vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptMalayalamInward.rpt", vAryInRpt
    frmCRViewer.Show 1
    
     'Serice Act register Report --changed by soumya vs
    ElseIf optserviceact.value = True Then
     vAryInRpt(0) = CStr(dtpFrom.value)
     vAryInRpt(1) = CStr(dtpTo.value)
    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptServiceActRegister.rpt", vAryInRpt
    frmCRViewer.Show 1
    
ElseIf optAck.value = True Then
        ReDim vAryInRpt2(2)
        Dim mSQl As String
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        If txtInwardNo.Text = "" Then
            MsgBox "Must Enter Inward No."
            Exit Sub
        End If
'        mSql = "SELECT FldFileID FROM TblTappaldetails "
'        mSql = mSql + " WHERE(year(FldDateofReceipt)=year(getdate()) AND (FldCurrentNo =( " & txtInwardNo.Text & ")))"
        mSQl = "SELECT numFileID FROM tInwardDetails "
        'changed by soumya VS oct 23
         mSQl = mSQl + " WHERE(year(dtDateofReceipt)=(" & txtyear.Text & ") AND (numCurrentNo =( " & txtInwardNo.Text & ")))"
        'mSql = mSql + " WHERE(year(dtDateofReceipt)=year(getdate()) AND (numCurrentNo =( " & txtInwardNo.Text & ")))"

        If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Soochika Connection is not present", vbCritical, "Common"
            Exit Sub
        End If
        Rec.Open mSQl, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            vAryInRpt2(0) = CStr(Rec!numFileID)
            vAryInRpt2(1) = 1
        End If
        Rec.Close
'    frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "AckSlip.rpt", vAryInRpt2
frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptAckSlip.rpt", vAryInRpt2
    frmCRViewer.Show 1
Else
    MsgBox "Select any of the Report"
End If
End Sub

Private Sub chkCurrent_Click()
chkSeatwise.value = False
chktappal.value = False
End Sub

Private Sub chkSeatwise_Click()
chktappal.value = False
chkCurrent.value = False
End Sub

Private Sub chktappal_Click()
chkCurrent.value = False
chkSeatwise.value = False
End Sub

Private Sub cmbseat_Change()
cmbseatid.ListIndex = cmbseat.ListIndex
End Sub

Private Sub cmbseat_Click()
cmbseatid.ListIndex = cmbseat.ListIndex
End Sub

Private Sub Form_Load()
    gSubCenterForm Me
    dtpFrom.value = Date
    dtpTo.value = Date
    'CHANGED by soumya vs
     
       txtInwardNo = ""
       txtyear.Text = ""
    Call FillSeats
     
    Call Disable

End Sub
Private Sub FillSeats()
        Dim mSQl As String
         Dim mCnn As New ADODB.Connection
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
    If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        
         End If
         'rec.Open "select chvSeatName,Right(Convert( VarChar(10),numSeatID), 3) as numSeatID From tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 6 and tUserDetails.intUserTypeID <> 99 and tUserDetails.tnyActive=0 order by chvSeatname", mCnn
         'If Not (rec.EOF Or rec.BOF) Then
         'cmbseatid.Text = rec!numSeatID
          'End If
         mSQl = "select chvSeatName,chvSeatName From tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 6 and tUserDetails.intUserTypeID <> 99 and tUserDetails.tnyActive=0 order by chvSeatname"
       Call PopulateList(cmbseat, mSQl, , True, True, True, enuSourceString.SoochikaUnicode)
        mSQl = " select Right(Convert( VarChar(10),numSeatID), 10) , chvSeatName From tSeatDetails left Join tUserDetails on tUserDetails.numUserID=tSeatDetails.numCurrentUserID where tUserDetails.intUserTypeID <> 6 and tUserDetails.intUserTypeID <> 99 and tUserDetails.tnyActive=0 order by chvSeatname"
      Call PopulateList(cmbseatid, mSQl, , True, True, True, enuSourceString.SoochikaUnicode)
        
        
End Sub


Private Sub optAck_Click()
     If optAck.value = True Then
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
        txtInwardNo.Enabled = True
        txtyear.Enabled = True
        txtyear.Text = Year(Now)
         cmbseat.Enabled = False
        
    Else
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
        txtyear.Enabled = False
        Call Disable
    End If
End Sub

Private Sub optBill_Click()
    If optBill.value = True Then
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
         cmbseat.Enabled = False
        Call Disable
    End If
End Sub

Private Sub optDR_Click()
   If optDR.value = True Then
        chkCurrent.Enabled = True
        chkSeatwise.Enabled = True
        chktappal.Enabled = True
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
        txtyear.Enabled = False
        cmbseat.Enabled = False
        
       
    End If
End Sub

Private Sub optDRSeatwise_Click()
'CHNAGED by soumya vs on 12 Nov 14

If optDRSeatwise.value = True Then
        
    Call FillSeats
        chkCurrent.Enabled = False
        chkSeatwise.Enabled = False
        chktappal.Enabled = False
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
        txtyear.Enabled = False
        optDR.value = False
        cmbseat.Enabled = True
        
        
    End If
End Sub

Private Sub optMalInw_Click()
cmbseat.Enabled = False
End Sub

Private Sub optPost_Click()
    If optPost.value = True Then
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
         cmbseat.Enabled = False
        Call Disable
    End If
End Sub

Private Sub optReceipt_Click()
   If optReceipt.value = True Then
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
         cmbseat.Enabled = False
        Call Disable
    End If
End Sub

Private Sub optRTI_Click()
    If optRTI.value = True Then
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
         cmbseat.Enabled = False
        Call Disable
    End If
End Sub
Private Sub optInward_Click()
  If optInward.value = True Then
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
        txtInwardNo.Enabled = False
        cmbseat.Enabled = False
        
        Call Disable
        
    End If
End Sub
Private Sub Disable()
        chkCurrent.Enabled = False
        chkSeatwise.Enabled = False
        chktappal.Enabled = False
End Sub

Private Sub optSecR_Click()
cmbseat.Enabled = False
End Sub

Private Sub optserviceact_Click()
'changed by soumya vs oct23
If optserviceact.value = True Then
dtpFrom.Enabled = True
dtpTo.Enabled = True
txtInwardNo.Enabled = False
   txtyear.Enabled = False
   cmbseat.Enabled = False
     Call Disable
End If
End Sub

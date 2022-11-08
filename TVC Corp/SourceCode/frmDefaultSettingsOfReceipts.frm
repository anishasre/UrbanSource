VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDefaultSettingsOfReceipts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Settings "
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmDefaultSettingsOfReceipts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Default Bank Settings"
      TabPicture(0)   =   "frmDefaultSettingsOfReceipts.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdApply0"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtBank"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDrawnBank"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtBankPlace"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdBank"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdClose0"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Integration "
      TabPicture(1)   =   "frmDefaultSettingsOfReceipts.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "cmdAppy1"
      Tab(1).Control(3)=   "chkRLB"
      Tab(1).Control(4)=   "chkPTax"
      Tab(1).Control(5)=   "chkKMBR"
      Tab(1).Control(6)=   "chkSugama"
      Tab(1).Control(7)=   "chkSevana"
      Tab(1).Control(8)=   "chkDOPFA"
      Tab(1).Control(9)=   "txtUrl"
      Tab(1).Control(10)=   "chkBandD"
      Tab(1).Control(11)=   "txtAllotmentNo"
      Tab(1).Control(12)=   "chkFinHo"
      Tab(1).Control(13)=   "cmdClose1"
      Tab(1).ControlCount=   14
      Begin VB.CommandButton cmdClose1 
         Caption         =   "Close"
         Height          =   345
         Left            =   -71760
         TabIndex        =   24
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdClose0 
         Caption         =   "Close"
         Height          =   345
         Left            =   3000
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkFinHo 
         Alignment       =   1  'Right Justify
         Caption         =   "Zonal Collection"
         Height          =   255
         Left            =   -71400
         TabIndex        =   22
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtAllotmentNo 
         Height          =   285
         Left            =   -73320
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkBandD 
         Alignment       =   1  'Right Justify
         Caption         =   "Birth And Death"
         Height          =   255
         Left            =   -71400
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   -73320
         TabIndex        =   16
         Top             =   2040
         Width           =   4695
      End
      Begin VB.CheckBox chkDOPFA 
         Alignment       =   1  'Right Justify
         Caption         =   "DOPFA"
         Height          =   255
         Left            =   -73080
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkSevana 
         Alignment       =   1  'Right Justify
         Caption         =   "Sevana"
         Height          =   255
         Left            =   -73080
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkSugama 
         Alignment       =   1  'Right Justify
         Caption         =   "Sugama"
         Height          =   255
         Left            =   -73080
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkKMBR 
         Alignment       =   1  'Right Justify
         Caption         =   "KMBR"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkPTax 
         Alignment       =   1  'Right Justify
         Caption         =   "Property Tax"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkRLB 
         Alignment       =   1  'Right Justify
         Caption         =   "Rent On Land"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdAppy1 
         Caption         =   "Apply"
         Height          =   345
         Left            =   -72960
         TabIndex        =   9
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtBankPlace 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtDrawnBank 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtBank 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdApply0 
         Caption         =   "Apply"
         Height          =   345
         Left            =   1800
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Staring Allotment No"
         Height          =   195
         Left            =   -74865
         TabIndex        =   21
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Service Url"
         Height          =   195
         Left            =   -74280
         TabIndex        =   17
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remittance Bank Place"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1740
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default Remittance Bank"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Default Bank"
         Height          =   195
         Left            =   1005
         TabIndex        =   1
         Top             =   720
         Width           =   930
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Default Configurations "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmDefaultSettingsOfReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub cmdApply0_Click()
        Dim mSQL As String
        Dim objDb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        If txtBank.Text = "" Then
            MsgBox "Please Select Default bank", vbApplicationModal
            Exit Sub
        End If
        mSQL = "Update faConfig Set intDefaultBankID=" & txtBank.Tag
        If txtDrawnBank.Text <> "" Then
            mSQL = mSQL + ",vchRemittingBank='" & txtDrawnBank.Text & "'"
        End If
        If txtBankPlace.Text <> "" Then
            mSQL = mSQL + ",vchRemittingPlaceOfBank='" & txtBankPlace.Text & "'"
        End If
        If (objDb.SetConnection(mCnn)) = True Then
            objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
        End If
        MsgBox "Please Log Out After Setting Default Configurations", vbInformation
        
      
    End Sub
    Private Sub cmdAppy1_Click()
        Dim mSQL As String
        Dim objDb   As New clsDB
        Dim mCnn    As New ADODB.Connection
        If (objDb.SetConnection(mCnn)) = True Then
            If txtUrl.Text <> "" Then
                mSQL = "Update faConfig Set vchDefaultUrl='" & txtUrl.Text & "'"
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set vchDefaultUrl='" & txtUrl.Text & "'"
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkRLB.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyRLB=" & chkRLB.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyRLB=" & chkRLB.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkPTax.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithPropertyTax=" & chkPTax.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithPropertyTax=" & chkPTax.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkKMBR.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithKMBR=" & chkKMBR.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithKMBR=" & chkKMBR.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkSugama.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithSugama=" & chkSugama.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithSugama=" & chkSugama.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkSevana.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithSevana=" & chkSevana.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithSevana=" & chkSevana.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkDOPFA.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithDandOPFA=" & chkDOPFA.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithDandOPFA=" & chkDOPFA.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkBandD.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithBandDSchedules=2"
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithBandDSchedules=" & chkBandD.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If Trim(txtAllotmentNo.Text) <> "" Then
                mSQL = "Update faConfig Set vchStartingAllotmentNo='" & Trim(txtAllotmentNo.Text) & "'"
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set vchStartingAllotmentNo='" & Trim(txtAllotmentNo.Text) & "'"
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If
            If chkFinHo.Value = vbChecked Then
                mSQL = "Update faConfig Set tnyLinkWithFinanceHO=" & chkFinHo.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            Else
                mSQL = "Update faConfig Set tnyLinkWithFinanceHO=" & chkFinHo.Value
                objDb.ExecuteSP mSQL, , , , mCnn, adCmdText
            End If

            MsgBox "Please Log Out After Setting Default Configurations", vbInformation
        End If
    End Sub

    Private Sub cmdBank_Click()
          Dim mSQL As String
            mSQL = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID From faAccountHeads Where intGroupID =2 And tinHiddenFlag <> 1 Order by vchAccountHeadCode"
            frmSearchAccountHeads.SQLString = mSQL
            frmSearchAccountHeads.Show vbModal
            txtBank.SetFocus
            If gbSearchID <> -1 Then
                setBank (gbSearchID)
            End If
    End Sub
    Private Sub setBank(mBankID As Integer)
        Dim objAcc  As New clsAccounts
        objAcc.SetAccountID (mBankID)
        txtBank.Text = objAcc.AccountHead
        txtBank.Tag = objAcc.AccountHeadID
    End Sub
''
''    Private Sub cmdSavePropertyTax_Click()
''        If cmbTransactionTypePT.ListIndex < 0 Then
''            MsgBox "Select Transaction Type", vbInformation, "Saankhya"
''            cmbTransactionTypePT.SetFocus
''            Exit Sub
''        End If
''        If cmbAccountHeadPT.ListIndex < 0 Then
''            MsgBox "Select Account Head", vbInformation, "Saankhya"
''            cmbAccountHeadPT.SetFocus
''            Exit Sub
''        End If
''        If cmbInstrumentTypePT.ListIndex < 0 Then
''            MsgBox "Select InstrumentType", vbInformation, "Saankhya"
''            cmbInstrumentTypePT.SetFocus
''            Exit Sub
''        End If
''        If cmbBankPT.ListIndex < 0 Then
''            MsgBox "Select Bank", vbInformation, "Saankhya"
''            cmbBankPT.SetFocus
''            Exit Sub
''        End If
''        If cmbZonePT.ListIndex < 0 Then
''            MsgBox "Select Zone", vbInformation, "Saankhya"
''            cmbZonePT.SetFocus
''            Exit Sub
''        End If
''        If Not cmbAccountHeadPT.ListIndex < 0 Then
''            Dim objAccounts As New clsAccounts
''            objAccounts.SetAccountID (cmbAccountHeadPT.ItemData(cmbAccountHeadPT.ListIndex))
''        End If
''
''        WriteINIfile "Receipt PTax", "DefaultTransactionTypeID", CStr(cmbTransactionTypePT.ItemData(cmbTransactionTypePT.ListIndex)), App.Path & "\Saankhya.INI"
''        WriteINIfile "Receipt PTax", "DefaultAccountHeadCode", CStr(objAccounts.AccountCode), App.Path & "\Saankhya.INI"
''        WriteINIfile "Receipt PTax", "DefaultInstumentID", CStr(cmbInstrumentTypePT.ItemData(cmbInstrumentTypePT.ListIndex)), App.Path & "\Saankhya.INI"
''        WriteINIfile "Receipt PTax", "DefaultBankID", CStr(cmbBankPT.ItemData(cmbBankPT.ListIndex)), App.Path & "\Saankhya.INI"
''        WriteINIfile "Receipt PTax", "DefaultZone", CStr(cmbZonePT.ItemData(cmbZonePT.ListIndex)), App.Path & "\Saankhya.INI"
''        MsgBox "Settings saved Successfully", vbInformation, "Saankhya"
''    End Sub



    Private Sub cmdClear_Click()
        Call FormInitialize
    End Sub
    Private Sub FormInitialize()
        txtBank.Text = ""
        txtBank.Tag = -1
        txtDrawnBank.Text = ""
        txtBankPlace.Text = ""
    End Sub

 


    Private Sub cmdClose0_Click()
        Unload Me
    End Sub

    Private Sub cmdClose1_Click()
        Unload Me
    End Sub

    Private Sub Form_Activate()
        Me.Top = 0
        Me.Width = 6945
        Me.Height = 5010
        Me.Left = (Screen.Width - Me.Width) / 2
    End Sub

    Private Sub Form_Load()
        Dim mSQL As String
        Dim objDb   As New clsDB
        Dim Rec     As New Recordset
        Dim mCnn    As New ADODB.Connection
        If gbSeatGroupID = gbSeatGroupAccountsSupt Then
            SSTab.TabVisible(0) = True
            SSTab.TabVisible(1) = False
            objDb.SetConnection mCnn
            Set Rec = objDb.ExecuteSP("Select * From faConfig", , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                If Not IsNull(Rec!intDefaultBankID) Then
                    Call setBank(CInt(Rec!intDefaultBankID))
                    txtBank.Text = IIf(IsNull(Rec!vchRemittingPlaceOfBank), "", Rec!vchRemittingPlaceOfBank)
                End If
                txtBankPlace.Text = IIf(IsNull(Rec!vchRemittingPlaceOfBank), "", Rec!vchRemittingPlaceOfBank)
                txtDrawnBank.Text = IIf(IsNull(Rec!vchRemittingBank), "", Rec!vchRemittingBank)
            End If
            Call setBank(CInt(gbDefaultBankID))
            txtDrawnBank.Text = gbRemittingBank
            txtBankPlace.Text = gbRemittingPlaceOfBank
        ElseIf gbUserTypeID = 1 Then
            SSTab.TabVisible(1) = True
            SSTab.TabVisible(0) = False
            mSQL = "Select * From faConfig"
            objDb.SetConnection mCnn
            Set Rec = objDb.ExecuteSP(mSQL, , , , mCnn, adCmdText)
            If Not (Rec.EOF And Rec.BOF) Then
                txtUrl = gbDefaultUrl
                chkRLB.Value = IIf(IsNull(Rec!tnyRLB), "", Rec!tnyRLB)
                chkDOPFA.Value = IIf(IsNull(Rec!tnyLinkWithDandOPFA), 0, Rec!tnyLinkWithDandOPFA)
                chkKMBR.Value = IIf(IsNull(Rec!tnyLinkWithKMBR), 0, Rec!tnyLinkWithKMBR)
                chkPTax.Value = IIf(IsNull(Rec!tnyLinkWithPropertyTax), 0, Rec!tnyLinkWithPropertyTax)
                chkFinHo.Value = IIf(IsNull(Rec!tnyLinkWithFinanceHO), 0, Rec!tnyLinkWithFinanceHO)
                chkSugama.Value = IIf(IsNull(Rec!tnyLinkWithSugama), 0, Rec!tnyLinkWithSugama)
                chkSevana.Value = IIf(IsNull(Rec!tnyLinkWithSevana), 0, Rec!tnyLinkWithSevana)
                chkBandD.Value = IIf(IsNull(Rec!tnyLinkWithBandDSchedules), 0, Rec!tnyLinkWithBandDSchedules)
                txtAllotmentNo.Text = IIf(IsNull(Rec!vchStartingAllotmentNo), "", Rec!vchStartingAllotmentNo)
                txtUrl.Text = IIf(IsNull(Rec!vchDefaultUrl), "", Rec!vchDefaultUrl)
            End If
        Else
            MsgBox "You Are Not Authorized to use This Menu", vbInformation
            SSTab.Enabled = False
            Exit Sub
        End If
        Call setBank(CInt(gbDefaultBankID))
        txtDrawnBank.Text = gbRemittingBank
        txtBankPlace.Text = gbRemittingPlaceOfBank
        
    End Sub
'''    Private Sub DefaultSettingsPT()
'''        Dim mIndex As Long
'''        Dim mTempString As String
'''        Dim mCount As Integer
'''        Dim objTransactionType As New clsTransactionType
'''        Dim objAccounts As New clsAccounts
'''        Dim objInstrument As New clsInstruments
'''        Dim objBank As New clsBank
'''        mTempString = ReadIniFile(App.Path & "\Saankhya.INI", "Receipt PTax", "DefaultTransactionTypeID")
'''        objTransactionType.SetTransactionType (CInt(mTempString))
'''        mTempString = objTransactionType.TransactionType
'''        mIndex = SendMyMessage(cmbTransactionTypePT.hwnd, LB_FINDSTRING, -1, mTempString)
'''        If mIndex <> -1 Then
'''            cmbTransactionTypePT.ListIndex = mIndex
'''        End If
'''        mIndex = -1
'''        mTempString = ""
'''        mTempString = ReadIniFile(App.Path & "\Saankhya.INI", "Receipt PTax", "DefaultAccountHeadCode")
'''        objAccounts.SetAccountCode (mTempString)
'''        mTempString = objAccounts.AccountHead
'''        mIndex = SendMyMessage(cmbAccountHeadPT.hwnd, LB_FINDSTRING, -1, mTempString)
'''        If mIndex <> -1 Then
'''            cmbAccountHeadPT.ListIndex = mIndex
'''        End If
'''        mIndex = -1
'''        mTempString = ""
'''        mTempString = ReadIniFile(App.Path & "\Saankhya.INI", "Receipt PTax", "DefaultInstumentID")
'''        objInstrument.SetInstrumentType (CInt(mTempString))
'''        mTempString = objInstrument.InstrumentType
'''        mIndex = SendMyMessage(cmbInstrumentTypePT.hwnd, LB_FINDSTRING, -1, mTempString)
'''        If mIndex <> -1 Then
'''            cmbInstrumentTypePT.ListIndex = mIndex
'''        End If
'''        mIndex = -1
'''        mTempString = ""
'''        mTempString = ReadIniFile(App.Path & "\Saankhya.INI", "Receipt PTax", "DefaultBankID")
'''        objBank.SetBankInfo (mTempString)
'''        mTempString = objBank.BankName
'''        mIndex = SendMyMessage(cmbBankPT.hwnd, LB_FINDSTRING, -1, mTempString)
'''        If mIndex <> -1 Then
'''            cmbBankPT.ListIndex = mIndex
'''        End If
'''        mTempString = ""
'''        mTempString = ReadIniFile(App.Path & "\Saankhya.INI", "Receipt PTax", "DefaultZone")
'''        For mCount = 0 To cmbZonePT.ListCount - 1
'''            If cmbZonePT.ItemData(mCount) = CInt(mTempString) Then
'''                cmbZonePT.ListIndex = mCount
'''                Exit For
'''            End If
'''        Next mCount
'''    End Sub


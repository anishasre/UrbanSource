VERSION 5.00
Begin VB.Form frmRegisterOfBills 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Register of Bills"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9375
      TabIndex        =   25
      Top             =   5205
      Width           =   9435
      Begin VB.CommandButton cmdDemand 
         Caption         =   "&Generate Demand"
         Enabled         =   0   'False
         Height          =   435
         Left            =   4575
         TabIndex        =   11
         Top             =   120
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Edit"
         Height          =   435
         Left            =   2880
         TabIndex        =   10
         Top             =   105
         Width           =   1470
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   30
      ScaleHeight     =   840
      ScaleWidth      =   9360
      TabIndex        =   0
      Top             =   -45
      Width           =   9360
   End
   Begin VB.Frame frm1 
      Height          =   4440
      Left            =   15
      TabIndex        =   13
      Top             =   750
      Width           =   9405
      Begin VB.ComboBox cmbRegType 
         Height          =   315
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2370
      End
      Begin VB.TextBox txtRegName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3375
         TabIndex        =   2
         Top             =   900
         Width           =   3930
      End
      Begin VB.CommandButton cmdSearchAccountHeads 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Left            =   7380
         TabIndex        =   9
         Top             =   3675
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchFunctionary 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Left            =   7380
         TabIndex        =   8
         Top             =   3285
         Width           =   300
      End
      Begin VB.CommandButton cmdSearchFunction 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Left            =   7380
         TabIndex        =   7
         Top             =   2895
         Width           =   300
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3690
         Width           =   3930
      End
      Begin VB.TextBox txtFunctionary 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3285
         Width           =   3930
      End
      Begin VB.TextBox txtFunction 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2895
         Width           =   3930
      End
      Begin VB.TextBox txtRefTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         TabIndex        =   4
         Top             =   1695
         Width           =   3930
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         TabIndex        =   3
         Top             =   1290
         Width           =   3930
      End
      Begin VB.ComboBox cmbPeriod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2085
         Width           =   2370
      End
      Begin VB.TextBox txtDueDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2505
         Width           =   2085
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Type :"
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
         Left            =   2790
         TabIndex        =   24
         Top             =   495
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reference Title :"
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
         Left            =   1935
         TabIndex        =   21
         Top             =   1695
         Width           =   1350
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Expenditure Account Head :"
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
         Left            =   1035
         TabIndex        =   20
         Top             =   3690
         Width           =   2250
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Functionary :"
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
         Left            =   2235
         TabIndex        =   19
         Top             =   3285
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Function :"
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
         Left            =   2490
         TabIndex        =   18
         Top             =   2895
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reference No :"
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
         Left            =   2085
         TabIndex        =   17
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Period :"
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
         Left            =   2655
         TabIndex        =   16
         Top             =   2085
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Due Day Of Month :"
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
         Left            =   1755
         TabIndex        =   15
         Top             =   2490
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Register Name :"
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
         Left            =   1980
         TabIndex        =   14
         Top             =   900
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmRegisterOfBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCheckDemand As Variant '1=if a demand is already generated
Dim mTypeID      As Variant

    Private Sub cmbPeriod_Change()
''''''''    Dim mCnn    As New ADODB.Connection
''''''''    Dim msQl    As String
''''''''    Dim ObjDb   As New clsDb
''''''''    Dim Rec     As New ADODB.Recordset
''''''''
''''''''        If ObjDb.SetConnection(mCnn) Then
''''''''          If cmbPeriod.ListIndex = 5 Then
''''''''            msQl = " SELECT tnyStatus From faBillRegisters Where intID = " & txtRegName.Tag & " "
''''''''            Rec.Open msQl, mCnn
''''''''            If Not (Rec.EOF And Rec.BOF) Then
''''''''                If Rec!tnyStatus = 3 Then
''''''''                    MsgBox "Payment already made.Demand cannot be edited", vbInformation
''''''''                    Exit Sub
''''''''                End If
''''''''            End If
''''''''            Rec.Close
''''''''          End If
''''''''        End If
    End Sub

    Private Sub cmbPeriod_Click()
    
    Dim mCnn    As New ADODB.Connection
    Dim msQl    As String
    Dim ObjDb   As New clsDB
    Dim Rec     As New ADODB.Recordset
       
        If cmbPeriod.ListIndex = 1 Or cmbPeriod.ListIndex = 2 Then
            txtDueDate.Enabled = True
            txtDueDate = ""
        Else
            txtDueDate = 1
            txtDueDate.Enabled = False
        End If
'         If ObjDb.SetConnection(mCnn) Then
'          If cmbPeriod.ListIndex = 5 Then
'            msQl = " SELECT tnyStatus From faBillRegisters Where intID = " & txtRegName.Tag & " "
'            Rec.Open msQl, mCnn
'            If Not (Rec.EOF And Rec.BOF) Then
'                If Rec!tnyStatus = 3 Then
'                    MsgBox "Payment already made.Demand cannot be edited", vbInformation
'                    Exit Sub
'                End If
'            End If
'            Rec.Close
'          End If
'     End If
   
    End Sub
    Private Sub cmdDemand_Click()
    Dim mCnn    As New ADODB.Connection
    Dim msQl    As String
    Dim ObjDb   As New clsDB
    Dim mArrIn  As Variant
    Dim mArrOut As Variant
    Dim mLoop       As Integer
    Dim mintID     As Variant
    'Dim mTypeID      As Variant
    Dim mPeriodID    As Variant
    Dim mDueDate     As Date
    Dim mDemandDueDt As Variant
    Dim mDemandDtMonth As Variant
    Dim mLoopContrl    As Variant
    Dim mLoopCondition     As Variant
    Dim mCount          As Variant
    Dim mYear    As Variant
    Dim Rec     As New ADODB.Recordset
    Dim i As Integer
    Dim mPeriod As Integer
    Dim mNewMonth As Date
    
    '---------TO GENERATE DEMAND------------------
    
'''''    If ObjDb.SetConnection(mCnn) Then
'''''        msQl = " SELECT tnyStatus From faBillRegisters Where intID = " & txtRegName.Tag & " "
'''''        Rec.Open msQl, mCnn
'''''        If Not (Rec.EOF And Rec.BOF) Then
'''''            If Rec!tnyStatus = 3 Then
'''''                MsgBox "Payment already made.Demand cannot be edited", vbInformation
'''''                Exit Sub
'''''            End If
'''''        End If
'''''        Rec.Close
'''''    End If
'''''
    
    
    frmListofBillRegister.CheckDemandID = 0
    If ObjDb.SetConnection(mCnn) Then
       If cmbPeriod.ListIndex = 5 Then
                mPeriodID = 1
            ElseIf cmbPeriod.ListIndex = 4 Then
                mPeriodID = 2
            Else
                mPeriodID = Null
        End If
        Select Case cmbPeriod.Text
            Case Is = "Monthly"
                mLoopContrl = 12
                mDemandDtMonth = 4
                mLoopCondition = 9 'just assigned
                mPeriod = 1
            Case Is = "Bymonthly"
                mLoopContrl = 6
                mDemandDtMonth = 4
                mLoopCondition = 6
                mPeriod = 2
            Case Is = "Quarterly"
                mLoopContrl = 4
                mDemandDtMonth = 4 '8
                mLoopCondition = 4
                mPeriod = 3
            Case Is = "Half Yearly"
                mLoopContrl = 2
                mDemandDtMonth = 4 '9
                mLoopCondition = 2
                mPeriod = 6
            Case Is = "Yearly"
                mLoopContrl = 1
                mDemandDtMonth = 4
                mPeriod = 12
        End Select
        mYear = gbFinancialYearID
        
        msQl = "Select Count(*) As Count From faBillRegisters Where intRegID = " & txtRegName.Tag & " And tnyStatus = 3"
        Rec.Open msQl, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mCount = IIf(IsNull(Rec!count), 0, Rec!count)
        End If
        Rec.Close
            
        For i = 1 To mCount Step mPeriod
            mLoopContrl = mLoopContrl - 1
        Next i
        
        If mCount > 0 Then
            msQl = "Select MAX(dtDemandDueDate) As LastPDate From faBillRegisters Where intRegID = " & txtRegName.Tag & " And tnyStatus = 3"
            Rec.Open msQl, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                mNewMonth = Month(IIf(IsNull(Rec!LastPDate), 0, Rec!LastPDate))
                mDemandDtMonth = mNewMonth + 1
                
            End If
            Rec.Close
        End If
        
        mCnn.Execute "Delete From faBillRegisters Where intRegID = " & txtRegName.Tag & " And tnyStatus <> 3"
        
        For mLoop = 1 To mLoopContrl Step 1
            mDueDate = val(txtDueDate.Text)
            mDemandDueDt = DateSerial(mYear, mDemandDtMonth, mDueDate)
            If mDemandDtMonth > 12 Then
                mDemandDtMonth = mID$(mDemandDueDt, 4, 2)
                mYear = gbFinancialYearID + 1
            End If
            
            mintID = -1
            mTypeID = txtRegName.Tag
            mArrIn = Array(mintID, mDemandDueDt, _
                            mTypeID, _
                            mYear, _
                            mDemandDtMonth, _
                            mPeriodID, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            Null, _
                            1, _
                            Null _
                            )
            ObjDb.ExecuteSP "spSaveBillRegisters", mArrIn, , , mCnn, adCmdStoredProc
            If mLoopCondition = 2 Then
            mDemandDtMonth = mDemandDtMonth + 6
            ElseIf mLoopCondition = 4 Then
            mDemandDtMonth = mDemandDtMonth + 4
            ElseIf mLoopCondition = 6 Then
            mDemandDtMonth = mDemandDtMonth + 2
            Else
            mDemandDtMonth = mDemandDtMonth + 1
            End If
           
        Next mLoop
         
        MsgBox "Demand Generated", vbInformation, "Saankhya"
        cmdDemand.Enabled = False
        frmListofBillRegister.CheckDemandID = 1
        frmListofBillRegister.CheckRegID = mTypeID
        frmListofBillRegister.Show vbModal
        Unload Me
    End If
    End Sub
    Private Sub cmdSave_Click()
    Dim mCnn    As New ADODB.Connection
    Dim msQl    As String
    Dim ObjDb   As New clsDB
    Dim mArrIn  As Variant
    Dim mArrOut As Variant
    Dim mintID     As Variant
    Dim mTypeID   As Variant
    Dim mPeriodID As Variant
    Dim Rec     As New ADODB.Recordset
    
    
        If SaveValidation = False Then Exit Sub
        
        If ObjDb.SetConnection(mCnn) Then
            If cmdSave.Caption <> "Save" Then
            
                msQl = " SELECT tnyStatus,tnyPreriodID From faBillRegisters Where intRegID = " & txtRegName.Tag & " "
                Rec.Open msQl, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    If Rec!tnyStatus = 3 Then
                        If (Rec!tnyPreriodID = 1) Then
                            MsgBox "Payment already made.Demand cannot be edited", vbInformation
                            Exit Sub
                        End If
                    End If
                End If
                Rec.Close
             End If

            
            mintID = IIf(txtRegName.Tag = "", -1, val(txtRegName.Tag))
            mTypeID = cmbRegType.ItemData(cmbRegType.ListIndex)
            mPeriodID = cmbPeriod.ItemData(cmbPeriod.ListIndex)
            mArrIn = Array(mintID, mTypeID, _
                            txtRegName, _
                            mPeriodID, _
                            txtDueDate.Text, _
                            txtRefNo.Text, _
                            txtRefTitle.Text, _
                            txtFunctionary.Tag, _
                            txtFunction.Tag, _
                            txtAccountHead.Tag, _
                            0 _
                         )
            ObjDb.ExecuteSP "spSaveRegisterOfBill", mArrIn, mArrOut, , mCnn, adCmdStoredProc
            MsgBox "Saved Successfully!", vbInformation, "Saankhya"
            If IsNumeric(mArrOut(0, 0)) Then
                txtRegName.Tag = mArrOut(0, 0)
            End If
            cmdSave.Enabled = False
            frm1.Enabled = False
            If mCheckDemand = 1 Then
                cmdDemand.Enabled = True 'false
                mCheckDemand = 0
                
            ElseIf mCheckDemand = 0 Then
                cmdDemand.Enabled = True
            End If
        Else
            MsgBox "Connection to Finance Does not Exist, Please contact your System Administrator", vbInformation
        End If
        
        'Unload Me
        'frmListOfRegisterOfBills.Show
        'Unload Me
    End Sub
   Private Sub cmdSearchAccountHeads_Click()
        Dim mToken   As String
        frmSearchAccountHeads.SQLString = "SELECT     ( vchAccountHeadCode + '  ' + vchAccountHead) AS vchAccountHeadCode,intAccountHeadID From faAccountHeads WHERE (vchAccountHeadCode + '  ' + vchAccountHead LIKE '2%')"
        frmSearchAccountHeads.Show vbModal
         mToken = Token(gbSearchStr, " ")
           If gbSearchID <> -1 Then
               txtAccountHead.Text = Trim(gbSearchStr)
               txtAccountHead.Tag = gbSearchID
               gbSearchID = -1
               gbSearchStr = ""
           End If
    End Sub
    Private Sub cmdSearchFunction_Click()
        Dim mToken   As String
        frmSearchFunction.Show vbModal
        mToken = Token(gbSearchStr, " ")
        txtFunction.Text = Trim(gbSearchStr)
        txtFunction.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub cmdSearchFunctionary_Click()
        Dim mToken As String
        frmSearchFunctionary.Show vbModal
        mToken = Token(gbSearchStr, " ")
        txtFunctionary.Text = Trim(gbSearchStr)
        txtFunctionary.Tag = gbSearchID
        gbSearchStr = ""
        gbSearchID = -1
    End Sub
    Private Sub Form_Activate()
'        Me.Top = 500
'        Me.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
    Private Sub Form_Load()
        Call FormInitialize
    End Sub
    Private Sub FormInitialize()
        Dim mCrl As Control
        For Each mCrl In Me.Controls
            If TypeOf mCrl Is TextBox Then
                mCrl.Text = ""
                mCrl.Tag = ""
            End If
        Next
        txtDueDate = 1
        txtDueDate.Enabled = False
        cmdSave.Enabled = True
        cmdDemand.Enabled = False
       ' mCheckDemand = 0
        Call PopulateList(cmbRegType, "Select vchRegType, intRegTypeID From faRegisterTypes Order By vchRegType", , True, True, True)
        Call PopulateList(cmbPeriod, "SELECT vchPeriodicity, intPeriodicityID From faPeriodicity WHERE intTypeID = 8", , True, True, True)
        
    End Sub
    Private Function SaveValidation() As Boolean
        If cmbRegType.ListIndex <= 0 Then
             MsgBox "Select the Register Type", vbInformation
             cmbRegType.SetFocus
             SaveValidation = False
             Exit Function
        End If
        If Trim(txtRegName.Text) = "" Then
            MsgBox "Enter the Register Name", vbInformation
            txtRegName.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If cmbPeriod.ListIndex <= 0 Then
            MsgBox "Select the Period", vbInformation
            cmbPeriod.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtRefNo.Text) = "" Then
            MsgBox "Enter the Reference Number", vbInformation
            txtRefNo.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtRefTitle.Text) = "" Then
            MsgBox "Enter the Reference Title", vbInformation
            txtRefTitle.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtDueDate.Text) = "" Or Trim(txtDueDate.Text) >= 31 Or Trim(txtDueDate.Text) = 0 Then
            MsgBox "Not a Valid Date ", vbInformation
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtFunction.Text) = "" Then
            MsgBox "Select the Function", vbInformation
            txtFunction.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtFunctionary.Text) = "" Then
            MsgBox "Enter the Functionary", vbInformation
            txtFunctionary.SetFocus
            SaveValidation = False
            Exit Function
        End If
        If Trim(txtAccountHead.Text) = "" Then
            MsgBox "Enter the Account Head", vbInformation
            txtAccountHead.SetFocus
            SaveValidation = False
            Exit Function
        End If
        SaveValidation = True
    End Function

    Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
        End If
    End Sub
    
     Public Property Let CheckDemand(mData As Variant)
        mCheckDemand = mData
    End Property
    Public Property Get CheckDemand() As Variant
        CheckDemand = mCheckDemand
    End Property
    
    
    
'    Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
'        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
'                KeyAscii = 0
'        End If
'    End Sub
'    Private Sub txtRefTitle_KeyPress(KeyAscii As Integer)
'        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
'                KeyAscii = 0
'        End If
'    End Sub
'
'    Private Sub txtRegName_KeyPress(KeyAscii As Integer)
'        If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
'                KeyAscii = 0
'        End If
'    End Sub

    
    Private Sub txtDueDate_LostFocus()
        If val(txtDueDate.Text) > 31 Then
            MsgBox "Enter a valid Date ", vbInformation, "Saankhya"
            txtDueDate.SetFocus
        End If
    End Sub

VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSplashForPanchayat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmSplashForPanchayat.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbNumSeatID 
      Height          =   315
      Left            =   -195
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox cmbSeats 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4275
      TabIndex        =   2
      Top             =   2175
      Width           =   1620
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5040
      TabIndex        =   4
      Top             =   2595
      Width           =   720
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4290
      TabIndex        =   3
      Top             =   2595
      Width           =   720
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4275
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1815
      Width           =   1605
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4275
      TabIndex        =   0
      Top             =   1455
      Width           =   1605
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   5895
      Top             =   2910
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3855
      TabIndex        =   11
      Top             =   2235
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3375
      TabIndex        =   10
      Top             =   1860
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3780
      TabIndex        =   9
      Top             =   1470
      Width           =   465
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6120
      TabIndex        =   8
      Top             =   3900
      Width           =   1050
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000009&
      Height          =   225
      Left            =   75
      TabIndex        =   7
      Top             =   3630
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright @ Information Kerala Mission"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   585
      TabIndex        =   6
      Top             =   3165
      Width           =   2835
   End
End
Attribute VB_Name = "frmSplashForPanchayat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbSeats_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
    End If
End Sub
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdLogin_Click()
    
    Dim objdb As New clsDB
    Dim objUser As New clsUser
    Dim objCounter As New clsCounter
    Dim mCon As New ADODB.Connection
    Dim mVarrIn As Variant
    Dim mShiftID As Integer
    Dim mSeatID As Variant
    Dim Rec As New ADODB.Recordset
    Dim Recs As New ADODB.Recordset
    Dim mServerDate As Date
    Dim mSql As String
    Dim mLDType As Integer
    Dim mVerID  As String
    Dim mSubVerID  As String
    Dim mDBVerID  As String
    Dim mDBSubVerID  As String
    Dim mClientDate As Date
    '-------------------------------------------------------------------------'
    '                                  Validations                            '
    '-------------------------------------------------------------------------'
    If Trim(txtLogin.Text) = "" Then
        MsgBox "Enter the Login Name", vbCritical
        txtLogin.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Enter the Password", vbCritical
        txtPassword.SetFocus
        Exit Sub
    End If
    If cmbSeats.ListIndex < 0 Then
        MsgBox "Select the Seat", vbCritical
        cmbSeats.SetFocus
        Exit Sub
    End If
    
    '-------------------------------------------------------------------------'
    '             V E R S I O N    C O N T R O L E R                          '
    '-------------------------------------------------------------------------'
    
    mSql = "Select * From faLBSettings"
    Set Rec = GetRecordSet(mSql)
    If Not (Rec.BOF And Rec.EOF) Then
        mLDType = IIf(IsNull(Rec!tnyLBTypeID), "", Rec!tnyLBTypeID)
    End If
    Rec.Close
    If mLDType = 3 Or mLDType = 4 Then
        mVerID = gbVerID
        mSubVerID = gbVerSubID
        mDBVerID = gbDBVerID
        mDBSubVerID = gbDBSubVerID
    Else
        mVerID = gbPVerID
        mSubVerID = gbPVerSubID
        mDBVerID = gbPDBVerID
        mDBSubVerID = gbPDBSubVerID
    End If
    If objdb.CreateNewConnection(mCon, enuSourceString.Saankhya) Then
        Rec.Open "spGetVersion", mCon, adOpenStatic, adLockReadOnly, adCmdStoredProc
        If Not (Rec.BOF And Rec.EOF) Then
            If Rec!vchVersionKey <> mVerID Then
                mSql = "Application Version Miss Match!" + vbCrLf
                MsgBox mSql, vbCritical
                End
            End If
            If Rec!vchSubVersionKey <> mSubVerID Then
                MsgBox "Application (Sub)Version Miss Match!", vbCritical
                End
            End If
        End If
        Rec.Close

        Rec.Open "spGetDBVersion", mCon, adOpenStatic, adLockReadOnly, adCmdStoredProc
        If Not (Rec.BOF And Rec.EOF) Then
            If Rec!vchDBVersionKey <> mDBVerID Then
                MsgBox "Database Version Miss Match!", vbCritical
                End
            End If
            If Rec!vchDBSubVersionKey <> mDBSubVerID Then
                MsgBox "Database (Sub)Version Miss Match!", vbCritical
                End
            End If
        End If
        Rec.Close
        
    Else
        MsgBox "Connection Failed, Check your ODBC, Please!", vbInformation
        End
    End If
'    ------------------------------------------------------------------------
    '--Urgent release For Kollam (Rent On Land)
    '------------------------------------------------------------------------
''''    mSQL = "Select intLocalBodyID From faVouchers"
''''    Rec.Open mSQL, mCon, adOpenStatic, adLockReadOnly, adCmdText
''''    If Not (Rec.BOF And Rec.EOF) Then
''''        If gbLocalBodyID <> 171 Then
''''            If Rec!intLocalBodyID <> 171 Then
''''                MsgBox "Build for Kzhikkode  2.2.4", vbCritical
''''                End
''''            End If
''''        End If
''''    End If
''''    Rec.Close
'    ------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------'
    ' Checking User Login and Password If found objUser will Set User Details '
    '-------------------------------------------------------------------------'
    If objUser.Login(Trim(txtLogin.Text), txtPassword.Text) = True Then
        
        '--------------------------------------------------------------------'
        ' CHECK ACTIVE USER OR NOT
        '--------------------------------------------------------------------'
        If gbUserActiveFlag > 0 Then
            MsgBox "This User is not Active user!", vbInformation
            Exit Sub
        End If
        '--------------------------------------------------------------------'
        ' Checking User permission to a Seat                                 '
        '--------------------------------------------------------------------'
        If cmbSeats.ListIndex > -1 And Not IsNull(gbSeatName) Then
            gbSeatID = cmbNumSeatID.List(cmbSeats.ListIndex)
            gbSeatName = cmbSeats.Text
        End If
        If Not objUser.LogonToSeat(gbSeatID, gbUserID) Then
            
            MsgBox "You are not Allowed to Logon to this Seat", vbCritical
            Exit Sub
            '-------------------------------------------------------------'
            ' Have to Validate Any other user Logon to the same seat      '
            '-------------------------------------------------------------'
            '? Send a Log Record to Administrator
            '?
            '?
            '-------------------------------------------------------------'
'            Rec.Open "Select * From faUserMovement Where numSeatID = " & gbSeatID & " And numUserID <> " & gbUserID & " And tnyStatus <> 3"
'            If Not (Rec.BOF And Rec.EOF) Then
'                MsgBox "The User on this Seat is already Logined !", vbInformation
'                mCon.Execute "Insert into faLoginLogFile (numUserID,dtLogRequestTime,intCounterID) Values (" & gbUserID & ", getDate() ," & gbCounterID & ")"
'                Exit Sub
'            End If
            '-------------------------------------------------------------'
            ' Need to validate Whether Counter and Seat and Sections      '
            '-------------------------------------------------------------'
            '?
            '?
            '?
            '?
            '-------------------------------------------------------------'
'            Rec.Open "Select * from faCounters Where intCounterID = " & gbCounterID
'            If Rec!intSectionID <> gbSectionID Then
'                MsgBox "Sections do not Match", vbInformation
'                Exit Sub
'            End If
        Else
            '-------------------------------------------------------------'
            ' ELSE PART IS ADDED BY AIBY ON 18-Jan-2009
            '-------------------------------------------------------------'
            gbSeatID = objUser.SeatID
            gbSeatName = objUser.SeatName
            gbSeatGroupID = objUser.SeatGroupID
            gbSectionID = objUser.SectionID
            gbSectionName = objUser.SectionName
            
            '-------------------------------------------------------------'
            ' Permission Checking to Login Janasevana Kendram Counters    '
            '-------------------------------------------------------------'
            If gbCounterSectionID = gbJSKSectionID Or gbSectionID = gbJSKSectionID Then
                If gbCounterSectionID <> gbSectionID Then
                    Dim mStrMsg As String
                    If gbSectionID = gbJSKSectionID Then
                        mStrMsg = mStrMsg + " Seats Assigned to Janasevana Kendram is" & vbNewLine
                        mStrMsg = mStrMsg + "            Permitted To Login" & vbNewLine
                        mStrMsg = mStrMsg + " From Janasevana Kendram Counters Only" & vbNewLine
                    Else
                        mStrMsg = mStrMsg + " Seats Assigned To Other Sections" & vbNewLine
                        mStrMsg = mStrMsg + " Are Not Permitted To Login" & vbNewLine
                        mStrMsg = mStrMsg + " Janasevana Kendram Counters" & vbNewLine
                    End If
                    MsgBox mStrMsg, vbInformation
                    Exit Sub
                End If
            End If
            
        End If
        
        '--------------------------------------------------------------------'
        ' User Login / Shift and User movements                              '
        '===================================================================='
        '------------------------------------------'
        ' Get Server Date and Time                 '
        '------------------------------------------'
        'mShiftID = 1
        objdb.SetConnection mCon
        Rec.Open "Select GetDate() as ServerDate ", mCon
        If Not (Rec.EOF And Rec.BOF) Then
            mServerDate = Rec!ServerDate
        '-----syalima on 13/sep/2018------------
            mClientDate = DateTime.Now
            If Abs(DateDiff("d", mServerDate, mClientDate)) > 0 Then
                MsgBox "Mismatch in Server date and Client date!!!", vbCritical
                cmdLogin.Enabled = False
                Exit Sub
            End If
        End If
        Rec.Close
        '--------------End-------------------
        '-------------------------------------------'
        ' Get Last Open UserMovement Details if any '
        '-------------------------------------------'
        gbLocalBodyID = objUser.LocalBodyID
        If IsEmpty(gbSeatID) Then
            mSql = " Select * From faUserMovement Where intID = "
            mSql = mSql + "(Select Max(intID) From faUserMovement "
            mSql = mSql + " Where numUserID = " & gbUserID & " And intLocalBodyID = " & objUser.LocalBodyID
            mSql = mSql + " And dbo.fnConvertDate(dtLoginDate) = '" & DdMmmYy(mServerDate) & "')"
        Else
            mSql = " Select * From faUserMovement Where intID = "
            mSql = mSql + "(Select Max(intID) From faUserMovement "
            mSql = mSql + " Where numUserID = " & gbUserID & " And intLocalBodyID = " & gbLocalBodyID
            mSql = mSql + " And dbo.fnConvertDate(dtLoginDate) = '" & DdMmmYy(mServerDate)
            mSql = mSql + "'  And numSeatID = " & gbSeatID & ")"
            mSql = mSql + " And tnyStatus In (1,2) "
        End If
        Rec.Open mSql, mCon, adOpenKeyset, adLockOptimistic
        '-------------------------------------------'
        ' Checking Previous Login in UserMovement   '
        '-------------------------------------------'
        If Not (Rec.BOF And Rec.EOF) Then
            'If Rec!tnyStatus = 1 Or Rec!tnyStatus = 2 Then
            If Rec!intCounterID = gbCounterID And Rec!numSeatID = val(gbSeatID) Then
                gbShiftID = Rec!intShiftID
                Set Recs = GetRecordSet("Select vchShift From faShifts Where intShiftID=" & IIf(IsNull(gbShiftID), 0, gbShiftID))
                If Not (Recs.EOF And Recs.BOF) Then
                    gbShiftName = Recs!vchShift
                End If
                mSql = "Update faUserMovement Set tnyStatus = 1 Where intID = " & Rec!intID
                mCon.Execute mSql
'                    objDB.ExecuteSP mSQL, , , , mCon
                GoTo SkipToUnload:
            End If
            'End If
        Else
            '-------------------------------------------'
            ' Checking Seat added in any Shift          '
            '-------------------------------------------'
            Rec.Close
            If IsEmpty(gbSeatID) Then
                gbSeatID = 0
            End If
            mSql = "Select Count(*) As ShiftCount From faShiftManager Inner Join "
            mSql = mSql + " faShifts On faShifts.intShiftID = faShiftManager.intShiftID"
            mSql = mSql + " Where numSeatID = " & gbSeatID
            Rec.Open mSql, mCon, adOpenKeyset, adLockOptimistic
            If Rec!ShiftCount > 0 Then
                '----------------------------------------------------------------'
                ' Check the Login Time whether it belongs in Shift Time or not   '
                '----------------------------------------------------------------'
                Rec.Close
                mSql = "Select * From faShiftManager Inner Join "
                mSql = mSql + " faShifts On faShifts.intShiftID = faShiftManager.intShiftID"
                mSql = mSql + " Where (Select CONVERT ( SmallDateTime,  Right(CONVERT(char(20) ,  GetDate()),8))) Between dtStartTime And dtEndTime"
                mSql = mSql + " And numSeatID = " & gbSeatID
                Rec.Open mSql, mCon, adOpenKeyset, adLockOptimistic
                
                If Not (Rec.BOF And Rec.EOF) Then
                    gbShiftID = Rec!intShiftID
                    gbShiftName = Rec!vchShift
                Else
                    '---------------------------------------------------'
                    ' Login time *NOT belongs in between the Shift time '
                    '---------------------------------------------------'
                    mSql = "Shift time is over! "
                    MsgBox mSql
                    'Exit Sub
                End If
                
            Else
                '---------------------------------------'
                ' Shift is not applicable for this Seat '
                '---------------------------------------'
            End If
        End If
        Rec.Close
        
       '----------------------------------------------'
         ''''''''For checkin AccAssistant'''''''''
         ''''''''Done By Sajith Kumar K V  On 5-7-12
       '----------------------------------------------'
         If gbSeatGroupID = 14 Then '''''for checkin accassistantseatgroup'''''
             Dim Checkintime As String
             Dim currentdate As String
             Dim mSQL1 As String
             Dim Rec1 As New ADODB.Recordset
             mSql = "Select count(*)count from faEmpLog where vchEmpCode='" & txtLogin.Text & "' "
             Rec.Open mSql, mCon
             If Not (Rec.BOF And Rec.EOF) Then
             
             If Rec!count = 0 Then
             ''''insert''''''''
             mVarrIn = Array(gbUserID, txtLogin.Text, Null, 1, 1)
             
             objdb.ExecuteSP "spSaveEmpLog", mVarrIn, , , mCon
             
             Else
             mSQL1 = "Select convert(varchar,dtDate,106) As dtDate,convert(varchar,getdate(),106)As currentdate from faEmpLog Where vchEmpCode='" & txtLogin.Text & "' "
             Rec1.Open mSQL1, mCon
               If Not (Rec1.BOF And Rec1.EOF) Then
               Checkintime = Rec1!dtDate
               currentdate = Rec1!currentdate
                If Checkintime = currentdate Then
                ''''do nothing'''''
                
                Else
                ''''update''''
                mVarrIn = Array(gbUserID, txtLogin.Text, Null, 2, 1)
             
                objdb.ExecuteSP "spSaveEmpLog", mVarrIn, , , mCon
                End If
               End If
             End If
             Else
              
             
             End If
             
             
             Rec.Close
         End If
       
        
        
'        gbShiftID = mShiftID
'        Set Recs = GetRecordSet("Select vchShift From faShifts Where intShiftID=" & gbShiftID)
'        gbShiftName = Recs!vchShift
        mVarrIn = Array(objUser.LocalBodyID, _
                        objUser.UserID, _
                        gbCounterID, _
                        IIf(val(gbShiftID) = 0, Null, val(gbShiftID)), _
                        gbSeatID)
                        
        Call objdb.ExecuteSP("spSaveUserLogin", mVarrIn, , , mCon, adCmdStoredProc)
        
SkipToUnload:

        '''===================================================================='
        ''' THIS FOR BLOCK BUILD FOR ONLY ONE LOCAL BODY ''
        '''===================================================================='
        'Dim mdtDate As Date
        'Dim mCnn As New ADODB.Connection
        'objDB.SetConnection mCnn
        'Set Rec = mCnn.Execute("Select GetDate()")
        'If IsDate(Rec.Fields(0)) Then
        '    mdtDate = DdMmmYy(Rec.Fields(0))
        'Else
        '    MsgBox "Didn't able to Access Server Date", vbInformation
        '    Exit Sub
        'End If
        'Rec.Close
        'mCnn.Close
        '
        'If gbLocalBodyID <> 1249 Then
        '    MsgBox "This Build is customized for NILABURE MUNICIPALITY", vbInformation
        '    Exit Sub
        'End If
        '
        'If Not (mdtDate >= "19-Nov-2013" And mdtDate <= "21-Nov-2013") Then
        '    MsgBox "This Build is not valide any more", vbInformation
        '    Exit Sub
        'End If
        '''===================================================================='
        '''
        '''===================================================================='
        
        Load frmMenu
        Unload Me
        frmMenu.Show
        
        ''Session Taking ''
        mCon.Close
        objdb.SetConnection mCon
        Rec.Open "Select max(intID) From faUserMovement", mCon
        gbSessionID = Rec.Fields(0)
        
    Else
        MsgBox "Login failed", vbInformation, "Saankhya"
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        'Call LoginByAiby
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    Load frmMenu
'    frmMenu.Show
'    Unload Me
End Sub

Private Sub Form_Load()
    Dim mSql As String
    Dim objCounter As New clsCounter
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objdb As New clsDB
    Dim mLbID As Long
    
    WindowsXPC.InitIDESubClassing
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    
    objdb.SetConnection mCnn
    Rec.Open "Select * From faLBSettings", mCnn, adOpenDynamic, adLockOptimistic
    If Not (Rec.EOF And Rec.BOF) Then
        If IsNumeric(Rec!intLBID) Then
            mLbID = Rec!intLBID
        Else
            mLbID = -1
        End If
    End If
    Rec.Close
    
    'fill Seat Combo
    PopulateList cmbSeats, "SELECT chvSeatTitle,numSeatID FROM GL_Seats Where intLocalBodyID = " & mLbID & " ORDER BY chvSeatTitle", , , True, , enuSourceString.DBMaster
    PopulateList cmbNumSeatID, "SELECT numSeatID FROM GL_Seats Where intLocalBodyID = " & mLbID & " ORDER BY chvSeatTitle", , , True, , enuSourceString.DBMaster
    '------------------------------------------------------------------------'
    ' Read and Check IP Address
    '------------------------------------------------------------------------'
    objCounter.SetCounterByIP (GetIPAddress())
    If objCounter.CounterID > 0 Then
        gbCounterNo = objCounter.CounterNo
        gbCounterName = objCounter.CounterDescription
        gbCounterID = objCounter.CounterID
        gbCounterIP = objCounter.CounterIP
        gbCounterSectionID = objCounter.CounterSectionID
        gbCounterSection = objCounter.CounterSection
        gbCounterOperationModeID = objCounter.CounterOperationModeID
        '------------------------------------------------------------'
        'Check Activated Counter
        '------------------------------------------------------------'
        If objCounter.CounterActive = False Then
            MsgBox "This Counter is Deactivated!", vbInformation
            Exit Sub
        End If
    Else
    
        mSql = mSql & "This Computer is not permitted to access " & vbNewLine
        mSql = mSql & "           Saankhya Database..!          " & vbNewLine
        mSql = mSql & " Please contact the System Administrator "
        MsgBox mSql, vbInformation
        cmdLogin.Enabled = False
    
        gbCounterNo = -1
        gbCounterName = ""
        gbCounterID = -1
        gbCounterIP = ""
    End If
    
    '------------------------------------------------------------------------'
    ' Read Mac Address
    '------------------------------------------------------------------------'
    gbCounterMacID = GetMacAddress
    
    
    '------------------------------------------------------------------------'
    ' Version Varification
    '------------------------------------------------------------------------'
    Set Rec = mCnn.Execute("spVarifyZig")
    Dim mMsg As String
    If Not (Rec.BOF And Rec.EOF) Then
        If Rec.Fields(0) <> 1 Then
            mMsg = ""
            mMsg = mMsg + "" + vbCrLf
            mMsg = mMsg + "                      Version Verification failed.               " + vbCrLf
            mMsg = mMsg + "************************************************" + vbCrLf
            mMsg = mMsg + "     This may be caused due to manual Rollback to                 " + vbCrLf
            mMsg = mMsg + "  Previous Version Or Attempt to Tamper Database Version.         " + vbCrLf
            mMsg = mMsg + " " + vbCrLf
            MsgBox mMsg, vbInformation
            cmdLogin.Enabled = False
        End If
    Else
            mMsg = ""
            mMsg = mMsg + "" + vbCrLf
            mMsg = mMsg + "                      Version Verification failed.               " + vbCrLf
            mMsg = mMsg + "************************************************" + vbCrLf
            mMsg = mMsg + "     System required to be update version using" + vbCrLf
            mMsg = mMsg + "     Authorized Version Update tools released from" + vbCrLf
            mMsg = mMsg + "     Software Division, IKM" + vbCrLf
            mMsg = mMsg + " " + vbCrLf
            MsgBox mMsg, vbInformation
            cmdLogin.Enabled = False
    End If
    Rec.Close
    
    Dim mMac As String
    Dim mIsIKM As Boolean
    mMac = GetMacAddress
    mIsIKM = IsIKMLAB(mMac)
    If mIsIKM Then
        lblInfo.Visible = True
    Else
        lblInfo.Visible = False
    End If
       'ServerDate
    Call CheckServerDate
    
    
End Sub
Private Sub Frame1_Click()
'    Load frmMenu
'    frmMenu.Show
'    Unload Me
End Sub
Private Function Authenticate(ByVal mLoginName As String, ByVal mPassword As String) As Boolean
    Dim objUser As New clsUser
    If objUser.Login(mLoginName, mPassword) = True Then
        Authenticate = True
    Else
        Authenticate = False
    End If
End Function

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
    End If
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
    End If
End Sub
Private Sub LoginByAiby()
    'To create a short cut for Login for Me(Aiby)
    ' While testing the Application during Development Stages
    
    txtLogin.Text = "Aiby"
    txtPassword.Text = "ib"
    cmbSeats.Text = "SWD1"
    cmbSeats.SetFocus
End Sub

Public Sub CheckServerDate()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mdtDate  As Date
        Dim mServerDate As Date
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "select top 1 Max(dtDate) as dtDate,GetDate() as ServerDate  from faVouchers  where tnyVoucherTypeID=10 And isNull(tnyStatus,0)<>4 And  intInstrumentTypeID=1 And isNull(tnyVoucherGroupID,0)<>4"
        Rec.Open mSql, mCnn
         If Not (Rec.EOF And Rec.BOF) Then
             If Not IsNull(Rec!dtDate) Then
                mdtDate = Rec!dtDate
                mServerDate = Rec!ServerDate
                Rec.Close
                If CDate(mServerDate) < CDate(mdtDate) Then
                    MsgBox "Please  Check your ServerDate ", vbInformation
                    cmdLogin.Enabled = False
                End If
                
                'If Not (mServerDate <= "04/Mar/2014") Then
                '    MsgBox "TEST VERSION", vbCritical
                '    End
                'End If
                
             End If
        End If
End Sub
   

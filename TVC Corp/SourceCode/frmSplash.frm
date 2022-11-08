VERSION 5.00
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5655
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":1CCA
   ScaleHeight     =   5655
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1425
      TabIndex        =   1
      Top             =   4215
      Width           =   2205
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
      Left            =   1425
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4560
      Width           =   2205
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
      Left            =   1620
      TabIndex        =   4
      Top             =   5295
      Width           =   720
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
      Left            =   2415
      TabIndex        =   5
      Top             =   5295
      Width           =   720
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
      Left            =   1425
      TabIndex        =   3
      Top             =   4905
      Width           =   2220
   End
   Begin VB.ComboBox cmbNumSeatID 
      Height          =   315
      Left            =   -120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4035
      Visible         =   0   'False
      Width           =   390
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   3945
      Top             =   5025
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   4
      Common_Dialog   =   0   'False
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000009&
      Height          =   225
      Left            =   75
      TabIndex        =   14
      Top             =   3630
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Information Kerala Mission"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   765
      TabIndex        =   13
      Top             =   30
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   405
      Picture         =   "frmSplash.frx":14D7B
      Stretch         =   -1  'True
      Top             =   15
      Width           =   330
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kerala Municipal Accounting System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   555
      TabIndex        =   12
      Top             =   765
      Width           =   3210
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saankhya"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4710
      TabIndex        =   11
      Top             =   540
      Width           =   2385
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
      TabIndex        =   10
      Top             =   3900
      Width           =   1050
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
      Left            =   900
      TabIndex        =   9
      Top             =   3900
      Width           =   2835
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4830
      TabIndex        =   8
      Top             =   1755
      Width           =   465
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4425
      TabIndex        =   7
      Top             =   2055
      Width           =   870
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4890
      TabIndex        =   6
      Top             =   2370
      Width           =   390
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub cmbSeats_LostFocus()
    Dim mIndex As Long
    Dim mStr As String
    mStr = cmbSeats.Text
    mIndex = SendMyMessage(cmbSeats.hwnd, CB_FINDSTRING, -1, ByVal mStr)
    If mIndex > -1 Then
        cmbSeats.ListIndex = mIndex
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdLogin_Click()
    Dim objDb As New clsDB
    Dim objUser As New clsUser
    Dim objCounter As New clsCounter
    Dim mCon As New ADODB.Connection
    Dim mVarrIn As Variant
    Dim mShiftID As Integer
    Dim mSeatID As Variant
    Dim Rec As New ADODB.Recordset
    Dim Recs As New ADODB.Recordset
    Dim mServerDate As Date
    Dim mSQL As String
    Dim mLDType As Integer
    Dim mVerID  As String
    Dim mSubVerID  As String
    Dim mDBVerID  As String
    Dim mDBSubVerID  As String
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
    
    mSQL = "Select * From faLBSettings"
    Set Rec = GetRecordSet(mSQL)
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
    If objDb.CreateNewConnection(mCon, enuSourceString.Saankhya) Then
        Rec.Open "spGetVersion", mCon, adOpenStatic, adLockReadOnly, adCmdStoredProc
        If Not (Rec.BOF And Rec.EOF) Then
            If Rec!vchVersionKey <> mVerID Then
                mSQL = "Application Version Miss Match!" + vbCrLf
                MsgBox mSQL, vbCritical
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
        objDb.SetConnection mCon
        Rec.Open "Select GetDate() as ServerDate ", mCon
        If Not (Rec.EOF And Rec.BOF) Then
            mServerDate = Rec!ServerDate
        End If
        Rec.Close
        
        '-------------------------------------------'
        ' Get Last Open UserMovement Details if any '
        '-------------------------------------------'
        gbLocalBodyID = objUser.LocalBodyID
        If IsEmpty(gbSeatID) Then
            mSQL = " Select * From faUserMovement Where intID = "
            mSQL = mSQL + "(Select Max(intID) From faUserMovement "
            mSQL = mSQL + " Where numUserID = " & gbUserID & " And intLocalBodyID = " & objUser.LocalBodyID
            mSQL = mSQL + " And dbo.fnConvertDate(dtLoginDate) = '" & DdMmmYy(mServerDate) & "')"
        Else
            mSQL = " Select * From faUserMovement Where intID = "
            mSQL = mSQL + "(Select Max(intID) From faUserMovement "
            mSQL = mSQL + " Where numUserID = " & gbUserID & " And intLocalBodyID = " & gbLocalBodyID
            mSQL = mSQL + " And dbo.fnConvertDate(dtLoginDate) = '" & DdMmmYy(mServerDate)
            mSQL = mSQL + "'  And numSeatID = " & gbSeatID & ")"
            mSQL = mSQL + " And tnyStatus In (1,2) "
        End If
        Rec.Open mSQL, mCon, adOpenKeyset, adLockOptimistic
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
                mSQL = "Update faUserMovement Set tnyStatus = 1 Where intID = " & Rec!intID
                mCon.Execute mSQL
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
            mSQL = "Select Count(*) As ShiftCount From faShiftManager Inner Join "
            mSQL = mSQL + " faShifts On faShifts.intShiftID = faShiftManager.intShiftID"
            mSQL = mSQL + " Where numSeatID = " & gbSeatID
            Rec.Open mSQL, mCon, adOpenKeyset, adLockOptimistic
            If Rec!ShiftCount > 0 Then
                '----------------------------------------------------------------'
                ' Check the Login Time whether it belongs in Shift Time or not   '
                '----------------------------------------------------------------'
                Rec.Close
                mSQL = "Select * From faShiftManager Inner Join "
                mSQL = mSQL + " faShifts On faShifts.intShiftID = faShiftManager.intShiftID"
                mSQL = mSQL + " Where (Select CONVERT ( SmallDateTime,  Right(CONVERT(char(20) ,  GetDate()),8))) Between dtStartTime And dtEndTime"
                mSQL = mSQL + " And numSeatID = " & gbSeatID
                Rec.Open mSQL, mCon, adOpenKeyset, adLockOptimistic
                
                If Not (Rec.BOF And Rec.EOF) Then
                    gbShiftID = Rec!intShiftID
                    gbShiftName = Rec!vchShift
                Else
                    '---------------------------------------------------'
                    ' Login time *NOT belongs in between the Shift time '
                    '---------------------------------------------------'
                    mSQL = "Shift time is over! "
                    MsgBox mSQL
                    'Exit Sub
                End If
                
            Else
                '---------------------------------------'
                ' Shift is not applicable for this Seat '
                '---------------------------------------'
            End If
        End If
        Rec.Close
'        gbShiftID = mShiftID
'        Set Recs = GetRecordSet("Select vchShift From faShifts Where intShiftID=" & gbShiftID)
'        gbShiftName = Recs!vchShift
        mVarrIn = Array(objUser.LocalBodyID, _
                        objUser.UserID, _
                        gbCounterID, _
                        IIf(val(gbShiftID) = 0, Null, val(gbShiftID)), _
                        gbSeatID)
                        
        Call objDb.ExecuteSP("spSaveUserLogin", mVarrIn, , , mCon, adCmdStoredProc)
        
SkipToUnload:
        Load frmMenu
        Unload Me
        frmMenu.Show
        
        ''Session Taking ''
        mCon.Close
        objDb.SetConnection mCon
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
    Dim mSQL As String
    Dim objCounter As New clsCounter
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objDb As New clsDB
    Dim mLbID As Long
    
    WindowsXPC.InitIDESubClassing
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
    objDb.SetConnection mCnn
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
    
        mSQL = mSQL & "This Computer is not permitted to access " & vbNewLine
        mSQL = mSQL & "           Saankhya Database..!          " & vbNewLine
        mSQL = mSQL & " Please contact the System Administrator "
        MsgBox mSQL, vbInformation
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

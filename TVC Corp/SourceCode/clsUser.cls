VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    
    Option Explicit
    
    '=================================================================================='
    '                                   [clsUser]                                      '
    '=================================================================================='
    '
    '----------------------------------------------------------------------------------'
    ' Class Object for User DB                                                         '
    '                                                                                  '
    ' List of Methods Implemented                                                      '
    '           SetSuperUser                                                           '
    '           SetUser                                                                '
    '           Login                                                                  '
    '           CreateNewUser                                                          '
    '                                                                                  '
    ' Private Members   :                                                              '
    '                      Encript                                                     '
    '                      Decript                                                     '
    '                                                                                  '
    ' Shared Object     :  clsDB                                                       '
    ' Shared Variables  :  gbLocalBodyID                                               '
    ' Default User Type :  0=Super User                                                '
    '                      1=Administrator                                             '
    '                      2=Approver 3:Accounts Office; 4:Operator                    '
    '----------------------------------------------------------------------------------'
    Private mvarUserID          As Variant
    Private mvarLoginName       As Variant
    Private mvarUserName        As Variant
    Private mvarPassWord        As Variant
    Private mvarUserTypeID      As Variant
    Private mvarUserType        As Variant
    Private mvarAddFlag         As Boolean
    Private mvarEditFlag        As Boolean
    Private mvarDeleteFlag      As Boolean
    Private mvarViewFlag        As Boolean
    Private mvarPrintFlag       As Boolean
    Private mUserCon            As New ADODB.Connection
    Private RecUser             As New ADODB.Recordset
    Private mvarLocalBodyID     As Variant
    Private mvarDesignationID   As Variant
    Private mvarDesignation     As Variant
    Private mvarActive          As Boolean
    Private mvarUserAddedDate   As Date
    
    Private mVarUserSeatID      As Variant
    Private mVarUserSeat        As Variant
    Private mVarSectionID       As Variant
    Private mVarSection         As Variant
    Private mVarSeatGroupID     As Variant
    
    Public Property Let UserActive(ByVal vData As Variant)
        mvarActive = vData
    End Property
    Public Property Get UserActive()
        UserActive = mvarActive
    End Property
    Public Property Let UserAddedDate(ByVal vData As Variant)
        mvarUserAddedDate = vData
    End Property
    Public Property Get UserAddedDate()
        UserAddedDate = mvarUserAddedDate
    End Property
    Public Property Let Designation(ByVal vData As Variant)
        mvarDesignation = vData
    End Property
    Public Property Get Designation()
        Designation = mvarDesignation
    End Property
    Public Property Let DesignationID(ByVal vData As Variant)
        mvarDesignationID = vData
    End Property
    Public Property Get DesignationID()
        DesignationID = mvarDesignationID
    End Property
    Public Property Let LocalBodyID(ByVal vData As Variant)
        mvarLocalBodyID = vData
    End Property

    Public Property Set LocalBodyID(ByVal vData As Variant)
        Set mvarLocalBodyID = vData
    End Property

    Public Property Get LocalBodyID() As Variant
        If IsObject(mvarLocalBodyID) Then
            Set LocalBodyID = mvarLocalBodyID
        Else
            LocalBodyID = mvarLocalBodyID
        End If
    End Property
    Public Property Get SeatID() As Variant
        SeatID = mVarUserSeatID
    End Property
    
    Public Property Get SeatGroupID() As Variant
        SeatGroupID = mVarSeatGroupID
    End Property
        
    Public Property Get SeatName() As Variant
        SeatName = mVarUserSeat
    End Property
    Public Property Get SectionID() As Variant
        SectionID = mVarSectionID
    End Property
    Public Property Get SectionName() As Variant
        SectionName = mVarSection
    End Property
    Public Function Login(strLoginName As String, strPassWord As String) As Variant
        Dim objDb As New clsDB
        Dim mSQL As String
        Dim mCon As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mVarrIn As Variant
        
        Login = False
        mVarrIn = Array(strLoginName, strPassWord)
        'If (objDB.SetExtDBConnection(mCon, objDB.GetConnectionString(enuSourceString.DBMaster))) Then
        If objDb.CreateNewConnection(mCon, enuSourceString.DBMaster) Then
            Set Rec = objDb.ExecuteSP("SpGM_User_S2", mVarrIn, , , mCon, adCmdStoredProc)
            If Not Rec.EOF Then
                gbUserID = IIf(IsNull(Rec!numUserID), False, Rec!numUserID)
                Login = SetUser(val(gbUserID))
                gbUserName = IIf(IsNull(mvarUserName), " ", mvarUserName)
                gbUserTypeID = IIf(IsNull(mvarUserTypeID), -1, mvarUserTypeID)
                gbUserActiveFlag = IIf(IsNull(Rec!tnyActive), 1, Rec!tnyActive)
            Else
                Login = False
            End If
        End If
    End Function
    
    
    Public Function LogonToSeat(mSeatID As Variant, numUserID As Variant) As Boolean
        Dim objDb As New clsDB
        Dim mSQL As String
        Dim mCon As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mVarrIn As Variant
        LogonToSeat = False
        objDb.SetConnection mCon
        If IsEmpty(mSeatID) Then
           mSQL = " Select faUserSeatAssign.*, A.chvSeatTitle, A.intSectionID, vchSectionName, A.intGroupID From faUserSeatAssign Inner Join"
            mSQL = mSQL + " DB_Masters..GL_Seats A On A.numSeatID =  faUserSeatAssign.numSeatID Left Join"
            mSQL = mSQL + " faSection On faSection.intSectionID = A.intSectionID "
            mSQL = mSQL + " Where faUserSeatAssign.numUserID = " & numUserID
        Else
            mSQL = " Select faUserSeatAssign.*, A.chvSeatTitle, A.intSectionID, vchSectionName, A.intGroupID From faUserSeatAssign Inner Join"
            mSQL = mSQL + " DB_Masters..GL_Seats A On A.numSeatID =  faUserSeatAssign.numSeatID Left Join"
            mSQL = mSQL + " faSection On faSection.intSectionID = A.intSectionID "
            mSQL = mSQL + " Where faUserSeatAssign.numSeatID = " & mSeatID & " AND faUserSeatAssign.numUserID = " & numUserID
        End If
        Rec.Open mSQL, mCon
        If Not (Rec.BOF And Rec.EOF) Then
            mVarUserSeatID = IIf(IsNull(Rec!numSeatID), -1, Rec!numSeatID)
            mVarUserSeat = IIf(IsNull(Rec!chvSeatTitle), "", Rec!chvSeatTitle)
            mVarSectionID = Rec!intSectionID
            mVarSection = Rec!vchSectionName
            mVarSeatGroupID = Rec!intGroupID
            LogonToSeat = True
        End If
        Rec.Close
    End Function
    Friend Function Encrypt(strPassWord As String) As Variant
       Dim mLength As Integer
        Dim strTemp As String
        mLength = Len(strPassWord)
        While mLength > 0
            strTemp = strTemp + Chr(Asc(mID$(strPassWord, Len(strPassWord) - mLength + 1, 1)) + 125)
            mLength = mLength - 1
        Wend
        Encrypt = strTemp
    End Function
    
    Friend Function Decrypt(strPassWord As String) As Variant
        Dim mLength As Integer
        Dim strTemp As String
        mLength = Len(strPassWord)
        While mLength > 0
            strTemp = strTemp + Chr$(Asc(mID$(strPassWord, Len(strPassWord) - mLength + 1, 1)) - 125)
            mLength = mLength - 1
        Wend
        Decrypt = strTemp
    End Function
    
    
    
    Public Function SetUser(numUserID As Long) As Boolean
        Dim objDb As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mCon As New ADODB.Connection
        Dim mSQL As String
        Dim mVarrIn As Variant
        SetUser = False
        mVarrIn = Array(numUserID)
        If (objDb.SetConnection(mCon)) Then
            Set Rec = objDb.ExecuteSP("spSelectUser", mVarrIn, , , mCon, adCmdStoredProc) 'faUser
            If Not Rec.EOF Then
                mvarUserID = IIf(IsNull(Rec!numUserID), "", Rec!numUserID)
                mvarUserName = IIf(IsNull(Rec!vchUserName), "", Rec!vchUserName)
                mvarUserTypeID = IIf(IsNull(Rec!intUserTypeID), "", Rec!intUserTypeID)
                
                mvarUserType = IIf(IsNull(Rec!vchUserType), "", Rec!vchUserType)
                Select Case mvarUserTypeID
                    Case Is = 0: mvarUserType = "Super User"
                    Case Is = 1: mvarUserType = "Administrator"
                    Case Is = 2: mvarUserType = "Approver "
                    Case Is = 3: mvarUserType = "Accounts Officer"
                    Case Is = 4: mvarUserType = "Operator"
                End Select
                mvarDesignationID = IIf(IsNull(Rec!intDesignationID), "", Rec!intDesignationID)
                mvarDesignation = IIf(IsNull(Rec!vchDesignation), "", Rec!vchDesignation)
                mvarUserAddedDate = IIf(IsNull(Rec!dtUserAddedDate), "", Rec!dtUserAddedDate)
                mvarActive = IIf(IsNull(Rec!tnyActive), False, Rec!tnyActive)
                mvarAddFlag = IIf(IsNull(Rec!tinAdd), False, Rec!tinAdd)
                mvarEditFlag = IIf(IsNull(Rec!tinEdit), False, Rec!tinEdit)
                mvarDeleteFlag = IIf(IsNull(Rec!tinDelete), False, Rec!tinDelete)
                mvarViewFlag = IIf(IsNull(Rec!tnyView), False, Rec!tnyView)
                mvarPrintFlag = IIf(IsNull(Rec!tinPrint), False, Rec!tinPrint)
                mvarLocalBodyID = IIf(IsNull(Rec!intLBID), "", Rec!intLBID)
                SetUser = True
            End If
        End If
        Set mCon = Nothing
        Set Rec = Nothing
        
    End Function
    
    Public Sub CreateNewUser()
        Dim arrInput As Variant
        Dim objDb As New clsDB
        '----------------------------------'
        ' Only Administrator and SuperUser '
        ' can add Users                    '
        '----------------------------------'
        If gbUserTypeID > 1 Then
            Exit Sub
        End If
        '----------------------------------'
        If mUserCon.State = 0 Then
            objDb.SetConnection mUserCon
        End If
        arrInput = Array(IIf(mvarUserID > 0, mvarUserID, Null), _
                        mvarUserName, _
                        mvarLoginName, _
                        mvarPassWord, _
                        mvarUserTypeID, _
                        IIf(mvarAddFlag, 1, 0), _
                        IIf(mvarEditFlag, 1, 0), _
                        IIf(mvarDeleteFlag, 1, 0), _
                        IIf(mvarViewFlag, 1, 0), _
                        IIf(mvarPrintFlag, 1, 0), _
                        IIf(IsNull(mvarLocalBodyID), gbLocalBodyID, mvarLocalBodyID) _
                        )
        objDb.ExecuteSP "spSaveUser", arrInput
    End Sub
    
    Public Sub SetSuperUser()
        mvarUserID = 0
        mvarLoginName = "Saankhya"
        mvarUserName = "ikm"
        mvarPassWord = "saankhya"
        mvarUserTypeID = 0
        mvarUserType = "Super User"
        mvarAddFlag = True
        mvarEditFlag = True
        mvarDeleteFlag = True
        mvarViewFlag = True
        mvarPrintFlag = True
        If IsNumeric(gbLocalBodyID) Then
            mvarLocalBodyID = gbLocalBodyID
        End If
        gbUserID = mvarUserTypeID
        gbUserName = CStr(mvarUserName)
        gbUserTypeID = mvarUserTypeID
    
    End Sub
    
    Public Function GetUserIDFromSthapanaEmpID(ByVal EmpID As Variant)
        On Error GoTo Err:
            Dim mCnn As New ADODB.Connection
            Dim objDb As New clsDB
            Dim Rec As New ADODB.Recordset
            Dim mSQL As String
            
            If objDb.CreateNewConnection(mCnn, enuSourceString.DBMaster) Then
                mSQL = "Select numUserID from GM_User Where numEmpID = " & EmpID
                Rec.Open mSQL, mCnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    GetUserIDFromSthapanaEmpID = Rec!numUserID
                End If
                If Rec.State = 1 Then Rec.Close
            Else
                MsgBox "Connection To Master Does not exit, Please contact your System Administrator", vbInformation
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
    
    Public Property Let PrintFlag(ByVal vData As Boolean)
        mvarPrintFlag = vData
    End Property
    
    Public Property Get PrintFlag() As Boolean
        PrintFlag = mvarPrintFlag
    End Property
    
    Public Property Let ViewFlag(ByVal vData As Boolean)
        mvarViewFlag = vData
    End Property
    
    Public Property Get ViewFlag() As Boolean
        ViewFlag = mvarViewFlag
    End Property
    
    Public Property Let DeleteFlag(ByVal vData As Boolean)
        mvarDeleteFlag = vData
    End Property
    
    Public Property Get DeleteFlag() As Boolean
        DeleteFlag = mvarDeleteFlag
    End Property
    
    Public Property Let EditFlag(ByVal vData As Boolean)
        mvarEditFlag = vData
    End Property
    
    Public Property Get EditFlag() As Boolean
        EditFlag = mvarEditFlag
    End Property
    
    Public Property Let AddFlag(ByVal vData As Boolean)
        mvarAddFlag = vData
    End Property
    
    Public Property Get AddFlag() As Boolean
        AddFlag = mvarAddFlag
    End Property
    
    Public Property Let UserType(ByVal vData As Variant)
        mvarUserType = vData
    End Property
    
    Public Property Set UserType(ByVal vData As Variant)
        Set mvarUserType = vData
    End Property
        
    Public Property Get UserType() As Variant
        If IsObject(mvarUserType) Then
            Set UserType = mvarUserType
        Else
            UserType = mvarUserType
        End If
    End Property
    
    Public Property Let UserTypeID(ByVal vData As Variant)
        mvarUserTypeID = vData
    End Property
    
    Public Property Set UserTypeID(ByVal vData As Variant)
        Set mvarUserTypeID = vData
    End Property
    
    Public Property Get UserTypeID() As Variant
        If IsObject(mvarUserTypeID) Then
            Set UserTypeID = mvarUserTypeID
        Else
            UserTypeID = mvarUserTypeID
        End If
    End Property
    
    Public Property Let PassWord(ByVal vData As Variant)
        mvarPassWord = vData
        mvarPassWord = Encrypt(CStr(mvarPassWord))
    End Property
    
    Public Property Get PassWord()
        PassWord = Decrypt(CStr(mvarPassWord))
    End Property
    
    Public Property Let UserName(ByVal vData As Variant)
        mvarUserName = vData
    End Property
    
    Public Property Set UserName(ByVal vData As Variant)
        Set mvarUserName = vData
    End Property
    
    Public Property Get UserName() As Variant
        If IsObject(mvarUserName) Then
            Set UserName = mvarUserName
        Else
            UserName = mvarUserName
        End If
    End Property
    
    Public Property Let LoginName(ByVal vData As Variant)
        mvarLoginName = vData
    End Property
    
    Public Property Set LoginName(ByVal vData As Variant)
        Set mvarLoginName = vData
    End Property
    
    Public Property Get LoginName() As Variant
        If IsObject(mvarLoginName) Then
            Set LoginName = mvarLoginName
        Else
            LoginName = mvarLoginName
        End If
    End Property
    
    Public Property Let UserID(ByVal vData As Variant)
        mvarUserID = vData
    End Property
    
    Public Property Set UserID(ByVal vData As Variant)
        Set mvarUserID = vData
    End Property
    
    Public Property Get UserID() As Variant
        If IsObject(mvarUserID) Then
            Set UserID = mvarUserID
        Else
            UserID = mvarUserID
        End If
    End Property
    
    Public Sub LogOut()
        Dim mConn As New ADODB.Connection
        Dim objDb As New clsDB
        Dim strQry As String
        Dim SQL As String
        mConn.Open (objDb.GetConnectionString(enuSourceString.Saankhya))
        SQL = "Select max(intid) as no from faUserMovement where numUserID=" & gbUserID
        Dim rss As New ADODB.Recordset
        rss.Open SQL, mConn
        strQry = "Update faUserMovement Set dtLogoutTime=getdate() where intID=" & rss!No
        mConn.Execute (strQry)
        MsgBox "LogOut Successfully"
    End Sub
    
    Private Sub Class_Initialize()
        mvarUserID = Null
        mvarLoginName = Null
        mvarUserName = Null
        mvarPassWord = Null
        mvarUserTypeID = Null
        mvarUserType = Null
        
        mVarUserSeatID = Null
        mVarUserSeat = Null
        mVarSectionID = Null
        mVarSection = Null
        mVarSeatGroupID = Null
        
        mvarAddFlag = False
        mvarEditFlag = False
        mvarDeleteFlag = False
        mvarViewFlag = False
        mvarPrintFlag = False
    End Sub

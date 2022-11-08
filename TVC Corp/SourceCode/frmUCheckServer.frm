VERSION 5.00
Begin VB.Form frmUCheckServer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   405
      Left            =   1245
      TabIndex        =   3
      Top             =   1980
      Width           =   1950
   End
   Begin VB.TextBox txtMac 
      Height          =   285
      Left            =   1110
      TabIndex        =   2
      Top             =   945
      Width           =   2355
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Top             =   615
      Width           =   2355
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Check MS SQL Server"
      Height          =   405
      Left            =   1245
      TabIndex        =   0
      Top             =   1380
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   675
      TabIndex        =   5
      Top             =   975
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   870
      TabIndex        =   4
      Top             =   645
      Width           =   195
   End
End
Attribute VB_Name = "frmUCheckServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    Dim mArr() As String
    Dim mLoop As Integer
    Dim mDSNDetails As tDSNAttrib
    Dim mErrString As String
    Dim mCn As New ADODB.Connection

    Dim Rec As New ADODB.Recordset
    Dim objDb As New clsDB
    Dim mIPString As String
    Dim mFoundFlag As Boolean
    
    objDb.SetConnection mCn
    mCn.Execute "If Exists(Select * From sysObjects Where name = 'IKMSWD') Begin Drop Table IKMSWD  END ;"
    mCn.Execute "SELECT intLBID INTO dbo.IKMSWD FROM  faLBSettings"
    mCn.Close
    
    Call GetMeListOfIPs(mArr)
    Call DSNDelete("dsnF", "SQL Server", True)
     
    mFoundFlag = CheckIPArray(mArr)
    If mFoundFlag = False Then
        Dim mADOCnn As New ADODB.Connection
        Dim mArrOut As Variant
        mADOCnn.ConnectionString = "PROVIDER=MSDASQL; dsn=dsnFA;uid=sa;pwd=007;database=DB_Finance;"
        mADOCnn.Open
        Set Rec = objDb.ExecuteSP("spGetCon", , mArrOut, , mADOCnn, adCmdStoredProc)
        ReDim mArr(UBound(mArrOut, 2))
        For mLoop = 0 To UBound(mArrOut, 2)
            mArr(mLoop) = mArrOut(0, mLoop)
        Next
        mFoundFlag = CheckIPArray(mArr)
        mADOCnn.Close
    End If
    Call DSNDelete("dsnF", "SQL Server", True)
    
End Sub

Private Function TryConnect(ByRef mAdoConn As ADODB.Connection, IPStr As String) As Boolean
        Dim mDSNDetails As tDSNAttrib
        Dim mErrString As String
        
        On Error GoTo SkipErr1:
        With mDSNDetails
            .Database = "DB_Finance"
            .Driver = "SQL Server"
            .Server = IPStr
            .TrustedConnection = False              'True = Use NT authentication
            .PassWord = ""
            .UserID = "FAUser"
            .Dsn = "dsnF"
            .Description = "Saankhay KMAS"
            .Type = ServerBased
            .SystemDSN = True                       'Create a System DSN
        End With
        mErrString = DSNCreate(mDSNDetails)
        'If mCn.State = 1 Then mCn.Close
        mAdoConn.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnF;uid=sa;pwd=007;database=DB_Finance;"
        mAdoConn.Open
        TryConnect = True
        Exit Function
SkipErr1:
        TryConnect = False
End Function

Private Function CheckTbl(ByRef AdoCon As ADODB.Connection) As Boolean
    Dim Rec As New ADODB.Recordset
    'On Error GoTo SkipErr1
    'Rec.Open "Select * From sysObjects Where name = 'IKMSWD'", AdoCon, adOpenStatic, adLockReadOnly
    Rec.Open "Select * From IKMSWD", AdoCon, adOpenStatic, adLockReadOnly
    If Not (Rec.BOF And Rec.EOF) Then
        MsgBox "Found " & Rec.Fields(0)
        'Debug.Print mArr(mLoop)
        'txtIP.Text = mArr(mLoop)
        'txtMac.Text = GetMeMacAddressOf(mArr(mLoop))
        'AdoCon.Execute "Drop Table IKMSWD"
        CheckTbl = True
    Else
        CheckTable = False
    End If
    Rec.Close
    Exit Function
SkipErr1:
    CheckTable = False
End Function

Private Sub Command1_Click()

End Sub

Private Sub cmdTest_Click()
    Dim objDb As New clsDB
    Dim mCn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    
    
    Call CreateSA
    mCn.ConnectionString = "PROVIDER=MSDASQL; dsn=dsnFA;uid=sa;pwd=007;database=DB_Finance;"
    mCn.Open
    If mCn.State Then
        Dim mArrOut As Variant
        'Set Rec = objDb.ExecuteSP("spGetCon", , , , mCn, adCmdStoredProc)
        Set Rec = objDb.ExecuteSP("spGetCon", , , , mCn, adCmdStoredProc)
        'If UBound(mArrOut, 2) > 0 Then
        If Not (Rec.BOF And Rec.EOF) Then
            Debug.Print Rec.Fields(0)
            'mArrOut = Rec.GetRows
        Else
        End If
        Rec.Close
    End If
End Sub

Private Function CheckIPArray(ByRef mInputArray As Variant) As Boolean
    Dim mIPString As String
    Dim mADOCnn As New ADODB.Connection
    
    For mLoop = 0 To UBound(mInputArray)
        mIPString = mInputArray(mLoop)
        If TryConnect(mADOCnn, mIPString) Then
            'mADOCnn.Close
            'mADOCnn.ConnectionString = "PROVIDER=MSDASQL;dsn=dsnF;uid=sa;pwd=007;database=DB_Finance;"
            'mADOCnn.Open
            If CheckTbl(mADOCnn) Then
                MsgBox "Found"
                CheckIPArray = True
                'Exit Function
            Else
                Debug.Print "IKMSWD not found - " & mIPString
            End If
        Else
            Debug.Print "X Connection Failed - " & mIPString
        End If
        If mADOCnn.State = 1 Then mADOCnn.Close
    Next mLoop
    CheckIPArray = False
End Function

Private Function CreateSA() As Boolean
        Dim mDSNDetails As tDSNAttrib
        Dim mErrString As String
        
        'On Error GoTo SkipErr1:
        With mDSNDetails
            .Database = "DB_Finance"
            .Driver = "SQL Server"
            .Server = IPStr
            .TrustedConnection = False              'True = Use NT authentication
            .PassWord = ""
            .UserID = "sa"
            .Dsn = "dsnF"
            .Description = "Saankhay KMAS"
            .Type = ServerBased
            .SystemDSN = True                       'Create a System DSN
        End With
        mErrString = DSNCreate(mDSNDetails)

End Function

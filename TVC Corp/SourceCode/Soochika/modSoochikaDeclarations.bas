Attribute VB_Name = "modSoochikaDeclarations"
'Declare all soochika application specific variables in this module
'Insert a comment for each variable which reflects the porpose and default value ranges
Option Explicit
Public gbLBID                   As Integer
Public gbDistID                 As Integer
Public gFunSetDBConnection      As New ADODB.Connection
Public gbShortname              As Variant
Public gbnumZonalID             As Variant
Public gbSeat                   As Variant
Public gbSubID                  As Variant
Public gbnumSeatID              As Variant
Public gbnumUserId              As Variant
Public gbSuitID                 As Variant
Public gbDeptID                 As Variant
Public InwardMode               As Integer
Public gbSoochikaVer            As Integer '1 for old and 2 for unicode
Public gbSoochikaDBVer          As Variant
Public gbSoochikaScriptVer      As Variant
Public gbSoochikaUserLogID      As Variant
Public gbSevanaMainTypeID       As Variant
Public gbSaankhya               As Variant
'Integration flags
Public gbSevanaIntegration          As Variant  'SevanaCR
Public gbSanchayaIntegration        As Variant  'Sanchaya
Public gbSaankhyaIntegration        As Variant  'Saankhy Double
Public gbSevanaPensionIntegraton    As Variant  'SevanaPension

Public Sub SetSoochkaEnvironment()
'    Dim objUser As New clsUser
'    objUser.SetUser (gbUserID)
'    gbLBID = gbLocalBodyID
'    gbnumSeatID = gbSeatID
'    gbShortname = IIf(IsNull(objUser.UserName), "", objUser.UserName)
'    gbnumZonalID = gbLocationID
'    gbSeat = gbSeatName
'    gbSubID = 1
'    gbSuitID = 105
'    gbSoochikaVer = 5
    
    gbLBID = gbLocalBodyID
    gbDistID = gbDistID
    gbnumSeatID = gbSeatID ' 3020801071#
    gbnumUserId = gbUserID          '102080344
    gbShortname = gbUserName
    gbnumZonalID = gbnumZonalID
    gbSeat = gbSeatName
    gbSubID = 1
    gbSuitID = 105
    gbSevanaMainTypeID = 0
    If gbSoochikaVer = 5 Then
        GetSoochikaLocalBodyDetails
    End If
    gbSevanaIntegration = 0
    gbSanchayaIntegration = 0
    gbSevanaPensionIntegraton = 0
    gbSaankhyaIntegration = 0
End Sub

Private Sub GetSoochikaLocalBodyDetails()
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Soochika Connection failure ", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
    
    Set Rec = mCnn.Execute("spSelectLBDetails")
    If Not (Rec.EOF Or Rec.BOF) Then
        gbDistID = Rec!intDistrictID
'        gbSoochikaDBVer = "5.0.0.2"
'        gbSoochikaScriptVer = "5.0.0.2.4"
        gbSoochikaDBVer = "5.0.0.7" 'paperless
        gbSoochikaScriptVer = "5.0.0.7" 'paperless
    End If
    If (mCnn.State = 1) Then
        mCnn.Close
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMajorAccountHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mvarMajorAccountHeadID As Long
Private mvarMajorAccountHeadCode As String
Private mvarMajorAccountHead As String
Private mvarTypeID As Long
Public Property Let TypeID(ByVal vData As Long)
    mvarTypeID = vData
End Property

Public Property Get TypeID() As Long
    TypeID = mvarTypeID
End Property

Public Property Let MajorAccountHead(ByVal vData As String)
    mvarMajorAccountHead = vData
End Property

Public Property Get MajorAccountHead() As String
    MajorAccountHead = mvarMajorAccountHead
End Property

Public Property Let MajorAccountHeadCode(ByVal vData As String)
    mvarMajorAccountHeadCode = vData
End Property

Public Property Get MajorAccountHeadCode() As String
    MajorAccountHeadCode = mvarMajorAccountHeadCode
End Property

Public Property Let MajorAccountHeadID(ByVal vData As Long)
    mvarMajorAccountHeadID = vData
End Property

Public Property Get MajorAccountHeadID() As Long
    MajorAccountHeadID = mvarMajorAccountHeadID
End Property
Friend Sub subClearVariables()
     mvarTypeID = -1
     mvarMajorAccountHeadID = -1
     mvarMajorAccountHeadCode = ""
     mvarMajorAccountHead = ""
End Sub

Public Function SetMajorAccountHead(ByVal strMajorHeadCode As String) As Boolean
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
         mSQL = "SELECT * "
                    mSQL = mSQL & " FROM faMajorAccountHeads "
                    mSQL = mSQL & " Where faMajorAccountHeads.vchMajorAccountHeadCode = '" & strMajorHeadCode & "'"
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSQL)
        If Rec.RecordCount > 0 Then
                mvarTypeID = Rec!tinType
                mvarMajorAccountHeadID = Rec!intMajorAccountHeadID
                mvarMajorAccountHeadCode = Rec!vchMajorAccountHeadCode
                mvarMajorAccountHead = Rec!vchMajorAccountHead
        Else
            subClearVariables
        End If
End Function




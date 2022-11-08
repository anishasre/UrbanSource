VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMinorAccountHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Private mvarMinorAccountHeadID As Long
Private mvarMinorAccountHeadCode As String
Private mvarMinorAccountHead As String
Private mvarTypeID As Variant
Private mvarMajorAccountHeadID As Long
Private mvarMajorAccountHeadCode As String
Private mvarMajorAccountHead As String
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

Public Property Let TypeID(ByVal vData As Long)
    mvarTypeID = vData
End Property

Public Property Get TypeID() As Long
    TypeID = mvarTypeID
End Property

Public Property Let MinorAccountHead(ByVal vData As String)
    mvarMinorAccountHead = vData
End Property

Public Property Get MinorAccountHead() As String
    MinorAccountHead = mvarMinorAccountHead
End Property

Public Property Let MinorAccountHeadCode(ByVal vData As String)
    mvarMinorAccountHeadCode = vData
End Property

Public Property Get MinorAccountHeadCode() As String
    MinorAccountHeadCode = mvarMinorAccountHeadCode
End Property

Public Property Let MinorAccountHeadID(ByVal vData As Long)
    mvarMinorAccountHeadID = vData
End Property

Public Property Get MinorAccountHeadID() As Long
    MinorAccountHeadID = mvarMinorAccountHeadID
End Property

Friend Sub subClearVariables()
     mvarMinorAccountHeadID = -1
     mvarMinorAccountHeadCode = ""
     mvarMinorAccountHead = ""
     mvarTypeID = -1
     mvarMajorAccountHeadID = -1
     mvarMajorAccountHeadCode = ""
     mvarMajorAccountHead = ""
End Sub

Public Function SetMinorAccountHead(ByVal strMInorHeadCode As String) As Boolean
        Dim mCn As ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
         mSQL = "SELECT * "
                    mSQL = mSQL & " FROM faMinorAccountHeads LEFT JOIN "
                    mSQL = mSQL & " faMajorAccountHeads ON faMinorAccountHeads.intMajorAccountHeadID= faMajorAccountHeads.intMajorAccountHeadID "
                    mSQL = mSQL & " Where faMinorAccountHeads.vchMinorAccountHeadCode = '" & strMInorHeadCode & "'"
        Rec.CursorLocation = adUseClient
        Set Rec = GetRecordSet(mSQL)
        If Rec.RecordCount > 0 Then
                mvarMinorAccountHeadID = Rec!intMinorAccountHeadID
                mvarMinorAccountHeadCode = Rec!vchMinorAccountHeadCode
                mvarMinorAccountHead = Rec!vchMinorAccountHead
                mvarTypeID = Rec!tinType
                mvarMajorAccountHeadID = Rec!intMajorAccountHeadID
                mvarMajorAccountHeadCode = Rec!vchMajorAccountHeadCode
                mvarMajorAccountHead = Rec!vchMajorAccountHead
        Else
            subClearVariables
        End If
    End Function

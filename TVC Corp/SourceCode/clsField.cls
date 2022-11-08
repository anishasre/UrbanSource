VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
    
    Private mvarFieldID         As Variant
    Private mvarFieldCode       As Variant
    Private mvarFieldName       As Variant
    Private mvarLocalBodyID     As Variant
    Private mvarDistrictID      As Variant
    Private mvarBlockID         As Variant
    Private mvarTotalWards      As Variant
    Private mvarLocalBody       As Variant
    Private mvarDistrict        As Variant
    
    Public Property Let District(ByVal vData As Variant)
        mvarDistrict = vData
    End Property
       
    Public Property Set District(ByVal vData As Variant)
        Set mvarDistrict = vData
    End Property
        
    Public Property Get District() As Variant
        If IsObject(mvarDistrict) Then
            Set District = mvarDistrict
        Else
            District = mvarDistrict
        End If
    End Property
    
    Public Property Let LocalBody(ByVal vData As Variant)
        mvarLocalBody = vData
    End Property
    
    Public Property Set LocalBody(ByVal vData As Variant)
        Set mvarLocalBody = vData
    End Property
        
    Public Property Get LocalBody() As Variant
        If IsObject(mvarLocalBody) Then
            Set LocalBody = mvarLocalBody
        Else
            LocalBody = mvarLocalBody
        End If
    End Property
   
    Public Property Let TotalWards(ByVal vData As Variant)
        mvarTotalWards = vData
    End Property
        
    Public Property Set TotalWards(ByVal vData As Variant)
        Set mvarTotalWards = vData
    End Property
        
    Public Property Get TotalWards() As Variant
        If IsObject(mvarTotalWards) Then
            Set TotalWards = mvarTotalWards
        Else
            TotalWards = mvarTotalWards
        End If
    End Property
    
    Public Property Let BlockID(ByVal vData As Variant)
        mvarBlockID = vData
    End Property
        
    Public Property Set BlockID(ByVal vData As Variant)
        Set mvarBlockID = vData
    End Property
    
    Public Property Get BlockID() As Variant
        If IsObject(mvarBlockID) Then
            Set BlockID = mvarBlockID
        Else
            BlockID = mvarBlockID
        End If
    End Property
    
    Public Property Let DistrictID(ByVal vData As Variant)
        mvarDistrictID = vData
    End Property
    
    Public Property Set DistrictID(ByVal vData As Variant)
        Set mvarDistrictID = vData
    End Property
    
    Public Property Get DistrictID() As Variant
        If IsObject(mvarDistrictID) Then
            Set DistrictID = mvarDistrictID
        Else
            DistrictID = mvarDistrictID
        End If
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
    
    Public Property Let FieldName(ByVal vData As Variant)
        mvarFieldName = vData
    End Property
    
    Public Property Set FieldName(ByVal vData As Variant)
        Set mvarFieldName = vData
    End Property
    
    Public Property Get FieldName() As Variant
        If IsObject(mvarFieldName) Then
            Set FieldName = mvarFieldName
        Else
            FieldName = mvarFieldName
        End If
    End Property
    
    Public Property Let FieldCode(ByVal vData As Variant)
        mvarFieldCode = vData
    End Property
    
    Public Property Set FieldCode(ByVal vData As Variant)
        Set mvarFieldCode = vData
    End Property
    
    Public Property Get FieldCode() As Variant
        If IsObject(mvarFieldCode) Then
            Set FieldCode = mvarFieldCode
        Else
            FieldCode = mvarFieldCode
        End If
    End Property
    
    Public Property Let FieldID(ByVal vData As Variant)
        mvarFieldID = vData
    End Property
    
    Public Property Set FieldID(ByVal vData As Variant)
        Set mvarFieldID = vData
    End Property
    
    Public Property Get FieldID() As Variant
        If IsObject(mvarFieldID) Then
            Set FieldID = mvarFieldID
        Else
            FieldID = mvarFieldID
        End If
    End Property

    Private Sub ClearVariables()
        mvarFieldID = Null
        mvarFieldCode = Null
        mvarFieldName = Null
        mvarLocalBodyID = Null
        mvarDistrictID = Null
        mvarBlockID = Null
        mvarTotalWards = Null
        mvarLocalBody = Null
        mvarDistrict = Null
    End Sub
    
    Public Sub SetField(mID As Long)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mID > 0 Then
            mSQL = mSQL + " SELECT * FROM faFields "
            mSQL = mSQL + " WHERE faFields.intFieldID= " & mID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFieldID = Rec!intFieldID
                mvarFieldCode = Rec!vchFieldCode
                mvarFieldName = Rec!vchField
                mvarLocalBodyID = Rec!intLocalBodyID
            End If
            Rec.Close
        End If
    End Sub

Public Sub SetFieldByCode(mFieldCode As Variant)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        Call ClearVariables
        If mFieldCode <> "" Then
            mSQL = mSQL + " SELECT * from faFields "
            mSQL = mSQL + " WHERE vchFieldCode = '" & CStr(mFieldCode) & "'"
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFieldID = Rec!intFieldID
                mvarFieldCode = Rec!vchFieldCode
                mvarFieldName = Rec!vchField
                mvarLocalBodyID = Rec!intLocalBodyID
            End If
            Rec.Close
        End If
    End Sub

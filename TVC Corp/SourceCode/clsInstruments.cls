VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstruments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit

    'Local variable(s) to hold property value(s)
    Private mvarGroupID As Variant
    Private mvarInstrumentTypeID As Variant
    Private mvarInstrumentType As Variant
    
    Private Sub Class_Initialize()
        Call ClearVariables
    End Sub
    Private Sub ClearVariables()
        mvarGroupID = Null
        mvarInstrumentTypeID = Null
        mvarInstrumentType = Null
    End Sub
    Public Sub SetInstrumentType(mintInstrumentTypeID As Long)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        
        Call ClearVariables
        If mintInstrumentTypeID > 0 Then
            mSQL = "SELECT * FROM faInstrumentTypes WHERE intInstrumentTypeID = " & mintInstrumentTypeID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarInstrumentTypeID = Rec!intInstrumentTypeID
                mvarInstrumentType = Rec!vchinstrumenttype
                mvarGroupID = Rec!intGroupID
            End If
            Rec.Close
        End If
    End Sub
    Public Property Let InstrumentType(ByVal vData As Variant)
        mvarInstrumentType = vData
    End Property
    
    Public Property Set InstrumentType(ByVal vData As Variant)
        Set mvarInstrumentType = vData
    End Property
    
    Public Property Get InstrumentType() As Variant
        If IsObject(mvarInstrumentType) Then
            Set InstrumentType = mvarInstrumentType
        Else
            InstrumentType = mvarInstrumentType
        End If
    End Property
    
    Public Property Let InstrumentTypeID(ByVal vData As Variant)
        mvarInstrumentTypeID = vData
    End Property
    
    Public Property Set InstrumentTypeID(ByVal vData As Variant)
        Set mvarInstrumentTypeID = vData
    End Property
    
    Public Property Get InstrumentTypeID() As Variant
        If IsObject(mvarInstrumentTypeID) Then
            Set InstrumentTypeID = mvarInstrumentTypeID
        Else
            InstrumentTypeID = mvarInstrumentTypeID
        End If
    End Property
    
    Public Property Let GroupID(ByVal vData As Variant)
        mvarGroupID = vData
    End Property
    
    Public Property Set GroupID(ByVal vData As Variant)
        Set mvarGroupID = vData
    End Property
    
    Public Property Get GroupID() As Variant
        If IsObject(mvarGroupID) Then
            Set GroupID = mvarGroupID
        Else
            GroupID = mvarGroupID
        End If
    End Property
    
    

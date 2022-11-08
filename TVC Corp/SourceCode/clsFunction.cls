VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
    
    Private mvarFunctionID          As Variant
    Private mvarFunctionCode        As Variant
    Private mvarFunctionName        As Variant
    Private mvarMajorFunctionID     As Variant
    Private mvarMajorFunctionName   As Variant
    Private mvarMajorFunctionCode   As Variant

    Private Sub ClearVariables()
        mvarFunctionID = Null
        mvarFunctionCode = Null
        mvarFunctionName = Null
        mvarMajorFunctionID = Null
        mvarMajorFunctionName = Null
        mvarMajorFunctionCode = Null
    End Sub
    
    Public Sub SetFunction(mFunctionCode As Long)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mFunctionCode > 0 Then
            mSQL = mSQL + " SELECT faFunctions.*, faMajorFunctions.vchMajorFunction, faMajorFunctions.vchMajorFunctionCode FROM faFunctions "
            mSQL = mSQL + " LEFT JOIN faMajorFunctions ON faMajorFunctions.intMajorFunctionID = faFunctions.intMajorFunctionID "
            mSQL = mSQL + " WHERE faFunctions.vchFunctionCode = " & mFunctionCode
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFunctionID = Rec!intFunctionID
                mvarFunctionCode = Rec!vchFunctionCode
                mvarFunctionName = Rec!vchFunction
                mvarMajorFunctionID = Rec!intMajorFunctionID
                mvarMajorFunctionName = Rec!vchMajorFunction
                mvarMajorFunctionCode = Rec!vchMajorFunctionCode
            End If
            Rec.Close
        End If
    End Sub
    
    Public Sub SetFunctionByID(mFunctionID As Long)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mFunctionID > 0 Then
            mSQL = mSQL + " SELECT faFunctions.*, faMajorFunctions.vchMajorFunction, faMajorFunctions.vchMajorFunctionCode FROM faFunctions "
            mSQL = mSQL + " LEFT JOIN faMajorFunctions ON faMajorFunctions.intMajorFunctionID = faFunctions.intMajorFunctionID "
            mSQL = mSQL + " WHERE faFunctions.intFunctionID = " & mFunctionID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFunctionID = Rec!intFunctionID
                mvarFunctionCode = Rec!vchFunctionCode
                mvarFunctionName = Rec!vchFunction
                mvarMajorFunctionID = Rec!intMajorFunctionID
                mvarMajorFunctionName = Rec!vchMajorFunction
                mvarMajorFunctionCode = Rec!vchMajorFunctionCode
            End If
            Rec.Close
        End If
    End Sub
    
    Private Sub Class_Initialize()
        Call ClearVariables
    End Sub
    
    Public Property Let MajorFunctionCode(ByVal vData As Variant)
        mvarMajorFunctionCode = vData
    End Property

    Public Property Set MajorFunctionCode(ByVal vData As Variant)
        Set mvarMajorFunctionCode = vData
    End Property

    Public Property Get MajorFunctionCode() As Variant
        If IsObject(mvarMajorFunctionCode) Then
            Set MajorFunctionCode = mvarMajorFunctionCode
        Else
            MajorFunctionCode = mvarMajorFunctionCode
        End If
    End Property

    Public Property Let MajorFunctionName(ByVal vData As Variant)
        mvarMajorFunctionName = vData
    End Property

    Public Property Set MajorFunctionName(ByVal vData As Variant)
        Set mvarMajorFunctionName = vData
    End Property

    Public Property Get MajorFunctionName() As Variant
        If IsObject(mvarMajorFunctionName) Then
            Set MajorFunctionName = mvarMajorFunctionName
        Else
            MajorFunctionName = mvarMajorFunctionName
        End If
    End Property

    Public Property Let MajorFunctionID(ByVal vData As Variant)
        mvarMajorFunctionID = vData
    End Property

    Public Property Set MajorFunctionID(ByVal vData As Variant)
        Set mvarMajorFunctionID = vData
    End Property

    Public Property Get MajorFunctionID() As Variant
        If IsObject(mvarMajorFunctionID) Then
            Set MajorFunctionID = mvarMajorFunctionID
        Else
            MajorFunctionID = mvarMajorFunctionID
        End If
    End Property

    Public Property Let FunctionName(ByVal vData As Variant)
        mvarFunctionName = vData
    End Property

    Public Property Set FunctionName(ByVal vData As Variant)
        Set mvarFunctionName = vData
    End Property

    Public Property Get FunctionName() As Variant
        If IsObject(mvarFunctionName) Then
            Set FunctionName = mvarFunctionName
        Else
            FunctionName = mvarFunctionName
        End If
    End Property

    Public Property Let FunctionCode(ByVal vData As Variant)
        mvarFunctionCode = vData
    End Property

    Public Property Set FunctionCode(ByVal vData As Variant)
        Set mvarFunctionCode = vData
    End Property

    Public Property Get FunctionCode() As Variant
        If IsObject(mvarFunctionCode) Then
            Set FunctionCode = mvarFunctionCode
        Else
            FunctionCode = mvarFunctionCode
        End If
    End Property
    
    Public Property Let FunctionID(ByVal vData As Variant)
        mvarFunctionID = vData
    End Property

    Public Property Set FunctionID(ByVal vData As Variant)
        Set mvarFunctionID = vData
    End Property

    Public Property Get FunctionID() As Variant
        If IsObject(mvarFunctionID) Then
            Set FunctionID = mvarFunctionID
        Else
            FunctionID = mvarFunctionID
        End If
    End Property

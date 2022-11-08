VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit
    Private mvarFunctionaryID As Variant        'local copy
    Private mvarFunctionaryName As Variant      'local copy
    Private mvarFunctionaryCode As Variant      'local copy
    Private mvarMajorFunctionaryID As Variant   'local copy
    Private mvarMajorFunctionaryCode As Variant 'local copy
    Private mvarMajorFunctionaryName As Variant 'local copy
    
    Public Property Let MajorFunctionaryName(ByVal vData As Variant)
        mvarMajorFunctionaryName = vData
    End Property

    Public Property Set MajorFunctionaryName(ByVal vData As Variant)
        Set mvarMajorFunctionaryName = vData
    End Property

    Public Property Get MajorFunctionaryName() As Variant
        If IsObject(mvarMajorFunctionaryName) Then
            Set MajorFunctionaryName = mvarMajorFunctionaryName
        Else
            MajorFunctionaryName = mvarMajorFunctionaryName
        End If
    End Property

    Public Property Let MajorFunctionaryCode(ByVal vData As Variant)
        mvarMajorFunctionaryCode = vData
    End Property

    Public Property Set MajorFunctionaryCode(ByVal vData As Variant)
        Set mvarMajorFunctionaryCode = vData
    End Property

    Public Property Get MajorFunctionaryCode() As Variant
        If IsObject(mvarMajorFunctionaryCode) Then
            Set MajorFunctionaryCode = mvarMajorFunctionaryCode
        Else
            MajorFunctionaryCode = mvarMajorFunctionaryCode
        End If
    End Property
    
    Public Property Let MajorFunctionaryID(ByVal vData As Variant)
        mvarMajorFunctionaryID = vData
    End Property
    
    Public Property Set MajorFunctionaryID(ByVal vData As Variant)
        Set mvarMajorFunctionaryID = vData
    End Property
        
    Public Property Get MajorFunctionaryID() As Variant
        If IsObject(mvarMajorFunctionaryID) Then
            Set MajorFunctionaryID = mvarMajorFunctionaryID
        Else
            MajorFunctionaryID = mvarMajorFunctionaryID
        End If
    End Property
    
    Public Property Let FunctionaryCode(ByVal vData As Variant)
        mvarFunctionaryCode = vData
    End Property
        
    Public Property Set FunctionaryCode(ByVal vData As Variant)
        Set mvarFunctionaryCode = vData
    End Property
        
    Public Property Get FunctionaryCode() As Variant
        If IsObject(mvarFunctionaryCode) Then
            Set FunctionaryCode = mvarFunctionaryCode
        Else
            FunctionaryCode = mvarFunctionaryCode
        End If
    End Property
          
    Public Property Let FunctionaryName(ByVal vData As Variant)
        mvarFunctionaryName = vData
    End Property
    
    Public Property Set FunctionaryName(ByVal vData As Variant)
        Set mvarFunctionaryName = vData
    End Property
    
    Public Property Get FunctionaryName() As Variant
        If IsObject(mvarFunctionaryName) Then
            Set FunctionaryName = mvarFunctionaryName
        Else
            FunctionaryName = mvarFunctionaryName
        End If
    End Property
           
    Public Property Let FunctionaryID(ByVal vData As Variant)
        mvarFunctionaryID = vData
    End Property
        
    Public Property Set FunctionaryID(ByVal vData As Variant)
        Set mvarFunctionaryID = vData
    End Property
        
    Public Property Get FunctionaryID() As Variant
        If IsObject(mvarFunctionaryID) Then
            Set FunctionaryID = mvarFunctionaryID
        Else
            FunctionaryID = mvarFunctionaryID
        End If
    End Property

    Private Sub ClearVariables()
        mvarFunctionaryID = Null
        mvarFunctionaryName = Null
        mvarFunctionaryCode = Null
        mvarMajorFunctionaryID = Null
        mvarMajorFunctionaryCode = Null
        mvarMajorFunctionaryName = Null
    End Sub
    
    Public Sub SetFunctionary(mFunctionaryCode As String)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mFunctionaryCode > 0 Then
            mSQL = mSQL + " SELECT faFunctionaries.*, faMajorFunctionaries.intDepartmentID, faMajorFunctionaries.vchDepartment, faMajorFunctionaries.vchDepartmentCode FROM faFunctionaries "
            mSQL = mSQL + " LEFT JOIN faMajorFunctionaries ON faMajorFunctionaries.intDepartmentID= faFunctionaries.intMajorFunctionaryID "
            mSQL = mSQL + " WHERE faFunctionaries.vchFunctionaryCode = '" & mFunctionaryCode & "'"
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFunctionaryID = Rec!intFunctionaryID
                mvarFunctionaryName = Rec!vchFunctionary
                mvarFunctionaryCode = Rec!vchFunctionaryCode
                mvarMajorFunctionaryID = Rec!intDepartmentID
                mvarMajorFunctionaryCode = Rec!vchDepartmentCode
                mvarMajorFunctionaryName = Rec!vchDepartment
            End If
            Rec.Close
        End If
    End Sub
    
    Public Sub SetFunctionaryByID(mFunctionaryID As Long)
        Dim objDB As New clsDB
        Dim Rec As New ADODB.Recordset
        Dim mSQL As String
        Call ClearVariables
        If mFunctionaryID > 0 Then
            mSQL = mSQL + " SELECT faFunctionaries.*, faMajorFunctionaries.intDepartmentID, faMajorFunctionaries.vchDepartment, faMajorFunctionaries.vchDepartmentCode FROM faFunctionaries "
            mSQL = mSQL + " LEFT JOIN faMajorFunctionaries ON faMajorFunctionaries.intDepartmentID= faFunctionaries.intMajorFunctionaryID "
            mSQL = mSQL + " WHERE faFunctionaries.intFunctionaryID = " & mFunctionaryID
            Set Rec = GetRecordSet(mSQL)
            If Not (Rec.EOF And Rec.BOF) Then
                mvarFunctionaryID = Rec!intFunctionaryID
                mvarFunctionaryName = Rec!vchFunctionary
                mvarFunctionaryCode = Rec!vchFunctionaryCode
                mvarMajorFunctionaryID = Rec!intDepartmentID
                mvarMajorFunctionaryCode = Rec!vchDepartmentCode
                mvarMajorFunctionaryName = Rec!vchDepartment
            End If
            Rec.Close
        End If
    End Sub
    
    Public Function GetFunctionaryList() As ADODB.Recordset
        Dim mSQL As String
        mSQL = "Select * From faFunctionaries"
        Set GetFunctionaryList = GetRecordSet(mSQL)
    End Function
    
    Private Sub Class_Initialize()
        Call ClearVariables
    End Sub

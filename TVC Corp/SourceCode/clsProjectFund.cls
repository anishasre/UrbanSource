VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
    Private intSourceOfFundID           As Variant
    Private strSourceCode               As Variant
    Private strSourceName               As Variant
    Private intAllocatedYearID          As Variant
    Private intSlNo                     As Variant
    Private fltSourceWiseAmount         As Variant
    Private fltSourceWiseUtilisedAmount As Variant
    
    Public Property Let SourceWiseUtilisedAmount(ByVal mData As Variant)
        fltSourceWiseUtilisedAmount = mData
    End Property
    Public Property Get SourceWiseUtilisedAmount() As Variant
        SourceWiseUtilisedAmount = fltSourceWiseUtilisedAmount
    End Property
    
    Public Property Let SourceWiseAmount(ByVal mData As Variant)
        fltSourceWiseAmount = mData
    End Property
    Public Property Get SourceWiseAmount() As Variant
        SourceWiseAmount = fltSourceWiseAmount
    End Property
    
    Public Property Let SlNo(ByVal mData As Variant)
        intSlNo = mData
    End Property
    Public Property Get SlNo() As Variant
        SlNo = intSlNo
    End Property
    
    Public Property Let AllocatedYearID(ByVal mData As Variant)
        intAllocatedYearID = mData
    End Property
    Public Property Get AllocatedYearID() As Variant
        AllocatedYearID = intAllocatedYearID
    End Property
    
    Public Property Let SourceName(ByVal mData As Variant)
        strSourceName = mData
    End Property
    Public Property Get SourceName() As Variant
        SourceName = strSourceName
    End Property
    
    Public Property Let SourceCode(ByVal mData As Variant)
        strSourceCode = mData
    End Property
    Public Property Get SourceCode() As Variant
        SourceCode = strSourceCode
    End Property
    
    Public Property Let SourceOfFundID(ByVal mData As Variant)
        intSourceOfFundID = mData
    End Property
    Public Property Get SourceOfFundID() As Variant
        SourceOfFundID = intSourceOfFundID
    End Property



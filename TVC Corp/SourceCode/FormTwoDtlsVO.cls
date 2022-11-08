VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormTwoDtlsVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit

    'id Long Unique id/Building Id
    'doorNo String door no
    'doorNo2 String second part of doorNo
    'zoneNo String zone no/Id
    'wardNo String ward no/Id
    'applicantName String owner name
    'ownerAddress String owneraddress
    'ownerPin String owner pin
    'applicationNo String Application No
    'statusText String current application Status
    'accountHeadCurrent String Acount head­Receivable for
    'accountHeadArrears String Acount head­Receivable for
    'accountHeadFine String Acount head­Fines
    'accountHeadLiberyCess String Acount head­Library Cess Payable
    'taxDetailsVO TaxDetailsVO TaxDetailsVO


    Private m_id As Long
    Private m_doorNo As String
    Private m_doorNo2 As Variant
    Private m_zoneNo As Variant
    Private m_wardNo As Variant
    Private m_applicantName As Variant
    Private m_ownerAddress As Variant
    Private m_ownerPin As Variant
    Private m_applicationNo As Variant
    Private m_statusText As Variant
    Private m_accountHeadCurrent As Variant
    Private m_accountHeadFine As Variant
    Private m_accountHeadLiberyCess As Variant
    Private m_taxDetailsVO As Variant


'    Public Property Let xxxx(ByVal vData As Variant)
'        m_ = vData
'    End Property
'    Public Property Get xxx()
'        xxx = m_
'    End Property
    
    
    
    Public Property Let id(ByVal vData As Variant)
        m_id = vData
    End Property
    Public Property Get id()
        id = m_id
    End Property



    Public Property Let doorNo(ByVal vData As Variant)
        m_doorNo = vData
    End Property
    Public Property Get doorNo()
        doorNo = m_doorNo
    End Property


    Public Property Let doorNo2(ByVal vData As Variant)
        m_doorNo = vData
    End Property
    Public Property Get doorNo2()
        doorNo = m_doorNo
    End Property
        
    
    Public Property Let zoneNo(ByVal vData As Variant)
        m_zoneNo = vData
    End Property
    Public Property Get zoneNo()
        zoneNo = m_zoneNo
    End Property
    
        
    Public Property Let wardNo(ByVal vData As Variant)
        m_wardNo = vData
    End Property
    Public Property Get wardNo()
        wardNo = m_wardNo
    End Property
        
    
    Public Property Let applicantName(ByVal vData As Variant)
        m_applicantName = vData
    End Property
    Public Property Get applicantName()
        applicantName = m_applicantName
    End Property
    
    
    Public Property Let ownerAddress(ByVal vData As Variant)
        m_ownerAddress = vData
    End Property
    Public Property Get ownerAddress()
        ownerAddress = m_ownerAddress
    End Property
    
    
    Public Property Let ownerPin(ByVal vData As Variant)
        m_ownerPin = vData
    End Property
    Public Property Get ownerPin()
        ownerPin = m_ownerPin
    End Property
    
    
    
    
    

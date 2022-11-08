Attribute VB_Name = "modMacAddress"

Option Explicit

' BEGIN - GetIpNetTable
Private Const MAXLEN_PHYSADDR = 8

Private Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr As Long
    dwType As Long
End Type
' END - GetIpNetTable


'---------------------------------------------------------------------------
' Used to get the MAC address.
'---------------------------------------------------------------------------


Private Const NCBNAMSZ As Long = 16
Private Const NCBENUM As Long = &H37
Private Const NCBRESET As Long = &H32
Private Const NCBASTAT As Long = &H33
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

'-------------------------------- '
' FOR REMOTEMAC                   '
'-------------------------------- '
    Private Const No_ERROR = 0
'--------------------------------


Private Type NET_CONTROL_BLOCK  'NCB
    ncb_command    As Byte
    ncb_retcode    As Byte
    ncb_lsn        As Byte
    ncb_num        As Byte
    ncb_buffer     As Long
    ncb_length     As Integer
    ncb_callname   As String * NCBNAMSZ
    ncb_name       As String * NCBNAMSZ
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_post       As Long
    ncb_lana_num   As Byte
    ncb_cmd_cplt   As Byte
    ncb_reserve(9) As Byte 'Reserved, must be 0
    ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Private Type NAME_BUFFER
    name_(0 To NCBNAMSZ - 1) As Byte
    name_num                 As Byte
    name_flags               As Byte
End Type

Private Type ASTAT
    adapt             As ADAPTER_STATUS
    NameBuff(0 To 29) As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
        (pncb As NET_CONTROL_BLOCK) As Byte

Private Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, ByVal _
        hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        lpMem As Any) As Long


' ------------------------------------------------------------------------------------ '
' To Get Remote MAC Address
' ------------------------------------------------------------------------------------ '
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function SendARP Lib "Iphlpapi.dll" (ByVal DestIp As Long, ByVal ScrIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
Private Declare Function GetIpNetTable Lib "Iphlpapi.dll" (pIpNetTable As Byte, pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemoryForIB Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function GetMacAddress() As String
    Dim l As Long
    Dim lngError As Long
    Dim lngSize As Long
    Dim pAdapt As Long
    Dim pAddrStr As Long
    Dim pASTAT As Long
    Dim strTemp As String
    Dim strAddress As String
    Dim strMACAddress As String
    Dim AST As ASTAT
    Dim NCB As NET_CONTROL_BLOCK

    '---------------------------------------------------------------------------
    ' Get the network interface card's MAC address.
    '----------------------------------------------------------------------------
    '
    On Error GoTo ErrorHandler
    GetMacAddress = ""
    strMACAddress = ""

    '
    ' Try to get MAC address from NetBios. Requires NetBios installed.
    '
    ' Supported on 95, 98, ME, NT, 2K, XP
    '
    ' Results Connected Disconnected
    ' ------- --------- ------------
    '   XP       OK         Fail (Fail after reboot)
    '   NT       OK         OK   (OK after reboot)
    '   98       OK         OK   (OK after reboot)
    '   95       OK         OK   (OK after reboot)
    '
    NCB.ncb_command = NCBRESET
    Call Netbios(NCB)

    NCB.ncb_callname = "*               "
    NCB.ncb_command = NCBASTAT
    NCB.ncb_lana_num = 0
    NCB.ncb_length = Len(AST)

    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or _
                       HEAP_ZERO_MEMORY, NCB.ncb_length)
    If pASTAT = 0 Then GoTo ErrorHandler

    NCB.ncb_buffer = pASTAT
    Call Netbios(NCB)

    Call CopyMemory(AST, NCB.ncb_buffer, Len(AST))

    strMACAddress = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)

    Call HeapFree(GetProcessHeap(), 0, pASTAT)

    GetMacAddress = strMACAddress
    GoTo NormalExit

ErrorHandler:
    Call MsgBox(err.Description, vbCritical, "Error")

NormalExit:
    End Function

'
'Private Sub Form_Load()
'Dim strMACAddress As String
'
'    strMACAddress = fGetMacAddress()
'
'    If strMACAddress <> "" Then
'        Call MsgBox(strMACAddress, vbInformation, "MAC Address")
'    End If
'
'End Sub


'============================================================================================================================== '
' TO Get Remote MAC address
'============================================================================================================================== '

''Option Explicit
''
''Private Const No_ERROR = 0
''
''Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
''Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIp As Long, ByVal ScrIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
''Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
''
''
''Public Function GetRemoteMACAddress(ByVal sRemoteIP As String, sRemoteMacAddress As String) As Boolean
''
''    Dim dwRemoteIp As Long
''    Dim pMacAddr As Long
''    Dim bpMacAddr As Byte
''    Dim PhyAddrLen As Long
''    Dim Cnt As Long
''    Dim tmp As String
''
''    dwRemoteIp = inet_addr(sRemoteIP)
''    If dwRemoteIp <> 0 Then
''        PhyAddrLen = 6
''        If SendARP(dwRemoteIp, 0&, pMacAddr, PhyAddrLen) = No_ERROR Then
''            If pMacAddr <> 0 And PhyAddrLen <> 0 Then
''                ReDim bpMacAddr(0 To PhyAddrLen - 1)
''                CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
''                For Cnt = 0 To PhyAddrLen - 1
''                    If bpMacAddr(Cnt) = 0 Then
''                        tmp = tmp & "00-"
''                    Else
''                        tmp = tmp & Hex$(bpMacAddr(Cnt)) & "-"
''                    End If
''                Next
''                If Len(tmp) > 0 Then
''                    sRemoteMacAddress = Left$(tmp, Len(tmp) - 1)
''                    GetRemoteMACAddress = True
''                End If
''
''                Exit Function
''            Else
''                GetRemoteMACAddress = False
''            End If
''        Else
''            GetRemoteMACAddress = False
''        End If
''    Else
''        GetRemoteMACAddress = False
''    End If
''End Function


''
''Private Sub btnGetMac_Click()
''Dim sRemoteMacAddress As String
''If Len(txtIpAddress.Text) > 0 Then
''    If GetRemoteMACAddress(txtIpAddress.Text, sRemoteMacAddress) Then
''        lblMacAddress.Caption = sRemoteMacAddress
''    Else
''        lblMacAddress.Caption = "(SendARP call failed)"
''    End If
''End If
''End Sub
''

Private Function GetMyIpNetTable() As String
    Dim Listing() As MIB_IPNETROW
    Dim Ret As Long
    Dim Cnt As Long
    Dim bBytes() As Byte
    Dim bTemp(0 To 3) As Byte
    
    'Call the function to retrieve how many bytes are needed
    GetIpNetTable ByVal 0&, Ret, False
    
    'If it failed, exit the sub
    If Ret <= 0 Then Exit Function
    
    'Redimension Our Buffer
    ReDim bBytes(0 To Ret - 1) As Byte
    
    'Retireve The Data
    GetIpNetTable bBytes(0), Ret, False
    
    'Copy the number of entries to the 'Ret' variable
    CopyMemory Ret, bBytes(0), 4
    
    'Redimension The Listing
    If Ret > 0 Then ReDim Listing(0 To Ret - 1) As MIB_IPNETROW
    
    'Print Data
    Debug.Print "Contents of address mapping table (items: " + CStr(Ret) + ")"
    For Cnt = 0 To Ret - 1
        CopyMemory Listing(Cnt), bBytes(4 + 24 * Cnt), 24
        CopyMemory bTemp(0), Listing(Cnt).dwAddr, 4
        Debug.Print " Item " + CStr(Listing(Cnt).dwIndex)
        Debug.Print " address " + ConvertAddressToString(bTemp(), 4)
        Debug.Print " physical address " + ConvertAddressToString(Listing(Cnt).bPhysAddr, Listing(Cnt).dwPhysAddrLen)
        Select Case Listing(Cnt).dwType
            Case 4 'Static
            Debug.Print " type: Static"
            Case 3 'Dynamic
            Debug.Print " type: Dynamic"
            Case 2 'Invalid
            Debug.Print " type: Invalid"
            Case 1 'Other
            Debug.Print " type: Other"
        End Select
    Next Cnt
End Function

Public Function GetMeMacAddressOf(ByVal mIP As String) As String
    Dim Listing() As MIB_IPNETROW
    Dim Ret As Long
    Dim Cnt As Long
    Dim bBytes() As Byte
    Dim bTemp(0 To 3) As Byte
    
    'Call the function to retrieve how many bytes are needed
    GetIpNetTable ByVal 0&, Ret, False
    
    'If it failed, exit the sub
    If Ret <= 0 Then Exit Function
    
    ' Redimension Our Buffer
    ReDim bBytes(0 To Ret - 1) As Byte
    
    'Retireve The Data
    GetIpNetTable bBytes(0), Ret, False
    
    'Copy the number of entries to the 'Ret' variable
    'CopyMemory Ret, bBytes(0), 4
    CopyMemoryForIB Ret, bBytes(0), 4
    
    'Redimension The Listing
    If Ret > 0 Then ReDim Listing(0 To Ret - 1) As MIB_IPNETROW
    
    For Cnt = 0 To Ret - 1
        CopyMemoryForIB Listing(Cnt), bBytes(4 + 24 * Cnt), 24
        CopyMemoryForIB bTemp(0), Listing(Cnt).dwAddr, 4
        If mIP = ConvertAddressToString(bTemp(), 4) Then ' IF IP Matching Then Read MacAddress
            GetMeMacAddressOf = ConvertAddressToHexString(Listing(Cnt).bPhysAddr, Listing(Cnt).dwPhysAddrLen)
        End If
    Next Cnt
End Function

'Converts a byte array to a string
Private Function ConvertAddressToString(bArray() As Byte, lLength As Long) As String
    Dim Cnt As Long
    For Cnt = 0 To lLength - 1
        ConvertAddressToString = ConvertAddressToString + CStr(bArray(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function ConvertAddressToHexString(bArray() As Byte, lLength As Long) As String
    Dim Cnt As Long
    Dim mStr As String
    For Cnt = 0 To lLength - 1
        mStr = Hex$(bArray(Cnt))
        If Len(mStr) = 1 And IsNumeric(mStr) Then
            mStr = "0" & mStr
        End If
        ConvertAddressToHexString = ConvertAddressToHexString + mStr + "."
    Next Cnt
    ConvertAddressToHexString = Left$(ConvertAddressToHexString, Len(ConvertAddressToHexString) - 1)
End Function


Public Sub GetMeListOfIPs(ByRef ArrIPs)
    Dim Listing() As MIB_IPNETROW
    Dim Ret As Long
    Dim Cnt As Long
    Dim bBytes() As Byte
    Dim bTemp(0 To 3) As Byte
    
    'Call the function to retrieve how many bytes are needed
    GetIpNetTable ByVal 0&, Ret, False
    
    'If it failed, exit the sub
    If Ret <= 0 Then Exit Sub
    
    ' Redimension Our Buffer
    ReDim bBytes(0 To Ret - 1) As Byte
    
    'Retireve The Data
    GetIpNetTable bBytes(0), Ret, False
    
    'Copy the number of entries to the 'Ret' variable
    'CopyMemory Ret, bBytes(0), 4
    CopyMemoryForIB Ret, bBytes(0), 4
    
    'Redimension The Listing
    If Ret > 0 Then ReDim Listing(0 To Ret - 1) As MIB_IPNETROW
    ReDim ArrIPs(Ret - 1)
    For Cnt = 0 To Ret - 1
        CopyMemoryForIB Listing(Cnt), bBytes(4 + 24 * Cnt), 24
        CopyMemoryForIB bTemp(0), Listing(Cnt).dwAddr, 4
        ArrIPs(Cnt) = ConvertAddressToString(bTemp(), 4)
    Next Cnt
End Sub

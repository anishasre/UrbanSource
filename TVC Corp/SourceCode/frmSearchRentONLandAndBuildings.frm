VERSION 5.00
Begin VB.Form frmSearchRentONLandAndBuildings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstItemData 
      Height          =   255
      Left            =   8175
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   780
      TabIndex        =   2
      Top             =   5100
      Width           =   5070
   End
   Begin VB.ListBox lstMasters 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   8820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   5130
      Width           =   510
   End
End
Attribute VB_Name = "frmSearchRentONLandAndBuildings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Private varnumWardID As Variant
    Private varintZoneID As Variant
    Private varnumAssetID As Variant
    Private varShowItem  As Variant
    Private Sub FillAssetMaster()
        '---------------------------------------------'
        ' Fills Assets From Sanchaya                  '
        ' Table Name : snRentAssetDetails             '
        '---------------------------------------------'
        Dim mSQL As String
        Dim mWhere As String
        mWhere = " WHERE "
        If Not IsEmpty(varnumWardID) Then
            mWhere = mWhere + " numWardID = " & varnumWardID
        ElseIf Not IsEmpty(varintZoneID) Then
            mWhere = mWhere + " numZoneID = " & varintZoneID
        End If
        If mWhere = " WHERE " Then mWhere = ""
        mSQL = "Select chvAssetName From snRentAssetDetails " & mWhere & " Order By chvAssetName"
        PopulateList lstMasters, mSQL, , , True, , enuSourceString.Sanchaya
        mSQL = "Select numRegNo From snRentAssetDetails " & mWhere & " Order By chvAssetName"
        PopulateList lstItemData, mSQL, , , True, , enuSourceString.Sanchaya
    End Sub
    Private Sub FillShops()
        '---------------------------------------------'
        ' Fills Shop Names From Sanchaya (Rent of L&B '
        ' Table Name : snRentDeedDetails1             '
        '---------------------------------------------'
        Dim mSQL As String
        Dim mWhere As String
        mWhere = " WHERE "
        If Not IsEmpty(varnumAssetID) Then
            mWhere = mWhere + " numMasterID = " & varnumAssetID
        ElseIf Not IsEmpty(varnumWardID) Then
            mWhere = mWhere + " numWardID = " & varnumWardID
        ElseIf Not IsEmpty(varintZoneID) Then
            'If Len(mWhere) > 7 Then mWhere = mWhere + " AND "
            mWhere = mWhere + " numZoneID = " & varintZoneID
        End If
        If mWhere = " WHERE " Then mWhere = ""
        mSQL = "Select chvShopName From snRentDeedDetails1 " & mWhere & " Order By chvShopName"
        PopulateList lstMasters, mSQL, , , True, , enuSourceString.Sanchaya
        mSQL = "Select numDeedRegNo From snRentDeedDetails1 " & mWhere & " Order By chvShopName"
        PopulateList lstItemData, mSQL, , , True, , enuSourceString.Sanchaya
    End Sub
    Private Sub FillLessee()
        '--------------------------------------------------'
        ' Fills Name of Lessee From Sanchaya (Rent of L&B) '
        ' Table Name : snRentDeedDetails1                  '
        '--------------------------------------------------'
        Dim mSQL As String
        Dim mWhere As String
        mWhere = " WHERE "
        If Not IsEmpty(varnumAssetID) Then
            mWhere = mWhere + " numMasterID = " & varnumAssetID
        ElseIf Not IsEmpty(varnumWardID) Then
            mWhere = mWhere + " numWardID = " & varnumWardID
        ElseIf Not IsEmpty(varintZoneID) Then
            'If Len(mWhere) > 7 Then mWhere = mWhere + " AND "
            mWhere = mWhere + " numZoneID = " & varintZoneID
        End If
        If mWhere = " WHERE " Then mWhere = ""
        mSQL = "Select chvLicenceeName From snRentDeedDetails1 " & mWhere & " Order By chvLicenceeName"
        PopulateList lstMasters, mSQL, , , True, , enuSourceString.Sanchaya
        mSQL = "Select numDeedRegNo From snRentDeedDetails1 " & mWhere & " Order By chvLicenceeName"
        PopulateList lstItemData, mSQL, , , True, , enuSourceString.Sanchaya
    End Sub
    
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
        End If
    End Sub
    Private Sub Form_Load()
        Select Case varShowItem
            Case Is = 1
                Me.Caption = " List of Assets ( Rent on Land & Buildings )"
                Call FillAssetMaster
            Case Is = 2
                Me.Caption = " List of Shops ( Rent on Land & Buildings )"
                Call FillShops
            Case Is = 3
                Me.Caption = " List of Name of Lessees ( Rent on Land & Buildings )"
                Call FillLessee
        End Select
    End Sub
    Private Sub lstMasters_DblClick()
        Dim mIndex As Integer
        mIndex = lstMasters.ListIndex
        If Len(lstMasters.Text) Then
            gbSearchStr = lstMasters.List(mIndex)
            gbSearchCode = lstItemData.List(mIndex)
            Unload Me
        End If
    End Sub
    Property Let WardID(Data As Variant)
        varnumWardID = Data
    End Property
    Property Let ZoneID(Data As Variant)
        varintZoneID = Data
    End Property
    Property Let ShowItem(Data As Variant)
        varShowItem = Data
    End Property
    Property Let AssetID(Data As Variant)
        varnumAssetID = Data
    End Property

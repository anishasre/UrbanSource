VERSION 5.00
Begin VB.Form frmLocalBodySettings 
   Caption         =   "Local Body Settings"
   ClientHeight    =   5160
   ClientLeft      =   1035
   ClientTop       =   2190
   ClientWidth     =   9825
   Icon            =   "frmLocalBodySettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   9825
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   2040
      Width           =   3015
   End
   Begin VB.ComboBox cmbLocationType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmLocalBodySettings.frx":1042
      Left            =   6600
      List            =   "frmLocalBodySettings.frx":1049
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3045
   End
   Begin VB.ComboBox cmbLBtype 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3045
   End
   Begin VB.ComboBox cmbDistrict 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmLocalBodySettings.frx":105A
      Left            =   6600
      List            =   "frmLocalBodySettings.frx":105C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   3045
   End
   Begin VB.ComboBox cmbLocalBody 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1110
      Width           =   3045
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      TabIndex        =   5
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdSetLB 
      Caption         =   "Set LB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tiltle"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "URL Address"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Location type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Body"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   1200
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4965
      TabIndex        =   8
      Top             =   195
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Body Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5055
      TabIndex        =   7
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label lblLBCode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7635
      TabIndex        =   6
      Top             =   1725
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   90
      Picture         =   "frmLocalBodySettings.frx":105E
      Top             =   630
      Width           =   4245
   End
End
Attribute VB_Name = "frmLocalBodySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim arrIn As Variant
         
    Private Sub FillbLBtype()
       Call PopulateList(cmbLBtype, "Select chvTypeDescEnglish,tnyLBTypeID From GM_LocalBodyType Order By tnyLBTypeID", , True, True, True, DBMaster)
    End Sub

    Private Sub cmbLBtype_Click()
        Dim mDistrictID As Integer
        Dim mLbID As Integer
        mLbID = cmbLBtype.ItemData(cmbLBtype.ListIndex)
        mDistrictID = cmbDistrict.ItemData(cmbDistrict.ListIndex)
        PopulateList cmbLocalBody, "Select chvLBNameEnglish,intLBID From GM_LocalBody where tnyLBTypeID=" & mLbID & " and tnyDistrictID=" & mDistrictID & " Order By chvLBNameEnglish", , True, True, True, DBMaster
    End Sub
    Private Sub cmbLocation_click()
        Dim LBID As Integer
        LBID = cmbLocalBody.ItemData(cmbLocalBody.ListIndex)
        Call PopulateList(cmbLocationType, "Select chvLocation, intLocationID From GM_Locations where IntLBID=" & LBID & " Order By chvLocation", , True, True, True, enuSourceString.Saankhya)
    End Sub
   
    Private Sub cmdSetLB_Click()
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim mConString As String
        ReDim arrIn(9)
        
        If objDB.CreateNewConnection(mCnn, DBMaster) Then
            If cmbDistrict.ListIndex < 0 Then
                MsgBox "Select District ", vbInformation, "Saankhya"
                cmbDistrict.SetFocus
                Exit Sub
            ElseIf cmbDistrict.ItemData(cmbDistrict.ListIndex) > 0 Then
                arrIn(2) = cmbDistrict.ItemData(cmbDistrict.ListIndex)
            End If
            
            If cmbLBtype.ListIndex < 0 Then
                MsgBox "Select Local Body Type ", vbInformation, "Saankhya"
                cmbLBtype.SetFocus
                Exit Sub
            ElseIf cmbLBtype.ItemData(cmbLBtype.ListIndex) > 0 Then
                arrIn(1) = cmbLBtype.ItemData(cmbLBtype.ListIndex)
'
            End If
            
            If cmbLocalBody.ListIndex < 0 Then
                MsgBox "Select Local Body  ", vbInformation, "Saankhya"
                cmbLocalBody.SetFocus
                Exit Sub
            ElseIf cmbLocalBody.ItemData(cmbLocalBody.ListIndex) > 0 Then
                
                arrIn(0) = cmbLocalBody.ItemData(cmbLocalBody.ListIndex)
                Rec.Open "Select chvLBNameEnglish ,tnyBlockID,chvLBCode,chvAddressEnglish from GM_LocalBody where intLBID=" & arrIn(0) & "", mCnn, adOpenStatic, adLockReadOnly
                arrIn(3) = Rec!tnyBlockID
                arrIn(4) = Rec!chvLBCode
                arrIn(5) = Rec!chvLBNameEnglish
                arrIn(7) = Rec!chvAddressEnglish
                mCnn.Close
            End If
               
        End If
        arrIn(6) = Trim(txtTitle.Text)
        objDB.SetConnection mCnn
        If objDB.SetConnection(mCnn) Then
            If cmbLocationType.ListIndex < 0 Then
                MsgBox "Select Location ", vbInformation, "Saankhya"
                cmbLocationType.SetFocus
                Exit Sub
            ElseIf cmbLocationType.ItemData(cmbLocationType.ListIndex) > 0 Then
                arrIn(8) = cmbLocationType.ItemData(cmbLocationType.ListIndex)
                Rec.Open "select chvLocation from GM_Locations where intLocationID=" & arrIn(8) & " ", mCnn, adOpenStatic, adLockReadOnly
                arrIn(9) = Rec!chvLocation
            End If
                objDB.ExecuteSP "spSaveLBSettings", arrIn, , , mCnn
            End If
            MsgBox "Saved Local Body Sttings", vbInformation, "Saankhya"
    End Sub
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
   
    Private Sub Form_Load()
        Call PopulateList(cmbDistrict, "Select chvDistrictEnglish, tnyDistrictID From GM_District Order By chvDistrictEnglish", , , , True, DBMaster)
        FillbLBtype
        
    End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmSugSaleofTender 
   BackColor       =   &H00EBFAFA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Search Work From List of Tender - "
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCopyToReceipt 
      Appearance      =   0  'Flat
      Caption         =   "Copy Of Receipt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7095
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2625
      Width           =   1710
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EBFAFA&
      Height          =   1545
      Left            =   180
      ScaleHeight     =   1485
      ScaleWidth      =   10650
      TabIndex        =   12
      Top             =   1035
      Width           =   10710
      Begin VB.CommandButton cmdAddParty 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         TabIndex        =   31
         Top             =   90
         Width           =   360
      End
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8430
         TabIndex        =   29
         Top             =   1035
         Width           =   2175
      End
      Begin VB.TextBox txtHousename 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   435
         Width           =   3800
      End
      Begin VB.TextBox txtPartyname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6630
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   90
         Width           =   3975
      End
      Begin VB.TextBox txtMainPlaceName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6630
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtLocalPlaceName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   750
         Width           =   3800
      End
      Begin VB.TextBox txtStreet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6630
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   405
         Width           =   3975
      End
      Begin VB.TextBox txtPostOffice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1065
         Width           =   2490
      End
      Begin VB.TextBox txtPin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         TabIndex        =   15
         Top             =   1065
         Width           =   975
      End
      Begin VB.ComboBox cmbPartyID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   90
         Width           =   3390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7665
         TabIndex        =   30
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   28
         Top             =   465
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5715
         TabIndex        =   26
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Place"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   345
         TabIndex        =   25
         Top             =   765
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Place"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5805
         TabIndex        =   23
         Top             =   735
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6105
         TabIndex        =   22
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post Office"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   21
         Top             =   1095
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3780
         TabIndex        =   20
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued To "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   14
         Top             =   135
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2625
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2625
      Width           =   915
   End
   Begin VB.PictureBox Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBFAFA&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   180
      ScaleHeight     =   525
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   450
      Width           =   10695
      Begin VB.CheckBox chkWithDuplicate 
         BackColor       =   &H00EBFAFA&
         Caption         =   "With Duplicate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9135
         TabIndex        =   11
         Top             =   45
         Width           =   1665
      End
      Begin VB.ComboBox cmbProjectID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   585
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   90
         Width           =   4770
      End
      Begin VB.TextBox txtCostOfTenderDocuments 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5805
         TabIndex        =   2
         Top             =   90
         Width           =   1590
      End
      Begin VB.TextBox txtVAT 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7830
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Work"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5355
         TabIndex        =   5
         Top             =   135
         Width           =   420
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7425
         TabIndex        =   4
         Top             =   135
         Width           =   510
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgTender 
      Height          =   2835
      Left            =   180
      TabIndex        =   10
      Top             =   3105
      Width           =   10605
      _cx             =   18706
      _cy             =   5001
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   5
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sale of Tender "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   10830
   End
End
Attribute VB_Name = "frmSugSaleofTender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCn As New ADODB.Connection

Private Sub chkWithDuplicate_Click()
    Dim Check As Integer
    Dim nCnt As Integer
    If cmbProjectID.ItemData(cmbProjectID.ListIndex) > 0 Then
    Check = cmbProjectID.ItemData(cmbProjectID.ListIndex)
    For nCnt = 1 To fgTender.Rows - 1
    If Check = fgTender.TextMatrix(nCnt, 10) Then  'fgTender.TextMatrix(fgTender.Row, 10) Then
    If chkWithDuplicate.Value <> 1 Then

                txtCostOfTenderDocuments.Text = fgTender.TextMatrix(nCnt, 6)
                txtVAT.Text = fgTender.TextMatrix(nCnt, 7)
              Else
                txtCostOfTenderDocuments.Text = fgTender.TextMatrix(nCnt, 11)
                txtVAT.Text = fgTender.TextMatrix(nCnt, 12)
             End If
    End If
    Next nCnt
    End If
End Sub

Private Sub cmbPartyId_Click()
    Dim vArryIn As Variant
    Dim vArryOut As Variant
    Dim objDB As New clsDb
    Dim mCnn As New ADODB.Connection

    If cmbPartyID.ListIndex > -1 Then
        ReDim vArryIn(0)
        vArryIn(0) = cmbPartyID.ItemData(cmbPartyID.ListIndex)
        objDB.CreateNewConnection mCnn, enuSourceString.Sugama
        objDB.ExecuteSP "SpGM_PartySelect23", vArryIn, vArryOut, , mCnn, adCmdStoredProc
        'gFunExecuteSP "SpGM_PartySelect23", RSelect, adCmdStoredProc, vArryIn, vArryOut, conSugama
        If IsArray(vArryOut) Then
            txtPartyname.Text = vArryOut(1, 0)
            txtHousename.Text = vArryOut(2, 0)
            txtStreet.Text = vArryOut(3, 0)
            txtLocalPlaceName.Text = vArryOut(4, 0)
            txtMainPlaceName.Text = vArryOut(5, 0)
            txtPostOffice.Text = vArryOut(6, 0)
            txtPin.Text = vArryOut(7, 0)
            txtPhoneNo.Text = vArryOut(8, 0)
        Else
            txtPartyname.Text = ""
            txtHousename.Text = ""
            txtStreet.Text = ""
            txtLocalPlaceName.Text = ""
            txtMainPlaceName.Text = ""
            txtPostOffice.Text = ""
            txtPin.Text = ""
            txtPhoneNo.Text = ""
        End If
    Else
        txtPartyname.Text = ""
        txtHousename.Text = ""
        txtStreet.Text = ""
        txtLocalPlaceName.Text = ""
        txtMainPlaceName.Text = ""
        txtPostOffice.Text = ""
        txtPin.Text = ""
        txtPhoneNo.Text = ""
    End If

End Sub

Private Sub cmbProjectID_Click()

    If cmbProjectID.ListIndex > -1 Then

        CallProc

    Else
            txtCostOfTenderDocuments.Text = ""
            txtVAT.Text = ""
            chkWithDuplicate.Value = 0
            txtPartyname.Text = ""
            txtHousename.Text = ""
            txtStreet.Text = ""
            txtLocalPlaceName.Text = ""
            txtMainPlaceName.Text = ""
            txtPostOffice.Text = ""
            txtPin.Text = ""
            txtPhoneNo.Text = ""
    End If

End Sub
Private Sub CallProc()
    Dim vArryIn, vArryOut As Variant
    Dim mCn As New ADODB.Connection
    Dim objDB As New clsDb

    If cmbProjectID.ListIndex > -1 Then
        ReDim vArryIn(1)
        vArryIn(0) = cmbProjectID.ItemData(cmbProjectID.ListIndex)
        vArryIn(1) = gbLocalBodyID
        Set vArryOut = Nothing

        objDB.CreateNewConnection mCn, enuSourceString.Sugama
        objDB.ExecuteSP "spTC_TenderPreparationSelectCounter", vArryIn, vArryOut, , mCn, adCmdStoredProc
        If IsArray(vArryOut) Then
           ' txtPAC.Text = vArryOut(0, 0)
           ' txtEMD.Text = vArryOut(1, 0)
          '  dtLastDateofTender.value = vArryOut(2, 0)
           ' txtFirmPeriodOfTender.Text = vArryOut(3, 0)
          '  txtPeriodOfCompletion.Text = vArryOut(4, 0)
            'txtCostOfTenderDocuments.Text = vArryOut(5, 0)
            'txtVAT.Text = vArryOut(6, 0)
          '  txtFileNo.Text = vArryOut(5, 0)
         '   txtProjectNo.Text = vArryOut(6, 0)
          If chkWithDuplicate.Value <> 1 Then
            txtCostOfTenderDocuments.Text = vArryOut(17, 0)
            txtVAT.Text = vArryOut(18, 0)
          Else
            txtCostOfTenderDocuments.Text = vArryOut(19, 0)
            txtVAT.Text = vArryOut(20, 0)
         End If
            'chkWithDuplicate.value = 0

    Else
       ' clearfields
    End If
End If
    End Sub

Private Sub cmdAddParty_Click()
'    frmPartySelect.Show vbModal
'    If gPartyID <> 0 Then
'        gSubSetComboItem cmbPartyID, gPartyID
'    End If
End Sub

Private Sub cmdClose_Click()
    Dim objTranType As New clsTransactionType
    objTranType.SetTransactionType (9999)
    frmReceiptsCounter.txtTransactiontype.Text = objTranType.TransactionType
    frmReceiptsCounter.txtTransactiontype.Tag = objTranType.TransactionTypeID
    Unload Me
End Sub


Private Sub callsub()
    Dim vArryIn As Variant
    Dim vArryOut As Variant
    ReDim vArryIn(1)
    
    Dim objDB As New clsDb
    Dim mCn As New ADODB.Connection
    
    vArryIn(0) = cmbPartyID.ItemData(cmbPartyID.ListIndex)
    vArryIn(1) = gbLocalBodyID
    Set vArryOut = Nothing
    
    
    objDB.CreateNewConnection mCn, enuSourceString.Sugama
    objDB.ExecuteSP "spGM_PartySelect1", vArryIn, vArryOut, , mCn, adCmdStoredProc
    'gFunExecuteSP "spGM_PartySelect1", RSelect, adCmdStoredProc, vArryIn, vArryOut, conSugama
    
    If IsArray(vArryOut) Then
            gSubSetComboItem cmbPartyID, vArryOut(0, 0)
            
''''            'txtPartTypeID.Text = vArryOut(1, nCnt)
''''            gSubSetComboItem cmbPartyType, Val(vArryOut(1, nCnt))
''''            txtPartyname.Text = vArryOut(2, nCnt)
''''            txtLocalPlaceName.Text = vArryOut(3, nCnt)
''''            txtMainPlaceName.Text = vArryOut(4, nCnt)
''''            gSubSetComboItem cmbDistrictID, vArryOut(5, nCnt)
''''            txtPhoneNo.Text = vArryOut(6, nCnt)
''''            txtEmail.Text = vArryOut(7, nCnt)
''''            gSubSetComboItem cmbPartyCategory, Val(vArryOut(9, nCnt))
''''            txtStreet.Text = vArryOut(10, nCnt)
''''            txtPostOffice.Text = vArryOut(11, nCnt)
''''            txtPin.Text = vArryOut(12, nCnt)
''''            dtDate.value = vArryOut(13, nCnt)
    Else
'''            'txtPartTypeID.Text = ""
'''            txtPartyname.Text = ""
'''            txtLocalPlaceName.Text = ""
'''            txtMainPlaceName.Text = ""
'''            txtPhoneNo.Text = ""
'''            txtEmail.Text = ""
'''            txtStreet.Text = ""
'''            txtPin.Text = ""
'''            txtPostOffice.Text = ""
'''            dtDate.value = Now
    End If
End Sub





Private Sub cmdCopyToReceipt_Click()
    
    '--------------------------------------------------------------------------------'
    ' Demand Details will be Transfered to Receipt Screen                            '
    '--------------------------------------------------------------------------------'
    Dim mTransactionType As Integer
    Dim mSaleOfTenderFormHeadCode As String
    Dim mVATHeadCode    As String
    Dim objAcc As New clsAccounts
    Dim mRows As Integer
    
    '--------------------------------------------------------------------------------'
    ' Validation    -                                                                '
    '--------------------------------------------------------------------------------'
    ' 1) Check Whether Project is selected
    
    If cmbProjectID.ListIndex = -1 Then
        MsgBox "Please select a Project/Work!", vbInformation
        cmbProjectID.SetFocus
        Exit Sub
    End If
    
    ' 2) Cost of Form and VAT should be fixed
    If Val(txtCostOfTenderDocuments.Text) <= 0 Or Val(txtVAT.Text) <= 0 Then
        MsgBox "Cost of Tender form + VAT amount should be fixed!", vbInformation
        txtCostOfTenderDocuments.SetFocus
        Exit Sub
    End If
    
    ' 3) Party ID shoul be fetched
    If cmbPartyID.ListIndex = -1 Then
        MsgBox "Please Select a party to which the form is Sold!", vbInformation
        cmbPartyID.SetFocus
        Exit Sub
    End If
    
    ' 4) Party Name
    If Trim(txtPartyname) = "" Then
        MsgBox "Please Update the Party Master with Name!", vbInformation
        txtPartyname.SetFocus
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------------'
    ' End of Validation
    '--------------------------------------------------------------------------------'
    
    
    
    '--------------------------------------------------------------------------------'
    ' Particulars - Party Name and Work/Project Info
    '--------------------------------------------------------------------------------'
    frmReceiptsCounter.txtName.Text = Trim(txtPartyname)
    frmReceiptsCounter.txtHouse.Text = Trim(txtHousename)
    frmReceiptsCounter.txtStreet.Text = Trim(txtStreet)
    frmReceiptsCounter.txtLocalPlace.Text = Trim(txtLocalPlaceName)
    frmReceiptsCounter.txtMainPlace.Text = Trim(txtMainPlaceName)
    frmReceiptsCounter.txtPost.Text = Trim(txtPostOffice)
    frmReceiptsCounter.txtPin.Text = Trim(txtPin)
    frmReceiptsCounter.txtPhone.Text = Trim(txtPhoneNo)
    
    frmReceiptsCounter.lblBuildingNo.Caption = "Work.Proj. No"
    frmReceiptsCounter.txtBuildingNo.Text = cmbProjectID.Text
    
    '--------------------------------------------------------------------------------'
    ' Demand Details To Grid - Sale of Tender Form ( Cost of Tender Form )           '
    '--------------------------------------------------------------------------------'
    mRows = 1
    mTransactionType = 30
    mSaleOfTenderFormHeadCode = "150110101"
    mVATHeadCode = "350300400"
    frmReceiptsCounter.vsGrid.Rows = 1
    frmReceiptsCounter.vsGrid.Rows = 2
    objAcc.SetAccountCode mSaleOfTenderFormHeadCode
    If Val(txtCostOfTenderDocuments.Text) > 0 Then
        If objAcc.AccountHeadID > 0 Then
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 3) = ""
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 5) = Format(Val(txtCostOfTenderDocuments), "0#")
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 8) = 3
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 9) = 0
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 10) = Val(cmbPartyID.Tag) '"Rec!intKeyID"
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 11) = Format(Val(txtCostOfTenderDocuments), "0#")
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 12) = Val(cmbProjectID.Tag) 'Rec!numDemandID"
            frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mRows, 12) = 1
            frmReceiptsCounter.vsGrid.MergeCol(12) = True
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 13) = "" '"Rec!numBatchID"
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 16) = "" '"DdMmmYy(Rec!dtDemandDate)"
        End If
    End If
    '--------------------------------------------------------------------------------'
    ' Value Added Tax                                                                '
    '--------------------------------------------------------------------------------'
    If Val(txtVAT.Text) > 0 Then
        frmReceiptsCounter.vsGrid.Rows = frmReceiptsCounter.vsGrid.Rows + 1
        mRows = frmReceiptsCounter.vsGrid.Rows - 1
        objAcc.SetAccountCode mVATHeadCode
        If objAcc.AccountHeadID > 0 Then
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 0) = objAcc.AccountCode
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 1) = objAcc.AccountHead
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 2) = str(gbFinancialYearID) & " - " & str(gbFinancialYearID + 1)
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 3) = ""
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 5) = Format(Val(txtCostOfTenderDocuments), "0#")
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 6) = objAcc.AccountHeadID
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 7) = gbFinancialYearID
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 8) = 3
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 9) = 0
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 10) = Val(cmbPartyID.Tag) '"Rec!intKeyID"
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 11) = Format(Val(txtCostOfTenderDocuments), "0#")
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 12) = Val(cmbProjectID.Tag) 'Rec!numDemandID"
            frmReceiptsCounter.vsGrid.Cell(flexcpChecked, mRows, 12) = 1
            frmReceiptsCounter.vsGrid.MergeCol(12) = True
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 13) = "" '"Rec!numBatchID"
            frmReceiptsCounter.vsGrid.Cell(flexcpText, mRows, 16) = "" '"DdMmmYy(Rec!dtDemandDate)"
        End If
    End If
    '--------------------------------------------------------------------------------'
    ' Call Function Calculate from Receipt Counter                                   '
    '--------------------------------------------------------------------------------'
    Call frmReceiptsCounter.Calculate
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim objDB As New clsDb
    Dim mCn As New ADODB.Connection
    
    Dim vArryIn As Variant
    If cmbProjectID.ListIndex = 0 Then
        MsgBox "Please Select Work", vbInformation, "Sugama Works"
        cmbProjectID.SetFocus
        Exit Sub
    End If
    If cmbPartyID.ListIndex = 0 Then
        MsgBox "Please Select Party", vbInformation, "Sugama Works"
        cmbPartyID.SetFocus
        Exit Sub
    End If
    ReDim vArryIn(4)
    vArryIn(0) = cmbProjectID.ItemData(cmbProjectID.ListIndex)
    vArryIn(1) = cmbPartyID.ItemData(cmbPartyID.ListIndex)
    vArryIn(2) = gbLocalBodyID
    vArryIn(3) = txtCostOfTenderDocuments.Text
    vArryIn(4) = txtVAT.Text
    
    objDB.CreateNewConnection mCn, enuSourceString.Sugama
    objDB.ExecuteSP "spTC_TenderIssueDetailsUpdateCounter", vArryIn, , , mCn, adCmdStoredProc
    'gFunExecuteSP "spTC_TenderIssueDetailsUpdateCounter", RUpdate, adCmdStoredProc, vArryIn, , conSugama
    cmbProjectID.ListIndex = 0
    cmbPartyID.ListIndex = 0
    txtPartyname.Text = ""
    txtHousename.Text = ""
    txtStreet.Text = ""
    txtLocalPlaceName.Text = ""
    txtMainPlaceName.Text = ""
    txtPostOffice.Text = ""
    txtPin.Text = ""
    txtPhoneNo.Text = ""


End Sub

Private Sub Command1_Click()

End Sub

Private Sub fgTender_Click()
Dim i As Integer
If fgTender.Rows = 1 Then
    Exit Sub
Else
    i = fgTender.Row
End If
callpopupTender (i)
End Sub
Private Sub callpopupTender(i As Integer)

    Dim varrin As Variant
    Dim varrOut As Variant
    Dim A As String
    Dim B As String
    Dim c As String
    Dim d As String
    Dim e As Integer
    'Dim nCnt As Integer
    Dim objDB As New clsDb
    Dim mCn As New ADODB.Connection
    ReDim varrin(1)
        varrin(0) = fgTender.TextMatrix(i, 10)
        varrin(1) = gbLocalBodyID
        objDB.CreateNewConnection mCn, enuSourceString.Sugama
        objDB.ExecuteSP "spTC_TenderPreparationSelect11", varrin, varrOut, , mCn, adCmdStoredProc
        If IsArray(varrOut) Then
            gSubSetComboItem cmbProjectID, varrOut(0, 0)
            CallProc
          Else

        End If
End Sub

Private Sub Form_Load()

       ' Dim vArryIn As Variant

        Me.Width = 10710
        Me.Height = 6225
        'Me.BackColor = gformbackcolor
        'Frame2.BackColor = gframebackcolor
        'cmdAddParty.BackColor = gbuttoncolor1
        'cmdClose.BackColor = gbuttoncolor
        'cmdOK.BackColor = gbuttoncolor
        'fgTender.BackColorFixed = gflexfixedcolor
        'chkWithDuplicate.BackColor = gframebackcolor
        'HeaderSettings1 Label5
        'CenterForm Me

        PopulateList cmbPartyID, "SELECT chvPartyName, intPartyID FROM GM_Party where tnyPartyTypeID=1", , , , True, enuSourceString.Sugama
        PopulateList cmbProjectID, "SELECT TR_Project.chvProjectName, TC_TenderPreparation.intProjectID FROM TC_TenderPreparation INNER JOIN TR_Project ON TC_TenderPreparation.intProjectID = TR_Project.intProjectID AND TC_TenderPreparation.tnyVersionID = TR_Project.tnyVersionID Where TC_TenderPreparation.intLBID=" & gbLocalBodyID, , , , True, enuSourceString.Sugama

        'FillComboWithZeroIndex cmbPartyID, "SELECT chvPartyName, intPartyID FROM GM_Party where tnyPartyTypeID=1", conSugama
        'FillComboWithZeroIndex cmbProjectID, "SELECT TR_Project.chvProjectName, TC_TenderPreparation.intProjectID FROM TC_TenderPreparation INNER JOIN TR_Project ON TC_TenderPreparation.intProjectID = TR_Project.intProjectID AND TC_TenderPreparation.tnyVersionID = TR_Project.tnyVersionID Where TC_TenderPreparation.intLBID=" & gbLocalBodyID, conSugama
        ' FillComb oWithZeroIndex cmbProjectID, "SELECT TR_Project.chvProjectName, TC_TenderPreparation.intProjectID FROM TC_TenderPreparation INNER JOIN TR_Project ON TC_TenderPreparation.intProjectID = TR_Project.intProjectID AND TC_TenderPreparation.tnyVersionID = TR_Project.tnyVersionID Where dtLastDateOfTender >=" & Now & " and dtSaleofTender <=" & Now & " and TC_TenderPreparation.intLBID=" & gbLocalBodyID, conSugama
        'ReDim vArryIn(0)
        'vArryIn(0) = gbLocalBodyID
        'gFunFillItemsInToCombo cmbProjectID, "spFillProject", adCmdStoredProc, conSugama, vArryIn

        chkWithDuplicate.Value = 0
        fgTenderfill
End Sub

Private Sub fgTenderfill()
    Dim objDB As New clsDb
    Dim mCn As New ADODB.Connection

    Dim vArryOut As Variant
    Dim nCnt As Integer
    Dim VarrIn1 As Variant
    Dim varrInn As Variant
    Set vArryOut = Nothing
    ReDim varrInn(0)

    varrInn(0) = gbLocalBodyID
    objDB.CreateNewConnection mCn, enuSourceString.Sugama
    objDB.ExecuteSP "spTC_TenderPreparationSelect", varrInn, vArryOut, , mCn, adCmdStoredProc
    'gFunExecuteSP "spTC_TenderPreparationSelect", RSelect, adCmdStoredProc, varrInn, vArryOut, conSugama
    If IsArray(vArryOut) Then
        fgTender.Rows = UBound(vArryOut, 2) + 2
        For nCnt = 0 To UBound(vArryOut, 2)
            If (nCnt + 1) Mod 2 = 0 Then
                  'fgTender.Cell(flexcpBackColor, nCnt + 1, 0, nCnt + 1, 12) = gFlexbetweenRowColor
            End If
            fgTender.Cell(flexcpText, nCnt + 1, 0) = vArryOut(0, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 1) = vArryOut(1, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 2) = vArryOut(2, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 3) = vArryOut(3, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 4) = vArryOut(4, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 5) = vArryOut(5, nCnt)
           fgTender.Cell(flexcpText, nCnt + 1, 6) = vArryOut(9, nCnt)
           fgTender.Cell(flexcpText, nCnt + 1, 7) = vArryOut(10, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 8) = vArryOut(6, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 9) = vArryOut(7, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 10) = vArryOut(8, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 11) = vArryOut(11, nCnt)
            fgTender.Cell(flexcpText, nCnt + 1, 12) = vArryOut(12, nCnt)
        Next nCnt
        'lSubAutoWardWrap fgTender
        Else
        fgTender.Rows = 1
      End If
End Sub

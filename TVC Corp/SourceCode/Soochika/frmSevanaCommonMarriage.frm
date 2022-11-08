VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSevanaCommonMarriageSearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "SEARCH FOR REGISTRATIONS UNDER  KRM(C) RULES"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SEARCH PARAMETERS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   0
      TabIndex        =   12
      Top             =   30
      Width           =   9795
      Begin VB.TextBox txthusname 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2340
         TabIndex        =   0
         Top             =   360
         Width           =   2370
      End
      Begin VB.TextBox txtwifename 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7155
         TabIndex        =   1
         Top             =   345
         Width           =   2370
      End
      Begin VB.TextBox txtmarriageplace 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2340
         TabIndex        =   4
         Top             =   1320
         Width           =   2370
      End
      Begin VB.TextBox txthusguard 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2340
         TabIndex        =   2
         Top             =   840
         Width           =   2370
      End
      Begin VB.TextBox txtwit1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2340
         TabIndex        =   6
         Top             =   1800
         Width           =   2370
      End
      Begin VB.TextBox txtregno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7155
         TabIndex        =   9
         Top             =   2265
         Width           =   855
      End
      Begin VB.CommandButton cmdsearch 
         Appearance      =   0  'Flat
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7890
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txtwifeguard 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7155
         TabIndex        =   3
         Top             =   795
         Width           =   2370
      End
      Begin VB.TextBox txtwit2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2340
         TabIndex        =   8
         Top             =   2280
         Width           =   2370
      End
      Begin VB.TextBox txtyear 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8790
         TabIndex        =   10
         Top             =   2250
         Width           =   750
      End
      Begin MSComCtl2.DTPicker dtpdor 
         Height          =   375
         Left            =   7155
         TabIndex        =   7
         Top             =   1770
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   62521345
         CurrentDate     =   38708
      End
      Begin MSComCtl2.DTPicker dtpdom 
         Height          =   375
         Left            =   7155
         TabIndex        =   5
         Top             =   1305
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   62521345
         CurrentDate     =   38708
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Name of Husband"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Place of Marriage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Father of Husband"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name of Witness1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date of Marriage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5175
         TabIndex        =   19
         Top             =   1365
         Width           =   2175
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date of Registration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5175
         TabIndex        =   18
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registration Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5175
         TabIndex        =   17
         Top             =   2385
         Width           =   2175
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name of Witness2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2370
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name of Wife"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5175
         TabIndex        =   15
         Top             =   405
         Width           =   2085
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Father of Wife "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5175
         TabIndex        =   14
         Top             =   855
         Width           =   1875
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8235
         TabIndex        =   13
         Top             =   2310
         Width           =   450
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgmarriage 
      Height          =   4230
      Left            =   0
      TabIndex        =   24
      Top             =   3180
      Width           =   9750
      _cx             =   17198
      _cy             =   7461
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14142647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   11587566
      BackColorAlternate=   8438015
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
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanaCommonMarriage.frx":0000
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
      Editable        =   2
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
      BackColorFrozen =   14868689
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmSevanaCommonMarriageSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
Dim con3 As New ADODB.Connection
Dim objDB As New clsDB
If objDB.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False Then
    MsgBox "Sevana connection failure", vbDefaultButton1
    Exit Sub
End If

Dim rs3 As New ADODB.Recordset
fgmarriage.Clear 1
fgmarriage.Rows = 1

 Dim Qry
      If Trim(txthusname.Text) = "" And Trim(txtwifename.Text) = "" And Trim(txtwit1.Text) = "" And Trim(txtwit2.Text) = "" And Trim(txtwifeguard.Text) = "" And Trim(txthusguard.Text) = "" And Trim(txtmarriageplace.Text) = "" And Trim(txtregno.Text) = "" And IIf(IsNull(dtpdor.value), "", dtpdor.value) = "" And IIf(IsNull(dtpdom.value), "", dtpdom.value) = "" And txtyear.Text = "" Then
      
      'Modified by Soumya V S on 25.02.16
      Qry = "set dateformat dmy select HusName,WifeName,HusNameMal,WifeNameMal,HusFather,WifeFather,MarriagePlaceMal,Witness1,Witness2,MarriageDate,RegnDate,RegnNo,BookNo  from COMMONMARRIAGESEARCHVIEW1"
      Else
      'Modified by Soumya V S on 25.02.16
      Qry = "set dateformat dmy select HusName,WifeName,HusNameMal,WifeNameMal,HusFather,WifeFather,MarriagePlaceMal,Witness1,Witness2,MarriageDate,RegnDate,RegnNo,BookNo  from COMMONMARRIAGESEARCHVIEW1 where "
      Qry = Qry + " HusName  like '" & Trim(txthusname.Text) & "%' "
      Qry = Qry + " and WifeName like '" & Trim(txtwifename.Text) & "%' "
      Qry = Qry + " and Witness1 like '" & Trim(txtwit1.Text) & "%' "
      Qry = Qry + " and Witness2  like '" & Trim(txtwit2.Text) & "%' "
      Qry = Qry + " and HusFather  like '" & Trim(txthusguard.Text) & "%' "
      Qry = Qry + " and WifeFather  like '" & Trim(txtwifeguard.Text) & "%' "

      Qry = Qry + " and MarriagePlaceEng like '" & Trim(txtmarriageplace.Text) & "%' "

      Qry = Qry + " and RegnNo like '" & txtregno.Text & "%' "
      If (txtyear.Text) <> "" Then
      Qry = Qry + " and year(RegnDate) like '" & (txtyear.Text) & "%' "
      End If


      If IIf(IsNull(dtpdom.value), "", (dtpdom.value)) <> "" Then
         Qry = Qry + " and Year(MarriageDate) = '" & IIf(IsNull(dtpdom.value), "", Year(dtpdom.value)) & "' and Month(MarriageDate) = '" & IIf(IsNull(dtpdom.value), "", Month(dtpdom.value)) & "' and Day(MarriageDate) = '" & IIf(IsNull(dtpdom.value), "", Day(dtpdom.value)) & "' "
      End If
      If IIf(IsNull(dtpdor.value), "", dtpdor.value) <> "" Then
         Qry = Qry + " and Year(RegnDate) = '" & IIf(IsNull(dtpdor.value), "", Year(dtpdor.value)) & "' and Month(RegnDate) = '" & IIf(IsNull(dtpdor.value), "", Month(dtpdor.value)) & "' and Day(RegnDate) = '" & IIf(IsNull(dtpdor.value), "", Day(dtpdor.value)) & "' "
         
      End If
      End If
 Set rs3 = con3.Execute(Qry)
 Dim J As Integer
 J = 1
 
While rs3.EOF = False
    fgmarriage.Rows = fgmarriage.Rows + 1
    fgmarriage.TextMatrix(J, 0) = J
    fgmarriage.TextMatrix(J, 1) = rs3(0)
    fgmarriage.TextMatrix(J, 2) = rs3(1)
    fgmarriage.TextMatrix(J, 3) = rs3(2)
    fgmarriage.Cell(flexcpFontName, J, 3) = "ML-TTRevathi"
    fgmarriage.Cell(flexcpFontSize, J, 3) = 11
    fgmarriage.TextMatrix(J, 4) = rs3(3)
    fgmarriage.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
    fgmarriage.Cell(flexcpFontSize, J, 4) = 11
    fgmarriage.TextMatrix(J, 5) = rs3(4)
    fgmarriage.TextMatrix(J, 6) = rs3(5)
    fgmarriage.TextMatrix(J, 7) = rs3(6)
    fgmarriage.TextMatrix(J, 8) = rs3(7)
    fgmarriage.TextMatrix(J, 9) = rs3(8)
    fgmarriage.TextMatrix(J, 10) = rs3(9)
If IsNull(rs3(10)) Then
    fgmarriage.TextMatrix(J, 11) = "Not Given"
Else
    fgmarriage.TextMatrix(J, 11) = rs3(10)
End If

fgmarriage.TextMatrix(J, 12) = rs3(11)
'Added by Soumya V S on 25.02.16
fgmarriage.TextMatrix(J, 13) = rs3(12)
J = J + 1
rs3.MoveNext
Wend
con3.Close

End Sub
Private Sub fgmarriage_Click()
'Dim SelectedCol As Long
'    Dim RelationshipComboIndex As Integer
''  '---for the Column wise Selection of Records  ' commented on 29/10/2011 By Poornima For Soochika
''    Select Case fgmarriage.ColSel
''      Case 1
''              SelectedCol = 1 'EngGroom
''              RelationshipComboIndex = 0
''      Case 2
''              SelectedCol = 2 'EngBride
''              RelationshipComboIndex = 1
'''      Case 4
'''              SelectedCol = 4 'MalBride
'''              RelationshipComboIndex = 1
''      Case Else
''              SelectedCol = 1 ' Default Name of groom
''              RelationshipComboIndex = 0
''    End Select
'        SelectedCol = 1 ' Default Name of groom
'        RelationshipComboIndex = 0
'
'
'    '*******************************************************************************************
'    ' Added by Akheel 09.03.11 for Unicode Version
'    If (gbSoochikaVer = 5) Then
'        frmUSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
'        If SelectedCol = 2 Then
'            frmUSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
'        Else
'            frmUSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 3))
'        End If
'        frmUSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 12))
'        frmUSevanaInward.txtbookno.Text = 1
'        frmUSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex
'
'        frmUSevanaInward.txtMalayalamname.Enabled = False
'        frmUSevanaInward.txtEnglishname.Enabled = False
'        frmUSevanaInward.txtregno.Enabled = False
'        frmUSevanaInward.txtbookno.Enabled = False
'        frmUSevanaInward.cboRelationship.Enabled = False
'    Else
'        frmSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
'        If SelectedCol = 2 Then
'            frmSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
'    '    ElseIf SelectedCol = 4 Then
'    '
'    '        frmSevanaInward.txtMalCertName.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
'        Else
'            frmSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 3))
'        End If
'        frmSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 12))
'        frmSevanaInward.txtbookno.Text = 1
'        frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex
'        ' for disabling Editing
'        frmSevanaInward.txtMalayalamname.Enabled = False
'        frmSevanaInward.txtEnglishname.Enabled = False
'        frmSevanaInward.txtregno.Enabled = False
'        frmSevanaInward.txtbookno.Enabled = False
'        frmSevanaInward.cboRelationship.Enabled = False
'    End If
'    Unload Me


    Dim SelectedCol As Long
    Dim RelationshipComboIndex As Integer
  '---for the Column wise Selection of Records
    Select Case fgmarriage.ColSel
      Case 1
              SelectedCol = 1 'EngGroom
              RelationshipComboIndex = 0
      Case 2
              SelectedCol = 2 'EngBride
              RelationshipComboIndex = 1
'      Case 4
'              SelectedCol = 4 'MalBride
'              RelationshipComboIndex = 1
      Case Else
              SelectedCol = 1 ' Default Name of groom
              RelationshipComboIndex = 0
    End Select
     'To make husname as relation for every combo click : - in page at the time of certificate printing it validate with hus name
             SelectedCol = 1 'EngGroom
              RelationshipComboIndex = 0
    'end of modification on 13.10.10
     If (gbSoochikaVer = 5) Then
    frmUSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
    If SelectedCol = 2 Then
        frmUSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
'    ElseIf SelectedCol = 4 Then
'
'        frmFeeParticulars.txtMalCertName.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
    Else
        frmUSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 3))
    End If
    frmUSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 12))
    'Modified by Soumya V S on 25.02.16
    'frmUSevanaInward.txtbookno.Text = 1
    frmUSevanaInward.txtBookNo.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 13))
    frmUSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex
    ' for disabling Editing
    frmUSevanaInward.txtMalayalamname.Enabled = False
    frmUSevanaInward.txtEnglishname.Enabled = False
    frmUSevanaInward.txtregno.Enabled = False
    frmUSevanaInward.txtBookNo.Enabled = False
    frmUSevanaInward.cboRelationship.Enabled = False
    
If frmUSevanaInward.txtEnglishname.Text = "Not Given" And frmUSevanaInward.cboLanguage.ListIndex = 1 And (frmUSevanaInward.txtMalayalamname.Text <> "\ðInbn«nñ") Then
  frmUSevanaInward.cboLanguage.ListIndex = 0
End If
If frmUSevanaInward.txtMalayalamname.Text = "\ðInbn«nñ" And frmUSevanaInward.cboLanguage.ListIndex = 0 And (frmUSevanaInward.txtEnglishname.Text <> "Not Given") Then
frmUSevanaInward.cboLanguage.ListIndex = 1
End If
Else

frmSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
    If SelectedCol = 2 Then
        frmSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
'    ElseIf SelectedCol = 4 Then
'
'        frmFeeParticulars.txtMalCertName.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 4))
    Else
        frmSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 3))
    End If
    frmSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 12))
    'Modified by soumya V S on 25.02.016
    'frmSevanaInward.txtbookno.Text = 1
    frmSevanaInward.txtBookNo.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 13))
    frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex
    ' for disabling Editing
    frmSevanaInward.txtMalayalamname.Enabled = False
    frmSevanaInward.txtEnglishname.Enabled = False
    frmSevanaInward.txtregno.Enabled = False
    frmSevanaInward.txtBookNo.Enabled = False
    frmSevanaInward.cboRelationship.Enabled = False
    
    If frmSevanaInward.txtEnglishname.Text = "Not Given" And frmSevanaInward.cboLanguage.ListIndex = 1 And (frmSevanaInward.txtMalayalamname.Text <> "\ðInbn«nñ") Then
      frmSevanaInward.cboLanguage.ListIndex = 0
    End If
    If frmSevanaInward.txtMalayalamname.Text = "\ðInbn«nñ" And frmSevanaInward.cboLanguage.ListIndex = 0 And (frmSevanaInward.txtEnglishname.Text <> "Not Given") Then
     frmSevanaInward.cboLanguage.ListIndex = 1
    End If
End If

    Unload Me

End Sub


Private Sub Form_Load()
    gSubCenterForm Me
End Sub

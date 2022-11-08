VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSevanadethsearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "DEATH SEARCH"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      Height          =   375
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "FIELDS TO SEARCH"
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10035
      Begin VB.TextBox txtyear 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8955
         TabIndex        =   6
         Top             =   1200
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtpdor 
         Height          =   345
         Left            =   6285
         TabIndex        =   8
         Top             =   1695
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   67043329
         CurrentDate     =   38708
      End
      Begin MSComCtl2.DTPicker dtpdod 
         Height          =   345
         Left            =   6285
         TabIndex        =   1
         Top             =   255
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   67043329
         CurrentDate     =   38708
      End
      Begin VB.TextBox txtplaceofdeath 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1965
         TabIndex        =   7
         Top             =   1725
         Width           =   2205
      End
      Begin VB.ComboBox Cmbsex 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6285
         TabIndex        =   3
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtregno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6285
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtmoorwif 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1965
         TabIndex        =   4
         Top             =   1200
         Width           =   2205
      End
      Begin VB.TextBox txtfaorhu 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1965
         TabIndex        =   2
         Top             =   720
         Width           =   2205
      End
      Begin VB.TextBox txtdeceased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1965
         TabIndex        =   0
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8385
         TabIndex        =   19
         Top             =   1275
         Width           =   435
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registration Number"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4425
         TabIndex        =   18
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date Of Registration"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4440
         TabIndex        =   17
         Top             =   1725
         Width           =   1755
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sex of Deceased"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4440
         TabIndex        =   16
         Top             =   795
         Width           =   1665
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date Of Death"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Place Of Death"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Mother/Wife"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Father/Husbund"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Deceased"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgdeath 
      Height          =   4485
      Left            =   0
      TabIndex        =   20
      Top             =   2205
      Width           =   9990
      _cx             =   17621
      _cy             =   7911
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanaDethSearch.frx":0000
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
Attribute VB_Name = "frmSevanadethsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varay As Variant
Dim vOut As Variant

Private Sub Command1_Click()
fgdeath.Clear 1
fgdeath.Rows = 1
Dim con3 As New ADODB.Connection
Dim objDB As New clsDB
Dim rs3 As New ADODB.Recordset
Dim gen As String

If objDB.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False Then
    MsgBox "Sevana Connection Failed", vbDefaultButton1
    Exit Sub
End If
If Cmbsex.ItemData(Cmbsex.ListIndex) = 1 Then
 gen = "male"
 
ElseIf Cmbsex.ItemData(Cmbsex.ListIndex) = 2 Then
 gen = "Female"
End If

  Dim Qry
    If Trim(txtdeceased.Text) = "" And Trim(txtfaorhu.Text) = "" And Trim(txtplaceofdeath.Text) = "" And Cmbsex.ItemData(Cmbsex.ListIndex) = 0 And txtregno.Text = "" And IIf(IsNull(dtpdor.value), "", dtpdor.value) = "" And IIf(IsNull(dtpdod.value), "", dtpdod.value) = "" And Trim(txtmoorwif.Text) = "" And (txtyear.Text) = "" Then
    'Added by Sreeja on 11.8.09
    ' Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH from DEATHSEARCHVIEW1 "
        Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH,MW from DEATHSEARCHVIEW1 "
    ' end here
    Else

      'Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH from DEATHSEARCHVIEW1 where  "
      Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH,MW from DEATHSEARCHVIEW1 where "
      Qry = Qry + " DecName  like '" & Trim(txtdeceased.Text) & "%' " ' and Father like '" & fname & "%' "
   
      Qry = Qry + " and DeathPlaceEng like '" & Trim(txtplaceofdeath.Text) & "%' "
      
      Qry = Qry + " and chvEngGender  like '" & gen & "%' "

      Qry = Qry + " and chvRegnNo like '" & txtregno.Text & "%' "
      
       Qry = Qry + " and year(dtmRegnDate) like '" & (txtyear.Text) & "%' "
      If IIf(IsNull(dtpdod.value), "", dtpdod.value) <> "" Then
         Qry = Qry + " and year(Date1) = '" & IIf(IsNull(dtpdod.value), "", Year(dtpdod.value)) & "' and Month(Date1) = '" & IIf(IsNull(dtpdod.value), "", Month(dtpdod.value)) & "' and Day(Date1) = '" & IIf(IsNull(dtpdod.value), "", Day(dtpdod.value)) & "' "
      End If
      
      If IIf(IsNull(dtpdor.value), "", dtpdor.value) <> "" Then
         Qry = Qry + " and year(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Year(dtpdor.value)) & "' and Month(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Month(dtpdor.value)) & "' and Day(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Day(dtpdor.value)) & "' "
      End If
 
      Qry = Qry + " and FH like '" & Trim(txtfaorhu.Text) & "%' "
      Qry = Qry + " and MW like '" & Trim(txtmoorwif.Text) & "%' "
    
    End If
 Set rs3 = con3.Execute(Qry)

Dim J As Variant
 J = 1
 
While rs3.EOF = False
fgdeath.Rows = fgdeath.Rows + 1
fgdeath.TextMatrix(J, 0) = J
fgdeath.TextMatrix(J, 1) = rs3(0)
fgdeath.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
fgdeath.TextMatrix(J, 2) = rs3(1)
fgdeath.TextMatrix(J, 3) = rs3(2)
fgdeath.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
fgdeath.TextMatrix(J, 4) = rs3(3)
fgdeath.TextMatrix(J, 5) = rs3(4)

'Added and modified by Sreeja on 11.8.09
fgdeath.Cell(flexcpFontName, J, 6) = "ML-TTRevathi"
fgdeath.TextMatrix(J, 6) = rs3(12)
'------------end

'fgdeath.TextMatrix(J, 7) = rs3(6)
'fgdeath.TextMatrix(J, 8) = rs3(7)
fgdeath.TextMatrix(J, 7) = rs3(5)
fgdeath.TextMatrix(J, 8) = IIf(IsNull(rs3(6)), 0, rs3(6))
If IsNull(rs3(7)) Then
fgdeath.TextMatrix(J, 9) = "Not Given"
Else
fgdeath.TextMatrix(J, 9) = rs3(7)
End If
fgdeath.TextMatrix(J, 10) = IIf(IsNull(rs3(8)), "Not Given", rs3(8))
fgdeath.TextMatrix(J, 11) = rs3(9)
fgdeath.TextMatrix(J, 12) = rs3(10)
'If IsNull(rs3(8)) Then
'fgdeath.TextMatrix(J, 9) = "Not Given"
'Else
'fgdeath.TextMatrix(J, 9) = rs3(8)
'End If
'fgdeath.TextMatrix(J, 10) = rs3(9)
'fgdeath.TextMatrix(J, 11) = rs3(10)
'fgdeath.TextMatrix(J, 12) = rs3(11)
J = J + 1
rs3.MoveNext
Wend
con3.Close
'
'Dim J As Integer
' J = 1
'If IsArray(vOut) Then
'For nCnt = 0 To UBound(vOut, 2)
' fgdeath.Rows = fgdeath.Rows + 1
' fgdeath.Cell(flexcpText, J, 0) = J
' fgdeath.Cell(flexcpText, J, 1) = vOut(0, nCnt)
' fgdeath.Cell(flexcpText, J, 2) = vOut(1, nCnt)
' fgdeath.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
' fgdeath.Cell(flexcpText, J, 3) = vOut(2, nCnt)
' fgdeath.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
' fgdeath.Cell(flexcpText, J, 4) = vOut(3, nCnt)
' fgdeath.Cell(flexcpText, J, 5) = vOut(4, nCnt)
' fgdeath.Cell(flexcpText, J, 6) = vOut(5, nCnt)
' fgdeath.Cell(flexcpText, J, 7) = vOut(6, nCnt)
' fgdeath.Cell(flexcpText, J, 8) = vOut(7, nCnt)
' fgdeath.Cell(flexcpText, J, 9) = vOut(8, nCnt)
' fgdeath.Cell(flexcpText, J, 10) = vOut(9, nCnt)
'J = J + 1
'Next nCnt
'End If
'
'Set vOut = Nothing
End Sub
Private Sub fgdeath_DblClick()
Dim SelectedCol As Long
Dim RelationshipComboIndex As Integer
  '-------------Modified by Arun A on 22/2/2007 for the Column wise Selection of Records--------
    Select Case fgdeath.ColSel
      Case 1, 2
              SelectedCol = 1 'EngDead
              RelationshipComboIndex = 0
      Case 3, 4
              SelectedCol = 3 'EngFatherOrHusband
              If InStr(1, fgdeath.TextMatrix(fgdeath.RowSel, 12), "Father") > 1 Then
                  RelationshipComboIndex = 1
              Else
                  RelationshipComboIndex = 2
              End If
    'Case 5 'Commented and added by Sreeja
     Case 5, 6
              SelectedCol = 5 'EngMotherOrWife
              'If InStr(1, fgdeath.TextMatrix(fgdeath.RowSel, SelectedCol), "Mother") > 1 Then
              If InStr(1, fgdeath.TextMatrix(fgdeath.RowSel, 12), "Mother") > 1 Then
                  RelationshipComboIndex = 3
              Else
                  RelationshipComboIndex = 4
              End If
      Case Else
              SelectedCol = 3 ' Default is EngFatherOrHusband
              If InStr(1, fgdeath.TextMatrix(fgdeath.RowSel, 12), "Father") > 1 Then
                  RelationshipComboIndex = 1
              Else
                  RelationshipComboIndex = 2
              End If
    End Select
        
    '*******************************************************************************************
    'Added by Akheel 09.03.11 for Unicode Version
  If gbSoochikaVer = 5 Then
      frmUSevanaInward.txtEnglishname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 1))
      frmUSevanaInward.txtMalayalamname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 2))
      frmUSevanaInward.txtregno.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 11))
      frmUSevanaInward.txtbookno.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 12))
      frmUSevanaInward.cboRelationship.ListIndex = 0
      frmUSevanaInward.txtEnglishname.Enabled = False
      frmUSevanaInward.txtMalayalamname.Enabled = False
      frmUSevanaInward.txtregno.Enabled = False
      frmUSevanaInward.txtbookno.Enabled = False
      frmUSevanaInward.cboRelationship.Enabled = False
      
      If frmUSevanaInward.txtEnglishname.Text = "Not Given" And frmUSevanaInward.cboLanguage.ListIndex = 1 And (frmUSevanaInward.txtMalayalamname.Text <> "\ðInbn«nñ") Then
        frmUSevanaInward.cboLanguage.ListIndex = 0
    End If

    If frmUSevanaInward.txtMalayalamname.Text = "\ðInbn«nñ" And frmUSevanaInward.cboLanguage.ListIndex = 0 And (frmUSevanaInward.txtEnglishname.Text <> "Not Given") Then
        frmUSevanaInward.cboLanguage.ListIndex = 1
    End If
      
    '*******************************************************************************************
  Else
      '-------------End--------------------------------------------------------------------------
      'Added by Sreeja on 18.8.09
      'frmSevanaInward.txtEnglishname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, SelectedCol))
      'frmSevanaInward.txtMalayalamname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, SelectedCol + 1))
      frmSevanaInward.txtEnglishname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 1))
      frmSevanaInward.txtMalayalamname.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 2))
      
      frmSevanaInward.txtregno.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 11))
      frmSevanaInward.txtbookno.Text = (fgdeath.TextMatrix(fgdeath.RowSel, 12))
      'Modified by Sreeja on 18.8.09
      'frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified by Arun A on 22/2/2007
      frmSevanaInward.cboRelationship.ListIndex = 0
      'Modified by Arun A for disabling data Editing
      frmSevanaInward.txtEnglishname.Enabled = False
      frmSevanaInward.txtMalayalamname.Enabled = False
      frmSevanaInward.txtregno.Enabled = False
      frmSevanaInward.txtbookno.Enabled = False
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
Cmbsex.Clear
    Cmbsex.AddItem "...."
    Cmbsex.ItemData(Cmbsex.NewIndex) = 0
    Cmbsex.AddItem "Male"
    Cmbsex.ItemData(Cmbsex.NewIndex) = 1
    Cmbsex.AddItem "Female"
    Cmbsex.ItemData(Cmbsex.NewIndex) = 2
    Cmbsex.ListIndex = 0
End Sub



Private Sub Search()
        
        fgdeath.Clear 1
        fgdeath.Rows = 1
        
        Dim con3 As New ADODB.Connection
        Dim objDB As New clsDB
        Dim rs3 As New ADODB.Recordset
        Dim gen As String
        Dim Qry As Variant
        Dim J As Variant
        
        If objDB.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False Then
            MsgBox "Sevana Connection Failed", vbDefaultButton1
            Exit Sub
        End If
        If Cmbsex.ItemData(Cmbsex.ListIndex) = 1 Then
            gen = "male"
        ElseIf Cmbsex.ItemData(Cmbsex.ListIndex) = 2 Then
            gen = "Female"
        End If
        
        
        If Trim(txtdeceased.Text) = "" And Trim(txtfaorhu.Text) = "" And Trim(txtplaceofdeath.Text) = "" And Cmbsex.ItemData(Cmbsex.ListIndex) = 0 And txtregno.Text = "" And IIf(IsNull(dtpdor.value), "", dtpdor.value) = "" And IIf(IsNull(dtpdod.value), "", dtpdod.value) = "" And Trim(txtmoorwif.Text) = "" And (txtyear.Text) = "" Then
            ' Added by Sreeja on 11.8.09
            ' Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH from DEATHSEARCHVIEW1 "
            Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH,MW from DEATHSEARCHVIEW1 "
            ' end here
        Else
            'Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH from DEATHSEARCHVIEW1 where  "
            Qry = "set dateformat dmy select DecName,DecNameMal,GurEng,GurMal,MW,DeathPlaceEng,Date1,chvEngGender,dtmRegnDate,chvRegnNo,bookno,FH,MW from DEATHSEARCHVIEW1 where "
            Qry = Qry + " DecName  like '" & Trim(txtdeceased.Text) & "%' " ' and Father like '" & fname & "%' "
            Qry = Qry + " and DeathPlaceEng like '" & Trim(txtplaceofdeath.Text) & "%' "
            Qry = Qry + " and chvEngGender  like '" & gen & "%' "
            Qry = Qry + " and chvRegnNo like '" & txtregno.Text & "%' "
            Qry = Qry + " and year(dtmRegnDate) like '" & (txtyear.Text) & "%' "
            If IIf(IsNull(dtpdod.value), "", dtpdod.value) <> "" Then
                Qry = Qry + " and year(Date1) = '" & IIf(IsNull(dtpdod.value), "", Year(dtpdod.value)) & "' and Month(Date1) = '" & IIf(IsNull(dtpdod.value), "", Month(dtpdod.value)) & "' and Day(Date1) = '" & IIf(IsNull(dtpdod.value), "", Day(dtpdod.value)) & "' "
            End If
            If IIf(IsNull(dtpdor.value), "", dtpdor.value) <> "" Then
                Qry = Qry + " and year(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Year(dtpdor.value)) & "' and Month(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Month(dtpdor.value)) & "' and Day(dtmRegnDate) = '" & IIf(IsNull(dtpdor.value), "", Day(dtpdor.value)) & "' "
            End If
            Qry = Qry + " and FH like '" & Trim(txtfaorhu.Text) & "%' "
            Qry = Qry + " and MW like '" & Trim(txtmoorwif.Text) & "%' "
        End If
        Set rs3 = con3.Execute(Qry)
        
        J = 1
        While rs3.EOF = False
            fgdeath.Rows = fgdeath.Rows + 1
            fgdeath.TextMatrix(J, 0) = J
            fgdeath.TextMatrix(J, 1) = rs3(0)
            fgdeath.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
            fgdeath.TextMatrix(J, 2) = rs3(1)
            fgdeath.TextMatrix(J, 3) = rs3(2)
            fgdeath.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
            fgdeath.TextMatrix(J, 4) = rs3(3)
            fgdeath.TextMatrix(J, 5) = rs3(4)
            
            'Added and modified by Sreeja on 11.8.09
            fgdeath.Cell(flexcpFontName, J, 6) = "ML-TTRevathi"
            fgdeath.TextMatrix(J, 6) = rs3(12)
            fgdeath.TextMatrix(J, 7) = rs3(5)
            fgdeath.TextMatrix(J, 8) = rs3(6)
            If IsNull(rs3(7)) Then
                fgdeath.TextMatrix(J, 9) = "Not Given"
            Else
                fgdeath.TextMatrix(J, 9) = rs3(7)
            End If
            fgdeath.TextMatrix(J, 10) = IIf(IsNull(rs3(8)), "Not Given", rs3(8))
            fgdeath.TextMatrix(J, 11) = rs3(9)
            fgdeath.TextMatrix(J, 12) = rs3(10)
            
            J = J + 1
            rs3.MoveNext
        Wend
        con3.Close
       
End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSevanaMarriageSearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "MARRIAGE SEARCH"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FIELDS TO BE SEARCH"
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
      Left            =   105
      TabIndex        =   12
      Top             =   -30
      Width           =   9675
      Begin VB.TextBox txtyear 
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
         Height          =   375
         Left            =   8775
         TabIndex        =   10
         Top             =   2250
         Width           =   750
      End
      Begin MSComCtl2.DTPicker dtpdor 
         Height          =   375
         Left            =   7155
         TabIndex        =   7
         Top             =   1770
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61931521
         CurrentDate     =   38708
      End
      Begin MSComCtl2.DTPicker dtpdom 
         Height          =   375
         Left            =   7155
         TabIndex        =   5
         Top             =   1305
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61931521
         CurrentDate     =   38708
      End
      Begin VB.TextBox txtwit2 
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
         Height          =   375
         Left            =   2340
         TabIndex        =   8
         Top             =   2280
         Width           =   2370
      End
      Begin VB.TextBox txtwifeguard 
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
         Height          =   375
         Left            =   7155
         TabIndex        =   3
         Top             =   795
         Width           =   2370
      End
      Begin VB.CommandButton cmdsearch 
         Appearance      =   0  'Flat
         Caption         =   "SEARCH"
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
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2715
         Width           =   975
      End
      Begin VB.TextBox txtregno 
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
         Height          =   375
         Left            =   7155
         TabIndex        =   9
         Top             =   2265
         Width           =   975
      End
      Begin VB.TextBox txtwit1 
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
         Height          =   375
         Left            =   2340
         TabIndex        =   6
         Top             =   1800
         Width           =   2370
      End
      Begin VB.TextBox txthusguard 
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
         Height          =   375
         Left            =   2340
         TabIndex        =   2
         Top             =   840
         Width           =   2370
      End
      Begin VB.TextBox txtmarriageplace 
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
         Height          =   375
         Left            =   2340
         TabIndex        =   4
         Top             =   1320
         Width           =   2370
      End
      Begin VB.TextBox txtwifename 
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
         Height          =   375
         Left            =   7155
         TabIndex        =   1
         Top             =   345
         Width           =   2370
      End
      Begin VB.TextBox txthusname 
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
         Height          =   375
         Left            =   2340
         TabIndex        =   0
         Top             =   360
         Width           =   2370
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         Height          =   300
         Left            =   8235
         TabIndex        =   23
         Top             =   2370
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Guardian Name(Wife)"
         Height          =   225
         Left            =   5175
         TabIndex        =   22
         Top             =   855
         Width           =   1875
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wife Name"
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
         Left            =   5175
         TabIndex        =   21
         Top             =   405
         Width           =   2085
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000004&
         Caption         =   "Witness2 Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2370
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   5175
         TabIndex        =   19
         Top             =   2370
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Of Registration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5175
         TabIndex        =   18
         Top             =   1860
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Of Marriage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5175
         TabIndex        =   17
         Top             =   1365
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
         Caption         =   "Witness1 Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         Caption         =   "Guardian Name(Hus)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   "Place Of Marriage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   "Husband Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgmarriage 
      Height          =   4230
      Left            =   120
      TabIndex        =   24
      Top             =   3225
      Width           =   9750
      _cx             =   17198
      _cy             =   7461
      Appearance      =   0
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
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanamarriagesearch.frx":0000
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
Attribute VB_Name = "frmSevanaMarriageSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varay As Variant
Dim vOut As Variant
Private Sub cmdSearch_Click()
Dim con3 As New ADODB.Connection
Dim objDB As New clsDB
If (objDB.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False) Then
    MsgBox "Sevena connection Failed", vbDefaultButton1
    Exit Sub
End If
Dim rs3 As New ADODB.Recordset
fgmarriage.Clear 1
fgmarriage.Rows = 1

 Dim Qry
      If Trim(txthusname.Text) = "" And Trim(txtwifename.Text) = "" And Trim(txtwit1.Text) = "" And Trim(txtwit2.Text) = "" And Trim(txtwifeguard.Text) = "" And Trim(txthusguard.Text) = "" And Trim(txtmarriageplace.Text) = "" And Trim(txtregno.Text) = "" And IIf(IsNull(dtpdor.value), "", dtpdor.value) = "" And IIf(IsNull(dtpdom.value), "", dtpdom.value) = "" And txtyear.Text = "" Then
      'Qry = "set dateformat dmy select GroomName,BrideName,GroomGuardian,BrideGuardian,MarriagePlaceEng,Witness1,Witness2,MarriageDate,RegnDate,RegnNo  from MARRIAGESEARCHVIEW1"
      
      'adding BooK no by Soumya V S on 25.02.016
      Qry = "set dateformat dmy select GroomName,GroomNameMal,BrideName,BrideNameMal,GroomGuardian,GroomGuardianMal,BrideGuardian,BrideGuardianMal,MarriagePlaceEng,Witness1,Witness1Mal,Witness2,Witness2Mal,MarriageDate,RegnDate,RegnNo,AchkNo,Bookno  from MARRIAGESEARCHVIEW1"
      Else
      'Qry = "set dateformat dmy select GroomName,BrideName,GroomGuardian,BrideGuardian,MarriagePlaceEng,Witness1,Witness2,MarriageDate,RegnDate,RegnNo  from MARRIAGESEARCHVIEW1 where "
      'adding BooK no by Soumya V S on 25.02.016
      Qry = "set dateformat dmy select GroomName,GroomNameMal,BrideName,BrideNameMal,GroomGuardian,GroomGuardianMal,BrideGuardian,BrideGuardianMal,MarriagePlaceEng,Witness1,Witness1Mal,Witness2,Witness2Mal,MarriageDate,RegnDate,RegnNo,AchkNo,Bookno  from MARRIAGESEARCHVIEW1 where "
      Qry = Qry + "     GroomName  like '" & Trim(txthusname.Text) & "%' "
      Qry = Qry + " and BrideName like '" & Trim(txtwifename.Text) & "%' "
      Qry = Qry + " and Witness1 like '" & Trim(txtwit1.Text) & "%' "
      Qry = Qry + " and Witness2  like '" & Trim(txtwit2.Text) & "%' "
      
      
      Qry = Qry + " and GroomGuardian  like '" & Trim(txthusguard.Text) & "%' "
      Qry = Qry + " and BrideGuardian  like '" & Trim(txtwifeguard.Text) & "%' "
      
      Qry = Qry + " and MarriagePlaceEng like '" & Trim(txtmarriageplace.Text) & "%' "
      
      Qry = Qry + " and RegnNo like '" & txtregno.Text & "%' "
      Qry = Qry + " and year(RegnDate) like '" & (txtyear.Text) & "%' "
      
      
      If IIf(IsNull(dtpdom.value), "", dtpdom.value) <> "" Then
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
'    fgmarriage.TextMatrix(J, 1) = rs3(0)
'    fgmarriage.TextMatrix(J, 2) = rs3(1)
'    fgmarriage.TextMatrix(J, 3) = rs3(2)
'    fgmarriage.TextMatrix(J, 4) = rs3(3)
'    fgmarriage.TextMatrix(J, 5) = rs3(4)
'    fgmarriage.TextMatrix(J, 6) = rs3(5)
'    fgmarriage.TextMatrix(J, 7) = rs3(6)
'    fgmarriage.TextMatrix(J, 8) = rs3(7)
'If IsNull(rs3(8)) Then
'    fgmarriage.TextMatrix(J, 9) = "Not Given"
'Else
'    fgmarriage.TextMatrix(J, 9) = rs3(8)
'End If
'
'fgmarriage.TextMatrix(J, 10) = rs3(9)
''Added by Sreeja on 11.8.09
'fgmarriage.TextMatrix(J, 11) = rs3(10)
'--------end
'Commented and added by Sreeja on 11.8.09
    fgmarriage.TextMatrix(J, 1) = rs3(0)
    fgmarriage.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 2) = rs3(1)
    fgmarriage.TextMatrix(J, 3) = rs3(2)
    fgmarriage.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 4) = rs3(3)
    fgmarriage.TextMatrix(J, 5) = rs3(4)
     fgmarriage.Cell(flexcpFontName, J, 6) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 6) = rs3(5)
    fgmarriage.TextMatrix(J, 7) = rs3(6)
    fgmarriage.Cell(flexcpFontName, J, 8) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 8) = rs3(7)
    fgmarriage.TextMatrix(J, 9) = rs3(8)
    fgmarriage.TextMatrix(J, 10) = rs3(9)
    fgmarriage.Cell(flexcpFontName, J, 11) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 11) = rs3(10)
    fgmarriage.TextMatrix(J, 12) = rs3(11)
    fgmarriage.Cell(flexcpFontName, J, 13) = "ML-TTRevathi"
    fgmarriage.TextMatrix(J, 13) = rs3(12)
    fgmarriage.TextMatrix(J, 14) = IIf(IsNull(rs3(13)), 0, rs3(13))
    
If IsNull(rs3(14)) Then
    fgmarriage.TextMatrix(J, 15) = "Not Given"
Else
    fgmarriage.TextMatrix(J, 15) = rs3(14)
End If

fgmarriage.TextMatrix(J, 16) = rs3(15)
'Added by Sreeja on 11.8.09
fgmarriage.TextMatrix(J, 17) = IIf(IsNull(rs3(16)), " ", rs3(16))

'-------end -------------------------------------------
J = J + 1
rs3.MoveNext
Wend
con3.Close

End Sub

Private Sub fgmarriage_DblClick()
    Dim SelectedCol As Long
    Dim RelationshipComboIndex As Integer
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim res As New ADODB.Recordset
        
    If (objDB.CreateNewConnection(mCnn, enuSourceString.SevanaRegn) = False) Then
        MsgBox "Sevena connection Failed", vbDefaultButton1
    Exit Sub
End If
  '-------------Modified by Arun A on 22/2/2007 for the Column wise Selection of Records--------
    Select Case fgmarriage.ColSel
      Case 1
              SelectedCol = 1 'EngGroom
              RelationshipComboIndex = 0
      Case 3
'              SelectedCol = 3 'EngBride
'              RelationshipComboIndex = 1
'add by  vp
                SelectedCol = 1             'EngBride
              RelationshipComboIndex = 0
      Case Else
              SelectedCol = 1 ' Default Name of groom
              RelationshipComboIndex = 0
    End Select
    
  
  '-------------End--------------------------------------------------------------------------
  '******************************************************************************************
  ' Modified by Akheel 09.03.11 for Unicode Version
    If (gbSoochikaVer = 5) Then
    
        Rec.Open "select count(name) as ExactExist from  dbo.syscolumns where name='chvbookno' and id=(select id from sysobjects where name='tMarriageEng')", mCnn
        
        If Rec(0) > 0 Then
          res.Open "Select chvbookno from tmarriageeng where chvackno='" & fgmarriage.TextMatrix(fgmarriage.RowSel, 17) & "'", mCnn
          If Not res.EOF Then
            frmUSevanaInward.txtBookNo.Text = IIf(IsNull(res(0)), 1, res(0)) 'Modified by Sreeja on 18.8.09
          Else
             'Modified by Soumya V S on 25.02.16
            'frmUSevanaInward.txtbookno.Text = 1
            frmUSevanaInward.txtBookNo.Text = fgmarriage.TextMatrix(fgmarriage.RowSel, 18)
            
          End If
          res.Close
        Else
        'Modified by Soumya V S on 25.02.16
            'frmUSevanaInward.txtbookno.Text = 1
            frmUSevanaInward.txtBookNo.Text = fgmarriage.TextMatrix(fgmarriage.RowSel, 18)
 
        End If
        Rec.Close
        
        frmUSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
        frmUSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol + 1))
        frmUSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 16)) 'Modified by Sreeja on 11.8.09
        frmUSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex
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
        
        Unload Me
    '********************************************************************************************************
    Else
        frmSevanaInward.txtEnglishname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol))
        'Added by Sreeja on 11.8.09
        frmSevanaInward.txtMalayalamname.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, SelectedCol + 1))
        '-------end
        'frmSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 10))
        frmSevanaInward.txtregno.Text = (fgmarriage.TextMatrix(fgmarriage.RowSel, 16)) 'Modified by Sreeja on 11.8.09
        'frmSevanaInward.txtbookno.Text = 1
        'Added by Sreeja on 11.8.09
    
        Rec.Open "select count(name) as ExactExist from  dbo.syscolumns where name='chvbookno' and id=(select id from sysobjects where name='tMarriageEng')", mCnn
        
        If Rec(0) > 0 Then
          res.Open "Select chvbookno from tmarriageeng where chvackno='" & fgmarriage.TextMatrix(fgmarriage.RowSel, 17) & "'", mCnn
          If Not res.EOF Then
            frmSevanaInward.txtBookNo.Text = IIf(IsNull(res(0)), 1, res(0)) 'Modified by Sreeja on 18.8.09
          Else
           'Modified by Soumya V S on 25.02.16
            ' frmSevanaInward.txtbookno.Text = 1
            frmSevanaInward.txtBookNo.Text = fgmarriage.TextMatrix(fgmarriage.RowSel, 18)
            
          End If
          res.Close
        Else
         'Modified by Soumya V S on 25.02.16
            ' frmSevanaInward.txtbookno.Text = 1
            frmSevanaInward.txtBookNo.Text = fgmarriage.TextMatrix(fgmarriage.RowSel, 18)
      
        
        End If
        Rec.Close
        '----------end 11.8.09
        frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified By Arun A on 22/2/2007
        'Modified by Arun A on 3.5.2006 for disabling Editing
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
        
        Unload Me
    End If
End Sub


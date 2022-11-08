VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSevanaStillBirth 
   BackColor       =   &H00C0E0FF&
   Caption         =   "STILLBIRTH REGISTER"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FIELDS TO SEARCH"
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9900
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
         Height          =   345
         Left            =   8805
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpdob 
         Height          =   345
         Left            =   7260
         TabIndex        =   15
         Top             =   1365
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   60489729
         CurrentDate     =   38708
      End
      Begin MSComCtl2.DTPicker dtpdor 
         Height          =   330
         Left            =   7260
         TabIndex        =   14
         Top             =   840
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   60489729
         CurrentDate     =   38708
      End
      Begin VB.ComboBox cmbsex 
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
         Left            =   2175
         TabIndex        =   13
         Top             =   1335
         Width           =   2385
      End
      Begin VB.TextBox txtplace 
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
         Left            =   2175
         TabIndex        =   11
         Top             =   1800
         Width           =   2385
      End
      Begin VB.CommandButton cmdsearch 
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
         Height          =   375
         Left            =   8385
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
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
         Height          =   345
         Left            =   7260
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox txtmother 
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
         Left            =   2175
         TabIndex        =   2
         Top             =   840
         Width           =   2385
      End
      Begin VB.TextBox txtfather 
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
         Left            =   2175
         TabIndex        =   1
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   8265
         TabIndex        =   17
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Place Of Still Birth"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1845
         Width           =   1695
      End
      Begin VB.Label Label6 
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
         Height          =   225
         Left            =   4995
         TabIndex        =   9
         Top             =   420
         Width           =   1890
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registration Date"
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
         Height          =   240
         Left            =   4995
         TabIndex        =   8
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Still Birth Date"
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
         Left            =   4995
         TabIndex        =   7
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sex Of Stillbirth"
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
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Mother"
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
         TabIndex        =   5
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Father"
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
         TabIndex        =   4
         Top             =   435
         Width           =   1695
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid fgstill 
      Height          =   4200
      Left            =   15
      TabIndex        =   18
      Top             =   2520
      Width           =   9870
      _cx             =   17410
      _cy             =   7408
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanaStillBirth.frx":0000
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
Attribute VB_Name = "frmSevanaStillBirth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varay As Variant
Dim vOut As Variant

Private Sub cmdSearch_Click()
    fgstill.Clear 1
    fgstill.Rows = 1
    Dim con3 As New ADODB.Connection
    Dim objDB As New clsDB
    Dim rs3 As New ADODB.Recordset
    Dim gen As String
    
    If objDB.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False Then
        MsgBox "Sevana connection failed", vbDefaultButton1
        Exit Sub
    End If
    
    If cmbsex.ItemData(cmbsex.ListIndex) = 1 Then
     gen = "male"
    ElseIf cmbsex.ItemData(cmbsex.ListIndex) = 2 Then
     gen = "Female"
    End If
    
     If Trim(txtfather.Text) = "" And Trim(txtmother.Text) = "" And Trim(txtPlace.Text) = "" And cmbsex.ItemData(cmbsex.ListIndex) = 0 And IIf(IsNull(dtpdob.value), "", dtpdob.value) = "" And IIf(IsNull(dtpDOR.value), "", dtpDOR.value) = "" And txtRegNo.Text = "" And (txtYear.Text) = "" Then
          Qry = "set dateformat dmy select FatherEng,FatherMal, MotherEng,MotherMal, Sex,Place,chvRegnNo, dtmRegnDate,StillBirthDate from StillBIRTHSEARCHVIEW1 where year(dtmRegnDate)=DATEPART(year, GETDATE())  "
          Else
            
          
          Qry = "set dateformat dmy select FatherEng,FatherMal, MotherEng,MotherMal, Sex,Place,chvRegnNo, dtmRegnDate,StillBirthDate from STILLBIRTHSEARCHVIEW1 where "
          Qry = Qry + "     FatherEng  like '" & Trim(txtfather.Text) & "%' " ' and Father like '" & fname & "%' "
          
          Qry = Qry + " and MotherEng like '" & Trim(txtmother.Text) & "%' "
         
          Qry = Qry + " and Place  like '" & Trim(txtPlace.Text) & "%' "
          Qry = Qry + " and Sex like '" & gen & "%' "
          Qry = Qry + " and chvRegnNo like '" & txtRegNo.Text & "%' "
          Qry = Qry + " and  year(dtmRegnDate) like '" & (txtYear.Text) & "%' "
    
          If IIf(IsNull(dtpdob.value), "", dtpdob.value) <> "" Then
             Qry = Qry + " and StillBirthDate = '" & IIf(IsNull(dtpdob.value), "", dtpdob.value) & "' "
          End If
          If IIf(IsNull(dtpDOR.value), "", dtpDOR.value) <> "" Then
             Qry = Qry + " and dtmRegnDate = '" & IIf(IsNull(dtpDOR.value), "", dtpDOR.value) & "' "
          End If
          
       End If
          
     Set rs3 = con3.Execute(Qry)
     
     Dim J As Integer
     J = 1
     
    While rs3.EOF = False
    fgstill.Rows = fgstill.Rows + 1
    fgstill.TextMatrix(J, 0) = J
    fgstill.TextMatrix(J, 1) = rs3(0)
    fgstill.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
    fgstill.TextMatrix(J, 2) = rs3(1)
    fgstill.TextMatrix(J, 3) = rs3(2)
    fgstill.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
    fgstill.TextMatrix(J, 4) = rs3(3)
    fgstill.TextMatrix(J, 5) = rs3(4)
    fgstill.TextMatrix(J, 6) = rs3(5)
    fgstill.TextMatrix(J, 7) = rs3(6)
    
    If IsNull(rs3(7)) Then
    fgstill.TextMatrix(J, 8) = "Not Given"
    Else
    fgstill.TextMatrix(J, 8) = rs3(7)
    End If
    fgstill.TextMatrix(J, 9) = rs3(8)
    
    
    J = J + 1
    rs3.MoveNext
    Wend
con3.Close
End Sub


Private Sub fgstill_DblClick()
 Dim SelectedCol As Long
    Dim RelationshipComboIndex As Integer
  '-------------Modified by Arun A on 22/2/2007 for the Column wise Selection of Records--------
  Select Case fgstill.ColSel
      Case 1, 2
              SelectedCol = 1 'Engfather
              RelationshipComboIndex = 0
      Case 3, 4
              SelectedCol = 2 'Engmother
              RelationshipComboIndex = 1
      Case Else
              SelectedCol = 1 ' Default Name of father
              RelationshipComboIndex = 0
    End Select
    
  
  '-------------End--------------------------------------------------------------------------
'*******************************************************************************************
' Added By Akheel 09.03.2011 for Unicode Version
If (gbSoochikaVer = 5) Then
     If (fgstill.TextMatrix(fgstill.RowSel, 3)) <> "Not Given" Then
        frmUSevanaInward.txtMalayalamname.Text = (fgstill.TextMatrix(fgstill.RowSel, SelectedCol + 1))
        frmUSevanaInward.txtEnglishname.Text = (fgstill.TextMatrix(fgstill.RowSel, SelectedCol))
        frmUSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified by Arun A on 22/2/2007
    Else
        frmUSevanaInward.txtMalayalamname.Text = (fgstill.TextMatrix(fgstill.RowSel, 4))
        frmUSevanaInward.txtEnglishname.Text = (fgstill.TextMatrix(fgstill.RowSel, 3))
        frmUSevanaInward.cboRelationship.ListIndex = 1
    End If
Else
    If (fgstill.TextMatrix(fgstill.RowSel, 3)) <> "Not Given" Then
        frmSevanaInward.txtMalayalamname.Text = (fgstill.TextMatrix(fgstill.RowSel, SelectedCol + 1))
        frmSevanaInward.txtEnglishname.Text = (fgstill.TextMatrix(fgstill.RowSel, SelectedCol))
        frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified by Arun A on 22/2/2007
    Else
        frmSevanaInward.txtMalayalamname.Text = (fgstill.TextMatrix(fgstill.RowSel, 4))
        frmSevanaInward.txtEnglishname.Text = (fgstill.TextMatrix(fgstill.RowSel, 3))
        frmSevanaInward.cboRelationship.ListIndex = 1
    End If
End If
 Unload Me
End Sub
Private Sub Form_Load()
cmbsex.Clear
    cmbsex.AddItem "...."
    cmbsex.ItemData(cmbsex.NewIndex) = 0
    cmbsex.AddItem "Male"
    cmbsex.ItemData(cmbsex.NewIndex) = 1
    cmbsex.AddItem "Female"
    cmbsex.ItemData(cmbsex.NewIndex) = 2
    cmbsex.ListIndex = 0
End Sub

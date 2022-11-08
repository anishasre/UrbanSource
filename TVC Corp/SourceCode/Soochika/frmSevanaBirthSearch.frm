VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSevanaBirthSearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Birth Register"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
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
   ScaleHeight     =   7005
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsearch 
      Appearance      =   0  'Flat
      Caption         =   "Search "
      Height          =   360
      Left            =   9015
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1860
      Width           =   1110
   End
   Begin VB.Frame Frabirthsearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Searching Fields"
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   10920
      Begin VB.TextBox txtbookno 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6510
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1500
         Width           =   945
      End
      Begin VB.TextBox txtyear 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3675
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1470
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpregdate 
         Height          =   285
         Left            =   9165
         TabIndex        =   9
         Top             =   1500
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   62259201
         CurrentDate     =   38708
      End
      Begin MSComCtl2.DTPicker dtpdob 
         Height          =   285
         Left            =   1410
         TabIndex        =   4
         Top             =   1065
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   62259201
         CurrentDate     =   38708
      End
      Begin VB.ComboBox cmbsex 
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
         Left            =   1410
         TabIndex        =   2
         Top             =   630
         Width           =   3120
      End
      Begin VB.TextBox txtregno 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1395
         TabIndex        =   6
         Top             =   1440
         Width           =   1725
      End
      Begin VB.TextBox txtnameofmother 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6495
         TabIndex        =   3
         Top             =   660
         Width           =   4305
      End
      Begin VB.TextBox txtnameoffather 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6495
         TabIndex        =   1
         Top             =   240
         Width           =   4305
      End
      Begin VB.TextBox txtplaceofbirth 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6495
         TabIndex        =   5
         Top             =   1095
         Width           =   4305
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         Top             =   240
         Width           =   3120
      End
      Begin VB.Label lblCnt 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   1920
         Width           =   3060
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Book No"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5070
         TabIndex        =   21
         Top             =   1545
         Width           =   735
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3180
         TabIndex        =   20
         Top             =   1515
         Width           =   450
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Registration Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7590
         TabIndex        =   19
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reg No"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   1485
         Width           =   690
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Mother"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5070
         TabIndex        =   17
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Father"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5070
         TabIndex        =   16
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Place Of Birth"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5070
         TabIndex        =   15
         Top             =   1125
         Width           =   1260
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date Of Birth"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sex Of Child"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   705
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name Of Child"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   75
         TabIndex        =   12
         Top             =   285
         Width           =   1320
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid flxbirth 
      Height          =   4575
      Left            =   60
      TabIndex        =   23
      Top             =   2385
      Width           =   10890
      _cx             =   19209
      _cy             =   8070
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
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSevanaBirthSearch.frx":0000
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
Attribute VB_Name = "frmSevanaBirthSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varay As Variant
Dim vOut As Variant
Private Sub cmdSearch_Click()
flxbirth.Clear 1
ReDim varay(7)
flxbirth.Rows = 1
'**********************************************
Dim con3 As New ADODB.Connection
Dim objDb As New clsDB
Dim rs3 As New ADODB.Recordset
Dim gen As String

        
If objDb.CreateNewConnection(con3, enuSourceString.SevanaRegn) = False Then
    MsgBox "Sevana connection failed", vbDefaultButton1
    Exit Sub
End If

If cmbsex.ItemData(cmbsex.ListIndex) = 1 Then
 gen = "male"
ElseIf cmbsex.ItemData(cmbsex.ListIndex) = 2 Then
 gen = "Female"
End If


Dim Qry

 If Trim(txtname.Text) = "" And cmbsex.ItemData(cmbsex.ListIndex) = 0 And IIf(IsNull(dtpdob.value), "", dtpdob.value) = "" And Trim(txtplaceofbirth.Text) = "" And Trim(txtnameoffather.Text) = "" And Trim(txtnameofmother.Text) = "" And IIf(IsNull(dtpregdate.value), "", dtpregdate.value) = "" And (txtregno.Text) = "" And (txtyear.Text) = "" And (txtbookno.Text) = "" Then  'misha
  '    Qry = "set dateformat dmy select top 100 Child,MalChild,Father,MalFather,Mother,MalMother,Sex,BirthDate,Place,dtmRegnDate,chvRegnNo,bookno from selectbirthdataVIEW1 "
     
   Qry = "set dateformat dmy select top 100 isnull(tBirthRep.chvEngChild, 'Not Given')as Child,"
 Qry = Qry + "  isnull(tBirthRep.chvMalChild, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalChild,"
 Qry = Qry + "  isnull(tBirthRep.chvEngFather, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Father,"
 Qry = Qry + "  isnull(tBirthRep.chvMalFather, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalFather,"
 Qry = Qry + "  isnull(tBirthRep.chvEngMother, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Mother,"
 Qry = Qry + "  isnull(tBirthRep.chvMalMother, 'tcJs¸Sp¯nbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalMother,"
 Qry = Qry + "  mGender.chvEngGender AS Sex,tBirthRep.dtmBirthDate1 as BirthDate,"
 Qry = Qry + "  CASE tBirthRep.intBirthPlaceID when  1   then isnull(mHospital.chvEngHospital, 'Not Given') when 2  then isnull(mInstitutionDetails.chvEngInstName, 'Not Given') "
  Qry = Qry + " when 3 then isnull(tBirthRep.chvBPEngHouse, 'Not Given') when 4 then isnull(tbirthrep.chvotherdetails, 'Not Given') else 'Not Given' End as Place,"
 
  Qry = Qry + " tBirthRep.dtmRegnDate as dtmRegnDate,isnull( tBirthRep.chvRegnNo, 'Not Given') as chvRegnNo,"
 Qry = Qry + "  isnull(tBirthRep.chvBookNo,'1') as bookno  from tBirthRep LEFT OUTER JOIN mPlace ON tBirthRep.intBirthPlaceID = mPlace.intID LEFT OUTER JOIN  mGender ON tBirthRep.inyGender = mGender.inyID LEFT OUTER JOIN"
 
  Qry = Qry + "  mHospital ON tBirthRep.intHospitalID = mHospital.intID LEFT OUTER JOIN  mInstitutionDetails ON tBirthRep.intInstID = mInstitutionDetails.intID LEFT OUTER JOIN "
   Qry = Qry + "  tInward ON tInward.chvAckNo =  tBirthRep.chvAckno where  isnull(bitadopted,0)=0"
                    
                    
                     
 
                     
     
     Else
                    
      Qry = "set dateformat dmy select  isnull(tBirthRep.chvEngChild, 'Not Given')as Child,"
 Qry = Qry + "  isnull(tBirthRep.chvMalChild, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalChild,"
 Qry = Qry + "  isnull(tBirthRep.chvEngFather, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Father,"
 Qry = Qry + "  isnull(tBirthRep.chvMalFather, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalFather,"
 Qry = Qry + "  isnull(tBirthRep.chvEngMother, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Mother,"
 Qry = Qry + "  isnull(tBirthRep.chvMalMother, 'tcJs¸Sp¯nbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalMother,"
 Qry = Qry + "  mGender.chvEngGender AS Sex,tBirthRep.dtmBirthDate1 as BirthDate,"
 Qry = Qry + "  CASE tBirthRep.intBirthPlaceID when  1   then isnull(mHospital.chvEngHospital, 'Not Given') when 2  then isnull(mInstitutionDetails.chvEngInstName, 'Not Given') "
  Qry = Qry + " when 3 then isnull(tBirthRep.chvBPEngHouse, 'Not Given') when 4 then isnull(tbirthrep.chvotherdetails, 'Not Given') else 'Not Given' End as Place,"
  Qry = Qry + " tBirthRep.dtmRegnDate as dtmRegnDate,isnull( tBirthRep.chvRegnNo, 'Not Given') as chvRegnNo,"
 Qry = Qry + "  isnull(tBirthRep.chvBookNo,'1') as bookno  from tBirthRep LEFT OUTER JOIN mPlace ON tBirthRep.intBirthPlaceID = mPlace.intID LEFT OUTER JOIN  mGender ON tBirthRep.inyGender = mGender.inyID LEFT OUTER JOIN"
 
  Qry = Qry + "  mHospital ON tBirthRep.intHospitalID = mHospital.intID LEFT OUTER JOIN  mInstitutionDetails ON tBirthRep.intInstID = mInstitutionDetails.intID LEFT OUTER JOIN "
   Qry = Qry + "  tInward ON tInward.chvAckNo =  tBirthRep.chvAckno where  isnull(bitadopted,0)=0"
                    
 Qry = Qry + " and  mGender.chvEngGender like '" & gen & "%' "
 
      If IIf(IsNull(txtname.Text), "", txtname.Text) <> "" Then
      Qry = Qry + "  and tBirthRep.chvEngChild  like '" & Trim(txtname.Text) & "%' "
       End If
      If IIf(IsNull(txtnameoffather.Text), "", txtnameoffather.Text) <> "" Then
      Qry = Qry + " and tBirthRep.chvEngFather like '" & Trim(txtnameoffather.Text) & "%' "
       End If
       If IIf(IsNull(txtnameofmother.Text), "", txtnameofmother.Text) <> "" Then
     Qry = Qry + " and tBirthRep.chvEngMother like '" & Trim(txtnameofmother.Text) & "%' "
      End If
    If IIf(IsNull(txtplaceofbirth.Text), "", txtplaceofbirth.Text) <> "" Then
    'Qry = Qry + " and Place like '" & Trim(txtplaceofbirth.Text) & "%' "
    Qry = Qry + " and (mHospital.chvEngHospital like '" & Trim(txtplaceofbirth.Text) & "%' or mInstitutionDetails.chvEngInstName  like '" & Trim(txtplaceofbirth.Text) & "%' or tBirthRep.chvBPEngHouse like '" & Trim(txtplaceofbirth.Text) & "%' or tbirthrep.chvotherdetails like '" & Trim(txtplaceofbirth.Text) & "%')"

    End If
      
  
     
     If IIf(IsNull(txtregno.Text), "", txtregno.Text) <> "" Then
      Qry = Qry + " and tBirthRep.chvRegnNo like '" & txtregno.Text & "%' "
     End If
      If IIf(IsNull(txtyear.Text), "", txtyear.Text) <> "" Then
      Qry = Qry + " and year(tBirthRep.dtmRegnDate) like '" & (txtyear.Text) & "%' "
      End If
      ' added bor book no manu 1-4-2006
      If IIf(IsNull(txtbookno.Text), "", txtbookno.Text) <> "" Then 'modified by Misha.S.V on 6 4 2006
         Qry = Qry + " and bookno = '" & IIf(IsNull(txtbookno.Text), "", txtbookno.Text) & "' "
      End If
      'chnaged Jun 06
      If IIf(IsNull(dtpdob.value), "", dtpdob.value) <> "" Then 'misha
         Qry = Qry + " and year(tBirthRep.dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Year(dtpdob.value)) & "' and Month(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Month(dtpdob.value)) & "' and day(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Day(dtpdob.value)) & "' " 'misha
      End If
      If IIf(IsNull(dtpregdate.value), "", dtpregdate.value) <> "" Then
         Qry = Qry + " and year(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Year(dtpregdate.value)) & "' and Month(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Month(dtpregdate.value)) & "' and Day(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Day(dtpregdate.value)) & "' "
      End If
      
      
       'CHANGED on Jul added on soumya vs
       '**************Adoption*********************
   Qry = Qry + "  Union All select isnull(tadoption.chvEngChild, 'Not Given') As Child,"
   Qry = Qry + "  isnull(tadoption.chvMalChild, '\ðInbn«nñ') As MalChild,"
   Qry = Qry + "  isnull(tadoption.chvGuardName1, 'Not Given') As Father,"
   Qry = Qry + "  isnull(tadoption.chvMalGuardName1, '\ðInbn«nñ') As MalFather,"
   Qry = Qry + "  isnull(tadoption.chvGuardName2, 'Not Given') As Mother,"
   Qry = Qry + "  isnull(tadoption.chvMalGuardName2, 'tcJs¸Sp¯nbn«nñ') As MalMother, "
     Qry = Qry + "  dbo.mGender.chvEngGender AS Sex,"
   Qry = Qry + "   convert(varchar,tadoption.dtmBirthDate,103)  As [BirthDate],"
   Qry = Qry + "  isnull(tadoption.chvBirthPlace,'Not Given') As Place,"
   Qry = Qry + "  convert(varchar,tBirthRep.dtmRegnDate,103) as dtmRegnDate,isnull( tBirthRep.chvRegnNo, 'Not Given') As chvRegnNo,"
   Qry = Qry + "  isnull(tBirthRep.chvBookNo,'1') As BookNo"

   
   Qry = Qry + "  FROM  tBirthRep LEFT OUTER JOIN tadoption on tBirthRep.chvackno=tadoption.chvackno LEFT OUTER JOIN mPlace ON tBirthRep.intBirthPlaceID = mPlace.intID LEFT OUTER JOIN"
   Qry = Qry + "  mGender ON tBirthRep.inyGender = mGender.inyID LEFT OUTER JOIN mHospital ON tBirthRep.intHospitalID = mHospital.intID LEFT OUTER JOIN"
   Qry = Qry + " mInstitutionDetails ON tBirthRep.intInstID = mInstitutionDetails.intID LEFT OUTER JOIN  tInward ON tInward.chvAckNo =  tBirthRep.chvAckno Where IsNull(bitadopted, 0) = 1 And tadoption.flgApproval = 1"
                    
    Qry = Qry + " and  mGender.chvEngGender like '" & gen & "%' "
                      
    If IIf(IsNull(txtname.Text), "", txtname.Text) <> "" Then
      Qry = Qry + "  and tBirthRep.chvEngChild  like '" & Trim(txtname.Text) & "%' "
       End If
      If IIf(IsNull(txtnameoffather.Text), "", txtnameoffather.Text) <> "" Then
      Qry = Qry + " and tBirthRep.chvEngFather like '" & Trim(txtnameoffather.Text) & "%' "
       End If
       If IIf(IsNull(txtnameofmother.Text), "", txtnameofmother.Text) <> "" Then
     Qry = Qry + " and tBirthRep.chvEngMother like '" & Trim(txtnameofmother.Text) & "%' "
      End If
    If IIf(IsNull(txtplaceofbirth.Text), "", txtplaceofbirth.Text) <> "" Then
    'Qry = Qry + " and Place like '" & Trim(txtplaceofbirth.Text) & "%' "
    Qry = Qry + " and (mHospital.chvEngHospital like '" & Trim(txtplaceofbirth.Text) & "%' or mInstitutionDetails.chvEngInstName  like '" & Trim(txtplaceofbirth.Text) & "%' or tBirthRep.chvBPEngHouse like '" & Trim(txtplaceofbirth.Text) & "%' or tbirthrep.chvotherdetails like '" & Trim(txtplaceofbirth.Text) & "%')"

    End If
      
  
     
     If IIf(IsNull(txtregno.Text), "", txtregno.Text) <> "" Then
      Qry = Qry + " and tBirthRep.chvRegnNo like '" & txtregno.Text & "%' "
     End If
      If IIf(IsNull(txtyear.Text), "", txtyear.Text) <> "" Then
      Qry = Qry + " and year(tBirthRep.dtmRegnDate) like '" & (txtyear.Text) & "%' "
      End If
      ' added for book no
      If IIf(IsNull(txtbookno.Text), "", txtbookno.Text) <> "" Then 'modified by Misha.S.V on 6 4 2006
         Qry = Qry + " and bookno = '" & IIf(IsNull(txtbookno.Text), "", txtbookno.Text) & "' "
      End If
      'chnaged Jun 06
      If IIf(IsNull(dtpdob.value), "", dtpdob.value) <> "" Then 'misha
         Qry = Qry + " and year(tBirthRep.dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Year(dtpdob.value)) & "' and Month(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Month(dtpdob.value)) & "' and day(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Day(dtpdob.value)) & "' " 'misha
      End If
      If IIf(IsNull(dtpregdate.value), "", dtpregdate.value) <> "" Then
         Qry = Qry + " and year(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Year(dtpregdate.value)) & "' and Month(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Month(dtpregdate.value)) & "' and Day(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Day(dtpregdate.value)) & "' "
      End If
        '**************Adoption*********************
      
 End If
    rs3.Open Qry, con3, adOpenStatic
     Set rs3 = con3.Execute(Qry)
'End If


''
''If Trim(txtname.Text) = "" And cmbsex.ItemData(cmbsex.ListIndex) = 0 And IIf(IsNull(dtpdob.value), "", dtpdob.value) = "" And Trim(txtplaceofbirth.Text) = "" And Trim(txtnameoffather.Text) = "" And Trim(txtnameofmother.Text) = "" And IIf(IsNull(dtpregdate.value), "", dtpregdate.value) = "" And (txtregno.Text) = "" And (txtyear.Text) = "" And (txtbookno.Text) = "" Then  'misha
''  '    Qry = "set dateformat dmy select top 100 Child,MalChild,Father,MalFather,Mother,MalMother,Sex,BirthDate,Place,dtmRegnDate,chvRegnNo,bookno from selectbirthdataVIEW1 "
''
''   Qry = "set dateformat dmy select top 100 isnull(tBirthRep.chvEngChild, 'Not Given')as Child,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalChild, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalChild,"
'' Qry = Qry + "  isnull(tBirthRep.chvEngFather, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Father,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalFather, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalFather,"
'' Qry = Qry + "  isnull(tBirthRep.chvEngMother, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Mother,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalMother, 'tcJs¸Sp¯nbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalMother,"
'' Qry = Qry + "  '' as Sex,tBirthRep.dtmBirthDate1 as BirthDate,"
'' Qry = Qry + "  '' as Place, tBirthRep.dtmRegnDate as dtmRegnDate,isnull( tBirthRep.chvRegnNo, 'Not Given') as chvRegnNo,"
'' Qry = Qry + "  isnull(tBirthRep.chvBookNo,'1') as bookno  from tBirthRep where  isnull(bitadopted,0)=0 "
''
''     Else
''
''      Qry = "set dateformat dmy select  isnull(tBirthRep.chvEngChild, 'Not Given')as Child,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalChild, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalChild,"
'' Qry = Qry + "  isnull(tBirthRep.chvEngFather, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Father,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalFather, '\ðInbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalFather,"
'' Qry = Qry + "  isnull(tBirthRep.chvEngMother, 'Not Given')  collate SQL_Latin1_General_CP1_CI_AS as Mother,"
'' Qry = Qry + "  isnull(tBirthRep.chvMalMother, 'tcJs¸Sp¯nbn«nñ')  collate SQL_Latin1_General_CP1_CI_AS as MalMother,"
'' Qry = Qry + "  '' as Sex,tBirthRep.dtmBirthDate1 as BirthDate,"
'' Qry = Qry + "  '' as Place, tBirthRep.dtmRegnDate as dtmRegnDate,isnull( tBirthRep.chvRegnNo, 'Not Given') as chvRegnNo,"
'' Qry = Qry + "  isnull(tBirthRep.chvBookNo,'1') as bookno  from tBirthRep where  isnull(bitadopted,0)=0 "
''     'Qry = Qry + " Sex like '" & gen & "%' "
''      If IIf(IsNull(txtname.Text), "", txtname.Text) <> "" Then
''      Qry = Qry + "  and tBirthRep.chvEngChild  like '" & Trim(txtname.Text) & "%' "
''       End If
''      If IIf(IsNull(txtnameoffather.Text), "", txtnameoffather.Text) <> "" Then
''      Qry = Qry + " and tBirthRep.chvEngFather like '" & Trim(txtnameoffather.Text) & "%' "
''       End If
''       If IIf(IsNull(txtnameofmother.Text), "", txtnameofmother.Text) <> "" Then
''     Qry = Qry + " and tBirthRep.chvEngMother like '" & Trim(txtnameofmother.Text) & "%' "
''      End If
''   '  If IIf(IsNull(txtplaceofbirth.Text), "", txtplaceofbirth.Text) <> "" Then
''      'Qry = Qry + " and Place like '" & Trim(txtplaceofbirth.Text) & "%' "
''    ' End If
''
''
''
''     If IIf(IsNull(txtregno.Text), "", txtregno.Text) <> "" Then
''      Qry = Qry + " and tBirthRep.chvRegnNo like '" & txtregno.Text & "%' "
''     End If
''      If IIf(IsNull(txtyear.Text), "", txtyear.Text) <> "" Then
''      Qry = Qry + " and year(tBirthRep.dtmRegnDate) like '" & (txtyear.Text) & "%' "
''      End If
''      ' added bor book no manu 1-4-2006
''      If IIf(IsNull(txtbookno.Text), "", txtbookno.Text) <> "" Then 'modified by Misha.S.V on 6 4 2006
''         Qry = Qry + " and bookno = '" & IIf(IsNull(txtbookno.Text), "", txtbookno.Text) & "' "
''      End If
''      If IIf(IsNull(dtpdob.value), "", dtpdob.value) <> "" Then 'misha
''         Qry = Qry + " and year(tBirthRep.dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Year(dtpdob.value)) & "' and Month(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Month(dtpdob.value)) & "' and day(dtmBirthDate1) = '" & IIf(IsNull(dtpdob.value), "", Day(dtpdob.value)) & "' " 'misha
''      End If
''      If IIf(IsNull(dtpregdate.value), "", dtpregdate.value) <> "" Then
''         Qry = Qry + " and year(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Year(dtpregdate.value)) & "' and Month(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Month(dtpregdate.value)) & "' and Day(tBirthRep.dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Day(dtpregdate.value)) & "' "
''      End If
'' End If
''    rs3.Open Qry, con3, adOpenStatic
''     Set rs3 = con3.Execute(Qry)

' If Trim(txtname.Text) = "" And cmbsex.ItemData(cmbsex.ListIndex) = 0 And IIf(IsNull(dtpdob.value), "", dtpdob.value) = "" And Trim(txtplaceofbirth.Text) = "" And Trim(txtnameoffather.Text) = "" And Trim(txtnameofmother.Text) = "" And IIf(IsNull(dtpregdate.value), "", dtpregdate.value) = "" And (txtregno.Text) = "" And (txtyear.Text) = "" And (txtbookno.Text) = "" Then  'misha
'      Qry = "set dateformat dmy select top 100 Child,MalChild,Father,MalFather,Mother,MalMother,Sex,BirthDate,Place,dtmRegnDate,chvRegnNo,bookno from selectbirthdataVIEW1 "
'     Else
'
'      Qry = "set dateformat dmy select  Child,MalChild,Father,MalFather,Mother,MalMother,Sex,BirthDate,Place,dtmRegnDate,chvRegnNo,bookno  from selectbirthdataVIEW1 where "
'      Qry = Qry + " Sex like '" & gen & "%' "
'      If IIf(IsNull(txtname.Text), "", txtname.Text) <> "" Then
'      Qry = Qry + "  and child  like '" & Trim(txtname.Text) & "%' "
'       End If
'      If IIf(IsNull(txtnameoffather.Text), "", txtnameoffather.Text) <> "" Then
'      Qry = Qry + " and father like '" & Trim(txtnameoffather.Text) & "%' "
'       End If
'       If IIf(IsNull(txtnameofmother.Text), "", txtnameofmother.Text) <> "" Then
'     Qry = Qry + " and Mother like '" & Trim(txtnameofmother.Text) & "%' "
'      End If
'     If IIf(IsNull(txtplaceofbirth.Text), "", txtplaceofbirth.Text) <> "" Then
'      Qry = Qry + " and Place like '" & Trim(txtplaceofbirth.Text) & "%' "
'     End If
'
'
'
'     If IIf(IsNull(txtregno.Text), "", txtregno.Text) <> "" Then
'      Qry = Qry + " and chvRegnNo like '" & txtregno.Text & "%' "
'     End If
'      If IIf(IsNull(txtyear.Text), "", txtyear.Text) <> "" Then
'      Qry = Qry + " and year(dtmRegnDate) like '" & (txtyear.Text) & "%' "
'      End If
'      ' added bor book no manu 1-4-2006
'      If IIf(IsNull(txtbookno.Text), "", txtbookno.Text) <> "" Then 'modified by Misha.S.V on 6 4 2006
'         Qry = Qry + " and bookno = '" & IIf(IsNull(txtbookno.Text), "", txtbookno.Text) & "' "
'      End If
'      If IIf(IsNull(dtpdob.value), "", dtpdob.value) <> "" Then 'misha
'         Qry = Qry + " and year(BirthDate) = '" & IIf(IsNull(dtpdob.value), "", Year(dtpdob.value)) & "' and Month(BirthDate) = '" & IIf(IsNull(dtpdob.value), "", Month(dtpdob.value)) & "' and day(BirthDate) = '" & IIf(IsNull(dtpdob.value), "", Day(dtpdob.value)) & "' " 'misha
'      End If
'      If IIf(IsNull(dtpregdate.value), "", dtpregdate.value) <> "" Then
'         Qry = Qry + " and year(dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Year(dtpregdate.value)) & "' and Month(dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Month(dtpregdate.value)) & "' and Day(dtmRegnDate) = '" & IIf(IsNull(dtpregdate.value), "", Day(dtpregdate.value)) & "' "
'      End If
' End If
'    rs3.Open Qry, con3, adOpenStatic
'     Set rs3 = con3.Execute(Qry)
'

'****************************************************




Dim J As Integer
 J = 1
 
While rs3.EOF = False
    flxbirth.Rows = flxbirth.Rows + 1
    flxbirth.TextMatrix(J, 0) = J
    flxbirth.TextMatrix(J, 1) = rs3(0)
    flxbirth.Cell(flexcpFontName, J, 2) = "ML-TTRevathi"
    flxbirth.TextMatrix(J, 2) = rs3(1)
    flxbirth.TextMatrix(J, 3) = rs3(2)
    flxbirth.Cell(flexcpFontName, J, 4) = "ML-TTRevathi"
    flxbirth.TextMatrix(J, 4) = rs3(3)
    flxbirth.TextMatrix(J, 5) = rs3(4)
    flxbirth.Cell(flexcpFontName, J, 6) = "ML-TTRevathi"
    flxbirth.TextMatrix(J, 6) = rs3(5)
    flxbirth.TextMatrix(J, 7) = IIf(IsNull(rs3(6)), "", rs3(6))
    If IsNull(rs3(7)) Then ' modified by misha
            flxbirth.TextMatrix(J, 8) = "Not Given"
    Else
            flxbirth.TextMatrix(J, 8) = rs3(7)
    End If
    flxbirth.TextMatrix(J, 9) = rs3(8)
    If IsNull(rs3(9)) Then
        flxbirth.TextMatrix(J, 10) = "Not Given"
    Else
        flxbirth.TextMatrix(J, 10) = rs3(9)
    End If
    flxbirth.TextMatrix(J, 11) = rs3(10)
    flxbirth.TextMatrix(J, 12) = rs3(11)
    
    J = J + 1
    rs3.MoveNext
Wend
   lblCnt.Caption = J - 1
   lblCnt.Caption = lblCnt.Caption + "  Record selected"
con3.Close



End Sub
Private Sub flxbirth_DblClick()
 Dim SelectedCol As Long
 Dim RelationshipComboIndex As Integer
 Dim BirthDate As Integer
 Dim Age As Integer
 If ((flxbirth.TextMatrix(flxbirth.RowSel, 8)) = "Not Given") Then
 Exit Sub
 End If
 BirthDate = Year((flxbirth.TextMatrix(flxbirth.RowSel, 8)))
 Age = Year(gbTransactionDate) - BirthDate
 If flxbirth.Rows = 1 Then
    MsgBox "No Selected records ", vbInformation
    Exit Sub
 End If
  '-------------Modified by Arun A on 22/2/2007 for the Column wise Selection of Records--------
    Select Case flxbirth.ColSel
      Case 1, 2
              SelectedCol = 1 'EngChild
              RelationshipComboIndex = 0
      Case 3, 4
              SelectedCol = 3 'EngFather
              RelationshipComboIndex = 1
      Case 5, 6
              SelectedCol = 5 'EngMother
              RelationshipComboIndex = 2
      Case Else
              SelectedCol = 3 ' Default is Name of Father
              RelationshipComboIndex = 1
    End Select
    
  
  '-------------End--------------------------------------------------------------------------
  
  '******************************************************************************************
  ' Added by Akheel for Unicode Version
  
  If (gbSoochikaVer = 5) Then
  
  If (frmUSevanaInward.txtSubTypeID = 74 Or frmUSevanaInward.txtSubTypeID = 59 Or frmUSevanaInward.txtSubTypeID = 8 Or frmUSevanaInward.txtSubTypeID = 120 Or frmUSevanaInward.txtSubTypeID = 122) Then   'Ranjitha 09/10
  
  
    If ((flxbirth.TextMatrix(flxbirth.RowSel, 1)) <> "Not Given" And (flxbirth.TextMatrix(flxbirth.RowSel, 2)) = "\ðInbn«nñ") Then
    MsgBox "Child Name Already Given"
    Exit Sub
    End If
    If ((flxbirth.TextMatrix(flxbirth.RowSel, 1)) <> "Not Given") Then
    MsgBox "Child Name Already Given"
    Exit Sub
    End If
    
    If ((flxbirth.TextMatrix(flxbirth.RowSel, 2)) <> "\ðInbn«nñ") Then
    MsgBox "Child Name Already Given"
    Exit Sub
    End If
  
    If (frmUSevanaInward.txtSubTypeID = 74) Then
     If (Age > 1) Then
      MsgBox "Date difference of Birth date  and current date should be within 1 year"
      Exit Sub
     End If
    End If
    
    'added on soumya dte:25/09/2015
    'purpose:cant take non registered records for name inclusion
    If (frmUSevanaInward.txtSubTypeID = 8 Or frmUSevanaInward.txtSubTypeID = 7 Or frmUSevanaInward.txtSubTypeID = 19 Or frmUSevanaInward.txtSubTypeID = 20 Or frmUSevanaInward.txtSubTypeID = 58 Or frmUSevanaInward.txtSubTypeID = 59 Or frmUSevanaInward.txtSubTypeID = 74 Or frmUSevanaInward.txtSubTypeID = 75 Or frmUSevanaInward.txtSubTypeID = 120 Or frmUSevanaInward.txtSubTypeID = 121 Or frmUSevanaInward.txtSubTypeID = 122 Or frmUSevanaInward.txtSubTypeID = 123) Then
            If ((flxbirth.TextMatrix(flxbirth.RowSel, 10)) = "Not Given") Then
            MsgBox "Non registered records"
            Exit Sub
            End If
  End If
  
    
  If (frmUSevanaInward.txtSubTypeID = 120) Then
   If (Age < 1) Then
   MsgBox "Date difference of Birth date and current date should be >1 year"
   Exit Sub
    End If
  End If
   End If
  
        'added on soumya dte:25/09/2015
    'purpose:cant take non registered records for name inclusion
    If (frmUSevanaInward.txtSubTypeID = 8 Or frmUSevanaInward.txtSubTypeID = 7 Or frmUSevanaInward.txtSubTypeID = 19 Or frmUSevanaInward.txtSubTypeID = 20 Or frmUSevanaInward.txtSubTypeID = 58 Or frmUSevanaInward.txtSubTypeID = 59 Or frmUSevanaInward.txtSubTypeID = 74 Or frmUSevanaInward.txtSubTypeID = 75 Or frmUSevanaInward.txtSubTypeID = 120 Or frmUSevanaInward.txtSubTypeID = 121 Or frmUSevanaInward.txtSubTypeID = 122 Or frmUSevanaInward.txtSubTypeID = 123) Then
            If ((flxbirth.TextMatrix(flxbirth.RowSel, 10)) = "Not Given") Then
            MsgBox "Non registered records"
            Exit Sub
            End If
  End If
  
 
  If (frmUSevanaInward.txtSubTypeID = 113 Or frmUSevanaInward.txtSubTypeID = 111 Or frmUSevanaInward.txtSubTypeID = 114 Or frmUSevanaInward.txtSubTypeID = 115 Or frmUSevanaInward.txtSubTypeID = 116 Or frmUSevanaInward.txtSubTypeID = 117 Or frmUSevanaInward.txtSubTypeID = 118 Or frmUSevanaInward.txtSubTypeID = 119) Then
  
    If (Age >= 6 And frmUSevanaInward.txtSubTypeID = 111) Then
    MsgBox "Age > 6..!! pet name correction not Allowd."
    Exit Sub
    End If
'    If (Age < 6 And frmUSevanaInward.txtSubTypeID = 113) Then
'    MsgBox "Age < 6..!! name  correction not  Allowd."
'    Exit Sub
'    End If
'    If (Age < 6 And frmUSevanaInward.txtSubTypeID = 114) Then
'    MsgBox "Age < 6..!! name  correction not  Allowd."
'    Exit Sub
'    End If
  
  If ((flxbirth.TextMatrix(flxbirth.RowSel, 1)) = "Not Given" And (flxbirth.TextMatrix(flxbirth.RowSel, 2)) = "\ðInbn«nñ") Then
  MsgBox "Child Name not Given"
  Exit Sub
  End If
  End If
  
  
  
    If (flxbirth.TextMatrix(flxbirth.RowSel, 3)) <> "Not Given" Then
'       frmUSevanaInward.txtMalayalamname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, SelectedCol + 1))
'       frmUSevanaInward.txtEnglishname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, SelectedCol))
       frmUSevanaInward.txtMalayalamname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 4))
       frmUSevanaInward.txtEnglishname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 3))
       frmUSevanaInward.txtregno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 11)) 'Modified by Misha.S.V 11 03 2006
       frmUSevanaInward.txtbookno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 12)) ' Manu on 1-4-2006
       'frmUSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified by Arun A on 22/2/2007
       frmUSevanaInward.cboRelationship.ListIndex = 1
    Else
       frmUSevanaInward.txtMalayalamname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 6))
       frmUSevanaInward.txtEnglishname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 5))
       frmUSevanaInward.cboRelationship.ListIndex = 2
       'Modified by Sreeja on 31.7.09---start
       frmUSevanaInward.txtregno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 11)) 'Modified by Misha.S.V 11 03 2006
       frmUSevanaInward.txtbookno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 12)) ' Manu on 1-4-2006
       '----------------------------------end
    End If
    'Modified by Arun A on 3.5.2006 for disabling Editing
    frmUSevanaInward.txtMalayalamname.Enabled = False
       frmUSevanaInward.txtEnglishname.Enabled = False
       frmUSevanaInward.txtregno.Enabled = False
       frmUSevanaInward.txtbookno.Enabled = False
       frmUSevanaInward.cboRelationship.Enabled = False
       
       If frmUSevanaInward.txtEnglishname.Text = "Not Given" And frmUSevanaInward.cboLanguage.ListIndex = 1 And (frmUSevanaInward.txtMalayalamname.Text <> "\ðInbn«nñ") Then

  frmUSevanaInward.cboLanguage.ListIndex = 0

End If

If frmUSevanaInward.txtMalayalamname.Text = "\ðInbn«nñ" And frmUSevanaInward.cboLanguage.ListIndex = 0 And (frmUSevanaInward.txtEnglishname.Text <> "Not Given") Then

frmUSevanaInward.cboLanguage.ListIndex = 1

End If
       
    Unload Me
    Exit Sub
  End If
  
  '******************************************************************************************
 If (flxbirth.TextMatrix(flxbirth.RowSel, 3)) <> "Not Given" Then
    frmSevanaInward.txtMalayalamname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, SelectedCol + 1))
    frmSevanaInward.txtEnglishname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, SelectedCol))
    frmSevanaInward.txtregno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 11)) 'Modified by Misha.S.V 11 03 2006
    frmSevanaInward.txtbookno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 12)) ' Manu on 1-4-2006
    frmSevanaInward.cboRelationship.ListIndex = RelationshipComboIndex 'Modified by Arun A on 22/2/2007
 Else
    frmSevanaInward.txtMalayalamname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 6))
    frmSevanaInward.txtEnglishname.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 5))
    frmSevanaInward.cboRelationship.ListIndex = 2
    'Modified by Sreeja on 31.7.09---start
    frmSevanaInward.txtregno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 11)) 'Modified by Misha.S.V 11 03 2006
    frmSevanaInward.txtbookno.Text = (flxbirth.TextMatrix(flxbirth.RowSel, 12)) ' Manu on 1-4-2006
    '----------------------------------end
 End If
 'Modified by Arun A on 3.5.2006 for disabling Editing
 frmSevanaInward.txtMalayalamname.Enabled = False
    frmSevanaInward.txtEnglishname.Enabled = False
    frmSevanaInward.txtregno.Enabled = False
    frmSevanaInward.txtbookno.Enabled = False
    frmSevanaInward.cboRelationship.Enabled = False
    
     If frmSevanaInward.txtEnglishname.Text = "Not Given" And frmSevanaInward.cboLanguage.ListIndex = 1 And (frmSevanaInward.txtMalayalamname.Text <> "\ðInbn«nñ") Then
        frmSevanaInward.cboLanguage.ListIndex = 0
    End If

    If frmSevanaInward.txtMalayalamname.Text = "\ðInbn«nñ" And frmSevanaInward.cboLanguage.ListIndex = 0 And (frmSevanaInward.txtEnglishname.Text <> "Not Given") Then
        frmSevanaInward.cboLanguage.ListIndex = 1
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


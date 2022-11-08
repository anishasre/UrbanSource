VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSoochikaSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S e a r c h"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8385
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   13425
      Begin VB.ListBox lstSubject 
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
         Height          =   2370
         Left            =   1950
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.TextBox txtSubID 
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
         Height          =   360
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtDtTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11190
         TabIndex        =   14
         Top             =   1680
         Width           =   1845
      End
      Begin VB.TextBox txtDtFrm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   13
         Top             =   1680
         Width           =   1845
      End
      Begin VB.ComboBox cmbSubType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1950
         TabIndex        =   4
         Top             =   1520
         Width           =   4575
      End
      Begin VB.CommandButton btnClear 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2850
         Width           =   1725
      End
      Begin VB.TextBox txtRefDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11190
         TabIndex        =   16
         Top             =   2100
         Width           =   1845
      End
      Begin VB.CommandButton btnClose 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2850
         Width           =   1725
      End
      Begin VB.CommandButton btnSearch 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2850
         Width           =   1725
      End
      Begin VB.ComboBox cboPriority 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   12
         Top             =   1200
         Width           =   3885
      End
      Begin VB.ComboBox cboInwardType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   11
         Top             =   720
         Width           =   3885
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "ML-TTRevathi"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         TabIndex        =   15
         Top             =   2100
         Width           =   1845
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2550
         TabIndex        =   3
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox txtlocality 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   10
         Top             =   3090
         Width           =   4515
      End
      Begin VB.TextBox txtWardNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   7
         Top             =   2700
         Width           =   1005
      End
      Begin VB.TextBox txtDoorNo2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5340
         TabIndex        =   9
         Top             =   2700
         Width           =   1125
      End
      Begin VB.TextBox txtDoorNo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4290
         TabIndex        =   8
         Top             =   2700
         Width           =   945
      End
      Begin VB.TextBox txtHName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   6
         Top             =   2310
         Width           =   4515
      End
      Begin VB.TextBox txtSender 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   5
         Top             =   1920
         Width           =   4515
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5220
         TabIndex        =   1
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtFileNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   0
         Top             =   690
         Width           =   2325
      End
      Begin MSComctlLib.ProgressBar pgbrSearch 
         Height          =   135
         Left            =   180
         TabIndex        =   37
         Top             =   3480
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VSFlex8LCtl.VSFlexGrid vsSearch 
         Height          =   4635
         Left            =   180
         TabIndex        =   39
         Top             =   3645
         Width           =   13035
         _cx             =   22992
         _cy             =   8176
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
         FormatString    =   $"frmSoochikaSearch.frx":0000
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
      Begin VB.Label lblSubtype 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sub Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   210
         TabIndex        =   36
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   90
         TabIndex        =   35
         Top             =   150
         Width           =   13230
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7050
         TabIndex        =   34
         Top             =   1208
         Width           =   885
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Correspondance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7050
         TabIndex        =   33
         Top             =   728
         Width           =   1605
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   10710
         TabIndex        =   32
         Top             =   2130
         Width           =   525
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Reference No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7050
         TabIndex        =   31
         Top             =   2160
         Width           =   1635
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   10710
         TabIndex        =   30
         Top             =   1703
         Width           =   315
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7050
         TabIndex        =   29
         Top             =   1703
         Width           =   1155
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   28
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Locality"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   27
         Top             =   3150
         Width           =   1635
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ward No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   26
         Top             =   2730
         Width           =   945
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Door No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   25
         Top             =   2730
         Width           =   1035
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "House Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   24
         Top             =   2370
         Width           =   1635
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Received From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   23
         Top             =   1980
         Width           =   1635
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4470
         TabIndex        =   22
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "File No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   21
         Top             =   720
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmSoochikaSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SevanaMainSubid As Variant
Private Sub btnClear_Click()
    ClearData
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub getSearchResult()
        Dim strQry As String
        Dim objDB As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        
        strQry = "set dateformat dmy" + Chr(13) + Chr(10) + "SELECT isnull(TblTappalDetails.FldLinkFileNo,TblTappalDetails.FldCurrentNo) as [FileNo],"
        strQry = strQry & " TblTappalDetails.FldDateOfReceipt as [DOR],left(TblTappalDetails.FldOutsideNo,10) as [RefNo],"
        strQry = strQry & " isnull(TblTappalDetails.FldSenderName,'')+' '+ isnull(TblTappalDetails.fldDesign,'')+' '+ isnull(TblTappalDetails.fldInstName,'')+' '+isnull(TblTappalDetails.FldHouseName,'')+' '+ isnull(TblTappalDetails.FldHouseNo,'')+' '+ isnull(TblTappalDetails.FldWardNo,'')+' '+ isnull(TblTappalDetails.FldLocality,'')as [Sender] ,"
        strQry = strQry & " left(TblTappalDetails.FldRemarks,40) as [Subject],chvSection,TblTappalDetails.FldFileId, "
        strQry = strQry & " (SELECT top 1 FldStatus FROM  TblTappalProcessing WHERE     FldFileId =TblTappalDetails.FldFileId order by TblTappalProcessing.FldMaxId desc) as status ,"
        strQry = strQry & " (SELECT top 1 FldNotes FROM  TblTappalProcessing WHERE     FldFileId =TblTappalDetails.FldFileId order by TblTappalProcessing.FldMaxId desc) as Note,"
        strQry = strQry & " (SELECT top 1 flgFont FROM  TblTappalProcessing WHERE     FldFileId =TblTappalDetails.FldFileId order by TblTappalProcessing.FldMaxId desc) as Flg"
        strQry = strQry & " FROM TblTappalDetails Inner join tblsection on tblsection.intID=TblTappalDetails.intForwardTo "
        strQry = strQry & " Where (TblTappalDetails.FldClosed=2 or TblTappalDetails.FldClosed =0 or TblTappalDetails.FldClosed =1 ) "
        If (txtFileNo.Text <> "") Then
            strQry = strQry & " and (FldCurrentNo='" & txtFileNo.Text & "') "
        End If
        If (txtyear.Text <> "") Then
            strQry = strQry & " and (Year(FldDateOfReceipt)='" & txtyear.Text & "') "
        End If
        If (txtSender.Text <> "") Then
            strQry = strQry & " and  (FldSenderName like'%" & txtSender.Text & "%')  "
        End If
        If (txtRefNo.Text <> "") Then
            strQry = strQry & " and (FldOutsideNo like'%" & txtRefNo.Text & "%')  "
        End If
        If (txtRefDate.Text <> "") Then
            strQry = strQry & " and  (FldRefDate='" & txtRefDate.Text & "')  "
        End If
        If (txtSubject.Text <> "") Then
            strQry = strQry & " and  (FldRemarks like '%" & Replace(txtSubject.Text, "'", "''") & "%')  "
        End If
        If (txtHName.Text <> "") Then
            strQry = strQry & " and  (FldHouseName like '%" & Replace(txtHName.Text, "'", "''") & "%')  "
        End If
        If (txtDoorNo1.Text <> "") Then
            strQry = strQry & " and  (FldHouseNo like '%" & txtDoorNo1.Text & "%')  "
        End If
        If (txtDoorNo2.Text <> "") Then
            strQry = strQry & " and  (FldHouseNo like '%" & Replace(txtDoorNo2.Text, "'", "''") & "%')  "
        End If
        If (txtWardNo.Text <> "") Then
            strQry = strQry & " and  (FldWardNo like '%" & Replace(txtWardNo.Text, "'", "''") & "%')  "
        End If
        If (txtLocality.Text <> "") Then
            strQry = strQry & " and  (FldLocality like '%" & Replace(txtLocality.Text, "'", "''") & "%') "
        End If
        If (cboInwardType.ListIndex >= 0) Then
            strQry = strQry & " and  (intInwardType='" & cboInwardType.ItemData(cboInwardType.ListIndex) & "') "
        End If
        If (cboPriority.ListIndex >= 0) Then
            strQry = strQry & " and  (FldPriority='" & cboPriority.ItemData(cboPriority.ListIndex) & "') "
        End If
        
        If (txtRefDate.Text <> "") Then
             strQry = strQry & " and TblTappalDetails.FldRefDate=convert(datetime,'" & txtRefDate.Text & "',103) "
        End If
        
        If (txtDtFrm.Text = "") And (txtDtTo.Text = "") Then
            strQry = strQry
        ElseIf (txtDtFrm.Text <> "") And (txtDtTo.Text <> "") Then
            strQry = strQry & " and TblTappalDetails.FldDateOfReceipt between convert(datetime,'" & txtDtFrm.Text & "',103) and convert(datetime,'" & txtDtTo.Text & "',103) "
        ElseIf (txtDtFrm.Text = "") Then
            strQry = strQry & " and TblTappalDetails.FldDateOfReceipt=convert(datetime,'" & txtDtTo.Text & "',103) "
        ElseIf (txtDtTo.Text = "") Then
            strQry = strQry & " and TblTappalDetails.FldDateOfReceipt=convert(datetime,'" & txtDtFrm.Text & "',103) "
        End If

        If strQry <> "" Then
             strQry = strQry & " order by FldCurrentNo"
        End If
        
        If (objDB.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
            MsgBox "Soochika Connection is not present", vbCritical, "Common"
            Exit Sub
        End If
          ' gSubSetFont vsEnclosure, 1, 2, vsEnclosure.Rows - 1, 2, "ML-TTRevathi"
            vsSearch.Rows = 2
            vsSearch.Clear 1
            vsSearch.TextMatrix(0, 3) = "Seat"
            vsSearch.TextMatrix(0, 4) = "Referecne No & Date"
            vsSearch.TextMatrix(0, 5) = "Sender"
            vsSearch.TextMatrix(0, 6) = "Subject"
            vsSearch.TextMatrix(0, 7) = "Status"
            vsSearch.TextMatrix(0, 8) = "Notes"
            vsSearch.TextMatrix(0, 9) = ""
            
            Rec.Open strQry, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                i = 0
'''''''''            vsSearch.TextMatrix(0, 7) = "Status"
'''''''''            vsSearch.TextMatrix(0, 8) = "Notes"
            Do While Not (Rec.EOF)
                vsSearch.Rows = vsSearch.Rows + 1
                vsSearch.TextMatrix(i + 1, 0) = i + 1
                vsSearch.TextMatrix(i + 1, 1) = Rec!FileNo
                vsSearch.TextMatrix(i + 1, 2) = Rec!DOR
                vsSearch.TextMatrix(i + 1, 3) = Rec!chvSection
                vsSearch.TextMatrix(i + 1, 4) = IIf(IsNull(Rec!RefNo), "", Rec!RefNo)
                vsSearch.TextMatrix(i + 1, 5) = IIf(IsNull(Rec!Sender), "", Rec!Sender)
                vsSearch.TextMatrix(i + 1, 6) = IIf(IsNull(Rec!Subject), "", Rec!Subject)
                vsSearch.TextMatrix(i + 1, 7) = IIf(IsNull(Rec!Status), "", Rec!Status)
                vsSearch.TextMatrix(i + 1, 8) = IIf(IsNull(Rec!Note), "", Rec!Note)
                vsSearch.TextMatrix(i + 1, 9) = FldFileId
                Rec.MoveNext
                i = i + 1
                Loop
            End If
            Rec.Close
            gSubSetFont vsSearch, 1, 1, vsSearch.Rows - 1, 1, "Verdana"
            gSubSetFont vsSearch, 1, 3, vsSearch.Rows - 1, 8, "Verdana"
            pgbrSearch.Value = 0
End Sub

Private Sub btnSearch_Click()
    If cmbSubType.Enabled = False Then
        getSearchResult
    Else
        GetSevanaSearch
    End If
End Sub
Public Sub GetSevanaSearch()
'''''''''''''''    Dim mSQL As String
'''''''''''''''    Dim mCnn As New ADODB.Connection
'''''''''''''''    Dim Rec As New ADODB.Recordset
'''''''''''''''    Dim Rec1 As New ADODB.Recordset
'''''''''''''''    Dim objdb As New clsDb
'''''''''''''''    Dim Count As Variant
'''''''''''''''
'''''''''''''''    objdb.CreateNewConnection mCnn, enuSourceString.SevanaRegn
'''''''''''''''    ''' For getting count for the progress bar
'''''''''''''''
'''''''''''''''    mSQL = mSQL & " select count(*) as count "
'''''''''''''''    mSQL = mSQL & " From tInward inner join mInwardTYpe on mInwardType.intID=tInward.inwRequest"
'''''''''''''''    mSQL = mSQL & " inner join mInwardSubType on mInwardSubType.intID=intInwardSubType"
'''''''''''''''    mSQL = mSQL & " where inwRequest is not null "
'''''''''''''''    If txtFileNo.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and inWNo=" & txtFileNo.Text
'''''''''''''''    End If
'''''''''''''''    If txtYear.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and year(InwDate)=" & txtYear.Text
'''''''''''''''    End If
'''''''''''''''    If txtSender.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvName='%" & txtSender.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    If SevanaMainSubid <> "" Then
'''''''''''''''        mSQL = mSQL & " and inwRequest=" & SevanaMainSubid
'''''''''''''''    End If
'''''''''''''''    If cmbSubType.ListIndex > -1 Then
'''''''''''''''        mSQL = mSQL & " and intInwardSubType=" & cmbSubType.ItemData(cmbSubType.ListIndex)
'''''''''''''''    End If
'''''''''''''''    If txtHName.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvHouseName='%" & txtHName.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    If txtDoorNo1.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvHouseNo='%" & txtDoorNo1.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    Rec.Open mSQL, mCnn
'''''''''''''''    If Not (Rec.EOF Or Rec.BOF) Then
'''''''''''''''        Count = Rec!Count
'''''''''''''''        If Count > 0 Then
'''''''''''''''            pgbrSearch.Max = Count
'''''''''''''''        End If
'''''''''''''''    Else
'''''''''''''''        Count = 0
'''''''''''''''    End If
'''''''''''''''    Rec.Close
'''''''''''''''    pgbrSearch.Min = 0
'''''''''''''''    pgbrSearch.Value = 0
'''''''''''''''    '''' code ends here
'''''''''''''''    mSQL = "set dateformat DMY "
'''''''''''''''    mSQL = mSQL & " select inWno as FileNo,convert(varchar,inwDate,103) as DOR,"
'''''''''''''''    mSQL = mSQL & " chvName as Sender,TypeofRequest as Subject,TypeofSubRequest as SubType"
'''''''''''''''    mSQL = mSQL & " From tInward inner join mInwardTYpe on mInwardType.intID=tInward.inwRequest"
'''''''''''''''    mSQL = mSQL & " inner join mInwardSubType on mInwardSubType.intID=intInwardSubType"
'''''''''''''''    mSQL = mSQL & " where inwRequest is not null "
'''''''''''''''    If txtFileNo.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and inWNo=" & txtFileNo.Text
'''''''''''''''    End If
'''''''''''''''    If txtYear.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and year(InwDate)=" & txtYear.Text
'''''''''''''''    End If
'''''''''''''''    If txtSender.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvName='%" & txtSender.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    If SevanaMainSubid <> "" Then
'''''''''''''''        mSQL = mSQL & " and inwRequest=" & SevanaMainSubid
'''''''''''''''    End If
'''''''''''''''    If cmbSubType.ListIndex > -1 Then
'''''''''''''''        mSQL = mSQL & " and intInwardSubType=" & cmbSubType.ItemData(cmbSubType.ListIndex)
'''''''''''''''    End If
'''''''''''''''    If txtHName.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvHouseName='%" & txtHName.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    If txtDoorNo1.Text <> "" Then
'''''''''''''''        mSQL = mSQL & " and chvHouseNo='%" & txtDoorNo1.Text & "%'"
'''''''''''''''    End If
'''''''''''''''    mSQL = mSQL & " order by inWno "
'''''''''''''''    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
'''''''''''''''    vsSearch.Clear 1
'''''''''''''''    vsSearch.Rows = 2
'''''''''''''''    'pgbrSearch.Max = Rec.RecordCount
'''''''''''''''    If Not (Rec.EOF Or Rec.BOF) Then
'''''''''''''''        vsSearch.TextMatrix(0, 7) = "Sub Type"
'''''''''''''''        vsSearch.TextMatrix(0, 8) = "Status"
'''''''''''''''        Do While Not Rec.EOF
'''''''''''''''            vsSearch.Rows = vsSearch.Rows + 1
'''''''''''''''                vsSearch.TextMatrix(i + 1, 0) = i + 1
'''''''''''''''                vsSearch.TextMatrix(i + 1, 1) = Rec!FileNo
'''''''''''''''                vsSearch.TextMatrix(i + 1, 2) = Rec!DOR
'''''''''''''''                vsSearch.TextMatrix(i + 1, 3) = ""
'''''''''''''''                vsSearch.TextMatrix(i + 1, 4) = ""
'''''''''''''''                vsSearch.TextMatrix(i + 1, 5) = IIf(IsNull(Rec!Sender), "", Rec!Sender)
'''''''''''''''                vsSearch.TextMatrix(i + 1, 6) = IIf(IsNull(Rec!Subject), "", Rec!Subject)
'''''''''''''''                vsSearch.TextMatrix(i + 1, 7) = IIf(IsNull(Rec!SubType), "", Rec!SubType)
'''''''''''''''                Rec1.Open "STATUSREPORT1 '" & Rec!FileNo & "','" & Year(Rec!DOR) & "'", mCnn
'''''''''''''''                If Not (Rec1.EOF Or Rec1.BOF) Then
'''''''''''''''                    vsSearch.TextMatrix(i + 1, 8) = IIf(IsNull(Rec1.Fields(0)), "", Rec1.Fields(0))
'''''''''''''''                Else
'''''''''''''''                    vsSearch.TextMatrix(i + 1, 8) = ""
'''''''''''''''                End If
'''''''''''''''                Rec1.Close
'''''''''''''''                vsSearch.TextMatrix(i + 1, 9) = ""
'''''''''''''''                If pgbrSearch.Value <> pgbrSearch.Max Then
'''''''''''''''                    pgbrSearch.Value = pgbrSearch.Value + 1
'''''''''''''''                End If
'''''''''''''''                Rec.MoveNext
'''''''''''''''                i = i + 1
'''''''''''''''        Loop
'''''''''''''''    Else
'''''''''''''''        vsSearch.Clear 1, 1
'''''''''''''''    End If
'''''''''''''''    Rec.Close
'''''''''''''''    gSubSetFont vsSearch, 1, 1, vsSearch.Rows - 1, vsSearch.Cols - 1, "Verdana"

    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim Rec1 As New ADODB.Recordset
    Dim objDB As New clsDb
    Dim count As Variant
    
    objDB.CreateNewConnection mCnn, enuSourceString.SevanaRegn
    ''' For getting count for the progress bar
         
    mSql = mSql & " select count(*) as count "
    mSql = mSql & " From tInward inner join mInwardTYpe on mInwardType.intID=tInward.inwRequest"
    mSql = mSql & " inner join mInwardSubType on mInwardSubType.intID=intInwardSubType"
    mSql = mSql & " left join tReceiptDetails on tReceiptDetails.intInwardID=tInward.intid"
    mSql = mSql & " where inwRequest is not null "
    If txtFileNo.Text <> "" Then
        mSql = mSql & " and tInward.inWNo=" & txtFileNo.Text
    End If
    If txtyear.Text <> "" Then
        mSql = mSql & " and year(InwDate)=" & txtyear.Text
    End If
    If txtSender.Text <> "" Then
        mSql = mSql & " and chvName='%" & txtSender.Text & "%'"
    End If
    If SevanaMainSubid <> "" Then
        mSql = mSql & " and inwRequest=" & SevanaMainSubid
    End If
    If cmbSubType.ListIndex > -1 Then
        mSql = mSql & " and intInwardSubType=" & cmbSubType.ItemData(cmbSubType.ListIndex)
    End If
    If txtHName.Text <> "" Then
        mSql = mSql & " and chvHouseName='%" & txtHName.Text & "%'"
    End If
    If txtDoorNo1.Text <> "" Then
        mSql = mSql & " and chvHouseNo='%" & txtDoorNo1.Text & "%'"
    End If
    Rec.Open mSql, mCnn
    If Not (Rec.EOF Or Rec.BOF) Then
        count = Rec!count
        If count > 0 Then
            pgbrSearch.Max = count
        End If
    Else
        count = 0
    End If
    Rec.Close
    pgbrSearch.Min = 0
    pgbrSearch.Value = 0
    '''' code ends here
    
    mSql = "set dateformat DMY "
    mSql = mSql & " select tInward.inWno as FileNo,convert(varchar,inwDate,103) as DOR,"
    mSql = mSql & " chvName as Sender,TypeofRequest as Subject,TypeofSubRequest as SubType,"
    mSql = mSql & " 'No:'+convert(varchar,tReceiptDetails.intReceiptNo)+', Date: '+"
    mSql = mSql & " convert(varchar,tReceiptDetails.dtReceiptDate,103)+', Amount: '+"
    mSql = mSql & " convert(varchar,treceiptDetails.fltAmount) as ReceiptDetails,"
    mSql = mSql & " tReceiptDetails.chvRegnNo as RegNo,tReceiptDetails.chvBookNo as BookNo"
    mSql = mSql & " From tInward inner join mInwardTYpe on mInwardType.intID=tInward.inwRequest"
    mSql = mSql & " inner join mInwardSubType on mInwardSubType.intID=intInwardSubType"
    mSql = mSql & " left join tReceiptDetails on tReceiptDetails.intInwardID=tInward.intid"
    mSql = mSql & " where inwRequest is not null "
    If txtFileNo.Text <> "" Then
        mSql = mSql & " and tInward.inWNo=" & txtFileNo.Text
    End If
    If txtyear.Text <> "" Then
        mSql = mSql & " and year(InwDate)=" & txtyear.Text
    End If
    If txtSender.Text <> "" Then
        mSql = mSql & " and chvName='%" & txtSender.Text & "%'"
    End If
    If SevanaMainSubid <> "" Then
        mSql = mSql & " and inwRequest=" & SevanaMainSubid
    End If
    If cmbSubType.ListIndex > -1 Then
        mSql = mSql & " and intInwardSubType=" & cmbSubType.ItemData(cmbSubType.ListIndex)
    End If
    If txtHName.Text <> "" Then
        mSql = mSql & " and chvHouseName='%" & txtHName.Text & "%'"
    End If
    If txtDoorNo1.Text <> "" Then
        mSql = mSql & " and chvHouseNo='%" & txtDoorNo1.Text & "%'"
    End If
    mSql = mSql & " order by tInward.inWno "
    Rec.Open mSql, mCnn, adOpenDynamic, adLockOptimistic
    
    vsSearch.Clear 1
    vsSearch.Rows = 2
    vsSearch.TextMatrix(0, 3) = "Sender"
    vsSearch.TextMatrix(0, 4) = "Subject"
    vsSearch.TextMatrix(0, 5) = "Sub Type"
    vsSearch.TextMatrix(0, 6) = "Status"
    vsSearch.TextMatrix(0, 7) = "Receipt Details"
    vsSearch.TextMatrix(0, 8) = "Reg No"
    vsSearch.TextMatrix(0, 9) = "Book No"
    
    If Not (Rec.EOF Or Rec.BOF) Then
        Do While Not Rec.EOF
            vsSearch.Rows = vsSearch.Rows + 1
                vsSearch.TextMatrix(i + 1, 0) = i + 1
                vsSearch.TextMatrix(i + 1, 1) = Rec!FileNo
                vsSearch.TextMatrix(i + 1, 2) = Rec!DOR
                vsSearch.TextMatrix(i + 1, 3) = IIf(IsNull(Rec!Sender), "", Rec!Sender)
                vsSearch.TextMatrix(i + 1, 4) = IIf(IsNull(Rec!Subject), "", Rec!Subject)
                vsSearch.TextMatrix(i + 1, 5) = IIf(IsNull(Rec!SubType), "", Rec!SubType)
                Rec1.Open "STATUSREPORT1 '" & Rec!FileNo & "','" & Year(Rec!DOR) & "'", mCnn
                If Not (Rec1.EOF Or Rec1.BOF) Then
                    vsSearch.TextMatrix(i + 1, 6) = IIf(IsNull(Rec1.Fields(0)), "", Rec1.Fields(0))
                Else
                    vsSearch.TextMatrix(i + 1, 6) = ""
                End If
                Rec1.Close
                vsSearch.TextMatrix(i + 1, 7) = IIf(IsNull(Rec!ReceiptDetails), "", Rec!ReceiptDetails)
                vsSearch.TextMatrix(i + 1, 8) = IIf(IsNull(Rec!RegNo), "", Rec!RegNo)
                vsSearch.TextMatrix(i + 1, 9) = IIf(IsNull(Rec!BookNo), "", Rec!BookNo)
                If pgbrSearch.Value <> pgbrSearch.Max Then
                    pgbrSearch.Value = pgbrSearch.Value + 1
                End If
                Rec.MoveNext
                i = i + 1
        Loop
    Else
        vsSearch.Clear 1, 1
    End If
    Rec.Close
    'gSubSetFont vsSearch, 1, 1, vsSearch.Rows - 1, 1, "ML-TTRevathi"
    'gSubSetFont vsSearch, 1, 3, vsSearch.Rows - 1, 8, "ML-TTRevathi"
    gSubSetFont vsSearch, 1, 1, vsSearch.Rows - 1, vsSearch.Cols - 1, "Verdana"
    pgbrSearch.Value = 0
End Sub
Private Sub Form_Load()
    gSubCenterForm Me
    getDeptID
    Call PopulateList(cboPriority, "Select chvPriority, bntPriorityId from tblPriority", , , , True, enuSourceString.SOOCHIKA)
    Call PopulateList(cboInwardType, "SELECT chvInwardType,intInwardType FROM TblInwardType ", , , , True, enuSourceString.SOOCHIKA)
    cboPriority.ListIndex = 4
    cboInwardType.ListIndex = 0
    txtyear.Text = CStr(Year(Date))
End Sub

Private Sub getDeptID()
        Dim mSql As String
        Dim objDB As New clsDb
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim varyOut As Variant
        objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        Set Rec = objDB.ExecuteSP("spSelectDepartment", , varyOut, , mCnn, adCmdStoredProc)
        If IsArray(varyOut) Then
            gbDeptID = varyOut(1, 0)
        End If
End Sub

Private Sub lstSubject_DblClick()
    txtSubject.Text = lstSubject.Text
    lstSubject.Visible = False
    GetSevanaSubType
End Sub
Private Sub GetSevanaSubType()
    Dim mSql As String
    Dim mCnn As New ADODB.Connection
    Dim objDB As New clsDb
    Dim Rec As New ADODB.Recordset
    objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
    mSql = "select intSubID,intMainSubID from tblSubjectCoding where chvsubject='" & txtSubject.Text & " '"
    Rec.Open mSql, mCnn
    cmbSubType.Clear
    If Not (Rec.EOF Or Rec.BOF) Then
        txtSubID.Text = Rec!intSubID
        SevanaMainSubid = Rec!intMainSubID
        If IsNull(Rec!intMainSubID) = False Then
            PopulateList cmbSubType, "select TypeofSubRequest,intID from TblSubjectSubType where intSubTypeID=" & Rec!intMainSubID & "", , False, True, True, enuSourceString.SOOCHIKA
            cmbSubType.Enabled = True
        Else
            cmbSubType.Enabled = False
        End If
    End If
End Sub
Private Sub txtDoorNo1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtDtFrm_LostFocus()
    If (txtDtFrm.Text <> "") Then
        If (gFunIsDMYDateBoolean(txtDtFrm.Text) = False) Then
            MsgBox "Check the date "
            txtDtFrm.SetFocus
        End If
    End If
End Sub

Private Sub txtDtTo_LostFocus()
    If (txtDtTo.Text <> "") Then
        If (gFunIsDMYDateBoolean(txtDtTo.Text) = False) Then
            MsgBox "Check the date "
            txtDtTo.SetFocus
        End If
    End If
End Sub

Private Sub txtRefDate_LostFocus()
    If (txtRefDate.Text <> "") Then
        If (gFunIsDMYDateBoolean(txtRefDate.Text) = False) Then
            MsgBox "Check the date "
            txtRefDate.SetFocus
        End If
    End If
End Sub

Private Sub txtSubID_KeyPress(KeyAscii As Integer)
     If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtSubID_LostFocus()
    If txtSubID.Text <> "" Then
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDb
        
        objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mSql = "Select chvSubject from tblsubjectcoding where intsubID=" & txtSubID.Text
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            txtSubject.Text = Rec!chvSubject
            GetSevanaSubType
        Else
            MsgBox "Invalid subject id", vbInformation
            txtSubID.Text = ""
        End If
    End If
End Sub

Private Sub txtSubject_KeyPress(KeyAscii As Integer)
    If txtSubject.Text <> "" Then
        lstSubject.Clear
        Dim mSql As String
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDb
        
        objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        
        mSql = "select chvSubject from tblSubjectcoding where chvsubject like '%" & txtSubject.Text & "%'"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF Or Rec.BOF) Then
            While Not Rec.EOF
                lstSubject.AddItem (Rec!chvSubject)
                Rec.MoveNext
            Wend
            lstSubject.Visible = True
        Else
            lstSubject.Visible = False
        End If
    End If
End Sub

Private Sub txtWardNo_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
            KeyAscii = 0
    End If
End Sub

Private Sub ClearData()
txtFileNo.Text = ""
txtyear.Text = ""
txtSender.Text = ""
txtSubject.Text = ""
txtLocality.Text = ""
txtHName.Text = ""
txtDoorNo1.Text = ""
txtDoorNo2.Text = ""
txtRefDate.Text = ""
txtRefNo.Text = ""
txtDtFrm.Text = ""
txtDtTo.Text = ""
txtWardNo.Text = ""
cboInwardType.ListIndex = -1
cboPriority.ListIndex = -1
vsSearch.Clear 1
cmbSubType.Clear
End Sub

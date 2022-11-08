VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSoochikaFrontOfficeDiary 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Front Office Diary"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   7605
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   14535
      Begin VB.TextBox txtyear 
         Height          =   315
         Left            =   7320
         TabIndex        =   32
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtInwardno 
         Height          =   315
         Left            =   5400
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton btnClose 
         BackColor       =   &H00C0C0C0&
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
         Left            =   12630
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   690
         Width           =   915
      End
      Begin VB.CommandButton btnPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   690
         Width           =   915
      End
      Begin VB.CommandButton btnSearch 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dptSearch 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   690
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60489729
         CurrentDate     =   40022
      End
      Begin VSFlex8LCtl.VSFlexGrid vsDiary 
         Height          =   5925
         Left            =   285
         TabIndex        =   7
         Top             =   1200
         Width           =   13935
         _cx             =   24580
         _cy             =   10451
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         BackColorAlternate=   14737632
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
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaFrontOfficeDiary.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
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
         Begin VB.Frame fraDiary 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   5085
            Left            =   2040
            TabIndex        =   8
            Top             =   720
            Visible         =   0   'False
            Width           =   8655
            Begin VB.CheckBox chkParty 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   6360
               TabIndex        =   19
               Top             =   960
               Width           =   495
            End
            Begin VB.CheckBox chkSeat 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   6360
               TabIndex        =   18
               Top             =   480
               Width           =   495
            End
            Begin VB.CommandButton cmdSave 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Save"
               Enabled         =   0   'False
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
               Top             =   2160
               Width           =   915
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Close"
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
               TabIndex        =   16
               Top             =   2640
               Width           =   915
            End
            Begin VB.TextBox txtRemarks 
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
               Height          =   735
               Left            =   3960
               MultiLine       =   -1  'True
               TabIndex        =   15
               Top             =   2400
               Width           =   3165
            End
            Begin VB.TextBox txtParty 
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
               Height          =   735
               Left            =   3960
               MultiLine       =   -1  'True
               TabIndex        =   14
               Top             =   1560
               Width           =   3165
            End
            Begin VB.TextBox txtsms 
               Height          =   1095
               Left            =   120
               TabIndex        =   13
               Top             =   3600
               Width           =   4935
            End
            Begin VB.CheckBox chksms 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SendSMS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   12
               Top             =   3240
               Width           =   1215
            End
            Begin VB.TextBox txtmobileno 
               Height          =   375
               Left            =   5160
               TabIndex        =   11
               Top             =   3600
               Width           =   1815
            End
            Begin VB.CommandButton btnsavesms 
               BackColor       =   &H8000000A&
               Caption         =   "SaveSMS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   10
               Top             =   4200
               Width           =   1335
            End
            Begin VB.CommandButton btnupdate 
               BackColor       =   &H000000FF&
               Caption         =   "UpdateMobileNo"
               Height          =   375
               Left            =   7080
               TabIndex        =   9
               Top             =   3600
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker dtpSection 
               Height          =   345
               Left            =   3960
               TabIndex        =   20
               Top             =   480
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   609
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   60489729
               CurrentDate     =   40022
            End
            Begin MSComCtl2.DTPicker dtpParty 
               Height          =   345
               Left            =   3960
               TabIndex        =   21
               Top             =   960
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   609
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   60489729
               CurrentDate     =   40022
            End
            Begin VB.Label lblSection 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   4080
               TabIndex        =   28
               Top             =   1560
               Width           =   1065
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   240
               TabIndex        =   27
               Top             =   2640
               Width           =   945
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Issued to"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   240
               TabIndex        =   26
               Top             =   1920
               Width           =   1020
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Issued to party "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   120
               TabIndex        =   25
               Top             =   1080
               Width           =   1725
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Received from Seat -"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   600
               Width           =   2355
            End
            Begin VB.Label lblmobileno 
               BackColor       =   &H00E0E0E0&
               Caption         =   "MobileNo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5520
               TabIndex        =   23
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Label lblcurrentno 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3960
               TabIndex        =   22
               Top             =   120
               Width           =   1575
            End
         End
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OR"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblyear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   6600
         TabIndex        =   31
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblinwardno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Inward No"
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
         Height          =   240
         Left            =   4200
         TabIndex        =   29
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delivery Date"
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
         TabIndex        =   3
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Front Office Diary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   14385
      End
   End
End
Attribute VB_Name = "frmSoochikaFrontOfficeDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lSoochikaFeildID As Variant
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
'    Dim vAryInRpt(1)
'    vAryInRpt(0) = CStr(dptSearch.value)
'    frmCRViewer.vShowReport App.Path & "\soochika\Reports", "RptDeliveryFileDetails.rpt", vAryInRpt
'    frmCRViewer.Show 1

 Dim vAryInRpt(2)
    'ReDim vAryInRpt(0)
    
    vAryInRpt(0) = CStr(dptSearch.value)
    vAryInRpt(1) = CStr(dptSearch.value)
    
    'vAryInRpt = Array(CStr(dptSearch.value), CStr(dptSearch.value))
    'frmCRViewer.vShowReport App.Path & "\soochika\Reports", "RptDeliveryFileDetails.rpt", vAryInRpt
   frmCRViewer.ShowUnicodeReport App.Path & "\soochika\Reports", "rptFrontofficeDiary.rpt", vAryInRpt
  
    frmCRViewer.Show 1
End Sub

Private Sub btnsavesms_Click()
'changed by soumya vs oct21
Dim Rec As New ADODB.Recordset
Dim arr As Variant
Dim objDB As New clsDB
Dim mCnn As New ADODB.Connection
If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
ReDim arr(4)
If (txtsms.Text <> "") Then
arr(0) = lSoochikaFeildID
'CHANGED
'arr(1) = Right((lSoochikaFeildID), 6)
arr(1) = 0
arr(2) = dtpSection.value
arr(3) = 0
arr(4) = txtsms.Text
Set Rec = objDB.ExecuteSP("SaveSMS_File", arr, , , mCnn, adCmdStoredProc)
MsgBox ("SMS Data Saved Successfully!!")
'CHANGED
chksms.value = 0
txtsms.Text = ""
txtmobileno.Text = ""
End If
End Sub

Private Sub btnSearch_Click()
    'FillDiary
    GetSearch
End Sub
'Private Sub FillDiary()
'    Dim objdb As New clsDB
'    Dim Rec As New ADODB.Recordset
'    Dim mCnn As New ADODB.Connection
'    Dim vAryIn As Variant
'    Dim varyOut As Variant
'        objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
'        ReDim vAryIn(0)
'        vAryIn(0) = str(dptSearch.value)
'        Set Rec = objdb.ExecuteSP("Sp_RptDeliveryReport", vAryIn, varyOut, , mCnn, adCmdStoredProc)
'        vsDiary.Rows = 2
'        vsDiary.Clear 1
'           If IsArray(varyOut) Then
'                For i = 0 To UBound(varyOut, 2)
'                If i > 0 Then
'                   vsDiary.Rows = vsDiary.Rows + 1
'                End If
'                vsDiary.TextMatrix(i + 1, 0) = i + 1
'                vsDiary.TextMatrix(i + 1, 1) = varyOut(2, i) 'Inward
'                vsDiary.TextMatrix(i + 1, 2) = varyOut(4, i) 'Sender
'                vsDiary.TextMatrix(i + 1, 3) = varyOut(3, i) 'DeliveryType
'                vsDiary.TextMatrix(i + 1, 4) = varyOut(6, i) 'Section
'                If Not IsNull(varyOut(9, i)) Then
'                     vsDiary.TextMatrix(i + 1, 5) = varyOut(9, i) 'Received Dt
'                End If
'                If Not IsNull(varyOut(12, i)) Then
'                     vsDiary.TextMatrix(i + 1, 6) = varyOut(12, i) 'Given Dt
'                End If
'                vsDiary.TextMatrix(i + 1, 7) = varyOut(8, i) 'Id
'               Next i
'           End If
'           gSubSetFont vsDiary, 1, 0, vsDiary.Rows - 1, 0, "Arial"
'           gSubSetFont vsDiary, 1, 5, vsDiary.Rows - 1, 7, "Arial"
'End Sub
Private Sub FillDiary()
   Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim vAryIn As Variant
    Dim varyOut As Variant
        'objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        objDB.CreateNewConnection mCnn, enuSourceString.SoochikaUnicode
        ReDim vAryIn(0)
'        vAryIn(0) = str(dptSearch.value)
         vAryIn = Array(CStr(dptSearch.value), CStr(dptSearch.value))
        'Set Rec = objdb.ExecuteSP("Sp_RptDeliveryReport", vAryIn, varyOut, , mCnn, adCmdStoredProc)
        Set Rec = objDB.ExecuteSP("Sp_RptFrontOfficeDiary", vAryIn, varyOut, , mCnn, adCmdStoredProc)
        vsDiary.Rows = 2
        vsDiary.Clear 1
           If IsArray(varyOut) Then
                For i = 0 To UBound(varyOut, 2)
                If i > 0 Then
                   vsDiary.Rows = vsDiary.Rows + 1
                End If
                vsDiary.TextMatrix(i + 1, 0) = i + 1
                'changed by soumya V S on 21
                vsDiary.TextMatrix(i + 1, 1) = varyOut(0, i) 'FileID
                vsDiary.TextMatrix(i + 1, 2) = varyOut(2, i) 'Inward
                'LATEST CHANGED
                 If Not IsNull(varyOut(4, i)) Then
                vsDiary.TextMatrix(i + 1, 3) = varyOut(4, i) 'Sender
                End If
                vsDiary.TextMatrix(i + 1, 4) = varyOut(5, i) 'DeliveryType
                'CHNAGED
                 If Not IsNull(varyOut(7, i)) Then
                vsDiary.TextMatrix(i + 1, 5) = varyOut(7, i) 'Section
                End If
                If Not IsNull(varyOut(3, i)) Then
                     vsDiary.TextMatrix(i + 1, 6) = varyOut(3, i) 'Received Dt
                End If
                If Not IsNull(varyOut(6, i)) Then
                     vsDiary.TextMatrix(i + 1, 7) = varyOut(6, i) 'Given Dt
                End If
                'chnaged by soumya VS oct17
                'CHNAGED
                If Not IsNull(varyOut(16, i)) Then
                vsDiary.TextMatrix(i + 1, 8) = varyOut(16, i) 'Id
                End If
               Next i
           End If
                'vsDiary.TextMatrix(i + 1, 1) = varyOut(2, i) 'Inward
                'vsDiary.TextMatrix(i + 1, 2) = varyOut(4, i) 'Sender
                'vsDiary.TextMatrix(i + 1, 3) = varyOut(5, i) 'DeliveryType
                'vsDiary.TextMatrix(i + 1, 4) = varyOut(7, i) 'Section
                'If Not IsNull(varyOut(3, i)) Then
                     'vsDiary.TextMatrix(i + 1, 5) = varyOut(3, i) 'Received Dt
                'End If
                'If Not IsNull(varyOut(6, i)) Then
                     'vsDiary.TextMatrix(i + 1, 6) = varyOut(6, i) 'Given Dt
                'End If
                'vsDiary.TextMatrix(i + 1, 7) = varyOut(8, i) 'Id
               'Next i
           'End If
           gSubSetFont vsDiary, 1, 0, vsDiary.Rows - 1, 0, "Arial"
           gSubSetFont vsDiary, 1, 5, vsDiary.Rows - 1, 7, "Arial"

End Sub
Private Sub GetSearch()
Dim strQry As String
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
                         
            strQry = "select tInwardDetails.numFileID as FileID,isnull(chvFileNo,numCurrentNo) as FileNo,convert(varchar,dtDateofReceipt,103) as DateofReceipt,"
            strQry = strQry & " tInwardDetails.chvApplicantName +isnull(','+tInwardDetails.chvHouseName,'')+isnull(',' +chvLocalPlace,'') as Address,tInwardDetails.chvSubject as Subject,"
            strQry = strQry & " convert(varchar,tInwardDetails.dtDeliveryDate,103) as DeliveryDate,tSeatDetails.chvSeatname as Seat,tSeatDetails.chvSeatNameMal as SeatMal,"
            strQry = strQry & " convert(varchar,tFrontofficeDiary.dtFromSeat,103) as dtReceivedFromSeat,tFrontofficeDiary.chvPartyDetails as PartyDetails,tFrontofficeDiary.chvRemarks as Remarks,"
            strQry = strQry & " convert(varchar,tFrontofficeDiary.dtIssueToParty,103) as dtIssuetoParty,tSeatDetails.numSeatID as SeatID "
            strQry = strQry & " From tInwardDetails left join tFrontofficeDiary on tInwardDetails.numFileID=tFrontofficeDiary.numFileID left join tSeatDetails on tSeatDetails.numSeatID=tInwardDetails.numCustodianSeatID "
            strQry = strQry & " where  "
            
       If (txtInwardno.Text = "") Then
        
        If (dptSearch.value <> "") Then
        strQry = strQry & " tInwardDetails.dtDeliveryDate between convert(datetime,'" & dptSearch.value & "',103) and convert(datetime,'" & dptSearch.value & "',103) "
        End If
        
    Else
        
        If (txtInwardno.Text <> "") Then
            strQry = strQry & "  tInwardDetails.numCurrentNo='" & txtInwardno.Text & "' "
        End If
        If (txtyear.Text <> "") Then
            strQry = strQry & " and year(tInwardDetails.dtDateofReceipt)= '" & txtyear.Text & "' "
        End If
      End If
        If strQry <> "" Then
             strQry = strQry & " order by tInwardDetails.numFileID "
        End If
        
          If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Connection not present", vbDefaultButton1, "SOOCHIKA"
            Exit Sub
        End If
          ' gSubSetFont vsEnclosure, 1, 2, vsEnclosure.Rows - 1, 2, "ML-TTRevathi"
          
            vsDiary.Rows = 2
            vsDiary.Clear 1
            vsDiary.TextMatrix(0, 3) = "Name of Applicant"
            vsDiary.TextMatrix(0, 4) = "Nature of Subject"
            vsDiary.TextMatrix(0, 5) = "Section"
            vsDiary.TextMatrix(0, 6) = "Date of received from seat"
            vsDiary.TextMatrix(0, 7) = "Service date"
            vsDiary.TextMatrix(0, 8) = ""
        
            
            Rec.Open strQry, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                i = 0
'''''''''            vsSearch.TextMatrix(0, 7) = "Status"
'''''''''            vsSearch.TextMatrix(0, 8) = "Notes"
            Do While Not (Rec.EOF)
                vsDiary.Rows = vsDiary.Rows + 1
                vsDiary.TextMatrix(i + 1, 0) = i + 1
                vsDiary.TextMatrix(i + 1, 1) = Rec!FileID
                vsDiary.TextMatrix(i + 1, 2) = Rec!FileNo
                vsDiary.TextMatrix(i + 1, 3) = Rec!Address
                vsDiary.TextMatrix(i + 1, 4) = IIf(IsNull(Rec!Subject), "", Rec!Subject)
                vsDiary.TextMatrix(i + 1, 5) = Rec!Seat
                vsDiary.TextMatrix(i + 1, 6) = IIf(IsNull(Rec!dtReceivedFromSeat), "", Rec!dtReceivedFromSeat)
                vsDiary.TextMatrix(i + 1, 7) = IIf(IsNull(Rec!deliveryDate), "", Rec!deliveryDate)
                vsDiary.TextMatrix(i + 1, 8) = IIf(IsNull(Rec!SeatID), "", Rec!SeatID)
                
                Rec.MoveNext
                i = i + 1
                Loop
            End If
            Rec.Close
            gSubSetFont vsDiary, 1, 1, vsDiary.Rows - 1, 1, "Verdana"
            gSubSetFont vsDiary, 1, 3, vsDiary.Rows - 1, 8, "Verdana"
            'pgbrSearch.value = 0
End Sub

Private Sub getDiaryDetails(numFileID As Variant)
        Dim objDB As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim vAryIn As Variant
        Dim varyOut As Variant
        Dim flgSave As Integer
        flgSave = 0
        'CHANGED
        If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
            MsgBox "Soochika Connection is not present", vbCritical, "Common"
            Exit Sub
        End If
        If numFileID <> "" Then
        ReDim vAryIn(0)
        vAryIn(0) = numFileID
        Set Rec = objDB.ExecuteSP("spSelectFrontOfficeDiary", vAryIn, varyOut, , mCnn, adCmdStoredProc)
        
        
            
            If Not (Rec.EOF And Rec.BOF) Then
             
            'added by soumya V S on 12.01.16
               If IsNull(varyOut(3, 0)) Then
                   chkSeat.Visible = True
                   chkSeat.Enabled = True
                   dtpSection.Visible = True
                   Label3.Visible = True
                   chkParty.Visible = False
                   dtpParty.Visible = False
                   txtParty.Visible = False
                   Label4.Visible = False
                   Label5.Visible = False
                   Label6.Visible = False
                   txtRemarks.Visible = False
                   lblSection.Visible = False
                Else
                
                   chkSeat.Visible = True
                   dtpSection.Visible = True
                   Label3.Visible = True
                   chkParty.Visible = True
                   dtpParty.Visible = True
                   txtParty.Visible = True
                   Label4.Visible = True
                   Label5.Visible = True
                   Label6.Visible = True
                   txtRemarks.Visible = True
                   End If
                   
            
                If IsNull(varyOut(5, 0)) Then
                    txtParty.Text = ""
                Else
                    txtParty.Text = varyOut(5, 0)
                    'CHANGED
                    chkParty.value = 1
                    chkSeat.Enabled = False
                    flgSave = 1
                End If
                If IsNull(varyOut(6, 0)) Then
                    txtRemarks.Text = ""
                Else
                    txtRemarks.Text = varyOut(6, 0)
                End If
                If IsNull(varyOut(3, 0)) Then
                    dtpSection.value = Date
                    flgSave = 0
                Else
                    dtpSection.value = varyOut(3, 0)
                    chkSeat.value = 1
                    chkSeat.Enabled = False
                    'CHANGED
                    dtpSection.Enabled = False
                    flgSave = 1
                End If
                If IsNull(varyOut(4, 0)) Then
                    dtpParty.value = Date
                    flgSave = 0
                Else
                    dtpParty.value = varyOut(4, 0)
                    chkParty.value = 1
                    chkParty.Enabled = False
                    flgSave = 1
                End If
                Else
                chkSeat.Visible = True
                   chkSeat.Enabled = True
                   dtpSection.Visible = True
                   Label3.Visible = True
                   chkParty.Visible = False
                   dtpParty.Visible = False
                   txtParty.Visible = False
                   Label4.Visible = False
                   Label5.Visible = False
                   Label6.Visible = False
                   txtRemarks.Visible = False
                   lblSection.Visible = False
                   
                   
            End If
            
            Rec.Close
        Else
            lblOfficerName.Caption = ""
        End If
        If flgSave = 1 Then
            cmdSave.Enabled = False
        End If
End Sub

Private Sub btnupdate_Click()
'changed by soumya Vs oct 21 14
Dim Rec As New ADODB.Recordset
Dim arr As Variant
Dim objDB As New clsDB
Dim mCnn As New ADODB.Connection
Dim mSql As String
If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
ReDim arr(1)

If (txtmobileno.Text = "") Then
MsgBox ("Please enter a valid mobile no.!!!")
End If
If Len(txtmobileno.Text) < 10 Or Len(txtmobileno.Text) > 11 Then
MsgBox ("Please enter a valid mobile no.!!!")
Else
arr(0) = lSoochikaFeildID
arr(1) = txtmobileno.Text
  mSql = "Update tInwardDetails set chvContactNo=" & txtmobileno & " WHERE numFileID=" & lSoochikaFeildID
    mCnn.Execute mSql
        If (mCnn.State = 1) Then
            mCnn.Close
        End If
MsgBox ("Update MobileNo Successfully")
btnupdate.Visible = False
End If
End Sub

Private Sub chkParty_Click()
    If chkParty.value = 1 Then
        dtpParty.Enabled = True
        cmdSave.Enabled = True
    Else
        dtpParty.Enabled = False
    End If
End Sub
Private Sub chkSeat_Click()
    If chkSeat.value = 1 Then
        dtpSection.Enabled = True
        cmdSave.Enabled = True
    Else
        dtpSection.Enabled = False
    End If
End Sub

Private Sub chksms_Click()
'chnaged by soumya vS on oct21
Dim arr As Variant
Dim mSql As String
Dim objDB As New clsDB
Dim mCnn As New ADODB.Connection
Dim Rec As New ADODB.Recordset

If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
If (chksms.value = 1) Then
txtsms.Visible = True
txtmobileno.Visible = True
btnupdate.Visible = True
btnsavesms.Visible = True
lblmobileno.Visible = True
Else
txtsms.Visible = False
txtmobileno.Visible = False
btnupdate.Visible = False
btnsavesms.value = False
lblmobileno.Visible = False
End If
  mSql = "select chvContactNo from tInwardDetails  WHERE  numFileID=" & lSoochikaFeildID
  Rec.Open mSql, mCnn
   If Not (Rec.EOF Or Rec.BOF) Then
   If (Rec!chvContactNo <> "") Then
   txtmobileno.Text = Rec!chvContactNo
   End If
   End If
End Sub

Private Sub cmdClose_Click()
    ClearFrDiary
End Sub

Private Sub cmdSave_Click()
If lSaveValidate = True Then
    SaveDiary
    fraDiary.Visible = False
    'FillDiary
    GetSearch
    
End If
End Sub
Private Sub SaveDiary()
'changed by soumya vs on oct 21
Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim objDB As New clsDB
 Dim Rec As New ADODB.Recordset
 Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(5)
     'changed by soumya VS
     If (objDB.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
        MsgBox "Connection Failure", vbInformation, "SOOCHIKA"
        Exit Sub
    End If
     'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
     
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        mVarrIn(1) = SeatID
        If chkSeat.value = 1 Then
            mVarrIn(2) = dtpSection.value
        Else
            mVarrIn(2) = Null
        End If
        If chkParty.value = 1 Then
            mVarrIn(3) = dtpParty.value
        Else
            mVarrIn(3) = Null
        End If
        If (txtParty.Text <> "") Then
        mVarrIn(4) = txtParty.Text
        Else
            mVarrIn(4) = Null
        End If
        If (txtRemarks.Text <> "") Then
        mVarrIn(5) = txtRemarks.Text
        Else
           mVarrIn(5) = Null
        End If
        
        Set Rec = objDB.ExecuteSP("spSaveFrontOfficeDiary", mVarrIn, , , mCnn, adCmdStoredProc)
        MsgBox ("Saved Successfully")
        'changed
        txtsms.Text = ""
        txtmobileno.Text = ""
        chksms.value = 0
 'Dim mVarrIn As Variant
 'Dim mVarrOut As Variant
 'Dim objDB As New clsDB
 'Dim Rec As New ADODB.Recordset
 'Dim mCnn As New ADODB.Connection
     'ReDim mVarrIn(4)
     'objDB.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        'mVarrIn(0) = lSoochikaFeildID 'FileId.
        'If chkSeat.value = 1 Then
            'mVarrIn(1) = dtpSection.value
        'Else
           ' mVarrIn(1) = Null
        'End If
        'If chkParty.value = 1 Then
            'mVarrIn(2) = dtpParty.value
        'Else
            'mVarrIn(2) = Null
       ' End If
        'mVarrIn(3) = txtParty.Text
        'mVarrIn(4) = txtRemarks.Text
        'Set Rec = objDB.ExecuteSP("spSaveFrontOfficeDiary", mVarrIn, , , mCnn, adCmdStoredProc)
End Sub





Private Sub Form_Load()
    gSubCenterForm Me
    dptSearch.value = Date
    dtpParty.value = Date
    dtpSection.value = Date
      'changed by soumya oct 21 2014
  txtsms.Visible = False
txtmobileno.Visible = False
btnupdate.Visible = False
btnsavesms.value = False
lblmobileno.Visible = False
btnsavesms.Visible = False
    FillDiary
End Sub

Private Sub txtmobileno_Change()
'changed by soumya V S 21 oct
Dim textval As String
Dim numval As String
textval = txtmobileno.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    txtmobileno.Text = CStr(numval)
  End If
End Sub

Private Sub vsDiary_Click()
    If vsDiary.Row > 0 Then
        ClearFrDiary
        If val(vsDiary.TextMatrix(vsDiary.Row, 7)) > 0 Then
        'CHANGED
           getDiaryDetails vsDiary.TextMatrix(vsDiary.Row, 1)
           fraDiary.Visible = True
           lblSection.Caption = vsDiary.TextMatrix(vsDiary.Row, 5)
           txtParty.Tag = vsDiary.TextMatrix(vsDiary.Row, 7)
           'changed by soumya vs on 21oct
           lSoochikaFeildID = vsDiary.TextMatrix(vsDiary.Row, 1)
           lblcurrentno.Caption = "CurrentNO:" + vsDiary.TextMatrix(vsDiary.Row, 2)
           SeatID = vsDiary.TextMatrix(vsDiary.Row, 8)
           'lSoochikaFeildID = val(vsDiary.TextMatrix(vsDiary.Row, 7))
        End If
    End If
End Sub
Private Sub vsDiary_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ClearFrDiary()
    txtParty.Text = ""
    txtRemarks.Text = ""
    dtpSection.value = Date
    dtpParty.value = Date
    fraDiary.Visible = False
    chkSeat.value = 0
    chkParty.value = 0
    chkParty.Enabled = True
    chkSeat.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Function lSaveValidate() As Boolean
    lSaveValidate = True
    If (dtpSection.value > Date) Then
        lSaveValidate = False
        MsgBox "Date of receipt from seat should not be greater than today"
        Exit Function
    ElseIf (dtpParty.value > Date) Then
        lSaveValidate = False
        MsgBox "Date of issue to party should not be greater than today"
        Exit Function
    ElseIf (dtpSection.value <> "") And (dtpParty.value <> "") Then
        If (dtpSection.value > dtpParty.value) Then
            lSaveValidate = False
            MsgBox "Date of issue to party should not be less than date of receipt from seat"
            Exit Function
        End If
    End If
End Function



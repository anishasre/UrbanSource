VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSoochikaDespatchDiary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despatch Diary"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1995
      End
      Begin VB.CommandButton btnPrint 
         BackColor       =   &H00C0E0FF&
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
         Left            =   11670
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   690
         Width           =   915
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
         Left            =   12630
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   690
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dptSearch 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
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
         Format          =   59899905
         CurrentDate     =   40022
      End
      Begin VSFlex8LCtl.VSFlexGrid vsDiary 
         Height          =   5925
         Left            =   300
         TabIndex        =   7
         Top             =   1230
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaDespatchDiary.frx":0000
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
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   3045
            Left            =   2670
            TabIndex        =   8
            Top             =   1170
            Visible         =   0   'False
            Width           =   7245
            Begin VB.CheckBox chkParty 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   6390
               TabIndex        =   14
               Top             =   480
               Width           =   495
            End
            Begin VB.CheckBox chkSeat 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   6390
               TabIndex        =   13
               Top             =   120
               Width           =   495
            End
            Begin VB.CommandButton cmdSave 
               BackColor       =   &H00C0E0FF&
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
               Left            =   4500
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   2520
               Width           =   915
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00C0E0FF&
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
               Left            =   5460
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   2550
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
               Left            =   3990
               MultiLine       =   -1  'True
               TabIndex        =   10
               Top             =   1680
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
               Left            =   3990
               MultiLine       =   -1  'True
               TabIndex        =   9
               Top             =   900
               Width           =   3165
            End
            Begin MSComCtl2.DTPicker dtpSection 
               Height          =   345
               Left            =   3990
               TabIndex        =   15
               Top             =   120
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
               Format          =   59899905
               CurrentDate     =   40022
            End
            Begin MSComCtl2.DTPicker dtpParty 
               Height          =   345
               Left            =   3990
               TabIndex        =   16
               Top             =   510
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
               Format          =   59899905
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
               Left            =   2580
               TabIndex        =   21
               Top             =   150
               Width           =   945
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   150
               TabIndex        =   20
               Top             =   2040
               Width           =   945
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Issued to"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   150
               TabIndex        =   19
               Top             =   1170
               Width           =   960
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Issued to party "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   150
               TabIndex        =   18
               Top             =   600
               Width           =   1605
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Received from Seat -"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   150
               TabIndex        =   17
               Top             =   120
               Width           =   2355
            End
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Postal Despatch"
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
         TabIndex        =   6
         Top             =   180
         Width           =   14385
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Date"
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
         TabIndex        =   5
         Top             =   690
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSoochikaDespatchDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lSoochikaFeildID As Variant
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim vAryInRpt(1)
    vAryInRpt(0) = CStr(dptSearch.Value)
    frmCRViewer.vShowReport App.Path & "\soochika\Reports", "RptDespatchFileDetails.rpt", vAryInRpt
    frmCRViewer.Show 1
End Sub

Private Sub btnSearch_Click()
    FillDiary
End Sub
Private Sub FillDiary()
    Dim objdb As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim vAryIn As Variant
    Dim varyOut As Variant
        objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        ReDim vAryIn(0)
        vAryIn(0) = str(dptSearch.Value)
        Set Rec = objdb.ExecuteSP("Sp_RptDespatchReport", vAryIn, varyOut, , mCnn, adCmdStoredProc)
        vsDiary.Rows = 2
        vsDiary.Clear 1
           If IsArray(varyOut) Then
                For i = 0 To UBound(varyOut, 2)
                If i > 0 Then
                   vsDiary.Rows = vsDiary.Rows + 1
                End If
                vsDiary.TextMatrix(i + 1, 0) = i + 1
                vsDiary.TextMatrix(i + 1, 1) = varyOut(2, i) 'Inward
                vsDiary.TextMatrix(i + 1, 2) = varyOut(4, i) 'Sender
                vsDiary.TextMatrix(i + 1, 3) = varyOut(3, i) 'DeliveryType
                vsDiary.TextMatrix(i + 1, 4) = varyOut(6, i) 'Section
                If Not IsNull(varyOut(9, i)) Then
                     vsDiary.TextMatrix(i + 1, 5) = varyOut(9, i) 'Received Dt
                End If
                If Not IsNull(varyOut(12, i)) Then
                     vsDiary.TextMatrix(i + 1, 6) = varyOut(12, i) 'Given Dt
                End If
                vsDiary.TextMatrix(i + 1, 7) = varyOut(8, i) 'Id
               Next i
           End If
           gSubSetFont vsDiary, 1, 0, vsDiary.Rows - 1, 0, "Arial"
           gSubSetFont vsDiary, 1, 5, vsDiary.Rows - 1, 7, "Arial"
End Sub

Private Sub getDiaryDetails(numSeatID As Variant)
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim vAryIn As Variant
        Dim varyOut As Variant
        Dim flgSave As Integer
        flgSave = 0
        If (objdb.CreateNewConnection(mCnn, enuSourceString.SOOCHIKA) = False) Then
            MsgBox "Soochika Connection is not present", vbCritical, "Common"
            Exit Sub
        End If
        If numSeatID <> "" Then
        ReDim vAryIn(0)
        vAryIn(0) = numSeatID
        Set Rec = objdb.ExecuteSP("spSelectFrontOfficeDiary", vAryIn, varyOut, , mCnn, adCmdStoredProc)
            If Not (Rec.EOF And Rec.BOF) Then
                If IsNull(varyOut(4, 0)) Then
                    txtParty.Text = ""
                Else
                    txtParty.Text = varyOut(4, 0)
                End If
                If IsNull(varyOut(5, 0)) Then
                    txtRemarks.Text = ""
                Else
                    txtRemarks.Text = varyOut(5, 0)
                End If
                If IsNull(varyOut(2, 0)) Then
                    dtpSection.Value = Date
                    flgSave = 0
                Else
                    dtpSection.Value = varyOut(2, 0)
                    chkSeat.Value = 1
                    chkSeat.Enabled = False
                    flgSave = 1
                End If
                If IsNull(varyOut(3, 0)) Then
                    dtpParty.Value = Date
                    flgSave = 0
                Else
                    dtpParty.Value = varyOut(3, 0)
                    chkParty.Value = 1
                    chkParty.Enabled = False
                    flgSave = 1
                End If
            End If
            Rec.Close
        Else
            lblOfficerName.Caption = ""
        End If
        If flgSave = 1 Then
            cmdSave.Enabled = False
        End If
End Sub
Private Sub chkParty_Click()
    If chkParty.Value = 1 Then
        dtpParty.Enabled = True
        cmdSave.Enabled = True
    Else
        dtpParty.Enabled = False
    End If
End Sub
Private Sub chkSeat_Click()
    If chkSeat.Value = 1 Then
        dtpSection.Enabled = True
        cmdSave.Enabled = True
    Else
        dtpSection.Enabled = False
    End If
End Sub
Private Sub cmdClose_Click()
    ClearFrDiary
End Sub
Private Sub ClearFrDiary()
    txtParty.Text = ""
    txtRemarks.Text = ""
    dtpSection.Value = Date
    dtpParty.Value = Date
    fraDiary.Visible = False
    chkSeat.Value = 0
    chkParty.Value = 0
    chkParty.Enabled = True
    chkSeat.Enabled = True
    cmdSave.Enabled = False
End Sub
Private Sub cmdSave_Click()
    SaveDiary
    fraDiary.Visible = False
    FillDiary
End Sub
Private Sub SaveDiary()
 Dim mVarrIn As Variant
 Dim mVarrOut As Variant
 Dim objdb As New clsDB
 Dim Rec As New ADODB.Recordset
 Dim mCnn As New ADODB.Connection
     ReDim mVarrIn(4)
     objdb.CreateNewConnection mCnn, enuSourceString.SOOCHIKA
        mVarrIn(0) = lSoochikaFeildID 'FileId.
        If chkSeat.Value = 1 Then
            mVarrIn(1) = dtpSection.Value
        Else
            mVarrIn(1) = Null
        End If
        If chkParty.Value = 1 Then
            mVarrIn(2) = dtpParty.Value
        Else
            mVarrIn(2) = Null
        End If
        mVarrIn(3) = txtParty.Text
        mVarrIn(4) = txtRemarks.Text
        Set Rec = objdb.ExecuteSP("spSaveFrontOfficeDiary", mVarrIn, , , mCnn, adCmdStoredProc)
End Sub
Private Sub Form_Load()
    gSubCenterForm Me
    dptSearch.Value = Date
    dtpParty.Value = Date
    dtpSection.Value = Date
    FillDiary
End Sub
Private Sub vsDiary_Click()
    If vsDiary.Row > 0 Then
        If Val(vsDiary.TextMatrix(vsDiary.Row, 7)) > 0 Then
            getDiaryDetails (Val(vsDiary.TextMatrix(vsDiary.Row, 7)))
           fraDiary.Visible = True
           lblSection.Caption = vsDiary.TextMatrix(vsDiary.Row, 4)
           txtParty.Tag = vsDiary.TextMatrix(vsDiary.Row, 7)
           lSoochikaFeildID = Val(vsDiary.TextMatrix(vsDiary.Row, 7))
        End If
    End If
End Sub
Private Sub vsDiary_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    KeyAscii = 0
End Sub


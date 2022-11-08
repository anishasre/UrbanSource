VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmProceedings 
   BackColor       =   &H00EDF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P R O C E E D I N G S"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmProceedings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin WinXPC_Engine.WindowsXPC winXPC 
      Left            =   6525
      Top             =   6435
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin TabDlg.SSTab tabProceedings 
      Height          =   6315
      Left            =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15595511
      ForeColor       =   128
      TabCaption(0)   =   "List of Proceedings"
      TabPicture(0)   =   "frmProceedings.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAdd"
      Tab(0).Control(1)=   "cmdsearch"
      Tab(0).Control(2)=   "txtProceedings"
      Tab(0).Control(3)=   "cmdClose"
      Tab(0).Control(4)=   "chkEdit"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "Label1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Define Proceedings"
      TabPicture(1)   =   "frmProceedings.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCancel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSave"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2385
         TabIndex        =   16
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3435
         TabIndex        =   17
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72495
         TabIndex        =   3
         Top             =   5175
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69285
         TabIndex        =   2
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtProceedings 
         Height          =   630
         Left            =   -73935
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   5580
         Width           =   4560
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -71445
         TabIndex        =   4
         Top             =   5175
         Width           =   975
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "Edit"
         Height          =   240
         Left            =   -74835
         TabIndex        =   19
         Top             =   5265
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EDF7F7&
         ForeColor       =   &H008080FF&
         Height          =   4635
         Left            =   315
         TabIndex        =   18
         Top             =   540
         Width           =   6645
         Begin MSComCtl2.DTPicker dtpProceedingsDate 
            Height          =   285
            Left            =   4185
            TabIndex        =   12
            Top             =   1800
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            _Version        =   393216
            Format          =   59244545
            CurrentDate     =   40564
         End
         Begin VB.TextBox txtRemarks 
            Height          =   885
            Left            =   2655
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   2205
            Width           =   3495
         End
         Begin VB.CheckBox chkCancel 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Remove Proceedings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   240
            Left            =   2385
            TabIndex        =   15
            Top             =   225
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.CheckBox chkUsed 
            BackColor       =   &H00EDF7F7&
            Caption         =   "Used"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2700
            TabIndex        =   20
            Top             =   990
            Width           =   825
         End
         Begin VB.TextBox txtProceedingsDate 
            Height          =   345
            Left            =   2655
            TabIndex        =   11
            Top             =   1755
            Width           =   1515
         End
         Begin VB.TextBox txtProceedingsNo 
            Height          =   345
            Left            =   2655
            TabIndex        =   9
            Top             =   1350
            Width           =   1875
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Left            =   1935
            TabIndex        =   13
            Top             =   2340
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proceedings Date"
            Height          =   195
            Left            =   1305
            TabIndex        =   10
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proceedings No"
            Height          =   195
            Left            =   1440
            TabIndex        =   8
            Top             =   1395
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EDF7F7&
         Height          =   4770
         Left            =   -74910
         TabIndex        =   6
         Top             =   360
         Width           =   7140
         Begin VSFlex8LCtl.VSFlexGrid vsGrid 
            Height          =   4410
            Left            =   45
            TabIndex        =   5
            Top             =   240
            Width           =   7020
            _cx             =   12382
            _cy             =   7779
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            BackColorFixed  =   15595511
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   15595511
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   15
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmProceedings.frx":0044
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   2
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceedings"
         Height          =   195
         Left            =   -74865
         TabIndex        =   0
         Top             =   5775
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmProceedings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mModule As Integer
    
    Private Sub cmdAdd_Click()
        Call SelectScreen(False)
        chkCancel.value = 0
        chkCancel.Visible = False
    End Sub
    
    Private Sub cmdCancel_Click()
        Call SelectScreen(True)
    End Sub
    
    Private Sub cmdClose_Click()
        Unload Me
    End Sub
    
    Private Sub cmdSave_Click()
        '   Save Validations
        If Trim(txtProceedingsNo.Text) = "" Then
            MsgBox "Please Enter the Proceedings Number", vbInformation
            txtProceedingsNo.SetFocus
            Exit Sub
        End If
        If Trim(txtProceedingsDate.Text) = "" Then
            MsgBox "Please Enter the Proceedings Date", vbInformation
            txtProceedingsDate.SetFocus
            Exit Sub
        End If
        If Trim(txtRemarks.Text) = "" Then
            If MsgBox("The Proceedings Remarks not Entered, Do you want to Continue", vbYesNo + vbInformation) = vbNo Then
                txtRemarks.SetFocus
                Exit Sub
            End If
        End If
        If chkUsed.value = 1 Then
            MsgBox "This Procedure is already Using, And Cannot be Edited", vbInformation
            Call SelectScreen(True)
            Exit Sub
        End If
        Dim objProceedings As New clsProceedings
        objProceedings.ProceedingsNo = Trim(txtProceedingsNo.Text)
        Call objProceedings.getProceedingsByNo
        If objProceedings.ProceedingsID <> -1 Then
            If objProceedings.ProceedingsID <> val(txtProceedingsNo.Tag) Then
                MsgBox "This Proceedings already Exists", vbInformation
                Exit Sub
            End If
        End If
        With objProceedings
            .ProceedingsID = val(txtProceedingsNo.Tag)
            .ProceedingsNo = txtProceedingsNo.Text
            .ProceedingsDate = Trim(txtProceedingsDate.Text)
            .Remarks = Trim(txtRemarks.Text)
            .Used = chkUsed.value
            .Removed = chkCancel.value
            .ModuleID = mModule
        End With
        Call objProceedings.SaveProceedings
        MsgBox "Proceedings saved Success fully", vbInformation
        Call SelectScreen(True)
        Call FillGrid
        vsGrid.SetFocus
    End Sub
    
    Private Sub cmdSearch_Click()
        Call FillGrid
    End Sub
    Private Sub dtpProceedingsDate_CloseUp()
        txtProceedingsDate.Text = CheckDateInMMM(Trim(txtProceedings.Text))
    End Sub
    
    Private Sub Form_Load()
        tabProceedings.Tab = 0
        tabProceedings.TabVisible(1) = False
        tabProceedings.TabsPerRow = 1
        winXPC.InitIDESubClassing
        Call FillGrid
    End Sub
    Private Sub Form_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            Call cmdSave_Click
        End If
    End Sub
    Private Sub FillGrid()
       Dim objDB As New clsDB
       Dim mCnn As New ADODB.Connection
       Dim Rec As New ADODB.Recordset
       Dim mCount As Integer
       Dim mSQL As String
       Dim mArrayInput As Variant
       On Error GoTo last
       '       Craeting Connction              '
       If objDB.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
           MsgBox "Connection to Saankhya Not Present", vbCritical
           Exit Sub
       End If
       mSQL = "SELECT intProceedingsID, vchProceedingsNo, dtProceedingsDate, intVoucherNo, vchRemarks, intVoucherID,tnyUsed,tnyRemoved FROM faProceedings " & vbNewLine
       mSQL = mSQL + " Where vchProceedingsNo Like '%" & Trim(txtProceedings.Text) & "%' Or Convert(varchar(11),dtProceedingsDate) Like '%" & txtProceedings.Text & "%' Or isNull(vchRemarks,'') Like '%" & txtProceedings.Text & "%'" & vbNewLine
       mSQL = mSQL + "Order By dtProceedingsDate Desc,vchProceedingsNo Desc,tnyRemoved Asc"
       Rec.Open mSQL, mCnn
       vsGrid.Rows = 1
       If Not (Rec.EOF And Rec.BOF) Then
           While Not Rec.EOF
               vsGrid.Rows = vsGrid.Rows + 1
               vsGrid.TextMatrix(vsGrid.Rows - 1, 0) = IIf(IsNull(Rec!intProceedingsID), -1, Rec!intProceedingsID)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 1) = IIf(IsNull(Rec!vchProceedingsNo), "", Rec!vchProceedingsNo)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 2) = IIf(IsNull(Rec!dtProceedingsDate), "", Rec!dtProceedingsDate)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 3) = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 4) = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 5) = IIf(IsNull(Rec!intVoucherID), -1, Rec!intVoucherID)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 6) = IIf(IsNull(Rec!tnyUsed), 0, Rec!tnyUsed)
               vsGrid.TextMatrix(vsGrid.Rows - 1, 7) = IIf(IsNull(Rec!tnyRemoved), 0, Rec!tnyRemoved)
               If Rec!tnyUsed <> 0 Then
                vsGrid.Cell(flexcpForeColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, vsGrid.Cols - 1) = vbGreen
               End If
               If Rec!tnyRemoved <> 0 Then
                vsGrid.Cell(flexcpForeColor, vsGrid.Rows - 1, 0, vsGrid.Rows - 1, vsGrid.Cols - 1) = vbRed
               End If
               Rec.MoveNext
           Wend
       End If
       If vsGrid.Rows < 15 Then
          vsGrid.Rows = 15
       End If
       Rec.Close
       mCnn.Close
       Exit Sub
last:
       MsgBox Err.Description, vbInformation
    End Sub
    
    Private Sub SelectScreen(TF As Boolean)
        tabProceedings.TabVisible(0) = True
        tabProceedings.TabVisible(1) = True
        txtProceedingsNo.Tag = -1
        txtProceedingsNo.Text = ""
        txtProceedingsDate.Text = ""
        txtRemarks.Text = ""
        chkUsed.value = 0
        chkCancel.value = 0
        tabProceedings.TabVisible(0) = TF
        tabProceedings.TabVisible(1) = Not TF
    End Sub
    
    Private Sub txtProceedingsDate_LostFocus()
        If Trim(txtProceedingsDate.Text) <> "" Then
            txtProceedingsDate.Text = CheckDateInMMM(Trim(txtProceedingsDate.Text))
        End If
    End Sub
    
    Private Sub vsGrid_Click()
        If vsGrid.Row > 0 Then
            If chkEdit.value = 1 Then
                If Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) <> "" Then
                    Call SelectScreen(False)
                    Call ShowDetails(Trim(vsGrid.TextMatrix(vsGrid.Row, 0)))
                End If
            End If
        End If
    End Sub
    
    Private Sub ShowDetails(ByVal mVal As Integer)
        Dim objProceedings As New clsProceedings
        With objProceedings
            .ProceedingsID = mVal
            Call .getProceedingsByID
            If .ProceedingsID <> -1 Then
                If .Removed <> 0 Then
                    MsgBox "This Proceedings Removed as Wrong Entry", vbInformation
                    Call SelectScreen(True)
                    Exit Sub
                End If
                txtProceedingsNo.Tag = .ProceedingsID
                txtProceedingsNo.Text = .ProceedingsNo
                txtProceedingsDate.Text = .ProceedingsDate
                txtRemarks.Text = .Remarks
                chkUsed.value = .Used
                chkCancel.Visible = True
                chkCancel.value = .Removed
            End If
        End With
    End Sub
    
    Private Sub VSGrid_DblClick()
        If val(vsGrid.TextMatrix(vsGrid.Row, 0)) > 0 Then
            If chkEdit.value = 0 Then
                gbSearchID = vsGrid.TextMatrix(vsGrid.Row, 0)
                gbSearchStr = vsGrid.TextMatrix(vsGrid.Row, 1)
                Unload Me
            End If
        End If
    End Sub

    Public Property Let Module(mData As Integer)
        mModule = mData
    End Property


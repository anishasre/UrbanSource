VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmCollectionRegister 
   BackColor       =   &H00EAFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collection Register"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin WinXPC_Engine.WindowsXPC WindowsXPC 
      Left            =   9780
      Top             =   5250
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1890
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1905
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1905
      Width           =   1065
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1905
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAFFFF&
      Height          =   1785
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   8655
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5295
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1350
         Width           =   2805
      End
      Begin VB.ComboBox cmbStaff 
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
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   630
         Width           =   5010
      End
      Begin VB.TextBox txtDate 
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2130
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox txtRegNo 
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
         Height          =   315
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtPageNo 
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
         Height          =   315
         Left            =   5490
         MaxLength       =   5
         TabIndex        =   2
         Top             =   255
         Width           =   1515
      End
      Begin VB.CheckBox chkBookClose 
         BackColor       =   &H00EAFFFF&
         Caption         =   "Closing the Book"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5265
         TabIndex        =   9
         Top             =   1035
         Width           =   1830
      End
      Begin VB.Label lblRemarks 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   210
         Left            =   4605
         TabIndex        =   15
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. of Pages"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   4080
         TabIndex        =   13
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Register No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1230
         TabIndex        =   12
         Top             =   300
         Width           =   870
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Staff"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   1095
         TabIndex        =   11
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Issue"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   1110
         TabIndex        =   10
         Top             =   1035
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAFFFF&
      Height          =   4095
      Left            =   30
      TabIndex        =   14
      Top             =   2310
      Width           =   8655
      Begin VSFlex8LCtl.VSFlexGrid fgCollectionRegister 
         Height          =   2685
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   8565
         _cx             =   15108
         _cy             =   4736
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
         BackColor       =   15400959
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   128
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   15400959
         BackColorAlternate=   15400959
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
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
         FormatString    =   $"frmCollectionRegister.frx":0000
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
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   1
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
      End
   End
End
Attribute VB_Name = "frmCollectionRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbStaff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call PressTabKey
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
  ClearDetails
End Sub

Private Sub cmdSave_Click()
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim arryin(6) As Variant
        If Trim(txtRegNo.Text) = "" Then
            MsgBox "Enter the Register Number"
            txtRegNo.SetFocus
        ElseIf Trim(txtPageNo) = "" Then
            MsgBox "Enter the Total No. of Pages "
            txtPageNo.SetFocus
        ElseIf cmbStaff.ListIndex <= 0 Then
            MsgBox "Select the staff's name from the list"
            cmbStaff.SetFocus
        ElseIf txtDate.Text = "" Then
            MsgBox "Enter the Date"
            txtDate.SetFocus
        Else
            arryin(0) = txtRegNo.Text
            arryin(1) = txtPageNo.Text
            arryin(2) = txtDate.Text
            arryin(3) = cmbStaff.ItemData(cmbStaff.ListIndex)
            arryin(4) = chkBookClose.Value
            arryin(5) = 0
            arryin(6) = txtRemarks.Text
            Set Rec = objDB.ExecuteSP("spSaveCollectionRegister", arryin)
            MsgBox "Saved"
            SelectDetails 'function name
        End If
ClearDetails 'function name
End Sub
Private Sub SelectDetails()
    Dim mCon As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim objDB As New clsDB
    Dim arryOut As Variant
    Dim Count As Integer
    If (objDB.SetConnection(mCon)) Then
        Set Rec = objDB.ExecuteSP("spGetCollectionRegister", , arryOut, , mCon, adCmdStoredProc)
        If IsArray(arryOut) Then
            fgCollectionRegister.Rows = UBound(arryOut, 2) + 2
            For Count = 0 To UBound(arryOut, 2)
                If arryOut(4, Count) = 1 Then
                    fgCollectionRegister.Cell(flexcpBackColor, Count + 1, 1, Count + 1, 6) = &HC0C0C0
                End If
                fgCollectionRegister.TextMatrix(Count + 1, 2) = arryOut(0, Count)
                fgCollectionRegister.TextMatrix(Count + 1, 3) = arryOut(1, Count)
                fgCollectionRegister.TextMatrix(Count + 1, 4) = arryOut(2, Count)
                fgCollectionRegister.TextMatrix(Count + 1, 6) = arryOut(3, Count)
            Next Count
        End If
    End If
End Sub
Private Sub fgCollectionRegister_Click()
    Dim mCon As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCount As Integer
    Dim arryOut As Variant
    chkBookClose.Enabled = True
    lblRemarks.Enabled = True
    txtRemarks.Enabled = True
    If (objDB.SetConnection(mCon)) Then
        Set Rec = objDB.ExecuteSP("spGetCollectionRegister", , arryOut, , mCon, adCmdStoredProc)
        If IsArray(arryOut) Then
            For mCount = 0 To UBound(arryOut, 2)
                If arryOut(0, mCount) = fgCollectionRegister.TextMatrix(fgCollectionRegister.RowSel, 2) Then
                    chkBookClose.Value = arryOut(4, mCount)
                    txtRegNo.Text = fgCollectionRegister.TextMatrix(fgCollectionRegister.RowSel, 2)
                    txtPageNo = fgCollectionRegister.TextMatrix(fgCollectionRegister.RowSel, 3)
                    txtDate = fgCollectionRegister.TextMatrix(fgCollectionRegister.RowSel, 4)
                    cmbStaff = fgCollectionRegister.TextMatrix(fgCollectionRegister.RowSel, 6)
                    If arryOut(5, mCount) <> " " Then
                    txtRemarks = arryOut(5, mCount)
                    Else: txtRemarks = ""
                    End If
                End If
            Next mCount
        End If
    End If
End Sub
Private Sub ClearDetails()
    chkBookClose.Enabled = False
    lblRemarks.Enabled = False
    txtRemarks.Enabled = False
    txtRegNo.Text = ""
    txtPageNo.Text = ""
    txtDate.Text = ""
    cmbStaff.ListIndex = 0
    chkBookClose.Value = 0
End Sub

Private Sub Form_Activate()
    Dim mCon As New ADODB.Connection
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    
    Me.Left = (frmMenu.Width - Me.Width) / 2
    Me.Top = 0
    chkBookClose.Enabled = False
    lblRemarks.Enabled = False
    txtRemarks.Enabled = False
    SelectDetails 'funct.name
    cmbStaff.AddItem ("...")
    If (objDB.SetConnection(mCon)) Then
        Rec.Open "Select numemployeeid,vchEmpName from faStaffs", mCon
        While Not Rec.EOF
            cmbStaff.AddItem (Rec(1))
            cmbStaff.ItemData(cmbStaff.NewIndex) = Rec(0)
            Rec.MoveNext
        Wend
    End If
End Sub
Private Sub Form_Load()
    WindowsXPC.InitIDESubClassing
End Sub



Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
    End If
         If KeyAscii <= 48 Or KeyAscii >= 56 Then
            KeyAscii = 0
            End If
End Sub
Private Sub txtDate_LostFocus()
    If Trim(txtDate.Text) <> "" Then
        txtDate.Text = CheckDateInMMM(txtDate.Text)
    End If
End Sub
Private Sub txtPageNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call PressTabKey
    End If
         If KeyAscii <= 48 Or KeyAscii >= 56 Then
            KeyAscii = 0
         End If
End Sub
Private Sub txtRegNo_KeyPress(KeyAscii As Integer)
     If KeyAscii <= 48 Or KeyAscii >= 56 Then
        KeyAscii = 0
    End If
        If KeyAscii = 13 Then
            Call PressTabKey
        End If
End Sub

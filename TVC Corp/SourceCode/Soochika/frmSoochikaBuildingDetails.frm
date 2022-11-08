VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{7B8F8FDE-7CAE-11D9-9F6C-FE443304477B}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmSoochikaBuildingDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Building Tax Details From Sachaya Revenue Module"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   405
      Top             =   4755
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7890
      TabIndex        =   10
      Top             =   4515
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tax Details"
      Height          =   3315
      Left            =   135
      TabIndex        =   8
      Top             =   1140
      Width           =   9060
      Begin VSFlex8LCtl.VSFlexGrid vsGrid 
         Height          =   2640
         Left            =   150
         TabIndex        =   9
         Top             =   450
         Width           =   8745
         _cx             =   15425
         _cy             =   4657
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSoochikaBuildingDetails.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Building Details"
      Height          =   960
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   9060
      Begin VB.TextBox txtassyear 
         Height          =   390
         Left            =   4530
         TabIndex        =   13
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox txtOwner 
         Height          =   345
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   2535
      End
      Begin VB.TextBox txtBuildingNo 
         Height          =   345
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   2535
      End
      Begin VB.TextBox txtDoorNo2 
         Height          =   390
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   345
         Width           =   795
      End
      Begin VB.TextBox txtDoorNo1 
         Height          =   390
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   345
         Width           =   915
      End
      Begin VB.TextBox txtWardNo 
         Height          =   390
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Lalblassyear 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Year"
         Height          =   360
         Left            =   4185
         TabIndex        =   14
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner Name"
         Height          =   270
         Left            =   5235
         TabIndex        =   12
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Building NO"
         Height          =   270
         Left            =   5310
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   3255
         X2              =   3120
         Y1              =   360
         Y2              =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DoorNo"
         Height          =   270
         Left            =   1560
         TabIndex        =   3
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ward No"
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   420
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSoochikaBuildingDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================='
'   Search Building Details for Soochika from Sevana Added On 5/1/10 '
'===================================================================='

Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitIDESubClassing
    Call FormInitialize
    'CHANGED MAY
    'txtWardNo.Text = frmSoochikaInward.txtWardNo.Text
    'txtDoorNo1.Text = frmSoochikaInward.txtDoorNo1.Text
    'txtDoorNo2.Text = frmSoochikaInward.txtDoorNo2.Text
    txtWardNo.Text = frmUSoochikaInward.txtWardNo.Text
    txtDoorNo1.Text = frmUSoochikaInward.txtDoorNo1.Text
    txtDoorNo2.Text = frmUSoochikaInward.txtDoorNo2.Text
    txtassyear.Text = frmUSoochikaInward.txtasssyear.Text
    Call SearchBuildDetailsFromSanchaya
End Sub

Private Sub FormInitialize()
    txtWardNo.Text = ""
    txtDoorNo1.Text = ""
    txtDoorNo2.Text = ""
    txtBuildingNo.Text = ""
    txtOwner.Text = ""
    vsGrid.Clear 1, 1
End Sub

Private Function SearchBuildDetailsFromSanchaya()
    On Error GoTo Err:
        
        Dim mCnn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim objDB As New clsDB
        Dim aryIn As Variant
        Dim mRowCnt As Integer
        
        If objDB.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
            aryIn = Array(val(txtWardNo.Text), val(txtDoorNo1.Text), txtDoorNo2.Text, txtassyear.Text)
            Set Rec = objDB.ExecuteSP("spSanchayaOwner", aryIn, , , mCnn, adCmdStoredProc)
            If Not (Rec.EOF Or Rec.BOF) Then
                txtBuildingNo.Text = IIf(IsNull(Rec!numBuildingID), "", Rec!numBuildingID)
                txtOwner.Text = IIf(IsNull(Rec!chvOwners), "", Rec!chvOwners)
                If Rec.State = 1 Then Rec.Close
                aryIn = ""
                aryIn = Array(txtBuildingNo.Text)
                
                Set Rec = objDB.ExecuteSP("sp_IndividualDemand_SBalancenew", aryIn, , , mCnn, adCmdStoredProc)
                
                vsGrid.Rows = 2
                mRowCnt = 1
                While Not (Rec.EOF Or Rec.BOF)
                    vsGrid.TextMatrix(mRowCnt, 0) = IIf(IsNull(Rec!chvYear), "", Rec!chvYear)
                    vsGrid.TextMatrix(mRowCnt, 1) = IIf(IsNull(Rec!Period), "", Rec!Period)
                    vsGrid.TextMatrix(mRowCnt, 2) = IIf(IsNull(Rec!PTax), "", Rec!PTax)
                    vsGrid.TextMatrix(mRowCnt, 3) = IIf(IsNull(Rec!LC), "", Rec!LC)
                    vsGrid.TextMatrix(mRowCnt, 4) = IIf(IsNull(Rec!fltDemand), "", Rec!fltDemand)
                    Rec.MoveNext
                    mRowCnt = mRowCnt + 1
                    vsGrid.Rows = vsGrid.Rows + 1
                Wend
                If Rec.State = 1 Then Rec.Close
            End If
        Else
            MsgBox "Connection To Sanchaya does not exist, Please contact your System Administrator"
        End If
    Exit Function
Err:
    MsgBox (Error$)
End Function

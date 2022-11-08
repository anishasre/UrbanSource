VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmInwardValuebles 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inward Valuables"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VSFlex8LCtl.VSFlexGrid fgDetails 
      Height          =   3525
      Left            =   600
      TabIndex        =   0
      Top             =   570
      Width           =   10035
      _cx             =   17701
      _cy             =   6218
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
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInwardValuebles.frx":0000
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
Attribute VB_Name = "frmInwardValuebles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fgDetails_DblClick()

    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    

    frmChequeDetails.Show

    objdb.SetConnection mCnn
    mSql = "Select A.chvSeatTitle,vchFileNo,vchDoorNo2,intDoorNo,chvSeatTitle,IntWardNo,vchHouseName,vchLocalPlace,vchStreet,vchMainPlace,numForwardedSeatID,vchsubject,fltAmount,dtReceivedDate,tnyStatus,intInstrumentTypeID,vchInstrumentNo,dtInstrumentDate,vchPost,vchPin,vchName,vchBankName,vchPhone From faSoochikaInward left join DB_Masters..GL_Seats A on faSoochikaInward.numForWardedSeatID=A.numSeatID Where A.intDepartMentID=3 and numForwardedSeatID='" & fgDetails.TextMatrix(fgDetails.Row, 6) & "'"
    Rec.Open mSql, mCnn
    If Not (Rec.EOF And Rec.BOF) Then
        frmChequeDetails.txtName.Text = Rec!vchName
        frmChequeDetails.txtHouseName.Text = Rec!vchHouseName
        frmChequeDetails.txtStreet.Text = Rec!vchStreet
        frmChequeDetails.txtLocalPlace.Text = Rec!vchLocalPlace
        frmChequeDetails.txtMainPlace.Text = Rec!vchMainPlace
        frmChequeDetails.txtWardNo.Text = Rec!intWardNo
        frmChequeDetails.txtDoorNo.Text = Rec!intDoorNo
        frmChequeDetails.txtBankName.Text = Rec!vchBankName
        frmChequeDetails.txtReceivedDate.Text = Rec!dtReceivedDate
        frmChequeDetails.txtInstrumentDate.Text = Rec!dtInstrumentDate
        frmChequeDetails.txtInstrumentNo.Text = Rec!vchInstrumentNo
        frmChequeDetails.txtPin.Text = Rec!vchPin
        frmChequeDetails.txtPost.Text = Rec!vchPost
        frmChequeDetails.txtPhone.Text = Rec!vchPhone
        frmChequeDetails.txtFileNo.Text = Rec!vchFileNo
        frmChequeDetails.txtForwardedSeatID.Text = Rec!chvSeatTitle
    End If
    Rec.Close
    
    
End Sub

Private Sub Form_Load()
    Dim objdb As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSql As String
    Dim mRowCount As Variant
    
    frmInwardValuebles.Height = 5450
    frmInwardValuebles.Width = 11490

        objdb.SetConnection mCnn
        mSql = "Select A.chvSeatTitle,numForwardedSeatID,vchsubject,fltAmount,dtReceivedDate,tnyStatus,intInstrumentTypeID,vchInstrumentNo,dtInstrumentDate From faSoochikaInward left join DB_Masters..GL_Seats A on faSoochikaInward.numForWardedSeatID=A.numSeatID Where A.intDepartMentID=3"
        mRowCount = 1
        Rec.Open mSql, mCnn
        
    While Not Rec.EOF
        If (Rec!tnyStatus = 1) Then
        fgDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtReceivedDate), "", Rec!dtReceivedDate)
        fgDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchSubject), "", Rec!vchSubject)
        fgDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo) & "          " & IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
        fgDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        fgDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        fgDetails.TextMatrix(mRowCount, 5) = "open"
        fgDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        End If
        If (Rec!tnyStatus = 2) Then
        fgDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtReceivedDate), "", Rec!dtReceivedDate)
        fgDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchSubject), "", Rec!vchSubject)
        fgDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo) & "          " & IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
        fgDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        fgDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        fgDetails.TextMatrix(mRowCount, 5) = "Demand"
        fgDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        End If
        If (Rec!tnyStatus = 3) Then
        fgDetails.TextMatrix(mRowCount, 0) = IIf(IsNull(Rec!dtReceivedDate), "", Rec!dtReceivedDate)
        fgDetails.TextMatrix(mRowCount, 1) = IIf(IsNull(Rec!vchSubject), "", Rec!vchSubject)
        fgDetails.TextMatrix(mRowCount, 2) = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo) & "          " & IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
        fgDetails.TextMatrix(mRowCount, 3) = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
        fgDetails.TextMatrix(mRowCount, 4) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        fgDetails.TextMatrix(mRowCount, 5) = "Received"
        fgDetails.TextMatrix(mRowCount, 6) = IIf(IsNull(Rec!numForwardedSeatID), "", Rec!numForwardedSeatID)
        End If
        Rec.MoveNext
        mRowCount = mRowCount + 1
    
    Wend
    Rec.Close
End Sub

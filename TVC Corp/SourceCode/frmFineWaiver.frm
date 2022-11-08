VERSION 5.00
Begin VB.Form frmFineWaiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penal Interest Waiver"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResult 
      Caption         =   "Penal Waive"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2115
      Width           =   1335
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1695
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1605
      Width           =   3615
   End
   Begin VB.TextBox txtChangedFine 
      Height          =   285
      Left            =   1695
      TabIndex        =   4
      Top             =   1215
      Width           =   1725
   End
   Begin VB.TextBox txtActualFine 
      Height          =   285
      Left            =   1695
      TabIndex        =   3
      Top             =   840
      Width           =   1710
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1665
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Changed Penal Interest"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1260
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Actual Penal Interest"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   900
      Width           =   1470
   End
End
Attribute VB_Name = "frmFineWaiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
        Dim mFineWaiveAmount As Variant
        Dim mMode As Variant              ' 1 =Property Tax 2 = Property Tax Calculator 3=Rent On Land And Building  4= Prof.Tax
    Private Sub cmdResult_Click()
    
        Dim mCnn         As New ADODB.Connection
        Dim objDB        As New clsDB
        Dim Rec          As New Recordset
        Dim mSQL         As Variant
        Dim mArrIn       As Variant
        Dim mArrOut      As Variant
        Dim mVoucher     As Variant
        Dim i            As Integer
        
        
        objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
         '******************validations**************************************
         
        If txtChangedFine = "" Then
           MsgBox "Please Enter the Actual Fine", vbInformation
           txtChangedFine.SetFocus
        Exit Sub
        End If
     
        If txtRemarks = "" Then
           MsgBox "Please Enter the Reason"
           txtRemarks.SetFocus
           Exit Sub
         End If
         
         mArrIn = Array(gbCounterID, _
                       val(frmReceiptsCounter.txtInstrument.Tag), _
                       gbFinancialYearID)
                       
         
         objDB.ExecuteSP "spGetNextReceiptNo", mArrIn, mArrOut, , mCnn, adCmdStoredProc
         If IsArray(mArrOut) Then
            mVoucher = mArrOut(0, 0)
        Else
            MsgBox "", vbInformation
            Exit Sub
        End If
         
                '@dtDate        SmallDateTime,
                '@intTransactionTypeID Int,
                '@numUserID    Numeric,
                '@numSeatID    Numeric,
                '@intCounterID     Int,
                '@intVoucherNo     Numeric,
                '@fltActualFine    Numeric,
                '@fltChangedFine   Numeric,
                '@vchReason    VarChar(100),
         
         mArrIn = Array(gbTransactionDate, _
                        gbTransactionTypePTax, _
                        gbUserID, _
                        gbSeatID, _
                        gbCounterID, _
                        mVoucher, _
                        val(txtActualFine.Text), _
                        val(txtChangedFine.Text), _
                        Trim(txtRemarks.Text))
        
         objDB.ExecuteSP "spSaveFineWaiver", mArrIn, , , mCnn, adCmdStoredProc
         MsgBox "Fine Waive Changed"
         
         If mMode = 1 Then
            frmPropertyTax.txtFine.Text = txtChangedFine.Text
               For i = 1 To frmPropertyTax.vsGrid.Rows - 1
                   If frmPropertyTax.vsGrid.TextMatrix(i, 0) = gbAcHeadCodePenalInterest Then
                       frmPropertyTax.vsGrid.TextMatrix(i, 5) = txtChangedFine.Text
                   End If
               Next i
        ElseIf mMode = 2 Then
            frmPropertyTaxCalculator.lblFine.Caption = txtChangedFine.Text
        ElseIf mMode = 3 Then
            frmRentOnLandBuildings.txtFine.Text = txtChangedFine.Text
            For i = 1 To frmRentOnLandBuildings.vsGrid.Rows - 1
                If frmRentOnLandBuildings.vsGrid.TextMatrix(i, 0) = gbAcHeadCodePenalInterest Then
                    frmRentOnLandBuildings.vsGrid.TextMatrix(i, 5) = val(txtChangedFine.Text)
                    frmRentOnLandBuildings.vsGrid.TextMatrix(i, 11) = val(txtChangedFine.Text)
                End If
            Next i
        ElseIf mMode = 4 Then
            For i = 1 To frmSearchProfTaxInstitutions.vsGridDemand.Rows - 1
                If frmSearchProfTaxInstitutions.vsGridDemand.TextMatrix(i, 0) = gbAcHeadCodePenalInterest Then
                    frmSearchProfTaxInstitutions.vsGridDemand.TextMatrix(i, 5) = val(txtChangedFine.Text)
                    frmSearchProfTaxInstitutions.vsGridDemand.TextMatrix(i, 11) = val(txtChangedFine.Text)
                    frmSearchProfTaxInstitutions.txtFine.Text = val(txtChangedFine.Text)
                    frmSearchProfTaxInstitutions.txtCurrentTotal.Text = val(txtChangedFine.Text)
                    frmSearchProfTaxInstitutions.txtDemandTotal = val(frmSearchProfTaxInstitutions.txtArrearTotal.Text) + val(frmSearchProfTaxInstitutions.txtCurrentTotal.Text)
                    
                End If
            Next i
        End If
        Unload Me
        End Sub
        Private Sub Form_Load()
            If mMode = 1 Then
                txtActualFine.Text = frmPropertyTax.txtFine
                txtActualFine.Enabled = False
                'Call frmPropertyTax.Calculate
            ElseIf mMode = 2 Then
                txtActualFine.Text = frmPropertyTaxCalculator.lblFine
                txtActualFine.Enabled = False
            ElseIf mMode = 4 Then
                txtActualFine.Text = frmSearchProfTaxInstitutions.txtFine
                txtActualFine.Enabled = False
            End If
        End Sub
        Property Let FineWaiveAmount(mAmount As Variant)
            mFineWaiveAmount = mAmount
        End Property
        Property Get FineWaiveAmount() As Variant
            FineWaiveAmount = mFineWaiveAmount
        End Property
        Property Let Mode(mData As Variant)
            mMode = mData
        End Property
        Property Get Mode() As Variant
            Mode = mMode
        End Property
        Private Sub txtChangedFine_KeyPress(KeyAscii As Integer)
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
                KeyAscii = 0
            End If
        End Sub

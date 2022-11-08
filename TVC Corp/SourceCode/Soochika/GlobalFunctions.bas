Attribute VB_Name = "GlobalFunctions"
Option Explicit



'*************************************************************************
'* Function: CenterForm(WhatForm As Form)
'*
'*
'*************************************************************************
'* Description: Center a form in the center of the screen.
'*
'*
'*************************************************************************
'* Parameters: Form to center
'*
'*************************************************************************
'* Notes:
'*
'*************************************************************************
'* Returns: None
'*************************************************************************
Sub gSubCenterForm(WhatForm As Form)

    If WhatForm.WindowState <> 0 Then Exit Sub
    
    WhatForm.Move (Screen.Width - WhatForm.Width) \ 2, (Screen.Height - WhatForm.Height) \ 2
    
End Sub
'*************************************************************************
'* Function: CenterMDIChild(frmParent As Form, frmChild As Form)
'*
'*
'*************************************************************************
'* Description: Centers a child form within a parent MDI form.
'*
'*
'*************************************************************************
'* Parameters: ParentMDI, Child Form
'*
'*************************************************************************
'* Notes:
'*
'*************************************************************************
'* Returns: None
'*************************************************************************

Sub gSubCenterMDIChild(frmParent As Form, frmChild As Form)
    Dim TTop As Integer, LLeft As Integer
    'Check to see if not minimized.
    If frmParent.WindowState <> 0 Or frmChild.WindowState <> 0 Then Exit Sub
    
    TTop = (frmParent.ScaleHeight - frmChild.Height) \ 2
    LLeft = (frmParent.ScaleWidth - frmChild.Width) \ 2
    
    If TTop And LLeft Then
        frmChild.Move LLeft, TTop
    End If
End Sub

Function gFunIsDMYDateBoolean(ByVal strDate As String) As Boolean
'*************************************************************
'Procedure:    Public Method gblnIsDMYDate
'Created on:   12/29/04
'Created by:   G.Krishna Kumar ,Binu K,Subash S and Abhilash.N
'Module:       Functions
'Project:      WaterSupply
'Note:Check whether input string is in dd/mm/yyyy format and year in between 1900 and 2079
'Parameters:
'strDate string
'Return Type:Boolean
'*************************************************************
    Dim intIndex As Integer
    Dim intDay As String
    Dim intMon As String
    Dim intYear As String
    Dim blnResult As Boolean
    blnResult = False
    If Len(strDate) >= 6 Then
        intIndex = InStr(1, strDate, "/")
        If intIndex > 1 Then
            If IsNumeric(Left(strDate, intIndex - 1)) Then
                intDay = Left(strDate, intIndex - 1)
                strDate = mID(strDate, intIndex + 1)
                intIndex = InStr(1, strDate, "/")
                If intIndex > 1 Then
                    If IsNumeric(Left(strDate, intIndex - 1)) Then
                         intMon = Left(strDate, intIndex - 1)
                         If IsNumeric(mID(strDate, intIndex + 1)) Then
                             intYear = mID(strDate, intIndex + 1)
                             If intMon <= 12 Then
                                 If intYear >= 1900 And intYear <= 2078 Then
                                    strDate = intDay & "/" & intMon & "/" & intYear
                                    If IsDate(strDate) Then
                                        blnResult = True
                                    End If
                                 End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    gFunIsDMYDateBoolean = blnResult
End Function

Private Function gFunSPdebubString(sp As String, Ary)
    If IsArray(Ary) Then
        gFunSPdebubString = sp & " "
        Dim i
        For Each i In Ary
            gFunSPdebubString = gFunSPdebubString & "'" & i & "',"
        Next
        gFunSPdebubString = Left(gFunSPdebubString, Len(gFunSPdebubString) - 1)
        Debug.Print gFunSPdebubString
    End If
End Function

Public Function gFunExecuteDBSP(ByVal strForExecute As String, ByVal ADOCmd As ADODB.CommandTypeEnum, _
                    Optional vAryIn, Optional varyOut, Optional adoConnection As ADODB.Connection)
    '****************************************************************************************************************
    'Procedure:    Public Method gFunExecuteDBSP                                                                    '
    'Created on:   03/12/2005                                                                                       '
    'Created by:   Pratheesh Kumar K.                                                                               '
    'Module:       Functions                                                                                        '
    'Project:      Sevana Pension                                                                                   '
    'Note:Executes the comand string using command objects and accepts parameter array to command parameters        '
    'Parameters:                                                                                                    '
    '****************************************************************************************************************
    Dim AdoCon As New ADODB.Connection
    Dim adocom As New ADODB.Command
    Dim adorst As New ADODB.Recordset
    Dim aryrec As Integer
    Dim iRecAffected As Variant
    adocom.ActiveConnection = adoConnection
    adocom.CommandType = ADOCmd
    adocom.CommandText = strForExecute
    If Not IsMissing(vAryIn) Then
        For aryrec = 0 To UBound(vAryIn)
            If IsEmpty(vAryIn(aryrec)) Or vAryIn(aryrec) = "" Then
                vAryIn(aryrec) = Null
            End If
            adocom.Parameters(aryrec + 1).value = vAryIn(aryrec)
        Next aryrec
        Set adorst = adocom.Execute(iRecAffected, adocom.Parameters)
    Else
        Set adorst = adocom.Execute(iRecAffected)
    End If
    If Not IsMissing(varyOut) Then
        If ((adorst.BOF = False) And (adorst.EOF = False)) Then
            varyOut = adorst.GetRows()
        End If
    End If
    Set adocom = Nothing
    Set adorst = Nothing
    gFunExecuteDBSP = iRecAffected
End Function

Function gSubFillItemsInToCombo1(ByVal cbo As ComboBox, ByVal SQL As String, _
                    ByVal ADOCmd As ADODB.CommandTypeEnum, Optional pInAry As Variant, Optional vClear As Variant)
    Dim AdoCon As New ADODB.Connection
    Dim arOt, Cnt, lCnt As Integer
    'Set adoCon = gFunSetDBConnection
    If IsMissing(pInAry) Then
        gFunExecuteDBSP SQL, ADOCmd, , arOt, AdoCon
    Else
        gFunExecuteDBSP SQL, ADOCmd, pInAry, arOt, AdoCon
    End If
    If IsMissing(vClear) Then
        cbo.Clear:    cbo.AddItem "...."
        cbo.ItemData(cbo.NewIndex) = 0
    End If
    If IsArray(arOt) Then
        For Cnt = 0 To UBound(arOt, 2)
            cbo.AddItem arOt(1, Cnt)
            cbo.ItemData(cbo.NewIndex) = arOt(0, Cnt)
        Next Cnt
    End If
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
    'adoCon.Close
End Function


Public Sub gsFilterKeys(ByRef iKeyAscii As Integer, ByVal blnDenyAlphabets As Boolean, ByVal blnDenyNumeric As Boolean, ByVal blnDenySpace As Boolean, ByVal blnDenyWildCharecters As Boolean, Optional sAllowedWildChar1 As String, Optional sAllowedWildChar2 As String)
'*************************************************************
'Procedure:    Public Method vFilterKeys
'Created on:   29/12/04
'Creared by:   Abdulla.P.P
'Module:       Functions
'Project:      WaterSupply
'Note:
'Parameters:
'ByRef iKeyAscii
'blnDenyAlphabets
'blnDenyNumeric
'blnDenySpace
'blnDenyWildCharecters
'Optional sAllowedWildChar1
'Optional sAllowedWildChar2
'*************************************************************
    If blnDenyAlphabets Then
        If (iKeyAscii >= 65 And iKeyAscii <= 90) Or (iKeyAscii >= 97 And iKeyAscii <= 122) Then iKeyAscii = 0
    End If
    If blnDenyNumeric Then
        If (iKeyAscii > 47 And iKeyAscii < 58) Then iKeyAscii = 0
    End If
    If blnDenyWildCharecters Then
        Dim iExcld1 As Integer
        Dim iExcld2 As Integer
        If Not (sAllowedWildChar1 = "") Then 'first arg is not missing
            iExcld1 = Asc(sAllowedWildChar1)
        Else
            iExcld1 = 0
        End If
        If Not (sAllowedWildChar2 = "") Then 'second arg is not missing
            iExcld2 = Asc(sAllowedWildChar2)
        Else
            iExcld2 = 0
        End If
        If (iKeyAscii < 48 Or (iKeyAscii > 57 And iKeyAscii < 65) Or (iKeyAscii > 90 And iKeyAscii < 97) Or iKeyAscii > 122) And (iKeyAscii <> 8 And iKeyAscii <> iExcld1 And iKeyAscii <> iExcld2 And iKeyAscii <> 32) Then
            iKeyAscii = 0
        End If
    End If
    If blnDenySpace Then
        If iKeyAscii = 32 Then iKeyAscii = 0
    End If
End Sub

Public Sub gsClearForm(ByRef frmForm As Form)
'*************************************************************
'Procedure:    Public Method gClearForm
'Created on:   12/29/04
'Created by:   Abhilash N
'Module:       Functions
'Project:      WaterSupply
'Note: clears controls in a form
'Parameters:
'ByRef frmForm
'*************************************************************
    Dim ctlControl As Control
    Dim J As Integer
    For Each ctlControl In frmForm.Controls
    Debug.Print TypeName(ctlControl)
       Select Case TypeName(ctlControl)
            Case "ListBox"
                If ctlControl.ListCount > 0 Then
                    If ctlControl.Style = 1 Then
                        For J = 0 To ctlControl.ListCount - 1
                            If ctlControl.Selected(J) = True Then
                                ctlControl.Selected(J) = False
                            End If
                        Next J
                    End If
                    ctlControl.ListIndex = 0
                End If
            Case "TextBox":
                ctlControl.Text = ""
            Case "ComboBox"
                If ctlControl.ListCount > 0 Then
                    ctlControl.ListIndex = 0
                End If
            Case "VSFlexGrid"
                ctlControl.Clear 1
            Case "CheckBox"
                If ctlControl.value = 1 Then
                    ctlControl.value = 0
                End If
            Case "NameBox"
                    ctlControl.value = ""
                    ctlControl.FirstNamePosition = 1
                    ctlControl.SecondNamePosition = 2
                    ctlControl.ThirdNamePosition = 3
                    ctlControl.FourthNamePosition = 4
                    ctlControl.FirstInitialPosition = 5
                    ctlControl.SecondInitialPosition = 6
                    ctlControl.ThirdInitialPosition = 7
                    ctlControl.FourthInitialPosition = 8
                    ctlControl.ReArrangeControlPositions
            Case "DTPicker"
                ctlControl.value = Now
        End Select
    Next ctlControl
End Sub

Public Function gFunGetListIndex(vValIn As Variant, vCtrl As ComboBox) As Variant
'For getting the list index value of current combo.text
'On Error Resume Next
    Dim intCnt As Integer
    For intCnt = 0 To vCtrl.ListCount - 1
        If vCtrl.List(intCnt) = vValIn Then
            gFunGetListIndex = intCnt: Exit For
        ElseIf val(vCtrl.ItemData(intCnt)) = vValIn Then
            gFunGetListIndex = intCnt: Exit For
        End If
    Next intCnt
End Function

Public Sub gsSelectText(ByRef txtTextBox As TextBox)
'*************************************************************
'Procedure:    Public Method vSelectText
'Created on:   02/04/05
'Creared by:   Abdulla.P.P
'Module:       Functions for selecting the contenet in a text
'Project:      WaterSupply
'Note:
'Parameters:
'txtTextBox
'*************************************************************
    If Trim(txtTextBox.Text) = "" Then
        txtTextBox.Text = ""
        Exit Sub
    End If
    txtTextBox.SelStart = 0
    txtTextBox.SelLength = Len(txtTextBox.Text)
End Sub

Public Function gFunGetCombovalue(cmbName As ComboBox)
    If cmbName.ListIndex <> -1 Then
        gFunGetCombovalue = cmbName.ItemData(cmbName.ListIndex)
    Else
        gFunGetCombovalue = 0
    End If
End Function



         
    Public Sub gSubSetComboItem2(Cmb As ComboBox, value As Variant)
        'Setting the selected Combo Item
        Dim i As Integer
            If Not value = "" Then
               For i = 0 To Cmb.ListCount - 1
                    If Cmb.ItemData(i) = value Then
                        Cmb.ListIndex = i
                        Exit For
                    End If
               Next i
            Else
               Cmb.ListIndex = -1
            End If
    End Sub
            
Public Sub gSubSetFont(fg As VSFlexGrid, StartRow As Integer, StartCol As Integer, EndRow As Integer, EndCol As Integer, FontName As String)
    Dim intRowCnt As Integer, intColCnt As Integer
    For intRowCnt = StartRow To EndRow
        For intColCnt = StartCol To EndCol
            fg.Cell(flexcpFontName, intRowCnt, intColCnt) = FontName
            fg.Cell(flexcpFontSize, intRowCnt, intColCnt) = 10
            fg.Cell(flexcpFontBold, intRowCnt, intColCnt) = False
        Next intColCnt
    Next intRowCnt
End Sub

    Public Function gDateValidation2(mDate As String) As Boolean
            '--------------------------------------------------
            'To Check Date Which is not Greater Than Transaction date and not Less than Minimun Date
            '-------------------------------------------------------
            Dim mSQL    As String
            Dim Rec     As New ADODB.Recordset
            Dim mCnn    As New ADODB.Connection
            Dim objDB   As New clsDB
            mSQL = "Select  top 1 dtStartingDate From faFinancialYear Order by intFinancialYear Asc"
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Format(Rec!dtStartingDate, "dd/MMM/YYYY") <= CDate(mDate) And CDate(mDate) <= CDate(gbTransactionDate) Then
                    gDateValidation2 = True
                Else
                    gDateValidation2 = False
                End If
            End If
     End Function
    Public Function gDateValidation(mDate As Date) As Boolean
            '--------------------------------------------------
            'To Check Date Which is not Greater Than Transaction date and not Less than Minimun Date
            '-------------------------------------------------------
            Dim mSQL    As String
            Dim Rec     As New ADODB.Recordset
            Dim mCnn    As New ADODB.Connection
            Dim objDB   As New clsDB
            mSQL = "Select  dtStartingDate, dtEndingDate From faFinancialYear Where intFinancialYear = " & gbFinancialYearID
            objDB.CreateNewConnection mCnn, enuSourceString.Saankhya
            Rec.Open mSQL, mCnn
            If Not (Rec.EOF And Rec.BOF) Then
                If Rec!dtStartingDate <= mDate And Rec!dtEndingDate >= mDate Then
                    gDateValidation = True
                Else
                    gDateValidation = False
                End If
            Else
                gDateValidation = False
            End If
     End Function
     Public Sub KeyPressNumber(ByRef mAscii As Integer, Optional mExtraKeys As String = "", Optional mBoolAlpha As Boolean = False, Optional mExcludes As String = "")
        '---------------------------------SINOJ-----------------------------------------'
        '      This Function is used to Give Extra Charectors to the TextEditor         '
        '-------------------------------------------------------------------------------'
        If Not (mAscii >= 48 And mAscii <= 57 Or mAscii = 8 Or mAscii = 13 Or InStr(1, mExtraKeys, Chr(mAscii)) > 0) Then
            If mBoolAlpha Then
                If Not (Asc(UCase(Chr(mAscii))) >= 65 And Asc(UCase(Chr(mAscii))) <= 90) Then
                    mAscii = 0
                End If
            Else
                mAscii = 0
            End If
        Else
            If InStr(1, mExcludes, Chr(mAscii)) > 0 Then
                mAscii = 0
            End If
        End If
    End Sub
   
   


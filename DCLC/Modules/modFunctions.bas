Attribute VB_Name = "modFunctions"
Option Explicit
Dim rsFees As New ADODB.Recordset

'Return a generated id
Public Function GenerateNextPK(ByVal srcField As String) As String
    Dim strYear As String
    Dim intLastNo As Integer
    Dim rs As New ADODB.Recordset
    strYear = Mid(Str(Year(Date)), 4, 2)
    
    'If rs.State = 1 Then Set rs = Nothing
    
    Set rs = GetRecordset("SELECT " & srcField & " FROM LastNo;")
    intLastNo = 0
    If rs.State = adStateOpen Then
        If rs.RecordCount > 0 Then
'            intLastNo = rs.Fields(""" & srcField & """).Value
            intLastNo = rs.Fields(0).Value
        End If
    End If
    intLastNo = intLastNo + 1
    If LCase(srcField) = "studentno" Then
        GenerateNextPK = strYear & "-" & Format(intLastNo, "0000000")
    ElseIf LCase(srcField) = "assessno" Then
        GenerateNextPK = Format(intLastNo, "0000000")
    ElseIf LCase(srcField) = "receiptno" Then
        GenerateNextPK = Format(intLastNo, "0000000")
    End If
    
    Set rs = Nothing
End Function

Public Sub ComputeFees(ByRef strCourseCode As String, _
                       ByRef intYearlevel As Integer, _
                       ByRef strAssessNo As String, _
                       ByRef strStudentId As String, _
                       ByRef strName As String, _
                       ByRef strStatus As String, _
                       ByRef blnCashBasis As Boolean, _
                       ByRef strTotalLec As String, _
                       ByRef strTotalLab As String)

    '--
    FeesBreakdown.AssessNo = strAssessNo
    FeesBreakdown.StudentId = strStudentId
    FeesBreakdown.Name = strName
    FeesBreakdown.Status = strStatus
    FeesBreakdown.CashBasis = blnCashBasis
    
    FeesBreakdown.TotalCashBasis = 0
    FeesBreakdown.TotalInstallmentBasis = 0
    FeesBreakdown.DownPayment = 0
            
            
    strSql = "SELECT * FROM dbo.Fees " & _
             "WHERE CourseCode = '" & strCourseCode & "' AND " & _
                   "CurrentSyId = " & SchoolInformation.CurrentSyId & " AND " & _
                   "CurrentSemesterId = " & SchoolInformation.CurrentSemesterId & " AND " & _
                   "CurrentYearLevelId = " & intYearlevel & " ;"
     
    
    Set rsFees = GetRecordset(strSql)
    If rsFees.State = adStateOpen Then
        If rsFees.RecordCount > 0 Then
            ' hands on = 350/unit ((3hrs each subject))
            ' laboratory = 350/unit ((3hrs each subject))
            ' tuition = (402 per unit)
            ' INSTALLMENT BASIS - sum of all fees from the downpayment up to the semi finals
            ' - down payment = miscellaneous fee * 20% + other fees
            ' - prelim       = miscellaneous fee * 40%
            ' - midterm      = miscellaneous fee * 25%
            ' - semi finals  = miscellaneoues fee * 20%
            ' CASH BASIS
            '  - miscellaneous fee + sum of other fees
            ' SUMMER
            ' other fees: - athletic and guidance and counselling fee (not included)
            ' miscellaneous fee: - only power fee is included in the computation

            FeesBreakdown.Entrance = Val(rsFees!Entrance)
            FeesBreakdown.TuitionFee = 402 * Val(strTotalLec) 'Val(txtTuitionFee)
            'OTHER FEES:
            FeesBreakdown.Registration = Val(rsFees!Registration)
            FeesBreakdown.Library = Val(rsFees!Library)
            FeesBreakdown.Laboratory = 350 * Val(strTotalLab) 'Val(txtLaboratory)
            FeesBreakdown.AthleticFee = 0
            FeesBreakdown.GuidanceAndCounselor = 0
            If LCase(SchoolInformation.Semester) <> "summer" Then
                FeesBreakdown.AthleticFee = Val(rsFees!AthleticFee) 'exclude from summer computation
                FeesBreakdown.GuidanceAndCounselor = Val(rsFees!GuidanceAndCounselor) 'exclude from summer computation
            End If
            'MISCELLANEOUS FEES:
            FeesBreakdown.Rle = Val(rsFees!Rle)
            FeesBreakdown.Affiliation = Val(rsFees!Affiliation)
            FeesBreakdown.NursingAudit = Val(rsFees!NursingAudit)
            FeesBreakdown.MarineLaboratory = Val(rsFees!MarineLaboratory)
            FeesBreakdown.SpeechLab = Val(rsFees!SpeechLab)
            FeesBreakdown.HrmLab = Val(rsFees!HrmLab)
            FeesBreakdown.Ojt = Val(rsFees!Ojt)
            FeesBreakdown.Rta = Val(rsFees!Rta)
            FeesBreakdown.HOn = 350 * Val(strTotalLab)  ' Val(txtHOn)
            FeesBreakdown.Mta = Val(rsFees!Mta)
            FeesBreakdown.IdNamePlate = Val(rsFees!IdNamePlate)
            FeesBreakdown.Sdf = Val(rsFees!Sdf)
            FeesBreakdown.PowerFee = 0
            If LCase(SchoolInformation.Semester) <> "summer" Then
                FeesBreakdown.PowerFee = Val(rsFees!PowerFee) 'exclude from summer computation
            End If
            FeesBreakdown.Internet = Val(rsFees!Internet)
            FeesBreakdown.Internship = Val(rsFees!Internship)
            FeesBreakdown.Waiver = Val(rsFees!Waiver)
            FeesBreakdown.Nstp = Val(rsFees!Nstp)
            '--
            
            'Totals
            FeesBreakdown.OtherFees = FeesBreakdown.Entrance + _
                                      FeesBreakdown.TuitionFee + _
                                      FeesBreakdown.Registration + _
                                      FeesBreakdown.Library + _
                                      FeesBreakdown.Laboratory + _
                                      FeesBreakdown.AthleticFee + _
                                      FeesBreakdown.GuidanceAndCounselor
            '
            FeesBreakdown.MiscFees = FeesBreakdown.Rle + _
                                     FeesBreakdown.Affiliation + _
                                     FeesBreakdown.NursingAudit + _
                                     FeesBreakdown.MarineLaboratory + _
                                     FeesBreakdown.SpeechLab + _
                                     FeesBreakdown.HrmLab + _
                                     FeesBreakdown.Ojt + _
                                     FeesBreakdown.Rta + _
                                     FeesBreakdown.HOn + _
                                     FeesBreakdown.Mta + _
                                     FeesBreakdown.IdNamePlate + _
                                     FeesBreakdown.Sdf + _
                                     FeesBreakdown.PowerFee + _
                                     FeesBreakdown.Internet + _
                                     FeesBreakdown.Internship + _
                                     FeesBreakdown.Waiver + _
                                     FeesBreakdown.Nstp
                                     
           
            '- DISCOUNT
            FeesBreakdown.Discount = 0
            If LCase(strStatus) = "with brother" Then
                FeesBreakdown.Discount = (FeesBreakdown.OtherFees + FeesBreakdown.MiscFees) * 0.1
            ElseIf LCase(strStatus) = "scholar" Then
                FeesBreakdown.Discount = 4000
            End If
            
            FeesBreakdown.TotalCashBasis = (FeesBreakdown.OtherFees + FeesBreakdown.MiscFees) - FeesBreakdown.Discount
                                    
            FeesBreakdown.DownPayment = (FeesBreakdown.MiscFees * 0.2) + FeesBreakdown.OtherFees
            FeesBreakdown.TotalInstallmentBasis = ((FeesBreakdown.MiscFees * 0.4 * 3) + _
                                                  (FeesBreakdown.MiscFees * 0.25 * 3) + _
                                                  (FeesBreakdown.MiscFees * 0.2 * 3)) - FeesBreakdown.Discount
                                                  
            
            
        Else
            MsgBox "No fee schedule found", vbInformation
        End If
    End If
    Set rsFees = Nothing
    
End Sub


'Check if the record exist or not.
Public Function RecordExist(strField As String, strTable As String, strValue As String) As Boolean
    Dim rs As New Recordset
    Dim strSql As String

    strSql = "SELECT " & strField & " FROM " & strTable & " WHERE LOWER(" & strField & ") = '" & LCase(strValue) & "';"
    rs.CursorLocation = adUseClient
    rs.Open strSql, adoConnection, adOpenStatic, adLockOptimistic
    If rs.RecordCount < 1 Then
        RecordExist = False
    Else
        RecordExist = True
    End If
    Set rs = Nothing
End Function

'Highlight text when focus
Public Sub FocusMe(ByRef sText)
    SendKeys "{HOME}+{END}"
End Sub

Public Sub EmulateEnter(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

'Clear the textbox content
Public Sub ClearEntries(ByRef thisForm As Form)
    Dim ctrlControl As Control
    For Each ctrlControl In thisForm.Controls
        If (TypeOf ctrlControl Is TextBox) Then ctrlControl = vbNullString
        If (TypeOf ctrlControl Is MaskEdBox) Then
            Dim strMask As String
            With ctrlControl
                strMask = .Mask
                .Mask = ""
                .Text = ""
                .Mask = strMask
            End With
        End If
    Next ctrlControl
    ValidateAccessLevel thisForm, "Insert"
    Set ctrlControl = Nothing
End Sub

'Procedure used to bind data combo
Public Sub BindDataCombo(ByVal srcSQL As String, _
                         ByVal srcBindField As String, _
                         ByRef srcDC As DataCombo, _
                         Optional srcColBound As String, _
                         Optional ShowFirstRec As Boolean)
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, adoConnection, adOpenStatic, adLockOptimistic
    
    With srcDC
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = rs
        'Display the first record
        If ShowFirstRec = True Then
            If Not rs.RecordCount < 1 Then
                .BoundText = rs.Fields(srcColBound) 'DataValueField
                .Tag = rs.RecordCount & "*~~~~~*" & rs.Fields(srcColBound) 'Text
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set rs = Nothing
End Sub

'Remove dreaded sql characters
Public Function CleanUpSql(ByVal strSql As AsyncProperty) As String
    Dim strRetVal As String
    strRetVal = strSql 'Replace(strSql, "'", "`")
    'Debug.Print strRetVal
    CleanUpSql = strRetVal
End Function

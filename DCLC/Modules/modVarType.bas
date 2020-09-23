Attribute VB_Name = "modVarType"
Public Type SchoolInfo
    SchoolName1 As String
    SchoolName2 As String
    Address As String
    TelNo1 As String
    TelNo2 As String
    FaxNo As String
    Sy As String
    Semester As String
    CurrentSyId As Integer
    CurrentSemesterId As Integer
End Type

'Variable structure for ReportInfo
Public Type ReportInfo
    Sy As String
    Semester As String
    YearLevel As String
    Course As String
    RoomNo As String
End Type

'Variable structure for Fees Breakdown
Public Type FeesBrkdwn
    AssessNo As String
    StudentId As String
    Name As String
    Status As String 'Scholar, With brother....
    CashBasis As Boolean
    
    Entrance As Integer
    TuitionFee As Integer '402 per unit
    
    Registration As Integer
    Library As Integer
    Laboratory As Integer '350/unit
    AthleticFee  As Integer 'exclude from summer computation
    GuidanceAndCounselor As Integer 'exclude from summer computation
    '--------
    OtherFees  As Integer
    '--------
    Rle As Integer
    Affiliation As Integer
    NursingAudit As Integer
    MarineLaboratory As Integer
    SpeechLab As Integer
    HrmLab As Integer
    Ojt As Integer
    Rta As Integer
    HOn As Integer '350/unit -> of lec or lab?
    Mta As Integer
    IdNamePlate As Integer
    Sdf As Integer
    PowerFee As Integer 'exclude from summer computation
    Internet As Integer
    Internship As Integer
    Waiver As Integer
    Nstp As Integer
    '--------
    MiscFees  As Integer
    '--------
    Discount As Integer
    
    DownPayment As Integer
    TotalInstallmentBasis As Integer
    
    TotalCashBasis As Integer
    
End Type

'Variable structure for user
Public Type UserInfo
    UserId As String
    UserName As String
    UserType As String
End Type

'Enumerator for form state
Public Enum FormState
    adStateAddMode = 0
    adStateEditMode = 1
    adStatePopupMode = 2
End Enum



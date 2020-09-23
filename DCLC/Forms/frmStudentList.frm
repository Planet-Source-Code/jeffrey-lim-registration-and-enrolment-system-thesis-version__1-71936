VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStudentList 
   Caption         =   "List of Students"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmStudentList.frx":0000
   ScaleHeight     =   4200
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo dcRoom 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   2970
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcSy 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   1215
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7470
      TabIndex        =   9
      Top             =   3705
      Width           =   7530
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   6315
         TabIndex        =   6
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   5220
         TabIndex        =   5
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00879693&
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   9915
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7530
      TabIndex        =   7
      Top             =   0
      Width           =   7530
      Begin VB.Image Image2 
         Height          =   375
         Left            =   7155
         Picture         =   "frmStudentList.frx":10380
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Students"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   990
         TabIndex        =   8
         Top             =   120
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmStudentList.frx":10830
         Top             =   60
         Width           =   915
      End
   End
   Begin MSDataListLib.DataCombo dcSemester 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Top             =   2115
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcYearLevel 
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   1665
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcCourses 
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   2565
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Room:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   15
      Top             =   2970
      Width           =   1140
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   14
      Top             =   2565
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   13
      Top             =   2115
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   12
      Top             =   1710
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "School Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   11
      Top             =   1260
      Width           =   1230
   End
End
Attribute VB_Name = "frmStudentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset

'Icon on Message vbCritical 16, vbQuestion 32, vbExclamation 48, vbInformation 64
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    ReportInformation.Sy = dcSy.Text
    ReportInformation.YearLevel = dcYearLevel.Text
    ReportInformation.Semester = dcSemester.Text
    ReportInformation.RoomNo = dcRoom.Text
    
    'Complex Sub-Queries
'    strSql = "SELECT s.StudentId, (s.Lastname + ', ' + s.Firstname + ' ' + s.Middlename) AS Name, s.Gender, s.DateEnrolled, " & _
'                    "yl.YearLevel " & _
'             "FROM dbo.Students s " & _
'             "INNER JOIN dbo.YearLevel yl ON (s.YearLevelId = yl.YearLevelId) " '& _
'''             "INNER JOIN dbo.Schedules sc ON (sc.SyId = sy.SyId) " & _
'''             "INNER JOIN dbo.Sy sy ON (sc.SyId = sy.SyId) " '& _
'''             "INNER JOIN dbo.Rooms rm ON (sc.RoomId = rm.RoomId) " '& _
'''             "WHERE s.CourseCode = '" & dcCourses.BoundText & "' AND " & _
'''                   "s.YearLevelId = " & dcYearLevel.BoundText & " AND " & _
'''                   "sy.SyId = " & dcSy.BoundText & " AND " & _
'''                   "s.SemesterId = " & dcSemester.BoundText & " AND " & _
'''                   "rm.RoomId = '" & dcRoom.BoundText & "' " & _
'''             "ORDER BY Name; "

    strSql = "SELECT s.StudentId, (s.Lastname + ', ' + s.Firstname + ' ' + s.Middlename) AS Name, s.Gender, s.DateEnrolled, " & _
                    "yl.YearLevel " & _
             "FROM dbo.Students s " & _
             "INNER JOIN dbo.YearLevel yl ON (s.YearLevelId = yl.YearLevelId) " & _
             "INNER JOIN dbo.Enrolled er ON (s.StudentId = er.StudentId) " & _
             "INNER JOIN dbo.Schedules sc ON (er.SchedCode = sc.SchedCode) " & _
             "INNER JOIN dbo.Sy sy ON (sc.SyId = sy.SyId) " & _
             "INNER JOIN dbo.Rooms rm ON (sc.RoomId = rm.RoomId) " & _
             "WHERE s.CourseCode = '" & dcCourses.BoundText & "' AND " & _
                   "s.YearLevelId = " & dcYearLevel.BoundText & " AND " & _
                   "sy.SyId = " & dcSy.BoundText & " AND " & _
                   "s.SemesterId = " & dcSemester.BoundText & " AND " & _
                   "rm.RoomId = '" & dcRoom.BoundText & "' " & _
             "ORDER BY Name; "
    'Debug.Print strSql
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            Set rptStudentList.DataSource = rsData
            rptStudentList.Show vbModal
        Else
            MsgBox "No report extracted.", vbInformation
        End If
    Else
        MsgBox "No report extracted.", vbInformation
    End If

    Set rsData = Nothing
End Sub

Private Sub Form_Load()
    BindDataCombo "SELECT * FROM dbo.Sy", "Sy", dcSy, "SyId", True 'Bind School Year
    BindDataCombo "SELECT * FROM dbo.Semester", "Semester", dcSemester, "SemesterId", True 'Bind Semester
    BindDataCombo "SELECT * FROM dbo.Courses", "CourseDesc", dcCourses, "CourseCode", True 'Bind Courses
    BindDataCombo "SELECT * FROM dbo.YearLevel", "YearLevel", dcYearLevel, "YearLevelId", True 'Bind Year Level
    BindDataCombo "SELECT * FROM dbo.Rooms", "RoomNo", dcRoom, "RoomId", True 'Bind Room #s
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
    Set rsData = Nothing
End Sub

Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 190 'Fade Form
End Sub


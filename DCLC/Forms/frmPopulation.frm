VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPopulation 
   Caption         =   "Population Report"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPopulation.frx":0000
   ScaleHeight     =   4200
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo dcSy 
      Height          =   315
      Left            =   1665
      TabIndex        =   13
      Top             =   1350
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
      TabIndex        =   4
      Top             =   3705
      Width           =   7530
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   6315
         TabIndex        =   2
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   5220
         TabIndex        =   1
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
         TabIndex        =   5
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
      TabIndex        =   0
      Top             =   0
      Width           =   7530
      Begin VB.Image Image2 
         Height          =   375
         Left            =   7155
         Picture         =   "frmPopulation.frx":10380
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "School Population"
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
         TabIndex        =   3
         Top             =   120
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmPopulation.frx":10830
         Top             =   60
         Width           =   915
      End
   End
   Begin MSDataListLib.DataCombo dcSemester 
      Height          =   315
      Left            =   1665
      TabIndex        =   10
      Top             =   2250
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
      Left            =   1665
      TabIndex        =   11
      Top             =   1800
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
      Left            =   1665
      TabIndex        =   12
      Top             =   2700
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
      Left            =   495
      TabIndex        =   9
      Top             =   2700
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
      Left            =   495
      TabIndex        =   8
      Top             =   2250
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
      Left            =   495
      TabIndex        =   7
      Top             =   1845
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
      Left            =   495
      TabIndex        =   6
      Top             =   1395
      Width           =   1230
   End
End
Attribute VB_Name = "frmPopulation"
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
    
    'Complex Sub-Queries
    strSql = "SELECT c.CourseDesc, " & _
              "(SELECT COUNT(*) FROM dbo.Enrolled e " & _
               "INNER JOIN dbo.Schedules sc ON (e.SchedCode = sc.SchedCode) " & _
               "INNER JOIN dbo.Students s ON (e.StudentId = s.StudentId) " & _
               "WHERE sc.CourseCode = c.CourseCode AND " & _
                     "sc.SyId = " & dcSy.BoundText & " AND " & _
                     "sc.SemesterId = " & dcSemester.BoundText & " AND " & _
                     "s.YearLevelId = " & dcYearLevel.BoundText & ") AS StudentCount " & _
             "FROM dbo.Courses c " & _
             "ORDER BY CourseDesc; "
    '"WHERE sc.CourseCode = '" & dcCourses.BoundText & "' AND "
    
    'Debug.Print strSql
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            Set rptPopulation.DataSource = rsData
            rptPopulation.Show vbModal
        Else
            MsgBox "No report extracted.", vbInformation
        End If
    Else
        MsgBox "No report extracted.", vbInformation
    End If

    Set rsData = Nothing
End Sub

Private Sub Form_Activate()
    MakeTransparent Me.hwnd, 255 'Unfade form
End Sub

Private Sub Form_Deactivate()
    MakeTransparent Me.hwnd, 190 'Fade Form
End Sub
Private Sub Form_Load()
    BindDataCombo "SELECT * FROM dbo.Sy", "Sy", dcSy, "SyId", True 'Bind School Year
    BindDataCombo "SELECT * FROM dbo.Semester", "Semester", dcSemester, "SemesterId", True 'Bind Semester
    BindDataCombo "SELECT * FROM dbo.Courses", "CourseDesc", dcCourses, "CourseCode", False 'Bind Courses
    BindDataCombo "SELECT * FROM dbo.YearLevel", "YearLevel", dcYearLevel, "YearLevelId", True 'Bind Year Level
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
    Set rsData = Nothing
End Sub


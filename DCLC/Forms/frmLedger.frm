VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLedger 
   Caption         =   "Student's Ledger"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLedger.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Height          =   405
      Left            =   4230
      Picture         =   "frmLedger.frx":10380
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Search"
      Top             =   1395
      Width           =   450
   End
   Begin MSMask.MaskEdBox meStudentId 
      Height          =   420
      Left            =   2025
      TabIndex        =   0
      Top             =   1395
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12640511
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##-#######"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5910
      TabIndex        =   5
      Top             =   2385
      Width           =   5970
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   315
         Left            =   4725
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   3600
         TabIndex        =   2
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
         TabIndex        =   6
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
      ScaleWidth      =   5970
      TabIndex        =   1
      Top             =   0
      Width           =   5970
      Begin VB.Image Image2 
         Height          =   375
         Left            =   5580
         Picture         =   "frmLedger.frx":10755
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ledger"
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
         TabIndex        =   4
         Top             =   120
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmLedger.frx":10C05
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   8
      Top             =   1485
      Width           =   1275
   End
End
Attribute VB_Name = "frmLedger"
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
    ReportInformation.Sy = SchoolInformation.Sy

    
    'Complex Sub-Queries
    strSql = "SELECT l1.LedgerId, l1.ReceiptNo, l1.TranDate, l1.Particular, l1.Debit, l1.Credit, SUM(l2.Debit + l2.Credit) AS Balance, l1.PostedBy, sy.Sy, yl.YearLevel, sem.Semester " & _
             "FROM dbo.Ledger AS l1 " & _
                  "INNER JOIN dbo.Sy AS sy ON (l1.SyId = sy.SyId) " & _
                  "INNER JOIN dbo.YearLevel AS yl ON (l1.YearLevelId = yl.YearLevelId) " & _
                  "INNER JOIN dbo.Semester AS sem ON (l1.SemesterId = sem.SemesterId) " & _
                  "INNER JOIN dbo.Ledger AS l2 ON (l1.TranDate >= l2.TranDate AND l1.LedgerId >= l2.LedgerId) " & _
            "WHERE l1.StudentId = '" & meStudentId & "' " & _
            "GROUP BY l1.LedgerId, l1.ReceiptNo, l1.TranDate, l1.Particular, l1.Debit, l1.Credit, l1.PostedBy, sy.Sy, yl.YearLevel, sem.Semester " & _
            "ORDER BY TranDate;"
            
    
    Set rsData = GetRecordset(strSql)
    If rsData.State = adStateOpen Then
        If rsData.RecordCount > 0 Then
            ReportInformation.YearLevel = rsData!YearLevel
            ReportInformation.Semester = rsData!Semester
            Set rptLedger.DataSource = rsData
            rptLedger.Show vbModal
        Else
            MsgBox "No report extracted.", vbInformation
        End If
    Else
        MsgBox "No report extracted.", vbInformation
    End If

    Set rsData = Nothing
End Sub

Private Sub cmdSearch_Click()
    frmSearchStudent.txtCallingForm = "frmLedger"
    frmSearchStudent.Show vbModal
End Sub

Private Sub Form_Load()
'    MsgBox "More info needed!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsData.Close
    Set rsData = Nothing
End Sub

Public Sub meStudentId_LostFocus()
'    strSql = "SELECT *, p.Address AS ParentsAddress " & _
'             "FROM dbo.Students AS s " & _
'                  "JOIN dbo.Courses AS c ON (s.CourseCode = c.CourseCode) " & _
'                  "JOIN dbo.YearLevel AS yl ON (s.YearLevelId = yl.YearLevelId) " & _
'                  "JOIN dbo.Semester AS sem ON (s.SemesterId = sem.SemesterId) " & _
'                  "JOIN dbo.Parents AS p ON (s.StudentId = p.StudentId) " & _
'                  "JOIN dbo.Credentials AS cr ON (s.StudentId = cr.StudentId) " & _
'            "WHERE s.StudentId = '" & meStudentId.Text & "';"
'
'    Set rsData = GetRecordset(strSql)
'    If rsData.State = adStateOpen Then
'        If rsData.RecordCount > 0 Then
'            txtLastname.Text = rsData!Lastname
'            txtFirstname.Text = rsData!Firstname
'            txtMiddlename.Text = rsData!Middlename
'        End If
'    End If
End Sub
